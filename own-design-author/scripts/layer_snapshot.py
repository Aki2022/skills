#!/usr/bin/env python3
"""Record a repository's check result as one comparable file, and compare two of them one record at a time.

snapshot: run the design check of a skills tree on a repository (the check is deterministic and calls no model) and
write the approvals (key, recorded hash, whether it still matches the current target), the target set with hashes,
the rule ids and the verdict (exit codes per stage, errors, warnings).
compare:  diff two snapshots record by record. Anything different is exit 1; a comparison that looked at nothing,
or at fewer records than stated, is exit 2 and never a pass.

Texts are normalised through the layout file (`layer-paths.json`) so a repository before and after a move compare:
each layer path is replaced by `<entry key>`.
"""
from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
from pathlib import Path

SKILL = Path(__file__).resolve().parents[1]
SKILLS = SKILL.parent
KINDS = ('ledger', 'targets', 'rules', 'stages', 'findings', 'judgments')


def normalise(text, layout):
    """Replace every layer path in `text` by `<key>` (longest path first; ties in key order)."""
    ordered = sorted(layout['entries'].items(), key=lambda kv: (-len(kv[1]), kv[0]))
    names = {}
    for key, rel in ordered:
        names.setdefault(rel, key)
    pattern = re.compile('|'.join(re.escape(rel) for _, rel in ordered))
    return pattern.sub(lambda m: '<' + names[m.group(0)] + '>', text)  # one pass: a replacement is never rescanned


def _rule_ids(manifest):
    web = (manifest or {}).get('web') or {}
    ids = []
    for row in web.get('rules') or []:
        ids.append(str(row.get('id')) if isinstance(row, dict) and 'id' in row else json.dumps(row, sort_keys=True))
    return sorted(ids)


def _ia(root, skills, layout):
    """The IA check on its own (findings by level, code and normalised path), when the repository has an infoarch."""
    script = Path(skills) / 'own-infoarch-author/scripts/infoarch_author.py'
    run = subprocess.run([sys.executable, str(script), 'check', '--root', str(root), '--date', '2026-01-01'],
                         capture_output=True, text=True, timeout=120, cwd=str(root))
    try:
        data = json.loads(run.stdout)
    except ValueError:
        return {'present': False, 'exit_code': run.returncode, 'counts': {}, 'findings': ['unreadable: ' + normalise(run.stdout[:200] + run.stderr[:200], layout)]}
    findings = sorted(f"{f['level']} {f['code']} {normalise(f.get('path', ''), layout)}" for f in data.get('findings', []))
    return {'present': bool(data.get('targets')), 'exit_code': data.get('exit_code'),
            'counts': data.get('targets') or {}, 'findings': findings}


def take(repo, skills=SKILLS, layout_file=None, catalog_dir=None):
    skills = Path(skills)
    scripts = skills / 'own-design-author/scripts'
    layout_file = Path(layout_file) if layout_file else skills / 'own-design-author/references/layer-paths.json'
    layout = json.loads(layout_file.read_text())
    sys.path.insert(0, str(scripts))
    import design_author as da  # the skills tree under test, not necessarily this one
    root = Path(repo).resolve()
    kwargs = {'catalog_dir': Path(catalog_dir)} if catalog_dir else {}
    result = da.check_design(root, **kwargs)
    targets = {k: v for k, v in sorted((result.get('approval_targets') or {}).items())}
    ledger_path = root / layout['entries']['ledger']
    rows = da.read_data(ledger_path) if ledger_path.exists() else []
    ledger = {}
    for row in rows or []:
        key = row['key']
        ledger[key] = {'value_hash': row['value_hash'], 'current_hash': targets.get(key),
                       'valid': targets.get(key) == row['value_hash']}
    ledger_text = (root / layout['entries']['judgments']).read_text(encoding='utf-8') if (root / layout['entries']['judgments']).exists() else ''
    judgments = {c['id']: {'hash': c['judgment_hash'], 'answered': c['judgment_hash'] in ledger_text}
                 for c in result.get('judgment_candidates') or []}
    stages = {s['id']: {'exit_code': s.get('exit_code'), 'targets': s.get('targets')}
              for s in (result.get('report') or {}).get('executed', [])}
    return {
        'schema_version': 1,
        'ia': _ia(root, skills, layout),
        'ledger': dict(sorted(ledger.items())),
        'targets': targets,
        'rules': _rule_ids(result.get('manifest')),
        'judgments': dict(sorted(judgments.items())),
        'verdict': {'exit_code': result['exit_code'], 'stages': dict(sorted(stages.items())),
                    'errors': sorted(normalise(str(e), layout) for e in result.get('errors', [])),
                    'warnings': sorted(normalise(str(w), layout) for w in result.get('warnings', []))},
    }


def compare(before, after):
    diffs = []

    def diff_maps(kind, a, b):
        for key in sorted(set(a) | set(b)):
            if key not in b:
                diffs.append(f'{kind}: {key} is missing after')
            elif key not in a:
                diffs.append(f'{kind}: {key} is new after')
            elif a[key] != b[key]:
                diffs.append(f'{kind}: {key} changed: {json.dumps(a[key], sort_keys=True)} -> {json.dumps(b[key], sort_keys=True)}')

    diff_maps('ledger', before['ledger'], after['ledger'])
    diff_maps('targets', before['targets'], after['targets'])
    diff_maps('judgments', before['judgments'], after['judgments'])
    for rule in sorted(set(before['rules']) ^ set(after['rules'])):
        diffs.append(f'rules: {rule} is not in both')
    diff_maps('stages', before['verdict']['stages'], after['verdict']['stages'])
    for field in ('present', 'exit_code', 'counts'):
        if before['ia'][field] != after['ia'][field]:
            diffs.append(f"ia.{field} changed: {before['ia'][field]!r} -> {after['ia'][field]!r}")
    for finding in sorted(set(before['ia']['findings']) ^ set(after['ia']['findings'])):
        diffs.append(f'ia.findings: {finding} is not in both')
    for field in ('exit_code', 'errors', 'warnings'):
        if before['verdict'][field] != after['verdict'][field]:
            diffs.append(f'verdict.{field} changed: {before["verdict"][field]!r} -> {after["verdict"][field]!r}')
    compared = {'ledger': len(before['ledger']), 'targets': len(before['targets']), 'rules': len(before['rules']),
                'stages': len(before['verdict']['stages']), 'findings': len(before['ia']['findings']), 'judgments': len(before['judgments'])}
    return {'differences': diffs, 'compared': compared}


def exit_code(result, minimum=None, expected=None):
    """0 equal, 1 different, 2 the comparison looked at too little (never a pass)."""
    if result['differences']:
        return 1
    compared = result['compared']
    if not any(compared.values()):
        return 2
    if minimum is not None and any(compared[k] < minimum for k in ('targets', 'rules', 'stages')):
        return 2
    for kind, count in (expected or {}).items():
        if compared.get(kind, 0) < count:
            return 2
    return 0


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = parser.add_subparsers(dest='command', required=True)
    s = sub.add_parser('snapshot')
    s.add_argument('--repo', required=True)
    s.add_argument('--skills', default=str(SKILLS), help='skills root whose check is run (default: this one)')
    s.add_argument('--layout', help='layer-paths.json used to normalise texts (default: the skills root\'s own)')
    s.add_argument('--catalog-dir')
    s.add_argument('--out', required=True)
    c = sub.add_parser('compare')
    c.add_argument('before')
    c.add_argument('after')
    c.add_argument('--min', type=int, default=1, help='each of targets, rules, stages must have at least this many records')
    c.add_argument('--expect', action='append', default=[], metavar='KIND=N', help='at least N records of KIND (repeatable)')
    args = parser.parse_args(argv)
    if args.command == 'snapshot':
        data = take(args.repo, args.skills, args.layout, args.catalog_dir)
        Path(args.out).write_text(json.dumps(data, ensure_ascii=False, indent=2, sort_keys=True) + '\n')
        print(json.dumps({'written': args.out, 'targets': len(data['targets']), 'ledger': len(data['ledger']),
                          'rules': len(data['rules']), 'stages': len(data['verdict']['stages'])}))
        return 0
    expected = {k: int(v) for k, v in (item.split('=') for item in args.expect)}
    result = compare(json.loads(Path(args.before).read_text()), json.loads(Path(args.after).read_text()))
    code = exit_code(result, args.min, expected)
    print(json.dumps({**result, 'exit_code': code}, ensure_ascii=False, sort_keys=True))
    return code


if __name__ == '__main__':
    raise SystemExit(main())
