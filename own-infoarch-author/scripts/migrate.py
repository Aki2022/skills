"""One atomic v2 -> v4 migration of design/info_arch/ (SPEC-information-architecture-standardization v5 requirement 18,
re-targeted at the v6 positioning shape; the v5 -> v6 bulk migration is ISSUE-17).

Mechanical moves only: problems (prb) -> pains (pan) with the same number, each audience gets a segment (seg, same
number) and an empty job for that segment, each pain belongs to its audience's segment, message addresses follow, outlines go to version 2 (problem -> pain),
rendered HTML outputs rename data-ia-problem / data-ia-role="problem", and design.yaml pins modules.ia 0.4.0.
What needs judgment (jobs, arrows, values, strengths, alternatives, categories, message serves, outline strategy) is
listed, never invented. Nothing is committed: undo with git (the report says how). A v4 tree is a no-op, so the tool
is rerunnable; a v3 (v5) tree is refused here.
"""
from __future__ import annotations

import json
import re
from pathlib import Path

import yaml

from check import POSITIONING, _glob_re

HEADER = ('# design/info_arch/info_arch.yaml — 伝え方の SSOT（スキーマ版 4）\n'
          '# v2 から infoarch-author migrate で移した。positioning（ジョブ・代替・カテゴリー・強み・提供価値と矢印）は下書きが要る。\n')
OLD = re.compile(r'\bprb-(\d{4,})\b')
LEFTOVER = re.compile(r'\bprb-\d{4,}\b|data-ia-problem|^problems:', re.M)
# In rendered HTML only the attributes are machine-read; an old id in a comment is reported, not counted as missed
# (D5 lets the migration touch data-ia-* attributes only).
LEFTOVER_HTML = re.compile(r'data-ia-problem|data-ia-[a-z-]+="prb-\d{4,}"|data-ia-role="problem"')


def _pan(text):
    return OLD.sub(lambda m: 'pan-' + m.group(1), text)


def _dump(data):
    return yaml.safe_dump(data, allow_unicode=True, sort_keys=False)


def plan(root):
    """-> (report, {path: new text}) without writing anything."""
    root = Path(root)
    base = root / 'design/info_arch'
    ssot_path = base / 'info_arch.yaml'
    ssot = yaml.safe_load(ssot_path.read_text(encoding='utf-8'))
    if ssot.get('schema_version') == 4:
        return {'status': 'already-v4', 'changes': 0, 'id_map': {}, 'missed': [], 'drafting_required': []}, {}
    if ssot.get('schema_version') != 2:
        raise ValueError('info_arch.yaml is not schema 2 (schema 3, the v5 shape, moves to v6 with the ISSUE-17 tool)')
    id_map, drafting, writes = {}, [], {}
    new = {'schema_version': 4, 'ledger': {}, **{b: {} for b in POSITIONING}, 'edges': []}
    nxt = dict(ssot['ledger']['next'])
    for aid, aud in ssot.get('audiences', {}).items():
        n = aid.split('-')[1]
        sid, jid = 'seg-' + n, 'job-' + n
        id_map[aid] = sid
        ok = aud.get('state') == 'proposed' and aud.get('label') and aud.get('situation')
        new['segments'][sid] = ({'state': 'proposed', 'provenance': 'proposal', 'qualifier': aud['label'], 'situation': aud['situation']}
                                if ok else {'state': 'empty', 'provenance': 'proposal'})
        new['jobs'][jid] = {'state': 'empty', 'provenance': 'proposal', 'segment': sid}
        drafting.append(f'jobs:{jid}')
    for pid, prb in ssot.get('problems', {}).items():
        nid = 'pan-' + pid.split('-')[1]
        id_map[pid] = nid
        pain = {k: v for k, v in prb.items() if k != 'audience'}
        if prb.get('audience'):
            pain['segment'] = 'seg-' + prb['audience'].split('-')[1]
        new['pains'][nid] = pain
    tomb = {}
    for old, target in (ssot['ledger'].get('tombstones') or {}).items():
        if old.startswith('prb-'):
            id_map.setdefault(old, 'pan-' + old.split('-')[1])
        tomb[_pan(old)] = _pan(target) if isinstance(target, str) else target
    new['ledger'] = {'next': {**{k: v for k, v in nxt.items() if k != 'prb'}, 'pan': nxt.get('prb', 1),
                              'seg': nxt.get('aud', 1), 'job': nxt.get('aud', 1),
                              **{p: 1 for p in ('gan', 'val', 'alt', 'str', 'cat')}},
                     'tombstones': tomb}
    if 'policy' in ssot:
        new['policy'] = ssot['policy']
    for block in ('audiences', 'messages', 'proofs', 'facts', 'entities', 'voices'):
        new[block] = ssot.get(block, {})
    for aid in new['audiences']:
        new['audiences'][aid]['segment'] = id_map[aid]
    for mid, msg in new['messages'].items():
        if 'addresses' in msg:
            msg['addresses'] = [_pan(x) for x in msg['addresses']]
        if msg.get('state') == 'proposed' and not msg.get('serves'):
            drafting.append(f'messages:{mid}.serves')
    drafting.append('values/strengths/alternatives/categories と edges（has_job・need・fulfills・member_of ほか）: 資料から下書きする（D4）')
    writes[ssot_path] = HEADER + _dump(new)
    outputs = []
    for f in sorted((base / 'outlines').glob('*.yaml')):
        o = yaml.safe_load(f.read_text(encoding='utf-8'))
        if o.get('schema_version') == 1:
            o['schema_version'] = 2
            for p in o['pages']:
                for s in p['sections']:
                    if s.get('role') == 'problem':
                        s['role'] = 'pain'
                    if 'problem' in s:  # rebuild in place so the key keeps its position
                        rebuilt = {}
                        for k, v in s.items():
                            rebuilt['pain' if k == 'problem' else k] = _pan(v) if k == 'problem' else v
                        s.clear(); s.update(rebuilt)
            writes[f] = _dump(o)
        if 'strategy' not in o:
            drafting.append(f'outlines:{o["id"]}.strategy')
    for f in sorted((base / 'channels').glob('*.yaml')):
        outputs += yaml.safe_load(f.read_text(encoding='utf-8')).get('outputs', [])
    globs = [_glob_re(g) for g in outputs]
    for f in sorted(p for p in root.rglob('*.htm*') if p.is_file() and '.git' not in p.parts and 'node_modules' not in p.parts):
        rel = f.relative_to(root).as_posix()
        if not any(g.match(rel) for g in globs):
            continue
        text = f.read_text(encoding='utf-8')
        moved = re.sub(r'data-ia-problem="prb-(\d{4,})"', r'data-ia-pain="pan-\1"', text).replace('data-ia-role="problem"', 'data-ia-role="pain"')
        if moved != text:
            writes[f] = moved
    design = root / 'design/design.yaml'
    if design.is_file():
        cfg = json.loads(design.read_text(encoding='utf-8'))
        if (cfg.get('modules') or {}).get('ia') == '0.2.0':
            cfg['modules']['ia'] = '0.4.0'
            raw = design.read_text(encoding='utf-8')
            writes[design] = json.dumps(cfg, ensure_ascii=False, indent=2 if raw.startswith('{\n') else None) + ('\n' if raw.endswith('\n') else '')
    ledger = root / 'design/judgments.json'
    ia_rows = 0
    if ledger.is_file():
        rows = json.loads(ledger.read_text(encoding='utf-8')).get('rows', [])
        ia_rows = sum(1 for r in rows if str(((r.get('candidate') or {}).get('input') or {}).get('kind', '')).startswith('ia-'))
    report = {'status': 'preview', 'changes': len(writes), 'files': sorted(p.relative_to(root).as_posix() for p in writes),
              'id_map': dict(sorted(id_map.items())), 'drafting_required': drafting,
              'invalidated_judgments': {'ia_rows': ia_rows},
              'recovery': '移行はコミットしない。戻すときは git checkout -- <files>（コミット後なら git revert <commit>）。'}
    return report, writes


def missed(root, writes):
    """-> (files with old names left where a machine reads them, files whose HTML comments still mention an old id)."""
    root = Path(root)
    out, comments = [], []
    candidates = [p for p in (root / 'design/info_arch').rglob('*') if p.is_file()] + [p for p in writes if p.suffix.startswith('.htm')]
    for p in sorted(set(candidates)):
        text = writes.get(p) if p in writes else p.read_text(encoding='utf-8', errors='replace')
        rel = p.relative_to(root).as_posix()
        if p.suffix.startswith('.htm'):
            if LEFTOVER_HTML.search(text):
                out.append(rel)
            elif OLD.search(text):
                comments.append(rel)
        elif LEFTOVER.search(text):
            out.append(rel)
    return out, comments


def run(root, write=False):
    report, writes = plan(root)
    if report['status'] == 'already-v4':
        return report, 0
    report['missed'], report['comment_mentions'] = missed(root, writes)
    if write:
        for p, text in writes.items():
            p.write_text(text, encoding='utf-8')
        report['status'] = 'migrated'
    return report, 2 if report['missed'] else 0
