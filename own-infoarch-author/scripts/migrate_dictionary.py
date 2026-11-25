"""One migration of ux/infoarch/ from schema 4 (facts, entities and a voice inside the SSOT) to schema 5 and the
docs dictionary (SPEC-deliverable-production-layers requirements 8 and 9).

- every fact (`fct-NNNN`) becomes a dictionary `value` entry and every entity (`ent-NNNN`) a `name` entry, under the id
  the caller maps it to: ids are names people read (`value-answer-time`), so the tool cannot invent them;
- a value entry must declare `expires` or `evergreen: true`: the old implicit lifetime (as_of plus 365 days) is gone, and
  which of the two a value gets is a decision about the repository's own content, so it is asked for, never defaulted;
- the SSOT loses `facts`, `entities` and `voices` (the voice is retired), `schema_version` becomes 5, the ledger and the
  tombstones lose those prefixes. Comments are kept: the blocks are removed from the text, and the result is read back
  and must equal what the schema 4 file lifts to, or nothing is written;
- every `{{fct-NNNN}}` in an outline becomes `{{<dictionary id>}}`, and a rendered output's `data-ia-fact` / `data-ia-entity`
  marks become `data-dict`.

What cannot be moved mechanically (a fact or entity with no mapping, a value with no lifetime, an id that is not a
dictionary id, a placeholder or mark that names nothing, an existing dictionary file) is listed under `unmigrated`:
one item is a failure (exit 2) and nothing is written. The dry run also checks the migrated tree in a scratch copy and
reports its failures next to the original's: a migration that makes the check fail is refused. Nothing is committed.
An already migrated tree is a no-op, so the tool is rerunnable. ux/approvals.yaml is not touched: approval hashes
cover the positioning paths, and facts, entities and voices are on none of them.
"""
from __future__ import annotations

import json
import re
import shutil
import tempfile
from pathlib import Path

import check
import common
import dictionary
import layer_paths  # on sys.path through common
import positioning_model as pm

DOMAIN = 'product'
TOP_LEVEL = re.compile(r'^[A-Za-z_]')
HTML_MARKS = (('data-ia-fact', 'fct'), ('data-ia-entity', 'ent'))
MAP_KEYS = {'id', 'expires', 'evergreen', 'variants'}


def _block_ranges(lines, names):
    """[(start, end)] of the top-level `name:` blocks (to the next top-level key, or to the end)."""
    starts = [i for i, ln in enumerate(lines) if TOP_LEVEL.match(ln)]
    out = []
    for n, i in enumerate(starts):
        if lines[i].split(':', 1)[0] in names:
            out.append((i, starts[n + 1] if n + 1 < len(starts) else len(lines)))
    return out


def _strip_ledger(text):
    """Drop fct/ent/vce from `ledger.next` (block or flow style) and from tombstones, textually."""
    out = []
    for ln in text.splitlines(keepends=True):
        if re.match(r'^\s+(fct|ent|vce):\s*\d+\s*$', ln):
            continue
        ln = re.sub(r'\b(?:fct|ent|vce):\s*\d+\s*,\s*', '', ln)
        ln = re.sub(r',\s*\b(?:fct|ent|vce):\s*\d+\s*(?=\})', '', ln)
        if re.match(r'^\s+((fct|ent|vce)-[0-9]+|"(fct|ent|vce)-[0-9]+"|\'(fct|ent|vce)-[0-9]+\'):', ln):
            continue
        out.append(ln)
    return ''.join(out)


def rewrite_ssot(text):
    lines = text.splitlines(keepends=True)
    drop = {i for a, b in _block_ranges(lines, ('facts', 'entities', 'voices')) for i in range(a, b)}
    text = ''.join(ln for i, ln in enumerate(lines) if i not in drop)
    text = re.sub(r'^schema_version:\s*4\s*$', 'schema_version: 5', text, count=1, flags=re.M)
    return _strip_ledger(text)


def entries_for(terms, mapping, unmigrated):
    """-> dictionary entries for the facts and entities of a schema 4 SSOT, and {old id: new id}."""
    out, ids = [], {}
    for block, kind, prefix in (('facts', 'value', 'fct'), ('entities', 'name', 'ent')):
        for old, atom in sorted(terms[block].items()):
            row = mapping.get(old)
            if not isinstance(row, dict) or not row.get('id'):
                unmigrated.append({'id': old, 'reason': 'not-mapped',
                                   'text': '辞書の id が決まっていない（--map に {"id": "<kind>-<slug>"} を書く）'})
                continue
            extra = set(row) - MAP_KEYS
            if extra:
                unmigrated.append({'id': old, 'reason': 'map-key', 'text': f'--map に知らないキー {sorted(extra)}'})
                continue
            entry = {'id': row['id'], 'kind': kind, 'text': atom['value'] if kind == 'value' else atom['name'],
                     'meaning': atom['label'] if kind == 'value' else (atom.get('type') or 'name')}
            if row.get('variants'):
                entry['variants'] = row['variants']
            if kind == 'value':
                if bool(row.get('expires')) == bool(row.get('evergreen')):
                    unmigrated.append({'id': old, 'reason': 'lifetime',
                                       'text': '値は expires（日付）か evergreen: true のどちらか 1 つを決める（暗黙の期限は無い）'})
                    continue
                if row.get('expires'):
                    entry['expires'] = row['expires']
                else:
                    entry['evergreen'] = True
                if atom.get('as_of'):
                    entry['as_of'] = atom['as_of']
            if not row['id'].startswith(kind + '-') or not dictionary.ID_RE.match(row['id']):
                unmigrated.append({'id': old, 'reason': 'dictionary-id', 'text': f'{row["id"]!r} は辞書の id（{kind}-<slug>）でない'})
                continue
            out.append(entry)
            ids[old] = row['id']
    return out, ids


def _ref_pattern(prefix):
    return re.compile(r'\{\{(' + prefix + r'-[0-9]{4,})\}\}')


def rewrite_outline_text(text, ids, where, unmigrated):
    def swap(m):
        if m.group(1) not in ids:
            unmigrated.append({'id': where, 'reason': 'placeholder', 'text': f'{{{{{m.group(1)}}}}} は SSOT に無い固定値'})
            return m.group(0)
        return '{{' + ids[m.group(1)] + '}}'
    return _ref_pattern('fct').sub(swap, text)


def rewrite_html(text, ids, where, unmigrated):
    for attr, prefix in HTML_MARKS:
        def swap(m, prefix=prefix):
            if m.group(1) not in ids:
                unmigrated.append({'id': where, 'reason': 'html-mark', 'text': f'{attr}="{m.group(1)}" は SSOT に無い'})
                return m.group(0)
            return f'data-dict="{ids[m.group(1)]}"'
        text = re.sub(re.escape(attr) + r'="(' + prefix + r'-[0-9]{4,})"', swap, text)
    return text


def run(root, mapping, domain=DOMAIN, write=False, date=None):
    import datetime as dt
    import yaml
    root = Path(root)
    date = date or dt.date.today().isoformat()
    report = common.Report('migrate-dictionary', date)
    ssot, terms = pm.load_ssot(root, report)
    if ssot is None:
        return {'exit_code': 2, 'written': False, 'unmigrated': [{'id': 'ssot', 'reason': 'unreadable',
                                                                   'text': [f.message for f in report.findings if f.level == 'fail']}]}
    if terms is None:
        return {'exit_code': 0, 'written': False, 'status': 'already-migrated',
                'note': 'スキーマ版 5 の木。移すものは無い（再実行しても何も変わらない）'}
    unmigrated, changes = [], {}
    new_entries, ids = entries_for(terms, mapping, unmigrated)
    target = Path(dictionary.DICTIONARY_DIR) / f'{domain}.json'
    if (root / target).exists() or (root / target).is_symlink():
        unmigrated.append({'id': str(target), 'reason': 'dictionary-exists', 'text': '辞書のファイルがすでにある（上書きしない。別の --domain にするか、手で足す）'})
    existing, _, derr = dictionary.load(root)
    clash = sorted({e['id'] for e in new_entries} & {e['id'] for e in existing})
    if clash:
        unmigrated.append({'id': ','.join(clash), 'reason': 'dictionary-id-taken', 'text': '辞書にすでにある id'})
    dict_doc = {'schema_version': dictionary.SCHEMA_VERSION, 'domain': domain,
                'summary': f'製品の正式名称と固定値（infoarch-author migrate-dictionary が {layer_paths.rel("infoarch")}/ から移した）',
                'entries': new_entries}
    shape = dictionary.shape_errors({f'{domain}.json': dict_doc}) if new_entries else []
    for e in shape:
        unmigrated.append({'id': str(target), 'reason': 'dictionary-shape', 'text': e})
    # the SSOT, textually, and read back
    ssot_path = layer_paths.path(root, 'positioning_file')
    new_text = rewrite_ssot(ssot_path.read_text(encoding='utf-8'))
    try:
        if common.safe_load_text(new_text) != pm.lift_v4(common.load_yaml(ssot_path))[0]:
            unmigrated.append({'id': str(pm.SSOT_PATH), 'reason': 'ssot-rewrite', 'text': '書き換えた SSOT が、スキーマ版 4 を持ち上げた形と一致しない（手で直す）'})
    except common.YamlError as exc:
        unmigrated.append({'id': str(pm.SSOT_PATH), 'reason': 'ssot-rewrite', 'text': str(exc)})
    changes[str(pm.SSOT_PATH)] = new_text
    # scripts and their rendered outputs (the retired outlines/ is not rewritten: check fails on it)
    base = layer_paths.path(root, 'infoarch')
    scripts = base / check.SCRIPTS_DIR
    for f in sorted(scripts.glob('*.yaml')) if scripts.is_dir() else []:
        rel = f.relative_to(root)
        text = f.read_text(encoding='utf-8')
        out = rewrite_outline_text(text, ids, str(rel), unmigrated)
        if out != text:
            changes[str(rel)] = out
        for output in (yaml.safe_load(text) or {}).get('outputs') or []:
            html_path = root / output
            if html_path.is_file() and html_path.suffix.lower() in ('.html', '.htm') and not html_path.is_symlink():
                html = html_path.read_text(encoding='utf-8')
                new_html = rewrite_html(html, ids, output, unmigrated)
                if new_html != html:
                    changes[output] = new_html
    result = {'exit_code': 0, 'written': False, 'id_map': ids, 'dictionary': {str(target): dict_doc},
              'changed': sorted(changes), 'unmigrated': unmigrated, 'dropped': ['voices（声は廃止）']}
    if unmigrated:
        return {**result, 'exit_code': 2}
    # prove the migrated tree in a scratch copy: its failures must not exceed the original's
    with tempfile.TemporaryDirectory() as tmp:
        scratch = Path(tmp)
        for sub in (layer_paths.LAYERS_ROOT, 'docs'):
            if (root / sub).is_dir():
                shutil.copytree(root / sub, scratch / sub, symlinks=True)
        (scratch / target).parent.mkdir(parents=True, exist_ok=True)
        (scratch / target).write_text(json.dumps(dict_doc, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
        for rel, text in changes.items():
            if not rel.startswith(layer_paths.LAYERS_ROOT + '/'):
                continue
            (scratch / rel).write_text(text, encoding='utf-8')
        before, _ = check.check(root, date)
        after, _ = check.check(scratch, date)
    fails_before = sorted({f.code for f in before.findings if f.level == 'fail'})
    fails_after = sorted({f.code for f in after.findings if f.level == 'fail'})
    result['check'] = {'before': {'exit_code': before.exit_code, 'fail_codes': fails_before},
                       'after': {'exit_code': after.exit_code, 'fail_codes': fails_after}}
    if fails_after != fails_before and not set(fails_after) <= set(fails_before):
        result['exit_code'] = 2
        result['unmigrated'] = [{'id': 'check', 'reason': 'check-regressed',
                                 'text': f'移行した木が新しい失格を出す: {sorted(set(fails_after) - set(fails_before))}'}]
        return result
    if write:
        for rel, text in changes.items():
            (root / rel).write_text(text, encoding='utf-8')
        (root / target).parent.mkdir(parents=True, exist_ok=True)
        (root / target).write_text(json.dumps(dict_doc, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
        result['written'] = True
    return result
