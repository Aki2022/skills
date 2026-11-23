"""The answer key for the meaning judgment (SPEC-information-architecture-standardization v6, requirement 12).

Every piece of copy is paired with the value its message points to (`own`, expected to convey it) and re-paired with
the other approved values of the same positioning (`swap`, candidates that should not). A positioning with a single
value is re-paired with another positioning's value. The machine only proposes: the owner's labels (conveys /
does not convey / unknown) decide. Agreement is counted against the shared judgment ledger (a score at or above theta
means Jev says "does not convey"). When the owner's labels hold no conveying pair or no non-conveying pair, there is
nothing to measure against, and the result is "unmeasurable", never a number.

Where Jev's answers live (ADR-20261123-ia-skill-completion-judgments-and-measurement):
- an `own` pair is the shipping question itself (the very candidate `judge` asks), so its answer is the shared judgment
  ledger's (`ux/judgments.json`): what is measured is what ships;
- a `swap` pair exists only to measure, so its answer is recorded in a measurement ledger beside the answer key
  (`measurement-judgments.json`, the shared ledger's shape). The ship gate never reads it, and `record` never writes the
  shared ledger.
Accuracy is measured from the two together, so either side being empty is "unmeasurable". Only `record` calls a model,
through own-design-author's bounded `judgment_run` (a preview unless `execute`).
"""
from __future__ import annotations

import html
import json
import os
import tempfile
from pathlib import Path

import judge
import outline as ol
import positioning as pos


MEASUREMENT_LEDGER = 'measurement-judgments.json'
ANSWERS = 'judgment-answers.json'  # the name judgment_run.py requires


def build(m, approved, policy, jl):
    """-> [item] sorted by id. An item: id (`<outline>/<sec>/<cpy>@<value>`), kind (own / swap), text, value, candidate."""
    atoms = pos.atoms(m)
    values = {v: a for v, a in atoms.items() if pos._prefix(v) == 'val' and v in approved}
    facts = m.ssot.get('facts') or {}
    out = []
    for oid, (_, o) in sorted(m.outlines.items()):
        seg = ol.segment_of(o)
        same = sorted(v for v, a in values.items() if a.get('segment') == seg)
        other = sorted(v for v, a in values.items() if a.get('segment') != seg)
        for sid, sec, copies in ol.ordered_sections(o):
            if sec['role'] in judge.VALUE_FIT_SKIP_ROLES:
                continue  # not asked whether it conveys a value, so nothing to measure
            for cid, c in copies:
                own = (o['messages'].get(c['message']) or {}).get('value')
                if own not in values:
                    continue  # an unapproved value has nothing settled to measure against
                swaps = [v for v in same if v != own] or other[:1]
                text = judge._expand(c['text'], facts)
                for v, kind in [(own, 'own')] + [(x, 'swap') for x in swaps]:
                    ident = f'{oid}/{sid}/{cid}@{v}'
                    value = {'id': v, 'text': values[v].get('text', '')}
                    # An own pair asks what `judge` asks (same id, so the same judgment hash and one answer for both).
                    asked = f'{oid}/{sid}/{cid}' if kind == 'own' else ident
                    out.append({'id': ident, 'kind': kind, 'text': text, 'value': value,
                                'candidate': judge.value_fit(jl, policy, asked, text, oid, c['message'], value)})
    return sorted(out, key=lambda i: i['id'])


def key_hash(items):
    """Binds a labels file to exactly these pairs (ids and judgment hashes)."""
    return pos.sha256([[i['id'], i['candidate']['judgment_hash']] for i in items])


def page(items):
    """A self-contained labeling page: one row per pair, conveys / does not convey / unknown, and a JSON download."""
    key = key_hash(items)
    rows = []
    for n, i in enumerate(items):
        name = f'p{n}'
        rows.append(
            f'<tr data-id="{html.escape(i["id"])}"><td class="id">{html.escape(i["id"])}</td>'
            f'<td>{html.escape(i["text"])}</td><td>{html.escape(i["value"]["text"])}</td><td class="choice">'
            f'<label><input type="radio" name="{name}" value="true">○ 伝わる</label> '
            f'<label><input type="radio" name="{name}" value="false">× 伝わらない</label> '
            f'<label><input type="radio" name="{name}" value="null" checked>？ 分からない</label></td></tr>')
    return f'''<!doctype html><html lang="ja"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>正解表のラベル付け</title><style>
:root{{--fg:#1b1b1b;--bg:#fff;--line:#d0d0d0;--muted:#666}}
@media (prefers-color-scheme:dark){{:root{{--fg:#eee;--bg:#161616;--line:#444;--muted:#aaa}}}}
body{{margin:0;padding:16px;font:15px/1.6 system-ui,sans-serif;color:var(--fg);background:var(--bg)}}
table{{border-collapse:collapse;width:100%}} td,th{{border-bottom:1px solid var(--line);padding:8px;vertical-align:top;text-align:left}}
.id{{font:12px ui-monospace,monospace;color:var(--muted)}} .choice label{{display:block;white-space:nowrap}}
button{{font:inherit;padding:8px 16px;margin:16px 0}}
.warn{{border:2px solid #c00;padding:8px 12px}}
textarea{{width:100%;box-sizing:border-box;font:12px ui-monospace,monospace;color:var(--fg);background:var(--bg);border:1px solid var(--line)}}
</style></head><body>
<h1>正解表のラベル付け</h1>
<p>文案を読んだ人に、その提供価値が伝わるかを選んでください。分からないものは「分からない」のままにします（数えません）。
選び終わったら「ラベルを保存」で labels.json を保存するか、「ラベルをコピー」で中身をコピーして、AI に渡してください。</p>
<p class="id">answer key: {key}（{len(items)} 組）</p>
<p id="no-script" class="warn">このページの「ラベルを保存」「ラベルをコピー」は、ページの中のプログラムで動きます。ファイルのプレビューのようにプログラムを動かさない画面では、押しても何も起きません。Safari や Chrome などのブラウザで開き直してください。この表示が消えていれば、保存とコピーは動きます。</p>
<table><thead><tr><th>組</th><th>文案</th><th>提供価値</th><th>伝わるか</th></tr></thead><tbody>
{chr(10).join(rows)}
</tbody></table>
<button id="save">ラベルを保存</button> <button id="copy">ラベルをコピー</button>
<p>保存したファイルが見つからないときは、下の欄の中身をそのままコピーして AI に渡してください（選ぶたびに更新されます）。</p>
<textarea id="labels-json" readonly rows="8" aria-label="labels.json の中身"></textarea>
<script>
function current() {{
  var labels = {{}};
  document.querySelectorAll('tr[data-id]').forEach(function (tr) {{
    var v = tr.querySelector('input:checked').value;
    labels[tr.getAttribute('data-id')] = v === 'null' ? null : v === 'true';
  }});
  return JSON.stringify({{answer_key: '{key}', labels: labels}}, null, 2);
}}
var box = document.getElementById('labels-json');
function show() {{ box.value = current(); }}
document.querySelectorAll('input[type=radio]').forEach(function (r) {{ r.addEventListener('change', show); }});
show();
document.getElementById('no-script').hidden = true;
document.getElementById('save').addEventListener('click', function () {{
  var blob = new Blob([current()], {{type: 'application/json'}});
  var a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = 'labels.json'; a.click();
}});
document.getElementById('copy').addEventListener('click', function () {{
  show(); box.select();
  if (navigator.clipboard) {{ navigator.clipboard.writeText(box.value).catch(function () {{ document.execCommand('copy'); }}); }}
  else {{ document.execCommand('copy'); }}
}});
</script></body></html>
'''


def _scores(ledger):
    return {(row.get('candidate') or {}).get('judgment_hash'): row.get('score') for row in (ledger or {}).get('rows') or []}


def measure(items, labels, ledger, theta, measurement=None):
    """-> the agreement between the owner's labels and Jev's answers, or 'unmeasurable'.

    own pairs are read from `ledger` (the shared one) and swap pairs from `measurement` only, so an answer in the wrong
    place is "uncovered", never silently counted."""
    if not isinstance(labels, dict) or labels.get('answer_key') != key_hash(items):
        raise ValueError('labels were made for another answer key (rebuild the page and label again)')
    by_id = {i['id']: i for i in items}
    given = labels.get('labels') or {}
    unknown_ids = sorted(set(given) - set(by_id))
    if unknown_ids:
        raise ValueError(f'labels name unknown pairs: {unknown_ids[:5]}')
    if any(v not in (True, False, None) for v in given.values()):
        raise ValueError('a label is true (conveys), false (does not convey) or null (unknown)')
    source = {'own': _scores(ledger), 'swap': _scores(measurement)}

    def score(key):
        item = by_id[key]
        return source[item['kind']].get(item['candidate']['judgment_hash'])

    confirmed = {k: v for k, v in given.items() if v is not None}
    covered = {k: v for k, v in confirmed.items() if score(k) is not None}
    uncovered = sorted(set(confirmed) - set(covered))
    counts = {'labeled': len(confirmed), 'conveys': sum(confirmed.values()),
              'does_not_convey': sum(not v for v in confirmed.values()),
              'unknown': sum(v is None for v in given.values()), 'uncovered': len(uncovered)}
    judged = {kind: sum(by_id[k]['kind'] == kind for k in covered) for kind in ('own', 'swap')}
    base = {'answer_key': key_hash(items), 'counts': counts, 'judged': judged, 'uncovered': uncovered, 'theta': theta}
    conveys = [k for k, v in covered.items() if v]
    not_conveys = [k for k, v in covered.items() if not v]
    missing = [f'{kind} pairs answered in the {where}: 0' for kind, where in
               (('own', 'shared judgment ledger'), ('swap', MEASUREMENT_LEDGER)) if not judged[kind]]
    if missing:
        return {**base, 'status': 'unmeasurable', 'reason': '; '.join(missing) + ' (agreement is not measured)'}
    if not conveys or not not_conveys:
        side = 'conveys' if not conveys else 'does not convey'
        return {**base, 'status': 'unmeasurable',
                'reason': f'judged pairs the owner labeled "{side}" are 0: agreement is not measured'}
    flagged = {k: score(k) >= theta for k in covered}
    agree = [k for k in covered if flagged[k] == (not covered[k])]
    return {**base, 'status': 'measured', 'agreement': len(agree) / len(covered),
            'misses': sorted(k for k in not_conveys if not flagged[k]),
            'false_detections': sorted(k for k in conveys if flagged[k])}


# ---- the measurement ledger and the tool that records into it -----------------------------------------------------

def _ledger_path(key_dir):
    return Path(key_dir) / MEASUREMENT_LEDGER


def read_measurement(key_dir):
    """-> the measurement ledger beside the answer key, or None when none was recorded. A damaged one raises."""
    path = _ledger_path(key_dir)
    if not judge._jr()._safe_path(path.absolute()):
        raise ValueError(f'{MEASUREMENT_LEDGER} が symlink を通る')
    if not path.exists():
        return None
    jl = judge._jl()
    try:
        data = json.loads(path.read_text(encoding='utf-8'))
        jl.validate_ledger(data, jl.load_policy())
    except (ValueError, KeyError, TypeError) as exc:
        raise ValueError(f'{MEASUREMENT_LEDGER} が読めない・今の policy と合わない（記録し直す）: {exc}') from None
    return data


def record(items, key_dir, limit, *, execute=False, _usage_path=None, _transport=None):
    """Ask Jev the swap pairs not yet answered and record them in the measurement ledger beside the answer key.

    Bounded by `limit` and the shared usage ledger. Without `execute` it only previews (no key read, no network, no
    writes). Own pairs are never asked here (they are the shipping question: `evaluate --acquire`), and nothing is
    written to the shared judgment ledger. Answers accumulate in `<key_dir>/judgment-answers.json`, so a later run
    continues where the last one stopped."""
    jr, jl = judge._jr(), judge._jl()
    policy = jl.load_policy()
    swaps = {i['candidate']['judgment_hash']: i['candidate'] for i in items if i['kind'] == 'swap'}
    if not swaps:
        raise ValueError('swap の組が 0 件（測定用の台帳に記録するものが無い）')
    existing = read_measurement(key_dir)
    done = {h for h in _scores(existing) if h in swaps}
    pending = [c for h, c in swaps.items() if h not in done]
    kwargs = {} if _usage_path is None else {'_usage_path': _usage_path}
    answers = Path(key_dir).absolute() / ANSWERS
    with tempfile.TemporaryDirectory(prefix='infoarch-measure-') as tmp:
        source = Path(tmp).resolve() / 'candidates.json'  # the shared reader refuses paths through a symlink (/var)
        source.write_text(json.dumps({'judgment_candidates': pending or list(swaps.values())}, ensure_ascii=False), encoding='utf-8')
        result = jr.run(source, answers, limit, dry_run=not execute, _transport=_transport, **kwargs)
    result['swap_pairs'] = len(swaps)
    if execute and answers.exists():
        got = {a['candidate']['judgment_hash']: {'candidate': a['candidate'], 'score': jl.probability(a['score'])}
               for a in judge.read_answers(answers, policy)['rows'] if a['candidate']['judgment_hash'] in swaps}
        rows = {**{r['candidate']['judgment_hash']: r for r in (existing or {}).get('rows', []) if r['candidate']['judgment_hash'] in swaps},
                **got}
        ledger = {'schema_version': jl.CURRENT_SCHEME, 'policy': policy, 'holdout': {'source': 'unmeasured'},
                  'rows': sorted(rows.values(), key=lambda r: r['candidate']['id'])}
        jl.validate_ledger(ledger, policy)
        target = _ledger_path(key_dir)
        tmp_file = target.with_name(target.name + '.tmp')
        if tmp_file.is_symlink() or tmp_file.exists():
            raise ValueError(f'一時ファイルが既にある: {tmp_file}')
        tmp_file.write_text(json.dumps(ledger, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
        os.replace(tmp_file, target)
        result['ledger'] = {'file': MEASUREMENT_LEDGER, 'rows': len(rows)}
    return result


def candidates_file(items):
    """The shape judgment_run.py reads (`judgment_candidates`): the swap pairs only. Own pairs are the shipping
    question (`judge --include-candidates`), so none is offered here to be recorded into the wrong ledger."""
    return {'judgment_candidates': [i['candidate'] for i in items if i['kind'] == 'swap']}


def key_file(items):
    return {'answer_key': key_hash(items),
            'items': [{k: i[k] for k in ('id', 'kind', 'text', 'value')} | {'judgment_hash': i['candidate']['judgment_hash']}
                      for i in items]}


def load_items(m, approved, policy, jl, key_path):
    """Rebuild the items and refuse a key file that no longer matches the current outlines and values."""
    items = build(m, approved, policy, jl)
    saved = json.loads(open(key_path, encoding='utf-8').read())
    if saved.get('answer_key') != key_hash(items):
        raise ValueError('the outlines or values changed since this answer key was built (rebuild it and label again)')
    return items

