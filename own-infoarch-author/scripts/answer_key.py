"""The answer key for the meaning judgment (SPEC-information-architecture-standardization v6, requirement 12).

Every piece of copy is paired with the value its message points to (`own`, expected to convey it) and re-paired with
the other approved values of the same positioning (`swap`, candidates that should not). A positioning with a single
value is re-paired with another positioning's value. The machine only proposes: the owner's labels (conveys /
does not convey / unknown) decide. Agreement is counted against the shared judgment ledger (a score at or above theta
means Jev says "does not convey"). When the owner's labels hold no conveying pair or no non-conveying pair, there is
nothing to measure against, and the result is "unmeasurable", never a number. Nothing here calls a model.
"""
from __future__ import annotations

import html
import json

import judge
import outline as ol
import positioning as pos


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
                    out.append({'id': ident, 'kind': kind, 'text': text, 'value': value,
                                'candidate': judge.value_fit(jl, policy, ident, text, oid, c['message'], value)})
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
textarea{{width:100%;box-sizing:border-box;font:12px ui-monospace,monospace;color:var(--fg);background:var(--bg);border:1px solid var(--line)}}
</style></head><body>
<h1>正解表のラベル付け</h1>
<p>文案を読んだ人に、その提供価値が伝わるかを選んでください。分からないものは「分からない」のままにします（数えません）。
選び終わったら「ラベルを保存」で labels.json を保存するか、「ラベルをコピー」で中身をコピーして、AI に渡してください。</p>
<p class="id">answer key: {key}（{len(items)} 組）</p>
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


def measure(items, labels, ledger, theta):
    """-> the agreement between the owner's labels and Jev's answers in the ledger, or 'unmeasurable'."""
    if not isinstance(labels, dict) or labels.get('answer_key') != key_hash(items):
        raise ValueError('labels were made for another answer key (rebuild the page and label again)')
    by_id = {i['id']: i for i in items}
    given = labels.get('labels') or {}
    unknown_ids = sorted(set(given) - set(by_id))
    if unknown_ids:
        raise ValueError(f'labels name unknown pairs: {unknown_ids[:5]}')
    if any(v not in (True, False, None) for v in given.values()):
        raise ValueError('a label is true (conveys), false (does not convey) or null (unknown)')
    scores = {}
    for row in (ledger or {}).get('rows') or []:
        cand = row.get('candidate') or {}
        scores[cand.get('judgment_hash')] = row.get('score')
    confirmed = {k: v for k, v in given.items() if v is not None}
    covered = {k: v for k, v in confirmed.items() if by_id[k]['candidate']['judgment_hash'] in scores}
    uncovered = sorted(set(confirmed) - set(covered))
    counts = {'labeled': len(confirmed), 'conveys': sum(confirmed.values()),
              'does_not_convey': sum(not v for v in confirmed.values()),
              'unknown': sum(v is None for v in given.values()), 'uncovered': len(uncovered)}
    base = {'answer_key': key_hash(items), 'counts': counts, 'uncovered': uncovered, 'theta': theta}
    conveys = [k for k, v in covered.items() if v]
    not_conveys = [k for k, v in covered.items() if not v]
    if not conveys or not not_conveys:
        side = 'conveys' if not conveys else 'does not convey'
        return {**base, 'status': 'unmeasurable',
                'reason': f'judged pairs the owner labeled "{side}" are 0: agreement is not measured'}
    flagged = {k: scores[by_id[k]['candidate']['judgment_hash']] >= theta for k in covered}
    agree = [k for k in covered if flagged[k] == (not covered[k])]
    return {**base, 'status': 'measured', 'agreement': len(agree) / len(covered),
            'misses': sorted(k for k in not_conveys if not flagged[k]),
            'false_detections': sorted(k for k in conveys if flagged[k])}


def candidates_file(items):
    """The shape judgment_run.py reads (`judgment_candidates`)."""
    return {'judgment_candidates': [i['candidate'] for i in items]}


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

