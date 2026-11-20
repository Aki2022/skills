"""The approval request the owner sees: one page, built from the same data the approval hash binds.

SPEC-deliverable-production-layers requirement 12: the display of an approval request is made by the skill that owns
the approval, by machine, from the data it approves (the hash of each value's whole path), and put where the owner can
open it. A hand-made page can drift from what the hash covers without anyone noticing; this one cannot, because
every card is rendered from the path it carries the hash of, and a test recomputes the hash from the page's own data.

The page is a private Artifact that declares the `db` capability. Each card has a ○ (approve) and a × (send back)
button and a note; the choice is saved as `decisions/<value id>` = {key, value_hash, decision, note, decided_at}, so
the owner's choice survives a reload and the AI reads it back (`ArtifactData` list of `decisions`). A decision whose
`value_hash` is not the card's current hash is stale (the path changed after it was made) and counts for nothing.
`decisions_from_rows` is that reading. The page never writes the approval ledger: only the owner's words, recorded
through own-design-author's approve, do (ADR-20261118-ai-records-owner-approvals).

A page without `db` (a preview that cannot run scripts, a signed-out view) says so and offers the choices as JSON to
copy to the AI. Nothing here calls a model.
"""
from __future__ import annotations

import html
import json

import positioning as pos
import positioning_model as pm

COLLECTION = 'decisions'
DECISIONS = ('ok', 'ng')


def build(m, approved=None, must=()):
    """-> [item] sorted by value id, one per value path. An item carries what the hash covers (members grouped by
    role, in the order requirement 3 reads them), the hash, whether the owner's current approval covers it, and the
    disqualifications that stop the path.

    `approved` is the set of value ids whose path has the owner's current approval (`positioning.approved_keys`); `must` the keys a shipped outline needs (from
    `infoarch-author targets`), shown as such. Both are optional: this layer alone does not know the outlines."""
    approved, must = set(approved or ()), set(must)
    hashes = pos.approval_targets(m)
    ps = pos.paths(m)
    atoms, edges = pos.atoms(m, proposed_only=False), pos.edges(m, proposed_only=False)
    items = []
    for key in sorted(hashes):
        vid = key[len(pm.APPROVAL_PREFIX):]
        p = ps.get(vid)
        if p is None:
            continue
        roles = []
        for role in pos.ROLES:
            rows = []
            for k in (k for k in p.members if p.roles[k] == role):
                if k in atoms:
                    a = atoms[k]
                    rows.append({'key': k, 'kind': pm.REVIEW_NAMES.get(pos._prefix(k), k),
                                 'text': a.get('text') or a.get('qualifier') or '',
                                 'detail': {label: ('、'.join(v) if isinstance(v, list) else str(v))
                                            for fld, label in pm.REVIEW_FIELDS.items()
                                            if label and fld not in ('provenance', 'segment') and (v := a.get(fld))}})
                else:
                    e = edges[k]
                    degree = f'（{e["degree"]}）' if 'degree' in e else ''
                    rows.append({'key': k, 'kind': '矢印', 'text': f'{e["from"]} → {e["to"]}：{pm.EDGE_NAMES.get(e["kind"], e["kind"])}{degree}',
                                 'detail': {}})
            if rows:
                roles.append({'role': role, 'label': pm.ROLE_NAMES[role], 'rows': rows})
        value = atoms.get(vid) or {}
        items.append({'id': vid, 'key': key, 'value_hash': hashes[key], 'hash12': hashes[key][:12],
                      'segment': p.segment, 'kind': p.kind, 'text': value.get('text', ''), 'roles': roles,
                      'comparison': list(p.comparison), 'targets': list(p.targets),
                      'problems': [{'code': c, 'message': msg} for c, msg in p.problems],
                      'approved': vid in approved, 'required': key in must})
    return items


def request_rows(items):
    """The approval-request rows (key and hash) of the cards: the shape own-design-author's `approve --prepare` binds."""
    return [{'key': i['key'], 'value_hash': i['value_hash']} for i in items]


def load_rows(path):
    """The owner's saved answers as rows [{id, data}], from any of the shapes they reach us in: a directory the
    `ArtifactData` read wrote with `out_dir` (one `<doc id>.json` per document, the body alone), a JSON list of rows (what
    the page offers to copy), one document body, or a {doc id: body} mapping."""
    import json
    from pathlib import Path
    path = Path(path)
    if path.is_dir():
        return [{'id': f.stem, 'data': json.loads(f.read_text(encoding='utf-8'))} for f in sorted(path.rglob('*.json'))]
    data = json.loads(path.read_text(encoding='utf-8'))
    if isinstance(data, list):
        return data
    if isinstance(data, dict) and 'key' in data:
        return [{'id': path.stem, 'data': data}]
    if isinstance(data, dict):
        return [{'id': k, 'data': v} for k, v in sorted(data.items())]
    raise ValueError('the saved answers are not a list, a document or a mapping of documents')


def decisions_from_rows(items, rows):
    """The owner's saved choices read back against the current cards.

    `rows` is what the page saved (a list of `decisions` documents, or the JSON the page offers to copy). ->
    {'ok': [key], 'ng': [key], 'stale': [key], 'undecided': [key], 'unknown': [doc id]}. A decision whose hash is not
    the card's current hash is stale and counts as neither; an unknown value id is reported, not guessed at."""
    by_key = {i['key']: i for i in items}
    out = {'ok': [], 'ng': [], 'stale': [], 'undecided': [], 'unknown': []}
    seen = set()
    for row in rows or []:
        data = row.get('data', row) if isinstance(row, dict) else {}
        key = data.get('key')
        if key not in by_key:
            out['unknown'].append(str(row.get('id', key)) if isinstance(row, dict) else str(row))
            continue
        seen.add(key)
        if data.get('value_hash') != by_key[key]['value_hash']:
            out['stale'].append(key)
        elif data.get('decision') in DECISIONS:
            out[data['decision']].append(key)
        else:
            out['unknown'].append(str(row.get('id', key)))
    out['undecided'] = sorted(set(by_key) - seen)
    for k in ('ok', 'ng', 'stale', 'unknown'):
        out[k] = sorted(out[k])
    return out


def _card(i):
    badges = []
    if i['approved']:
        badges.append('<span class="tag done">承認済み</span>')
    if i['required']:
        badges.append('<span class="tag must">出荷に必要</span>')
    sections = []
    for r in i['roles']:
        rows = ''.join(
            f'<li><span class="kind">{html.escape(x["kind"])}</span> <span class="text">{html.escape(x["text"])}</span>'
            + (f'<span class="detail">{html.escape(" ／ ".join(f"{k}：{v}" for k, v in x["detail"].items()))}</span>' if x['detail'] else '')
            + f' <code>{html.escape(x["key"])}</code></li>' for x in r['rows'])
        sections.append(f'<h4>{html.escape(r["label"])}（{len(r["rows"])}）</h4><ul>{rows}</ul>')
    problems = ''.join(f'<li><code>{html.escape(p["code"])}</code> {html.escape(p["message"])}</li>' for p in i['problems'])
    todo = (f'<div class="todo"><strong>決めてほしいこと（承認だけでは出荷できない欠け）</strong><ul>{problems}</ul></div>'
            if problems else '')
    meta = (f'比べる集合：{html.escape("、".join(i["comparison"]) or "なし")}　'
            f'満たす先：{html.escape("、".join(i["targets"]) or "なし")}')
    return f'''<article class="card" data-key="{html.escape(i["key"])}" data-id="{html.escape(i["id"])}">
<header><h3>{html.escape(i["text"] or i["id"])}</h3>
<p class="meta"><code>{html.escape(i["key"])}</code> <span class="hash" title="この経路全体のハッシュ">{i["hash12"]}</span> {html.escape(i["segment"])}・{html.escape(i["kind"])} {"".join(badges)}</p></header>
<p class="meta">{meta}</p>
{todo}{"".join(sections)}
<div class="choice" role="group" aria-label="{html.escape(i["id"])} の判断">
<button type="button" data-d="ok" aria-pressed="false">○ 承認する</button>
<button type="button" data-d="ng" aria-pressed="false">× 差し戻す</button>
<span class="state" role="status"></span></div>
<label class="note">理由・メモ（任意）<textarea rows="2"></textarea></label>
</article>'''


def page(items, title='承認依頼'):
    """A self-contained page body (the Artifact tool wraps it in its own document: title, meta and style first, then the content; no
    doctype, html, head or body): a card per value path, ○ and × saved to the `db` collection `decisions`."""
    data = json.dumps({'collection': COLLECTION, 'cards': {i['id']: {'key': i['key'], 'value_hash': i['value_hash']} for i in items}},
                      ensure_ascii=False).replace('</', '<\\/')
    cards = '\n'.join(_card(i) for i in items) or '<p>承認の対象がありません。</p>'
    return f'''<title>{html.escape(title)}</title>
<meta name="claude:viewport-fit-sides" content="cover">
<style>
:root{{--fg:#1b1b1b;--bg:#fafaf8;--card:#fff;--line:#d6d6d0;--muted:#5f5f5a;--ok:#1a6b3a;--ng:#a3262a;--warn:#8a5a00}}
@media (prefers-color-scheme:dark){{:root:not([data-theme="light"]){{--fg:#ececea;--bg:#151515;--card:#1e1e1e;--line:#3c3c3a;--muted:#a3a39e;--ok:#5ac17f;--ng:#ff7b7f;--warn:#e8b24a;color-scheme:dark}}}}
:root[data-theme="dark"]{{--fg:#ececea;--bg:#151515;--card:#1e1e1e;--line:#3c3c3a;--muted:#a3a39e;--ok:#5ac17f;--ng:#ff7b7f;--warn:#e8b24a;color-scheme:dark}}
body{{margin:0;padding-block:16px;padding-left:max(16px,env(safe-area-inset-left,0px));padding-right:max(16px,env(safe-area-inset-right,0px));background:var(--bg);color:var(--fg);font:15px/1.65 system-ui,"Hiragino Sans",sans-serif}}
main{{max-width:760px;margin:0 auto}}
h1{{font-size:20px;margin:0 0 4px}} h3{{font-size:17px;margin:0}} h4{{font-size:13px;margin:14px 0 4px;color:var(--muted)}}
.card{{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:14px 16px;margin:16px 0}}
.meta{{color:var(--muted);font-size:13px;margin:4px 0}} code,.hash{{font:12px ui-monospace,monospace}}
.hash{{background:var(--line);padding:1px 6px;border-radius:4px}}
.tag{{font-size:12px;border:1px solid currentColor;border-radius:999px;padding:0 8px;margin-left:4px}} .tag.done{{color:var(--ok)}} .tag.must{{color:var(--warn)}}
ul{{margin:0;padding-left:20px}} li{{margin:2px 0}} .kind{{color:var(--muted);font-size:12px}} .detail{{display:block;color:var(--muted);font-size:12px}}
.todo{{border-left:4px solid var(--warn);padding:4px 12px;margin:10px 0;background:color-mix(in srgb,var(--warn) 10%,transparent)}}
.choice{{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-top:14px}}
button{{font:inherit;padding:8px 18px;border-radius:8px;border:1.5px solid var(--line);background:var(--card);color:var(--fg);cursor:pointer;min-height:44px}}
button[data-d=ok][aria-pressed=true]{{border-color:var(--ok);color:var(--ok);font-weight:700}} button[data-d=ng][aria-pressed=true]{{border-color:var(--ng);color:var(--ng);font-weight:700}}
.state{{font-size:13px;color:var(--muted)}} .note{{display:block;margin-top:8px;font-size:13px;color:var(--muted)}}
textarea{{width:100%;box-sizing:border-box;font:inherit;color:var(--fg);background:var(--bg);border:1px solid var(--line);border-radius:6px}}
.warn{{border:2px solid var(--ng);padding:8px 12px;border-radius:8px}}
</style>
<main>
<h1>{html.escape(title)}</h1>
<p>提供価値ごとに、それを支える経路の全体（主張→差別化の根拠→前提→証拠）を見せています。内容を読んで、○（承認する）か ×（差し戻す）を選んでください。選ぶとこのページに保存され、AI が読み戻します。承認の台帳に書くのは、あなたの承認の言葉を受けた AI です。</p>
<p class="meta">{len(items)} 件。各カードの英数字はその経路のハッシュの先頭 12 桁で、経路の中身が変わるとハッシュも変わり、保存した判断は古いものとして数えません。</p>
<p id="nodb" class="warn" hidden>この画面では選んだ判断を保存できません（スクリプトが動かない、またはサインインしていない画面です）。ブラウザで開き直してください。保存できないときは、下の欄の中身をコピーして AI に渡してください。</p>
{cards}
<h2 style="font-size:15px">保存できないときの控え</h2>
<textarea id="copy" readonly rows="6" aria-label="判断の控え（JSON）"></textarea>
<script id="data" type="application/json">{data}</script>
<script>
(function () {{
  var meta = JSON.parse(document.getElementById('data').textContent), cards = meta.cards, state = {{}}, col = null;
  var box = document.getElementById('copy');
  function rows() {{
    return Object.keys(state).map(function (id) {{ return {{id: id, data: state[id]}}; }});
  }}
  function show() {{ box.value = JSON.stringify(rows(), null, 2); }}
  function paint(card) {{
    var id = card.getAttribute('data-id'), s = state[id], st = card.querySelector('.state');
    card.querySelectorAll('button').forEach(function (b) {{ b.setAttribute('aria-pressed', String(!!s && s.decision === b.getAttribute('data-d') && s.value_hash === cards[id].value_hash)); }});
    if (!s) st.textContent = '未回答';
    else if (s.value_hash !== cards[id].value_hash) st.textContent = '古い判断（経路の中身が変わりました。選び直してください）';
    else st.textContent = s.decision === 'ok' ? '○ 承認する（保存済み）' : '× 差し戻す（保存済み）';
    if (s && s.note !== undefined && card.querySelector('textarea') !== document.activeElement) card.querySelector('textarea').value = s.note;
  }}
  function paintAll() {{ document.querySelectorAll('.card').forEach(paint); show(); }}
  function save(card, decision) {{
    var id = card.getAttribute('data-id');
    var row = {{key: cards[id].key, value_hash: cards[id].value_hash, decision: decision,
               note: card.querySelector('textarea').value, decided_at: new Date().toISOString()}};
    state[id] = row; paintAll();
    if (col) col.doc(id).set(row).catch(function () {{ card.querySelector('.state').textContent = '保存できませんでした（下の控えをコピーしてください）'; }});
  }}
  document.querySelectorAll('.card').forEach(function (card) {{
    card.querySelectorAll('button').forEach(function (b) {{ b.addEventListener('click', function () {{ save(card, b.getAttribute('data-d')); }}); }});
    card.querySelector('textarea').addEventListener('change', function () {{
      var id = card.getAttribute('data-id'); if (state[id]) save(card, state[id].decision);
    }});
  }});
  paintAll();
  var timer = setTimeout(function () {{ document.getElementById('nodb').hidden = false; }}, 10000);
  (window.claude ? window.claude.use('db') : Promise.resolve(null)).then(function (db) {{
    clearTimeout(timer);
    if (!db) {{ document.getElementById('nodb').hidden = false; return; }}
    col = db.collection(meta.collection);
    col.onSnapshot(function (snap) {{
      snap.docs.forEach(function (d) {{ if (d.exists && cards[d.id]) state[d.id] = d.data(); }});
      paintAll();
    }}, function () {{ document.getElementById('nodb').hidden = false; }});
  }}).catch(function () {{ clearTimeout(timer); document.getElementById('nodb').hidden = false; }});
}})();
</script></main>
'''
