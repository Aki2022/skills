"""The approval request for the layout overrides: one card per row, the same page and the same ledger as the product items.

SPEC-deliverable-production-layers requirement 4 (a placement against the allocation rules is the owner's, row by row). The page, its
○ and × saved to the Artifact db and the reading back are the positioning layer's (own-positioning-author/scripts/approval_request.py,
the page the product items use); what differs is the card: one override row, and the hash that binds it is the row's own
(`layout_overrides.Row.hash`: the row and the declarations of the two regions it points at), so a card changes only when its row, or a
region it points at, does.

Recording is not here: the owner's words go through `design-author approve --prepare --target <key>` and `approve --request ...
--owner-words ... --approved-on ...`, which writes the ledger (the only writer).
"""
from __future__ import annotations

import sys
from pathlib import Path

import layer_paths
import layout_overrides

_POSITIONING = Path(__file__).resolve().parents[2] / 'own-positioning-author' / 'scripts'
if not (_POSITIONING / 'approval_request.py').is_file():
    raise RuntimeError('own-positioning-author/scripts/approval_request.py が無い（承認依頼のページは positioning layer の道具を使う）')
if str(_POSITIONING) not in sys.path:
    sys.path.append(str(_POSITIONING))
import approval_request as ar  # noqa: E402

INTRO = ('画面の領域の置き方の規則（優先度の高い領域を上に置く）に反する配置を、1 行ずつ見せています。1 行が許すのは、書いてある画面・領域の組・幅だけです。'
         '1 行ずつ読んで、○（承認する）か ×（差し戻す）を選んでください。選ぶとこのページに保存され、AI が読み戻します。'
         '承認の台帳に書くのは、あなたの承認の言葉を受けた AI です。理由が書いてあるだけでは、どの行も通りません。')
HASH_NOTE = ('各カードの英数字はその行のハッシュの先頭 12 桁です。行の中身、または行が指す領域の宣言が変わるとハッシュも変わり、保存した判断は古いものとして数えません。'
             'ほかの行の承認には影響しません。')


def _region_label(regions, rid):
    reg = regions.get(rid) or {}
    return f'{rid}（優先度 {reg.get("priority", "?")}・{reg.get("text") or reg.get("shape") or "名前なし"}）'


def build(root, evaluation=None):
    """-> [item], one per row that can be approved (its structure is sound and its screen and regions exist), in file order."""
    state = evaluation or layout_overrides.evaluate_repo(root)
    items = []
    for r in state.rows:
        if r.status not in ('unapproved', 'approved'):
            continue
        d = r.data
        regions = layout_overrides._blueprint_regions(root, d['screen']) or {}
        items.append({
            'id': r.id, 'key': r.key, 'value_hash': r.hash, 'hash12': r.hash[:12], 'kind': '配置の例外',
            'text': f'{d["screen"]} で、{d["above"]} を {d["below"]} より上に置く（{"・".join(d["viewports"])}）',
            'fields': [('画面', d['screen']), ('上に置く領域', _region_label(regions, d['above'])),
                       ('規則では上の領域', _region_label(regions, d['below'])), ('許す幅', '、'.join(d['viewports'])), ('理由', d['reason'])],
            'roles': [], 'problems': [], 'approved': r.status == 'approved', 'required': False,
            'hash_note': 'この行 1 つ（と、指す 2 つの領域の宣言）のハッシュ'})
    return items


def page(items, title='配置の例外の承認依頼'):
    return ar.page(items, title, intro=INTRO, hash_note=HASH_NOTE)
