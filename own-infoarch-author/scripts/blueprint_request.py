"""The approval request for the product items of a blueprint (SPEC-deliverable-production-layers requirement 12).

The page, its ○ and × saved to the Artifact db and the reading back are the positioning layer's
(own-positioning-author/scripts/approval_request.py); what differs is the card: one product item, and the hash that
binds it is that item's own (`blueprint.item_hash`), so a card for a story, a principle, a constraint, an object or an
entry changes only when that item does.
"""
from __future__ import annotations

import approval_request as ar
import blueprint

KIND_NAMES = {'sto': 'ストーリー', 'pri': '原則', 'con': '制約', 'ojc': 'オブジェクト', 'ntr': '入口'}
INTRO = ('プロダクト UI の土台になる項目を、1 件ずつ見せています（ストーリー・原則・制約・オブジェクトと主役・入口）。1 件ずつ読んで、'
         '○（承認する）か ×（差し戻す）を選んでください。選ぶとこのページに保存され、AI が読み戻します。画面ごとの優先度や配置は承認の対象ではなく、'
         'ここで承認した項目から AI が作ります。承認の台帳に書くのは、あなたの承認の言葉を受けた AI です。')
HASH_NOTE = '各カードの英数字はその項目のハッシュの先頭 12 桁です。項目の中身が変わるとハッシュも変わり、保存した判断は古いものとして数えません。ほかの項目の承認には影響しません。'
FIELD_NAMES = {'actor': 'だれが', 'text': '内容', 'kind': '種類', 'serves': '応えるストーリー', 'values': '引く提供価値', 'name': '名前',
               'lead': '主役', 'screen': '入る画面'}
KIND_VALUES = {'arrangement': '並べ方の原則', 'behavior': 'ふるまいの原則'}


def build(bp, approved=(), must=()):
    """-> [item] sorted by item id, one per proposed product item, in the shape the page renders."""
    approved, must = set(approved), set(must)
    out = []
    for key, digest in sorted(blueprint.targets(bp).items(), key=lambda kv: kv[0]):
        iid = key[len(blueprint.ITEM_PREFIX):]
        item = blueprint.items(bp)[iid]
        fields = []
        for fld in ('actor', 'kind', 'name', 'lead', 'text', 'serves', 'values', 'screen'):
            v = item.get(fld)
            if v in (None, [], ''):
                continue
            v = KIND_VALUES.get(v, v) if fld == 'kind' else ('主役' if fld == 'lead' and v is True else v)
            if fld == 'lead' and v is False:
                continue
            fields.append((FIELD_NAMES[fld], '、'.join(v) if isinstance(v, list) else str(v)))
        title = item.get('text') or item.get('name') or iid
        out.append({'id': iid, 'key': key, 'value_hash': digest, 'hash12': digest[:12], 'kind': KIND_NAMES[iid.split('-')[0]],
                    'text': title, 'fields': fields, 'roles': [], 'problems': [], 'approved': iid in approved,
                    'required': key in must, 'hash_note': 'この項目 1 件のハッシュ'})
    return out


def page(items, title='プロダクト UI の承認依頼'):
    return ar.page(items, title, intro=INTRO, hash_note=HASH_NOTE)
