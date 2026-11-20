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
    order = {p: n for n, p in enumerate(blueprint.ITEM_BLOCKS.values())}  # stories, principles, constraints, objects, entries
    named = lambda ident: ((blueprint.items(bp).get(ident) or {}).get('text') or (blueprint.items(bp).get(ident) or {}).get('name')
                           or (bp.screens[ident][1]['title'] if ident in bp.screens else '') or '')
    for key, digest in sorted(blueprint.targets(bp).items(), key=lambda kv: (order[kv[0].split(':')[2].split('-')[0]], kv[0])):
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
            shown = [f'{x}（{named(x)}）' if named(x) and x != named(x) and fld in ('serves', 'screen') else x for x in (v if isinstance(v, list) else [v])] if fld in ('serves', 'screen') else None
            fields.append((FIELD_NAMES[fld], '、'.join(shown) if shown else ('、'.join(v) if isinstance(v, list) else str(v))))
        title = item.get('text') or item.get('name') or (f'入口：{named(item["screen"]) or item["screen"]}' if iid.startswith('ntr-') else iid)
        out.append({'id': iid, 'key': key, 'value_hash': digest, 'hash12': digest[:12], 'kind': KIND_NAMES[iid.split('-')[0]],
                    'text': title, 'fields': fields, 'roles': [], 'problems': [], 'approved': iid in approved,
                    'required': key in must, 'hash_note': 'この項目 1 件のハッシュ'})
    return out


LAYER_TITLES = {0: '提供価値（positioning）', 1: 'ストーリー', 2: '原則・制約', 3: 'オブジェクト', 4: '入口', 5: '画面（AI が作る。承認の対象ではない）'}


def overview_graph(bp, values):
    """The whole product as one graph: which story each principle and entry answers, which positioning value each principle
    and the lead object stand on, which screen an entry opens, which stories a screen serves. `values`: {value id: text}."""
    p = bp.product
    nodes, edges = {}, []

    def node(iid, layer, text, role):
        nodes.setdefault(iid, {'id': iid, 'layer': layer, 'role': role, 'text': text, 'kind': ''})

    def edge(a, b, label, kind='cites'):
        edges.append({'key': f'{kind}:{a}:{b}', 'from': a, 'to': b, 'kind': kind, 'degree': None, 'role': 'premise', 'label': label})

    for iid, s in p['stories'].items():
        if s['state'] == 'proposed':
            node(iid, 1, f'{s["actor"]}：{s["text"]}', 'premise')
    for iid, x in p['principles'].items():
        if x['state'] == 'proposed':
            node(iid, 2, x['text'], 'basis')
            for st in x.get('serves', []):
                edge(iid, st, '応える')
            for v in x.get('values', []):
                node(v, 0, values.get(v, v), 'claim')
                edge(iid, v, '引く')
    for iid, c in p['constraints'].items():
        if c['state'] == 'proposed':
            node(iid, 2, c['text'], 'basis')
    for iid, o in p['objects'].items():
        if o['state'] == 'proposed':
            node(iid, 3, ('主役：' if o.get('lead') else '') + o['name'], 'claim' if o.get('lead') else 'premise')
            for v in o.get('values', []):
                node(v, 0, values.get(v, v), 'claim')
                edge(iid, v, '引く')
    for iid, e in p['entries'].items():
        if e['state'] == 'proposed':
            node(iid, 4, f'入口：{e["screen"]}', 'premise')
            for st in e.get('serves', []):
                edge(iid, st, '応える')
            node(e['screen'], 5, bp.screens[e['screen']][1]['title'] if e['screen'] in bp.screens else e['screen'], 'evidence')
            edge(iid, e['screen'], '入る')
    for sid, (_, s) in bp.screens.items():
        node(sid, 5, s['title'], 'evidence')
        for st in s['serves']:
            if st in nodes:
                edge(sid, st, '応える')
        for con in s.get('constraints', []):
            if con in nodes:
                edge(sid, con, '守る')
    return {'nodes': list(nodes.values()), 'edges': edges, 'layer_titles': LAYER_TITLES}


def overview(bp, values):
    graph = overview_graph(bp, values)
    return ('<section class="card"><h3>プロダクト全体の関係図</h3>'
            '<p class="meta">箱は項目、線は「どれがどれを引いているか」です。各項目の中身は、この下のカードで 1 件ずつ読みます。</p>'
            + ar._diagram(graph) + '</section>')


def page(items, title='プロダクト UI の承認依頼', overview_html=''):
    return ar.page(items, title, intro=INTRO, hash_note=HASH_NOTE, overview=overview_html)
