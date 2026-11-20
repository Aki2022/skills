"""Fixtures of the product UI's infoarch (blueprint): a product and two screens standing on the positioning fixture."""
import copy
import sys
from pathlib import Path

import yaml

SCRIPTS = Path(__file__).resolve().parents[1]
SKILL = SCRIPTS.parent
for _p in (SCRIPTS, SKILL.parent / 'own-positioning-author/scripts', SKILL.parent / 'own-positioning-author/scripts/tests',
           SKILL.parent / 'own-doc-update/scripts'):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))
from ia_fixtures import SSOT, confirm, dump  # noqa: E402


def _i(**kw):
    return {'state': 'proposed', 'provenance': 'owner', **kw}


PRODUCT = {
    'schema_version': 1,
    'ledger': {'next': {'sto': 3, 'pri': 3, 'con': 3, 'ojc': 4, 'ntr': 2}, 'tombstones': {}},
    'stories': {
        'sto-0001': _i(actor='情報部門の責任者', text='組織の運用の状態を、ひと目で経営に報告したい'),
        'sto-0002': _i(actor='情報部門の担当者', text='一人ひとりの運用の中身を確かめたい')},
    'principles': {
        'pri-0001': _i(kind='arrangement', text='いちばん知りたい状態を、画面のいちばん上に置く', serves=['sto-0001'], values=['val-0001']),
        'pri-0002': _i(kind='behavior', text='今どの状態にあるかを、必ず画面に示す', serves=['sto-0001', 'sto-0002'], values=['val-0003'])},
    'constraints': {
        'con-0001': _i(text='個人を他の個人と並べて順位づけしない'),
        'con-0002': _i(text='承認されていない値を表示しない')},
    'objects': {
        'ojc-0001': _i(name='組織の運用の状態', lead=True, values=['val-0001']),
        'ojc-0002': _i(name='部署'),
        'ojc-0003': _i(name='担当者')},
    'entries': {'ntr-0001': _i(serves=['sto-0001'], screen='scr-overview')},
}

OVERVIEW = {
    'schema_version': 1, 'id': 'scr-overview', 'ship': True, 'title': '概要', 'serves': ['sto-0001'],
    'constraints': ['con-0001'],
    'groups': {'grp-0001': {'text': '全体の状態', 'principle': 'pri-0001'}, 'grp-0002': {'text': '内訳', 'principle': 'pri-0001'}},
    'regions': {
        'rgn-0001': {'group': 'grp-0001', 'priority': 1, 'shape': 'value', 'objects': ['ojc-0001'], 'text': '組織の運用の状態の総合値'},
        'rgn-0002': {'group': 'grp-0002', 'priority': 2, 'shape': 'comparison', 'objects': ['ojc-0002'], 'text': '部署ごとの比較'},
        'rgn-0003': {'group': 'grp-0002', 'priority': 3, 'shape': 'list', 'objects': ['ojc-0003'], 'text': '担当者の一覧'}},
    'actions': {'act-0001': {'text': '担当者を見る', 'region': 'rgn-0003', 'primary': True, 'to': 'scr-person'}},
    'states': {'loading': {'text': '読み込み中'}, 'empty': {'text': 'まだ回答がありません', 'region': 'rgn-0001'},
               'error': {'text': '読み込めませんでした'}},
}

PERSON = {
    'schema_version': 1, 'id': 'scr-person', 'ship': True, 'title': '担当者の詳細', 'serves': ['sto-0002'],
    'groups': {'grp-0001': {'text': '本人の状態', 'principle': 'pri-0001'}},
    'regions': {
        'rgn-0001': {'group': 'grp-0001', 'priority': 1, 'shape': 'value', 'objects': ['ojc-0003'], 'text': '本人の総合値'},
        'rgn-0002': {'group': 'grp-0001', 'priority': 2, 'shape': 'trend', 'objects': ['ojc-0003'], 'text': '本人の推移'}},
    'actions': {'act-0001': {'text': '概要に戻る', 'region': 'rgn-0001', 'primary': True, 'to': 'scr-overview'}},
    'states': {'loading': {'text': '読み込み中'}, 'error': {'text': '読み込めませんでした'}},
}

SCREENS = (OVERVIEW, PERSON)
BP_DIR = 'design/info_arch/blueprint'


def item_hashes(product):
    import blueprint
    return {blueprint.ITEM_PREFIX + iid: blueprint.item_hash(iid, item)
            for block in blueprint.ITEM_BLOCKS for iid, item in product[block].items() if item.get('state') == 'proposed'}


def confirm_items(root, product, skip=()):
    """Record the owner's approval of every proposed product item (as the owner would with design-author approve)."""
    rows = [{'key': key, 'value_hash': digest, 'approver': 'fixture-owner', 'date': '2026-11-20'}
            for key, digest in sorted(item_hashes(product).items()) if key[len('ia:item:'):] not in skip]
    path = root / 'design/approvals.yaml'
    kept = yaml.safe_load(path.read_text(encoding='utf-8')) if path.is_file() else []
    kept = [r for r in (kept or []) if not str(r.get('key', '')).startswith('ia:item:')]
    dump(path, kept + rows)


def write_blueprint(root, product=PRODUCT, screens=SCREENS, ssot=SSOT, confirmed=True, skip=()):
    """A repository with a positioning and a blueprint, and no outline (a product UI only)."""
    dump(root / 'design/info_arch/info_arch.yaml', ssot)
    confirm(root, ssot)
    if product is not None:
        dump(root / BP_DIR / 'product.yaml', product)
        if confirmed:
            confirm_items(root, product, skip=skip)
    for s in screens:
        name = s.get('_file', s['id'])  # `_file` keeps the file name when a case changes the id
        dump(root / BP_DIR / 'screens' / f'{name}.yaml', {k: v for k, v in s.items() if k != '_file'})


def apply_ops(ops, product=PRODUCT, screens=SCREENS):
    """Apply a red case's declarative operations to copies of the fixtures. -> (product, screens, ids left unapproved).

    op: {file: product | <screen id>, path: [..], set: value} sets a value; {file, path, del: true} removes one;
    {drop_product: true} removes product.yaml; {add_screen: {...}} adds a screen; {drop_screen: id} removes one; {unapprove: <item id>} leaves an item unapproved."""
    product = copy.deepcopy(product)
    screens = {s['id']: copy.deepcopy(s) for s in screens}
    unapproved = []
    for op in ops:
        if 'unapprove' in op:
            unapproved.append(op['unapprove'])
            continue
        if op.get('drop_product'):
            product = None
            continue
        if 'add_screen' in op:
            screens[op['add_screen']['id']] = copy.deepcopy(op['add_screen'])
            continue
        if 'drop_screen' in op:
            del screens[op['drop_screen']]
            continue
        target = product if op['file'] == 'product' else screens[op['file']]
        *head, last = op['path']
        for key in head:
            target = target[key]
        if op.get('del'):
            del target[last]
        else:
            target[last] = op['set']
    for key, screen in screens.items():
        screen.setdefault('_file', key)
    return product, list(screens.values()), tuple(unapproved)
