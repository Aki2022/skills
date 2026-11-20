"""blueprint: the infoarch of a product UI (token 0).

SPEC-product-ui-infoarch (and SPEC-deliverable-production-layers requirements 3, 4 and 5). Two kinds of files under
design/info_arch/blueprint/:

- `product.yaml`, one per product: stories, principles, constraints, objects (one of them the lead) and entries. The
  owner approves each item on its own: the approval row is `ia:item:<item id>` and its hash covers that item's
  content alone, so adding an item, or changing another, never moves an item's approval.
- `screens/<screen id>.yaml`, one per screen: the stories it serves, groups and regions with a priority, the primary
  action, the states it declares, where its actions lead. The AI writes them; a screen may only cite approved items.

Only what does not change with the screen width is written here (position, width and parts are design's: a region says
what kind of thing it shows, `shape`, and design's allocation rules choose the part). Order is the `priority` field.
What is checked is structure and references, never meaning.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

from common import INFO_ARCH_DIR, YamlError, load_yaml, validator_for  # first: it puts the sibling skills on sys.path

import positioning as pos  # noqa: E402

BLUEPRINT_DIR = INFO_ARCH_DIR / 'blueprint'
PRODUCT_PATH = BLUEPRINT_DIR / 'product.yaml'
SCREENS_DIR = BLUEPRINT_DIR / 'screens'
ITEM_BLOCKS = {'stories': 'sto', 'principles': 'pri', 'constraints': 'con', 'objects': 'ojc', 'entries': 'ntr'}
ITEM_PREFIX = pos.ITEM_PREFIX  # ia:item:<item id>
SHAPES = ('value', 'list', 'comparison', 'trend', 'distribution', 'status', 'form', 'text')
STATES = ('loading', 'empty', 'error', 'partial', 'offline', 'denied')
ITEM_ID = re.compile(r'(sto|pri|con|ojc|ntr)-((?!0000)[0-9]{4}|[1-9][0-9]{4,})')  # always fullmatch
CONSTRAINT_FORM = re.compile(r'.*ない[。.]?\Z', re.S)  # the negative form (〜しない, 〜伝えない): what is not done


@dataclass
class Blueprint:
    root: Path
    present: bool = False  # design/info_arch/blueprint/ exists
    product: dict | None = None
    product_rel: Path = PRODUCT_PATH
    screens: dict = field(default_factory=dict)  # screen id -> (rel path, data)

    def item(self, iid):
        m = ITEM_ID.fullmatch(iid or '')
        if not m or not self.product:
            return None
        block = next(b for b, p in ITEM_BLOCKS.items() if p == m.group(1))
        return (self.product.get(block) or {}).get(iid)


def load(root, report):
    """Read product.yaml and screens/*.yaml. Nothing is read when the directory is absent (a repository with no
    product UI). A file that does not parse or fit its schema is a failure."""
    root = Path(root)
    bp = Blueprint(root)
    base = root / BLUEPRINT_DIR
    if not base.is_dir():
        return bp
    bp.present = True
    product_file = root / PRODUCT_PATH
    if product_file.is_file():
        try:
            data = load_yaml(product_file)
        except YamlError as exc:
            report.add('fail', 'IA_YAML', f'YAML として読めない: {exc}', PRODUCT_PATH)
            data = None
        if data is not None:
            errs = pm_errors(validator_for('blueprint-product.v1.schema.json'), data)
            for e in errs:
                report.add('fail', 'IA_SCHEMA', e.message, PRODUCT_PATH, pointer='/' + '/'.join(map(str, e.absolute_path)))
            if not errs:
                bp.product = data
    screens_dir = root / SCREENS_DIR
    validator = validator_for('blueprint-screen.v1.schema.json')
    for f in sorted(screens_dir.iterdir()) if screens_dir.is_dir() else []:
        rel = f.relative_to(root)
        if f.name.startswith('.') and f.is_file():
            continue  # editor and OS dotfiles are not declarations
        if f.is_dir() or f.suffix != '.yaml':
            report.add('fail', 'IA_BLUEPRINT_SCREEN_FILE', f'{f.name} は読まれない（screens/<画面 id>.yaml だけ。黙って無視しない）', rel)
            continue
        try:
            data = load_yaml(f)
        except YamlError as exc:
            report.add('fail', 'IA_YAML', f'YAML として読めない: {exc}', rel)
            continue
        errs = pm_errors(validator, data)
        for e in errs:
            report.add('fail', 'IA_SCHEMA', e.message, rel, pointer='/' + '/'.join(map(str, e.absolute_path)))
        if errs:
            continue
        if data['id'] != f.stem:
            report.add('fail', 'IA_BLUEPRINT_SCREEN_ID', f'画面の id {data["id"]!r} がファイル名 {f.stem!r} と違う', rel)
            continue
        bp.screens[data['id']] = (rel, data)
    return bp


def pm_errors(validator, data):
    return sorted(validator.iter_errors(data), key=lambda e: list(e.absolute_path))


# ---- approval: one item, one row ---------------------------------------------------------------------------------

def item_hash(iid, item):
    """The hash an approval row binds: the item's own content (its id and everything written on it except where it
    came from). Nothing about any other item, so an unrelated change never moves it."""
    return pos.sha256([iid, pos.content(item)])


def items(bp):
    """{item id: item} of every item written in product.yaml, in one table."""
    if not bp.product:
        return {}
    return {iid: item for block in ITEM_BLOCKS for iid, item in (bp.product.get(block) or {}).items()}


def targets(bp):
    """{ia:item:<id>: hash} for every proposed item (a draft that is empty is not asked for)."""
    return {ITEM_PREFIX + iid: item_hash(iid, item) for iid, item in sorted(items(bp).items()) if item.get('state') == 'proposed'}


def approved_items(bp, rows):
    """Item ids whose recorded row matches the item's current hash. `rows`: {approval key: value_hash}."""
    return {k[len(ITEM_PREFIX):] for k, h in targets(bp).items() if rows.get(k) == h}


def screen_items(screen):
    """Every product item a screen cites, as (where, item id)."""
    out = [('serves', s) for s in screen['serves']]
    out += [('constraints', c) for c in screen.get('constraints', [])]
    out += [(f'groups.{g}.principle', grp['principle']) for g, grp in sorted(screen['groups'].items()) if 'principle' in grp]
    for r, reg in sorted(screen['regions'].items()):
        out += [(f'regions.{r}.objects', o) for o in reg.get('objects', [])]
    return out


def dependencies(bp, screen):
    """Every product item a screen stands on (what a change to an item reaches)."""
    return {iid for _, iid in screen_items(screen)}


def required_keys(bp):
    """The approval keys shipping needs: every proposed item a shipped screen cites."""
    need = set()
    for _, (_, s) in bp.screens.items():
        if s['ship']:
            need |= {ITEM_PREFIX + iid for iid in dependencies(bp, s)}
    return sorted(need & set(targets(bp)))


# ---- disqualification ---------------------------------------------------------------------------------------------

def _num(iid):
    return int(ITEM_ID.fullmatch(iid).group(2))


def check_product(bp, report, values, approved_values):
    p, rel = bp.product, bp.product_rel
    nxt, tomb = p['ledger']['next'], p['ledger']['tombstones']
    for block, prefix in ITEM_BLOCKS.items():
        for iid in (p.get(block) or {}):
            if _num(iid) >= nxt[prefix]:
                report.add('fail', 'IA_BLUEPRINT_ID_NOT_ISSUED', f'{iid} は台帳の次番号 {nxt[prefix]} 以上で、台帳が発行していない', rel, id=iid)
            if iid in tomb:
                report.add('fail', 'IA_BLUEPRINT_ID_REUSED', f'{iid} は墓標にある（id は再利用しない）', rel, id=iid)
    for sid in bp.screens:
        if sid in tomb:
            report.add('fail', 'IA_BLUEPRINT_ID_REUSED', f'画面 {sid} は墓標にある（id は再利用しない）', rel, id=sid)
    for old, new in tomb.items():
        if new is not None and ITEM_ID.fullmatch(new) and bp.item(new) is None:
            report.add('fail', 'IA_BLUEPRINT_REF_BROKEN', f'墓標 {old} → {new} の先が無い', rel)
    stories = {i: s for i, s in p['stories'].items() if s['state'] == 'proposed'}
    if not stories:
        report.add('fail', 'IA_BLUEPRINT_NO_STORIES', 'ストーリーが 0 件（対象 0 件を合格にしない）', rel)
    leads = [i for i, o in p['objects'].items() if o['state'] == 'proposed' and o.get('lead')]
    if len(leads) != 1:
        report.add('fail', 'IA_BLUEPRINT_LEAD_COUNT', f'主役のオブジェクトが {len(leads)} 件 {sorted(leads)}（ちょうど 1 件）', rel)
    for iid, c in p['constraints'].items():
        if c['state'] == 'proposed' and not CONSTRAINT_FORM.fullmatch(c['text'].strip()):
            report.add('fail', 'IA_BLUEPRINT_CONSTRAINT_FORM', f'{iid}: 制約は「〜しない」「〜ない」の形で書く（やらないことを書く）', rel, id=iid)
    # principles and the lead object stand on positioning values: that is what makes them the concept (requirement 2)
    holders = [(i, x) for i, x in p['principles'].items() if x['state'] == 'proposed'] + \
              [(i, p['objects'][i]) for i in leads]
    for iid, item in holders:
        refs = item.get('values') or []
        if not refs:
            report.add('fail', 'IA_BLUEPRINT_VALUE_NONE', f'{iid}: positioning の提供価値の id を 1 つも引いていない（values）', rel, id=iid)
        for v in refs:
            atom = values.get(v)
            if atom is None or pos._prefix(v) != 'val' or atom.get('state') == 'empty':
                report.add('fail', 'IA_BLUEPRINT_REF_BROKEN', f'{iid}: 提供価値 {v} が positioning に無いか空欄', rel, id=iid)
            elif v not in approved_values:
                report.add('warn', 'IA_BLUEPRINT_VALUE_UNAPPROVED', f'{iid}: 提供価値 {v} の経路にオーナーの承認が無い', rel, id=iid)
    for iid, item in list(p['principles'].items()) + list(p['entries'].items()):
        if item['state'] != 'proposed':
            continue
        for s in item.get('serves', []):
            if bp.item(s) is None:
                report.add('fail', 'IA_BLUEPRINT_REF_BROKEN', f'{iid}: 応えるストーリー {s} が無い', rel, id=iid)
    for iid, e in p['entries'].items():
        if e['state'] == 'proposed' and e['screen'] not in bp.screens:
            report.add('fail', 'IA_BLUEPRINT_REF_BROKEN', f'{iid}: 入口の画面 {e["screen"]} が無い', rel, id=iid)


def check_screen(bp, report, sid, rel, s, approved):
    shipped = s['ship']
    if not s['serves']:
        report.add('fail', 'IA_BLUEPRINT_SCREEN_NO_STORY', f'{sid}: どのストーリーにも応えない（serves が空。優先度の基準が無い画面は作らない）', rel)
    for where, iid in screen_items(s):
        item = bp.item(iid)
        expected = {'serves': 'sto', 'constraints': 'con'}.get(where.split('.')[0], 'pri' if where.endswith('principle') else 'ojc')
        if item is None or not ITEM_ID.fullmatch(iid) or ITEM_ID.fullmatch(iid).group(1) != expected:
            report.add('fail', 'IA_BLUEPRINT_REF_BROKEN', f'{sid}.{where}: 引く項目 {iid} が無い（または種類が違う）', rel, id=iid)
        elif item['state'] == 'empty':
            report.add('fail', 'IA_BLUEPRINT_REF_EMPTY', f'{sid}.{where}: 空欄の {iid} を引いている', rel, id=iid)
        elif iid not in approved:
            report.add('fail' if shipped else 'warn', 'IA_BLUEPRINT_ITEM_UNAPPROVED',
                       f'{sid}.{where}: 引く項目 {iid} にオーナーの承認が無いか外れている'
                       '（design-author approve --prepare で束を作り、オーナーの承認の言葉で記録する）', rel, id=iid)
    groups, regions, actions = s['groups'], s['regions'], s['actions']
    for g, grp in sorted(groups.items()):
        principle = bp.item(grp.get('principle'))
        if principle is None and 'principle' not in grp:
            report.add('fail', 'IA_BLUEPRINT_GROUP_NO_PRINCIPLE', f'{sid}.{g}: グループを正当化する並べ方の原則が無い（principle）', rel, id=g)
        elif principle is not None and principle.get('kind') != 'arrangement':
            report.add('fail', 'IA_BLUEPRINT_GROUP_NO_PRINCIPLE', f'{sid}.{g}: {grp["principle"]} は並べ方の原則でない（kind: arrangement）', rel, id=g)
    if not regions:
        report.add('fail', 'IA_BLUEPRINT_NO_REGIONS', f'{sid}: 領域が 0 件', rel)
    priorities = {}
    for r, reg in sorted(regions.items()):
        if reg['group'] not in groups:
            report.add('fail', 'IA_BLUEPRINT_REF_BROKEN', f'{sid}.{r}: グループ {reg["group"]} が無い', rel, id=r)
        if 'priority' not in reg:
            report.add('fail', 'IA_BLUEPRINT_PRIORITY_MISSING', f'{sid}.{r}: 優先度が無い（小さいほど上位。優先度の無い領域は位置を導けない）', rel, id=r)
        else:
            priorities.setdefault(reg['priority'], []).append(r)
    for n, rs in sorted(priorities.items()):
        if len(rs) > 1:
            report.add('fail', 'IA_BLUEPRINT_PRIORITY_DUPLICATE', f'{sid}: 領域 {sorted(rs)} の優先度がどれも {n}（順が決まらない）', rel)
    primary = [a for a, act in actions.items() if act.get('primary')]
    if len(primary) > 1:
        report.add('fail', 'IA_BLUEPRINT_PRIMARY_MANY', f'{sid}: 主操作が {len(primary)} つ {sorted(primary)}（1 つだけ）', rel)
    if s.get('no_primary') and primary:
        report.add('fail', 'IA_BLUEPRINT_PRIMARY_CONTRADICTS', f'{sid}: 「主操作なし」と宣言したのに主操作 {sorted(primary)} がある', rel)
    if not primary and not s.get('no_primary'):
        report.add('fail', 'IA_BLUEPRINT_PRIMARY_NONE', f'{sid}: 主操作が無い（主操作を持たない画面なら no_primary: true と書く）', rel)
    for a, act in sorted(actions.items()):
        if act['region'] not in regions:
            report.add('fail', 'IA_BLUEPRINT_REF_BROKEN', f'{sid}.{a}: 操作の領域 {act["region"]} が無い', rel, id=a)
        if 'to' in act and act['to'] not in bp.screens:
            report.add('fail', 'IA_BLUEPRINT_REF_BROKEN', f'{sid}.{a}: 移動先の画面 {act["to"]} が無い', rel, id=a)
    if not s['states']:
        report.add('fail', 'IA_BLUEPRINT_STATES_MISSING', f'{sid}: 状態の宣言が無い（読み込み中・空・エラーなど、出る状態を書く）', rel)
    for st, decl in sorted(s['states'].items()):
        if 'region' in decl and decl['region'] not in regions:
            report.add('fail', 'IA_BLUEPRINT_REF_BROKEN', f'{sid}.states.{st}: 領域 {decl["region"]} が無い', rel, id=st)


def check(bp, report, values, approved_values, approved):
    """Disqualification of the product and every screen. `values`: positioning atoms {id: atom}; `approved_values`:
    value ids whose path the owner approved; `approved`: product item ids whose approval row is current."""
    if not bp.present:
        return
    if bp.product is None:
        if not (bp.root / PRODUCT_PATH).is_file():
            report.add('fail', 'IA_BLUEPRINT_PRODUCT_MISSING', f'{PRODUCT_PATH} が無い（画面は、承認された項目を引く）', PRODUCT_PATH)
        return  # an unreadable product.yaml was reported by load
    check_product(bp, report, values, approved_values)
    if not bp.screens:
        report.add('fail', 'IA_BLUEPRINT_NO_SCREENS', 'blueprint/screens/*.yaml が 0 件（対象 0 件を合格にしない）', SCREENS_DIR)
    for sid, (rel, s) in sorted(bp.screens.items()):
        check_screen(bp, report, sid, rel, s, approved)
    report.add('info', 'IA_BLUEPRINT_CHECKED', f'blueprint: 画面 {len(bp.screens)} 件・項目 {len(items(bp))} 件を検査した', BLUEPRINT_DIR,
               screens=len(bp.screens), items=len(items(bp)))
