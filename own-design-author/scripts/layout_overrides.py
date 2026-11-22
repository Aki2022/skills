"""Layout overrides: a placement against the allocation rules, approved one row at a time.

SPEC-deliverable-production-layers requirement 4, ADR-20261121-layers-overrides-measurement-ux-folder. The file is
`layout_overrides` of layer-paths.json (ux/design/layout-overrides.yaml): a list of rows

    - id: ovr-0001          # the row's own name; two rows never share one
      screen: scr-overview  # the screen the row is about
      above: rgn-0002       # the region the owner wants on top (the rules put it lower)
      below: rgn-0001       # the region the rules put higher
      viewports: [desktop]  # the widths the row allows. Never left out: leaving it out used to mean every width
      reason: ...           # why. A reason alone never passes a row: anyone can write one

A row allows the screen, the pair of regions and the widths it names, and nothing else. The owner approves each row in the
approval ledger (`ds:web:override:<id>`, the same ledger and the same shape as a page exclusion or a product item). The hash
binds the row's own content AND the declarations of the two regions it points at (as the blueprint's screen file writes them),
so:

- changing the row anywhere (a width, the reason, the pair, the screen) is a new hash: that row alone is unapproved again;
- a region the row points at changing or disappearing, or its screen disappearing, drops the approval (nothing to carry over);
- copying a row keeps nothing: the copy needs its own id (a repeated id or a repeated scope is disqualified), and its id has no approval;
- adding a row, or changing another, never moves the approval of a row (the hash holds nothing of any other row).

Only the owner approves (`design-author approve`, words and date recorded). This module only reads: it computes what a row stands on and
says which rows hold a current approval.
"""
from __future__ import annotations

import json
import re
from pathlib import Path

import layer_paths
from approvals import verify_rows
from dtcg import value_hash

FIELDS = ('id', 'screen', 'above', 'below', 'viewports', 'reason')
ID = re.compile(r'ovr-[0-9]{4,}')
NAME = re.compile(r'[a-z][a-z0-9-]*')  # a screen or a region id (also a file name, so never a path)
KEY_PREFIX = 'ds:web:override:'
ALLOCATION_RULES = Path(__file__).resolve().parents[1] / 'references' / 'allocation-rules.json'

# Every code is a disqualification, except UNAPPROVED, which the ship gate turns into one (a check run without `--ship` warns).
UNAPPROVED = 'OVERRIDE_UNAPPROVED'


def key_for(row_id):
    return KEY_PREFIX + row_id


def widths():
    """The width names a row may allow: the breakpoints of the allocation rules."""
    rules = json.loads(ALLOCATION_RULES.read_text(encoding='utf-8'))
    return tuple(b['name'] for b in rules['breakpoints'])


def _unique_yaml(text):
    try:
        import yaml
    except ImportError as exc:
        raise ValueError('PyYAML is required to read the layout overrides and the blueprint') from exc

    class Loader(yaml.SafeLoader):
        pass

    def construct(loader, node):
        out = {}
        for k, v in node.value:
            key = loader.construct_object(k)
            if key in out:
                raise ValueError(f'duplicate key: {key}')
            out[key] = loader.construct_object(v)
        return out

    Loader.add_constructor(yaml.resolver.BaseResolver.DEFAULT_MAPPING_TAG, construct)
    try:
        return yaml.load(text, Loader=Loader)
    except yaml.YAMLError as exc:
        raise ValueError('malformed YAML') from exc


def _load(path):
    return _unique_yaml(Path(path).read_text(encoding='utf-8'))


def read_rows(path):
    """The rows of an overrides file; [] when there is no file. A file that is not a list of rows is a ValueError."""
    if not path or not Path(path).is_file():
        return []
    data = _load(path)
    if data is None:
        return []
    if not isinstance(data, list):
        raise ValueError('layout-overrides は行の一覧（- id: …）でなければならない')
    return data


def read_approvals(root):
    """{ledger key: value hash} of the repository's approval ledger; {} when it has none."""
    path = layer_paths.path(root, 'ledger')
    if not path.is_file():
        return {}
    return verify_rows(_load(path) or [], root)


def _blueprint_regions(root, screen):
    """The regions of the blueprint screen (`ux/infoarch/blueprint/screens/<screen>.yaml`) -> dict, or None when it cannot be read."""
    path = layer_paths.path(root, 'infoarch', 'blueprint', 'screens', f'{screen}.yaml')
    if not path.is_file():
        return None
    try:
        data = _load(path)
    except (ValueError, OSError):
        return None
    regions = data.get('regions') if isinstance(data, dict) else None
    return regions if isinstance(regions, dict) else None


def _text(v):
    return isinstance(v, str) and bool(v.strip())


class Row:
    """One row of the file and what it stands on."""

    def __init__(self, index, data):
        self.index = index
        self.data = data
        self.errors = []  # [(code, message)]
        self.id = None
        self.hash = None
        self.status = 'invalid'  # invalid | unapproved | approved

    @property
    def key(self):
        return key_for(self.id) if self.id else None

    @property
    def where(self):
        return self.id or f'{self.index + 1} 行目'

    def fail(self, code, message):
        self.errors.append((code, f'{self.where}: {message}'))


class Evaluation:
    def __init__(self, rows):
        self.rows = rows

    @property
    def findings(self):
        """Every disqualification (code, message), in file order."""
        return [e for r in self.rows for e in r.errors]

    @property
    def targets(self):
        """{ledger key: hash} of the rows that can be approved (read well, pointing at a real screen and regions)."""
        return {r.key: r.hash for r in self.rows if r.status in ('unapproved', 'approved')}

    @property
    def declared_keys(self):
        """The ledger key of every row that has a usable id, whether or not the rest of it is sound (a stale approval is not an unknown one)."""
        return {r.key for r in self.rows if r.id}

    def unapproved(self):
        return [r for r in self.rows if r.status == 'unapproved']

    def approved_for(self, screen_id):
        """The rows, as read, that hold a current approval for this screen."""
        return [dict(r.data) for r in self.rows if r.status == 'approved' and r.data.get('screen') == screen_id]

    def problems_for(self, screen_id):
        """Findings the audit of `screen_id` must fail on: every structural error anywhere in the file (a row with no screen is about
        every screen until it is read), and a row of this screen that holds no current approval."""
        out = list(self.findings)
        for r in self.rows:
            if r.status == 'unapproved' and r.data.get('screen') == screen_id:
                out.append((UNAPPROVED, f'{r.id}: オーナーの承認が無いか、承認した後に行（または指す領域）が変わった。承認の無い行は何も許さない'))
        return out


def evaluate(root, rows, approvals):
    """Read `rows` against the repository's blueprint and `approvals` ({key: hash}) -> Evaluation. Nothing is written."""
    allowed_widths = widths()
    results = [Row(i, r) for i, r in enumerate(rows)]
    for r in results:
        d = r.data
        if not isinstance(d, dict):
            r.fail('OVERRIDE_ROW', '行が対応表（キーと値）でない')
            continue
        if isinstance(d.get('id'), str) and ID.fullmatch(d['id']):
            r.id = d['id']
        else:
            r.fail('OVERRIDE_ID', f'id が無い、または ovr-0001 の形でない（行の名前。承認はこの名前に付く）: {d.get("id")!r}')
        extra = sorted(set(d) - set(FIELDS), key=str)
        if extra:
            r.fail('OVERRIDE_UNKNOWN_FIELD', f'知らない欄 {extra}（許す欄は {list(FIELDS)} だけ。黙って無視しない）')
        for field in ('screen', 'above', 'below'):
            if not (isinstance(d.get(field), str) and NAME.fullmatch(d[field])):
                r.fail('OVERRIDE_NO_TARGET', f'{field}（画面・領域）が無い、または id の形でない: {d.get(field)!r}')
        v = d.get('viewports')
        if not isinstance(v, list) or not v:
            r.fail('OVERRIDE_NO_WIDTH', f'viewports（許す幅）が無い。幅は省けない（省くと全部の幅に広がるため）。{list(allowed_widths)} から選ぶ')
        elif any(w not in allowed_widths for w in v) or len(set(v)) != len(v):
            r.fail('OVERRIDE_WIDTH_UNKNOWN', f'viewports {v!r} が、知らない幅か重複を含む（{list(allowed_widths)}）')
        if not _text(d.get('reason')):
            r.fail('OVERRIDE_NO_REASON', 'reason（理由）が無い')
        if d.get('above') is not None and d.get('above') == d.get('below'):
            r.fail('OVERRIDE_SAME_REGION', 'above と below が同じ領域')
    ids = {}
    for r in results:
        if r.id:
            ids.setdefault(r.id, []).append(r)
    for rid, group in ids.items():
        if len(group) > 1:
            for r in group:
                r.fail('OVERRIDE_DUPLICATE', f'id {rid} が {len(group)} 行にある（写した行は同じ名前を引き継げない。別の名前を付け、オーナーが承認し直す）')
    scopes = {}
    for r in results:
        d = r.data
        if not r.errors and isinstance(d, dict):
            for w in d['viewports']:
                scopes.setdefault((d['screen'], d['above'], d['below'], w), []).append(r)
    for (screen, above, below, w), group in scopes.items():
        if len({id(r) for r in group}) > 1:
            for r in group:
                r.fail('OVERRIDE_DUPLICATE_SCOPE', f'画面 {screen}・{above} を {below} より上・幅 {w} が、{len(group)} 行で重なる（写した行。1 つの範囲に 1 行）')
    for r in results:
        d = r.data
        if r.errors or not isinstance(d, dict):
            continue
        regions = _blueprint_regions(root, d['screen'])
        if regions is None:
            r.fail('OVERRIDE_SCREEN_MISSING', f'画面 {d["screen"]} の blueprint が無い（または読めない）。画面が消えた行は何も許さない')
            continue
        missing = [x for x in (d['above'], d['below']) if x not in regions or not isinstance(regions[x], dict)]
        if missing:
            r.fail('OVERRIDE_REGION_MISSING', f'画面 {d["screen"]} に領域 {missing} が無い。領域が消えた行は何も許さない')
            continue
        try:
            r.hash = value_hash({'override': r.id, 'screen': d['screen'], 'above': d['above'], 'below': d['below'],
                                 'viewports': sorted(d['viewports']), 'reason': d['reason'],
                                 'regions': {d['above']: regions[d['above']], d['below']: regions[d['below']]}})
        except (TypeError, ValueError):
            r.fail('OVERRIDE_REGION_MISSING', f'画面 {d["screen"]} の領域 {d["above"]}・{d["below"]} の宣言が、ハッシュに束ねられない形（文字・数・真偽・一覧・対応表だけ）')
            continue
        r.status = 'approved' if approvals.get(r.key) == r.hash else 'unapproved'
    return Evaluation(results)


def evaluate_repo(root, approvals=None):
    """The repository's own overrides file against its own ledger (or the `approvals` given)."""
    path = layer_paths.path(root, 'layout_overrides')
    return evaluate(root, read_rows(path), read_approvals(root) if approvals is None else approvals)


def stale_keys(evaluation, approvals):
    """Ledger rows for an override that is no longer in the file: inert (nothing reads them), but never silent."""
    return sorted(k for k in approvals if k.startswith(KEY_PREFIX) and k not in evaluation.declared_keys)
