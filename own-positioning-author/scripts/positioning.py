"""Positioning: atoms and arrows per segment, owner approval per value path, derived gaps and parity, ids (token 0).

SPEC-information-architecture-standardization v6, requirements 2 (atoms and arrows), 3 (approval per value path),
4 (derivation), 5 (disqualification) and 8 (ids); ADR-20261116-ia-v6-positioning-and-outline and
ADR-20261118-ia-approval-per-value-path (biz_ops).

- An element is an atom (positioning blocks, and proofs that belong to a segment) or an arrow. An arrow's key is
  `<kind>:<from>:<to>`; a `supports` arrow may point at another arrow's key.
- Approval is per value: design/approvals.yaml rows `ia:path:<value id>` whose value_hash is the hash of every atom
  and arrow on that value's path (claim, basis of the difference, premises, evidence). The path is extracted from
  every proposed element, so adding or removing an alternative or a degree arrow moves it. Per-element rows
  (`ia:positioning:`) are not read; `migrate-approvals` replaces them once.
- A positioning is every approved path of one segment taken together. Elements on no approved path stay drafts.
- Gaps and points of parity are derived per path, never recorded: a gap compares every alternative of the segment,
  a point of parity the alternatives of the member category only (ADR item 5).
- Ids hash content only: the order things are written in, where a copy came from (`copied_from`) and who wrote an
  element (`provenance`) never move an id or an approval.
"""
from __future__ import annotations

import hashlib
import json
from dataclasses import dataclass, field

import ia_common  # noqa: F401  (first: it puts own-design-author/scripts on sys.path)
import layer_paths

ATOM_BLOCKS = {'segments': 'seg', 'jobs': 'job', 'pains': 'pan', 'gains': 'gan', 'alternatives': 'alt',
               'categories': 'cat', 'strengths': 'str', 'values': 'val'}
EVIDENCE_BLOCK = 'proofs'  # a proof with `segment` is positioning evidence; one without supports outline messages
NEEDS = {'pan', 'gan'}
ATOM_PREFIXES = set(ATOM_BLOCKS.values()) | {'prf'}
EDGE_KINDS = {  # kind -> (allowed `from` prefixes, allowed `to` prefixes; 'edge' = another arrow's key)
    'has_job': ({'seg'}, {'job'}),
    'need': ({'job'}, NEEDS),
    'fulfills': ({'alt'}, NEEDS),
    'recalls': ({'job'}, {'cat'}),
    'includes': ({'cat'}, {'alt'}),
    'member_of': ({'seg'}, {'cat'}),
    'lacks': ({'str'}, {'alt'}),
    'produces': ({'str'}, {'val'}),
    'satisfies': ({'val'}, NEEDS),
    'supports': ({'prf'}, (ATOM_PREFIXES - {'prf', 'seg'}) | {'edge'}),
}
DEGREES = ('full', 'partial', 'none')
PATH_PREFIX = 'ia:path:'             # one approval row per value: its key names the value, its hash binds the path
ITEM_PREFIX = 'ia:item:'              # one approval row per product item (own-infoarch-author's blueprint): the item's own hash
LEGACY_PREFIX = 'ia:positioning:'    # the per-element rows before ADR-20261118-ia-approval-per-value-path
APPROVAL_PREFIX = PATH_PREFIX
ID_PREFIX = 'pos-'
PATH_ID_PREFIX = 'pth-'


def canonical(value):
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(',', ':'))


def sha256(value):
    return hashlib.sha256(canonical(value).encode()).hexdigest()


def edge_key(edge):
    return f'{edge["kind"]}:{edge["from"]}:{edge["to"]}'


METADATA = ('copied_from', 'provenance')  # where a copy came from and who wrote it: not what the element says


def content(element):
    """What an element says: everything written on it except its metadata."""
    return {k: v for k, v in element.items() if k not in METADATA}


def element_hash(key, element):
    return sha256([key, content(element)])


def _prefix(aid):
    return aid.split('-')[0] if isinstance(aid, str) and '-' in aid else None


def atoms(m, proposed_only=True):
    """{atom id: atom} of every positioning atom (and positioning evidence)."""
    out = {}
    for block in ATOM_BLOCKS:
        for aid, a in (m.ssot.get(block) or {}).items():
            if not proposed_only or a.get('state') == 'proposed':
                out[aid] = a
    for aid, a in (m.ssot.get(EVIDENCE_BLOCK) or {}).items():
        if 'segment' in a and (not proposed_only or a.get('state') == 'proposed'):
            out[aid] = a
    return out


def edges(m, proposed_only=True):
    """{edge key: edge}; the first one wins when a key repeats (the duplicate is reported by the checker)."""
    out = {}
    for e in m.ssot.get('edges') or []:
        if not proposed_only or e.get('state') == 'proposed':
            out.setdefault(edge_key(e), e)
    return out


def segment_map(m):
    """{key: segment} for every atom and arrow: an atom's `segment`; an arrow's is its `from` atom's."""
    out = {aid: aid if _prefix(aid) == 'seg' else a.get('segment') for aid, a in atoms(m, proposed_only=False).items()}
    for key, e in edges(m, proposed_only=False).items():
        out.setdefault(key, out.get(e['from']))
    return out


def segment_of(m, key):
    return segment_map(m).get(key)


# ---- the path of one value (requirement 3) ---------------------------------------------------------------------

ROLES = ('claim', 'basis', 'premise', 'evidence')  # the order an approval request shows them in


@dataclass
class ValuePath:
    """The atoms and arrows one value stands on, extracted from all current data (drafts included)."""
    value: str
    segment: str | None
    kind: str | None
    roles: dict = field(default_factory=dict)        # {member key: claim | basis | premise | evidence}
    targets: list = field(default_factory=list)      # the pains and gains the value satisfies
    comparison: list = field(default_factory=list)   # the alternatives it is compared against
    categories: list = field(default_factory=list)   # member_of targets (exactly one is required)
    missing_fulfillment: list = field(default_factory=list)  # [alternative, need] pairs with no degree arrow
    points: list | None = None                       # the gaps (difference) or points of parity (parity) it can meet
    problems: list = field(default_factory=list)     # [(code, message)] that disqualify the path
    hash: str = ''

    @property
    def id(self):
        return PATH_ID_PREFIX + self.hash

    @property
    def members(self):
        return sorted(self.roles)

    def to_dict(self):
        return {'id': self.id, 'segment': self.segment, 'kind': self.kind, 'targets': self.targets,
                'comparison': self.comparison, 'points': self.points, 'members': self.members,
                'missing_fulfillment': self.missing_fulfillment, 'problems': [c for c, _ in self.problems]}


def extract(m, vid, all_atoms=None, all_edges=None, segs=None):
    """Requirement 3: the path of value `vid`. Read from every proposed element, never from the approved ones only,
    so adding or removing an alternative, a degree arrow or a recalled category always moves the hash."""
    A = atoms(m) if all_atoms is None else all_atoms
    E = edges(m) if all_edges is None else all_edges
    segs = segment_map(m) if segs is None else segs
    v = A[vid]
    seg, kind = v.get('segment'), v.get('kind')
    p = ValuePath(vid, seg, kind)

    def mine(x):
        return x in A and segs.get(x) == seg

    def add(key, role):
        p.roles.setdefault(key, role)

    def arrows(kind_, frm=None, to=None):
        return [(k, e) for k, e in sorted(E.items()) if e['kind'] == kind_
                and (frm is None or e['from'] in frm) and (to is None or e['to'] in to)]

    # The claim: the value, the strengths that produce it and what those strengths have that alternatives lack.
    add(vid, 'claim')
    strengths = set()
    for k, e in arrows('produces', to={vid}):
        if mine(e['from']):
            strengths.add(e['from'])
            add(k, 'claim'), add(e['from'], 'claim')
    lacked = []
    for k, e in arrows('lacks', frm=strengths):
        if mine(e['to']):
            add(k, 'claim')
            lacked.append(e['to'])  # the alternative itself is shown with the basis when it is compared there
    # The basis of the difference: what it satisfies, and every alternative of the comparison set with its degrees.
    for k, e in arrows('satisfies', frm={vid}):
        if mine(e['to']):
            p.targets.append(e['to'])
            add(k, 'basis'), add(e['to'], 'basis')
    p.targets.sort()
    member_of = [(k, e) for k, e in arrows('member_of', frm={seg}) if mine(e['to'])]
    p.categories = sorted(e['to'] for _, e in member_of)
    if kind == 'parity':
        includes = [(k, e) for k, e in arrows('includes', frm=set(p.categories)) if mine(e['to'])]
        for k, _ in includes:
            add(k, 'basis')
        p.comparison = sorted({e['to'] for _, e in includes})
    else:
        p.comparison = sorted(a for a in A if _prefix(a) == 'alt' and segs.get(a) == seg)
    degree = {}
    for alt in p.comparison:
        add(alt, 'basis')
        for need in p.targets:
            k = f'fulfills:{alt}:{need}'
            if k in E:
                add(k, 'basis')
                degree[(alt, need)] = E[k].get('degree')
            else:
                p.missing_fulfillment.append([alt, need])
    for alt in lacked:
        add(alt, 'claim')
    # The premises: the segment, the jobs behind what the value satisfies, the member category and what the jobs recall.
    if seg in A:  # a value whose segment is missing is IA_REF_BROKEN; its path holds what is there
        add(seg, 'premise')
    jobs, has_need = set(), set()
    for k, e in arrows('need', to=set(p.targets)):
        if mine(e['from']):
            jobs.add(e['from'])
            has_need.add(e['to'])
            add(k, 'premise'), add(e['from'], 'premise')
    for k, e in arrows('has_job', frm={seg}, to=jobs):
        add(k, 'premise')
    for k, e in member_of:
        add(k, 'premise'), add(e['to'], 'premise')
    for k, e in arrows('recalls', frm=jobs):
        if mine(e['to']):
            add(k, 'premise'), add(e['to'], 'premise')
    # Evidence on anything above (a supports arrow may support another supports arrow: follow until nothing is added).
    grew = True
    while grew:
        grew = False
        for k, e in arrows('supports'):
            if k not in p.roles and e['to'] in p.roles and mine(e['from']):
                add(k, 'evidence'), add(e['from'], 'evidence')
                grew = True
    _disqualify(p, A, E, strengths, jobs, has_need, degree)
    p.hash = sha256([[k, content(A.get(k) or E[k])] for k in p.members])
    return p


def _disqualify(p, A, E, strengths, jobs, has_need, degree):
    """Requirements 3 and 5 for one path. A missing input is reported, never read as "does not meet"."""
    seg, vid = p.segment, p.value
    on = [E[k] for k in p.roles if k in E]
    touched = {e['from'] for e in on} | {e['to'] for e in on}
    for aid in sorted(k for k in p.roles if k in A):
        if aid not in touched:
            p.problems.append(('IA_POSITIONING_ISOLATED_ATOM', f'{vid}: 経路の {aid} がどの矢印にも繋がらない'))
    for need in p.targets:
        if need not in has_need:
            p.problems.append(('IA_POSITIONING_JOB_WITHOUT_NEED',
                               f'{vid}: 満たす先の {need} を持つジョブが無い（ジョブ→痛み・得たいもの の矢印が要る）'))
    unjobbed = sorted(j for j in jobs if f'has_job:{seg}:{j}' not in E)
    missing = [name for name, ok in (('ジョブ', jobs), ('痛みか得たいもの', p.targets)) if not ok]
    if missing or unjobbed:
        why = f'{"・".join(missing)} が 0 件' if missing else f'ジョブ {unjobbed} がセグメントに繋がらない（has_job）'
        p.problems.append(('IA_POSITIONING_INCOMPLETE', f'{vid}: 経路の{why}（中身の無い経路を合格にしない）'))
    if len(p.categories) != 1:
        p.problems.append(('IA_POSITIONING_CATEGORY_COUNT',
                           f'{seg}: 所属カテゴリー（member_of）が {len(p.categories)} 本 {p.categories}（ちょうど 1 本）'))
    if not p.comparison:
        if p.kind == 'parity':
            if len(p.categories) == 1:
                p.problems.append(('IA_POSITIONING_NO_CATEGORY_ALTERNATIVES',
                                   f'{vid}: 所属カテゴリー {p.categories[0]} に入っている代替が 0 件（同等であるべき点を導けない）'))
        else:
            p.problems.append(('IA_POSITIONING_NO_ALTERNATIVES', f'{vid}: 隙間の比較に使う {seg} の代替が 0 件'))
    if p.missing_fulfillment:
        p.problems.append(('IA_POSITIONING_FULFILLMENT_MISSING',
                           f'{vid}: 比べる集合の満たす度合いの矢印が無い組がある {p.missing_fulfillment}（未調査を「満たさない」とみなさない）'))
    for s in sorted(strengths):
        if not any(e['kind'] == 'lacks' and e['from'] == s for e in on):
            p.problems.append(('IA_POSITIONING_STRENGTH_UNCOMPARED',
                               f'{vid}: 強み {s} がどの代替とも比べていない（lacks の矢印が無い）'))
    if p.comparison and not p.missing_fulfillment:
        if p.kind == 'parity':
            p.points = [n for n in p.targets if all(degree[(a, n)] == 'full' for a in p.comparison)]
        else:
            p.points = [n for n in p.targets if all(degree[(a, n)] != 'full' for a in p.comparison)]
        if not p.points:
            name = '同等であるべき点' if p.kind == 'parity' else '隙間'
            p.problems.append(('IA_POSITIONING_VALUE_UNANCHORED',
                               f'{vid}: 提供価値（{p.kind}）がどの{name}も満たさない（満たす先: {p.targets}）'))


def paths(m):
    """{value id: ValuePath} for every proposed value."""
    A, E, segs = atoms(m), edges(m), segment_map(m)
    return {vid: extract(m, vid, A, E, segs) for vid in sorted(A) if _prefix(vid) == 'val'}


def approval_targets(m):
    """{ia:path:<value id>: path hash} for every proposed value (requirement 3)."""
    return {PATH_PREFIX + v: p.hash for v, p in paths(m).items()}


def approved_keys(m, rows):
    """Value ids whose path hash matches the owner's recorded row. `rows`: {approval key: value_hash}. A row of one
    path never approves another (its key names the value), and per-element rows (ia:positioning:) are not read."""
    return {k[len(PATH_PREFIX):] for k, h in approval_targets(m).items() if rows.get(k) == h}


@dataclass
class Positioning:
    """Requirement 4: every approved path of one segment, taken together."""
    segment: str
    paths: dict = field(default_factory=dict)        # {value id: path id} of the approved paths
    unapproved: list = field(default_factory=list)   # values of the segment whose path waits for the owner
    atoms: list = field(default_factory=list)
    edges: list = field(default_factory=list)
    categories: list = field(default_factory=list)   # member_of targets on the approved paths
    gaps: list = field(default_factory=list)
    parity: list = field(default_factory=list)
    id: str = ''

    @property
    def category(self):
        return self.categories[0] if len(self.categories) == 1 else None

    def to_dict(self):
        return {'id': self.id, 'paths': self.paths, 'unapproved': self.unapproved, 'atoms': self.atoms,
                'edges': self.edges, 'category': self.category, 'gaps': self.gaps, 'parity': self.parity}


def derive(m, approved, all_paths=None):
    """{segment: Positioning} for every segment with at least one approved path. Elements on no approved path stay
    drafts: they are not part of any positioning and an outline cannot stand on them."""
    ps = paths(m) if all_paths is None else all_paths
    A, E = atoms(m), edges(m)
    out = {}
    for vid, p in ps.items():
        if p.segment is None:
            continue
        if vid not in approved:
            out.setdefault(p.segment, Positioning(p.segment)).unapproved.append(vid)
            continue
        q = out.setdefault(p.segment, Positioning(p.segment))
        q.paths[vid] = p.id
        q.atoms = sorted(set(q.atoms) | {k for k in p.roles if k in A})
        q.edges = sorted(set(q.edges) | {k for k in p.roles if k in E})
        q.categories = sorted(set(q.categories) | set(p.categories))
        if p.points:
            target = q.parity if p.kind == 'parity' else q.gaps
            target[:] = sorted(set(target) | set(p.points))
    out = {s: q for s, q in out.items() if q.paths}
    for q in out.values():
        q.id = ID_PREFIX + sha256([[k, content(A.get(k) or E[k])] for k in q.atoms + q.edges])
    return out


# ---- checks ----------------------------------------------------------------------------------------------------

SSOT_REL = layer_paths.rel('positioning_file')  # repository-relative name; reading goes through layer_paths.path


def check_shape(m, report, path=SSOT_REL):
    """Requirement 2, over every proposed element (drafts too: a malformed draft cannot be approved into shape)."""
    all_atoms, all_edges = atoms(m), edges(m)
    every_atom, every_edge = atoms(m, proposed_only=False), edges(m, proposed_only=False)
    segs, seen = segment_map(m), set()
    for e in m.ssot.get('edges') or []:
        key = edge_key(e)
        if key in seen:
            report.add('fail', 'IA_POSITIONING_EDGE_DUPLICATE', f'矢印 {key} が 2 本ある（同じ種類・始点・終点は 1 本だけ）', path, key=key)
        seen.add(key)
    for key, e in all_edges.items():
        froms, tos = EDGE_KINDS[e['kind']]
        ends = []
        for side, allowed in (('from', froms), ('to', tos)):
            target = e[side]
            if target in every_atom:
                kind = _prefix(target)
            elif side == 'to' and target in every_edge and 'edge' in allowed:
                kind = 'edge'
            else:
                report.add('fail', 'IA_REF_BROKEN', f'矢印 {key} の{"始点" if side == "from" else "終点"} {target} が無い', path, key=key)
                continue
            if kind not in allowed:
                report.add('fail', 'IA_POSITIONING_EDGE_SHAPE',
                           f'矢印 {key}：{e["kind"]} の{"始点" if side == "from" else "終点"}は {sorted(allowed)} の原子（{target} は {kind}）', path, key=key)
                continue
            if (every_atom.get(target) or every_edge.get(target) or {}).get('state') == 'empty':
                report.add('fail', 'IA_EMPTY_REFERENCED', f'矢印 {key} が空欄の {target} を指している', path, key=key)
            ends.append(target)
        if (e['kind'] == 'fulfills') != ('degree' in e):
            report.add('fail', 'IA_POSITIONING_EDGE_SHAPE',
                       f'矢印 {key}：満たす度合い（degree）は fulfills にだけ、必ず書く', path, key=key)
        ends_seg = {segs.get(x) for x in ends}
        if len(ends) == 2 and len(ends_seg) > 1:
            report.add('fail', 'IA_POSITIONING_CROSS_SEGMENT',
                       f'矢印 {key} が別のセグメントの原子を結ぶ {sorted(s or "?" for s in ends_seg)}（原子は共有しない。写しを作る）', path, key=key)
        if 'copied_from' in e and e['copied_from'] not in every_edge:
            report.add('fail', 'IA_POSITIONING_COPY_SOURCE', f'矢印 {key} の写しの出どころ {e["copied_from"]} が無い', path, key=key)
    for aid, a in all_atoms.items():
        src = a.get('copied_from')
        if src is None:
            continue
        origin = every_atom.get(src)
        why = ('が無い' if origin is None else
               'が別の種類の原子' if _prefix(src) != _prefix(aid) else
               'が同じセグメントの原子（写しは別のセグメントで使うときに作る）' if origin.get('segment') == a.get('segment') else None)
        if why:
            report.add('fail', 'IA_POSITIONING_COPY_SOURCE', f'{aid} の写しの出どころ {src} {why}', path, id=aid)


def check_positioning(m, report, approved, referenced=None, path=SSOT_REL):
    """Requirement 2 (shape, every element) and requirements 3 and 5 (per path). -> ({segment: Positioning}, paths).

    `referenced` is {value id: True when a shipped outline points at it, False for a draft outline}. A disqualified
    path fails only when a shipped outline points at it; otherwise it warns when it is approved or pointed at, and a
    draft nobody points at is not judged (requirement 5: only the outlines that point at a path stop)."""
    referenced = referenced or {}
    check_shape(m, report, path)
    ps = paths(m)
    derived = derive(m, approved, ps)
    for vid, p in ps.items():
        if vid not in referenced and vid not in approved:
            continue
        level = 'fail' if referenced.get(vid) else 'warn'
        for code, message in p.problems:
            report.add(level, code, message, path, id=vid, path_id=p.id,
                       **({'pairs': p.missing_fulfillment} if code == 'IA_POSITIONING_FULFILLMENT_MISSING' else {}))
    for seg, q in derived.items():
        pending = [v for v in q.unapproved if v not in referenced]
        if pending:
            report.add('warn', 'IA_POSITIONING_PATH_PENDING',
                       f'{seg}: 提供価値 {pending} の経路が未承認（どの outline も指していないので止めない）', path,
                       segment=seg, values=pending)
    return derived, ps


def _verify_rows():
    """own-design-author's approval-ledger validator: IA reads the ledger exactly as design does (no differential)."""
    import sys
    from pathlib import Path
    design = Path(__file__).resolve().parents[2] / 'own-design-author/scripts'
    if not (design / 'approvals.py').is_file():
        raise RuntimeError('own-design-author/scripts/approvals.py が無い（承認台帳は design と共有）')
    if str(design) not in sys.path:
        sys.path.insert(0, str(design))
    from approvals import verify_rows
    return verify_rows


def read_approvals(root, report, load_yaml, yaml_error):
    """The owner's ia:path rows (and the blueprint's ia:item rows) of design/approvals.yaml: {key: value_hash}.

    Validated by own-design-author's verify_rows (fields, key namespace, hash, approver, ISO date, no duplicate key)
    plus a symlink refusal. Any violation fails closed: IA_APPROVALS_INVALID and nothing counts as approved.
    """
    rel = layer_paths.rel('ledger')
    path = layer_paths.path(root, 'ledger')  # a repository that moved stops here, never "no approvals"
    if path.is_symlink():
        report.add('fail', 'IA_APPROVALS_INVALID', f'{rel} が symlink（承認台帳は repo の実ファイルだけを読む）', rel)
        return {}
    if not path.is_file():
        return {}
    try:
        rows = load_yaml(path)
    except yaml_error as exc:
        report.add('fail', 'IA_APPROVALS_INVALID', f'{rel} が読めない: {exc}', rel)
        return {}
    if rows is None:
        return {}
    try:
        approved = _verify_rows()(rows, root)
    except ValueError as exc:
        report.add('fail', 'IA_APPROVALS_INVALID', f'{rel} が承認台帳の形でない: {exc}（何も承認済みとして数えない）', rel)
        return {}
    legacy = sorted(k for k in approved if k.startswith(LEGACY_PREFIX))
    if legacy:
        report.add('warn', 'IA_APPROVALS_LEGACY',
                   f'{rel} に原子と矢印ごとの旧い行が {len(legacy)} 件ある（承認として数えない。positioning-author migrate-approvals で経路の行へ移す）',
                   rel, rows=len(legacy))
    return {k: v for k, v in approved.items() if k.startswith((PATH_PREFIX, ITEM_PREFIX))}
