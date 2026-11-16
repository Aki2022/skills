"""Positioning: atoms and arrows per segment, owner approval per element, derived gaps and parity, ids (token 0).

SPEC-information-architecture-standardization v6, requirements 2 (atoms and arrows), 4 (derivation),
5 (disqualification of an approved positioning) and 8 (ids); ADR-20261116-ia-v6-positioning-and-outline.

- An element is an atom (positioning blocks, and proofs that belong to a segment) or an arrow. An arrow's key is
  `<kind>:<from>:<to>`; a `supports` arrow may point at another arrow's key.
- Approval is per element: design/approvals.yaml rows `ia:positioning:<key>` whose value_hash is the element's own
  content hash. Changing one element withdraws only its own approval.
- A positioning is every approved atom of one segment and every approved arrow between approved atoms of it.
  Drafts (proposed, not approved) are never part of it and are never disqualified.
- Gaps and points of parity are derived, never recorded. Ids hash content only: the order things are written in,
  where a copy came from (`copied_from`) and who wrote an element (`provenance`) never move an id or an approval.
- Points of parity compare the alternatives of the member category only, gaps every alternative of the segment
  (ADR item 5), so a need every category alternative meets is parity even when an outside alternative does not.
"""
from __future__ import annotations

import hashlib
import json
from dataclasses import dataclass, field

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
APPROVAL_PREFIX = 'ia:positioning:'
ID_PREFIX = 'pos-'


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


def approval_targets(m):
    """{ia:positioning:<key>: content hash} for every proposed atom and arrow."""
    out = {APPROVAL_PREFIX + aid: element_hash(aid, a) for aid, a in atoms(m).items()}
    out.update({APPROVAL_PREFIX + k: element_hash(k, e) for k, e in edges(m).items()})
    return dict(sorted(out.items()))


def approved_keys(m, rows):
    """Keys whose current content hash matches the owner's recorded row. `rows`: {approval key: value_hash}."""
    return {k[len(APPROVAL_PREFIX):] for k, h in approval_targets(m).items() if rows.get(k) == h}


@dataclass
class Positioning:
    segment: str
    atoms: list = field(default_factory=list)          # approved atom ids of the segment, sorted
    edges: list = field(default_factory=list)          # approved arrow keys between approved atoms, sorted
    unapproved: list = field(default_factory=list)     # proposed elements of the segment waiting for the owner
    alternatives: list = field(default_factory=list)
    categories: list = field(default_factory=list)     # member_of targets (requirement 5 wants exactly one)
    category_alternatives: list = field(default_factory=list)
    missing_fulfillment: list = field(default_factory=list)  # [alternative, need] pairs without an approved arrow
    gaps: list | None = None
    parity: list | None = None
    id: str = ''

    @property
    def category(self):
        return self.categories[0] if len(self.categories) == 1 else None

    def to_dict(self):
        return {'id': self.id, 'atoms': self.atoms, 'edges': self.edges, 'unapproved': self.unapproved,
                'alternatives': self.alternatives, 'category': self.category, 'categories': self.categories,
                'category_alternatives': self.category_alternatives, 'gaps': self.gaps, 'parity': self.parity,
                'missing_fulfillment': self.missing_fulfillment}


def derive(m, approved):
    """{segment: Positioning} for every segment with at least one approved element (requirement 4)."""
    all_atoms, all_edges = atoms(m), edges(m)
    ok_atoms = {aid for aid in all_atoms if aid in approved}
    ok_edges = {k for k, e in all_edges.items() if k in approved and e['kind'] != 'supports'
                and e['from'] in ok_atoms and e['to'] in ok_atoms}
    # A supports arrow counts only when what it supports is itself in the positioning (an atom or an arrow).
    ok_edges |= {k for k, e in all_edges.items() if k in approved and e['kind'] == 'supports'
                 and e['from'] in ok_atoms and (e['to'] in ok_atoms or e['to'] in ok_edges)}
    out, segs = {}, segment_map(m)
    for key in sorted(set(all_atoms) | set(all_edges)):
        seg = segs.get(key)
        if seg is None:
            continue
        p = out.setdefault(seg, Positioning(seg))
        if key in ok_atoms:
            p.atoms.append(key)
        elif key in ok_edges:
            p.edges.append(key)
        else:
            p.unapproved.append(key)
    out = {s: p for s, p in out.items() if p.atoms or p.edges}
    for seg, p in out.items():
        _derive_one(m, p, all_atoms, all_edges)
    return out


def _derive_one(m, p, all_atoms, all_edges):
    on = set(p.edges)
    by_kind = {}
    for k in sorted(on):
        e = all_edges[k]
        by_kind.setdefault(e['kind'], []).append(e)
    p.alternatives = sorted(a for a in p.atoms if _prefix(a) == 'alt')
    p.categories = sorted(e['to'] for e in by_kind.get('member_of', []) if e['from'] == p.segment)
    p.category_alternatives = sorted({e['to'] for e in by_kind.get('includes', []) if e['from'] == p.category}
                                     & set(p.alternatives)) if p.category else []
    needs = sorted(a for a in p.atoms if _prefix(a) in NEEDS)
    degree = {(e['from'], e['to']): e.get('degree') for e in by_kind.get('fulfills', [])}
    p.missing_fulfillment = [[alt, need] for alt in p.alternatives for need in needs if (alt, need) not in degree]
    if p.missing_fulfillment or not p.alternatives:
        p.gaps = p.parity = None  # not derived: an unknown degree is never read as "does not meet"
    else:
        p.gaps = [n for n in needs if all(degree[(a, n)] != 'full' for a in p.alternatives)]
        p.parity = ([n for n in needs if all(degree[(a, n)] == 'full' for a in p.category_alternatives)]
                    if p.category_alternatives else None)
    body = {'atoms': {a: content(all_atoms[a]) for a in p.atoms},
            'edges': {k: content(all_edges[k]) for k in p.edges}}
    p.id = ID_PREFIX + sha256(body)


# ---- checks ----------------------------------------------------------------------------------------------------

SSOT_REL = 'design/info_arch/info_arch.yaml'


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


def check_positioning(m, report, approved, path=SSOT_REL):
    """Requirement 2 (shape) and requirement 5 (every approved positioning). -> {segment: Positioning}."""
    check_shape(m, report, path)
    derived = derive(m, approved)
    all_edges = edges(m)
    for seg, p in derived.items():
        on = [all_edges[k] for k in p.edges]
        touched = {e['from'] for e in on} | {e['to'] for e in on}
        for aid in p.atoms:
            if aid not in touched:
                report.add('fail', 'IA_POSITIONING_ISOLATED_ATOM', f'{seg}: 承認済みの {aid} がどの承認済みの矢印にも繋がらない', path, id=aid)
            if _prefix(aid) == 'job' and not any(e['kind'] == 'need' and e['from'] == aid for e in on):
                report.add('fail', 'IA_POSITIONING_JOB_WITHOUT_NEED', f'{seg}: ジョブ {aid} が痛みも得たいものも持たない', path, id=aid)
            if _prefix(aid) == 'str' and not any(e['kind'] == 'lacks' and e['from'] == aid for e in on):
                report.add('fail', 'IA_POSITIONING_STRENGTH_UNCOMPARED', f'{seg}: 強み {aid} がどの代替とも比べていない（lacks の矢印が無い）', path, id=aid)
        missing = [name for name, prefixes in (('ジョブ', {'job'}), ('痛みか得たいもの', NEEDS), ('提供価値', {'val'}))
                   if not any(_prefix(a) in prefixes for a in p.atoms)]
        if missing:
            report.add('fail', 'IA_POSITIONING_INCOMPLETE',
                       f'{seg}: 承認済みの {"・".join(missing)} が 0 件（中身の無い positioning を合格にしない）', path, segment=seg)
        if len(p.categories) != 1:
            report.add('fail', 'IA_POSITIONING_CATEGORY_COUNT',
                       f'{seg}: 所属カテゴリー（member_of）が {len(p.categories)} 本 {p.categories}（ちょうど 1 本）', path, segment=seg)
        if not p.alternatives:
            report.add('fail', 'IA_POSITIONING_NO_ALTERNATIVES', f'{seg}: 隙間の比較に使う承認済みの代替が 0 件', path, segment=seg)
        if p.category and not p.category_alternatives:
            report.add('fail', 'IA_POSITIONING_NO_CATEGORY_ALTERNATIVES',
                       f'{seg}: 所属カテゴリー {p.category} に入っている承認済みの代替が 0 件（同等であるべき点を導けない）', path, segment=seg)
        if p.missing_fulfillment:
            report.add('fail', 'IA_POSITIONING_FULFILLMENT_MISSING',
                       f'{seg}: 満たす度合いの承認済みの矢印が無い組がある {p.missing_fulfillment}（未調査を「満たさない」とみなさない）',
                       path, segment=seg, pairs=p.missing_fulfillment)
        for aid in p.atoms:
            if _prefix(aid) != 'val':
                continue
            kind = atoms(m)[aid].get('kind')
            points = p.gaps if kind == 'difference' else p.parity
            if points is None:
                continue  # not derivable: the missing input is reported above
            hit = sorted(e['to'] for e in on if e['kind'] == 'satisfies' and e['from'] == aid and e['to'] in points)
            if not hit:
                name = '隙間' if kind == 'difference' else '同等であるべき点'
                report.add('fail', 'IA_POSITIONING_VALUE_UNANCHORED',
                           f'{seg}: 提供価値 {aid}（{kind}）がどの{name}も満たさない（{name}: {points}）', path, id=aid)
    return derived


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
    """The owner's ia:positioning rows of design/approvals.yaml: {key: value_hash}.

    Validated by own-design-author's verify_rows (fields, key namespace, hash, approver, ISO date, no duplicate key)
    plus a symlink refusal. Any violation fails closed: IA_APPROVALS_INVALID and nothing counts as approved.
    """
    rel = 'design/approvals.yaml'
    path = root / rel
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
    return {k: v for k, v in approved.items() if k.startswith(APPROVAL_PREFIX)}
