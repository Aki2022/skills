"""The positioning SSOT (design/info_arch/info_arch.yaml, schema 5; the place is layer-paths.json's `positioning_file`): loading, id and ledger checks, approval reading, and
the readable form of what an approval request binds.

What is checked here is structure and references, never meaning. The positioning layer knows nothing of outlines:
whether a disqualified path stops anything depends on which outlines point at it (requirement 5), which is the
infoarch layer's check; it passes the answer in (`positioning.check_positioning(..., referenced)`).

Schema 4 trees (the format before the dictionary: facts, entities and a voice inside the SSOT) are still read, with a
warning, so a repository keeps working until it moves its fixed values and official names to docs/dictionary/
(`migrate-dictionary`). The lift is one function (`lift_v4`) and nothing else in the checks knows the old shape.
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

import positioning as pos
from ia_common import YamlError, load_yaml, validator_for  # first: it puts own-design-author/scripts on sys.path
import layer_paths  # noqa: E402

SCHEMA = 5
LEGACY_SCHEMA = 4
LEGACY_PREFIXES = ('fct', 'ent', 'vce')  # ids of the blocks schema 5 does not hold
LEGACY_TERM_BLOCKS = ('facts', 'entities')  # lifted into the dictionary view
LEGACY_DROPPED_BLOCKS = ('voices',)  # retired (SPEC-deliverable-production-layers requirement 8): read, never used

# Positioning (v6): segment, job, pain, gain, alternative, category, strength, value (+ evidence in proofs).
POSITIONING = dict(pos.ATOM_BLOCKS)
BLOCKS = {'proofs': 'prf', **POSITIONING}
PREFIX_BLOCK = {v: k for k, v in BLOCKS.items()}
ID_RE = re.compile(r'(' + '|'.join(BLOCKS.values()) + r')-((?!0000)[0-9]{4}|[1-9][0-9]{4,})')  # always fullmatch: $ would accept a trailing newline
SSOT_PATH = Path(layer_paths.rel('positioning_file'))  # repository-relative name; reading goes through layer_paths.path


@dataclass
class Model:
    root: Path
    ssot: dict | None = None

    def atom(self, aid):
        m = ID_RE.fullmatch(aid or '')
        return (self.ssot.get(PREFIX_BLOCK[m.group(1)]) or {}).get(aid) if m and self.ssot else None

    def tombstone(self, aid):
        return (self.ssot['ledger'].get('tombstones') or {}).get(aid, '__live__')

    def resolve(self, aid):
        """Follow tombstones to the live id (None when deleted). Bounded so a cycle cannot loop."""
        for _ in range(len(self.ssot['ledger'].get('tombstones') or {}) + 1):
            t = self.tombstone(aid)
            if t == '__live__' or t is None:
                return aid if t == '__live__' else None
            aid = t
        return None


def errors_of(validator, data):
    return sorted(validator.iter_errors(data), key=lambda e: list(e.absolute_path))


def lift_v4(data):
    """A schema 4 SSOT -> (the schema 5 shape, {'facts', 'entities'} for the dictionary view).

    Facts and entities are the dictionary's now; the voice is retired. Ids of those blocks are dropped from the
    ledger and from tombstones, so the positioning checks see only positioning ids."""
    out = {k: v for k, v in data.items() if k not in LEGACY_TERM_BLOCKS + LEGACY_DROPPED_BLOCKS}
    out['schema_version'] = SCHEMA
    ledger = data['ledger']
    out['ledger'] = {
        'next': {k: v for k, v in ledger['next'].items() if k not in LEGACY_PREFIXES},
        'tombstones': {k: v for k, v in (ledger.get('tombstones') or {}).items()
                       if k.split('-')[0] not in LEGACY_PREFIXES and (v is None or v.split('-')[0] not in LEGACY_PREFIXES)}}
    return out, {block: data.get(block) or {} for block in LEGACY_TERM_BLOCKS}


def load_ssot(root, report):
    """-> (ssot in the schema 5 shape or None, legacy terms or None). Failures go to the report (exit 2)."""
    ssot_file = layer_paths.path(Path(root), 'positioning_file')  # a repository that moved stops here, never "missing"
    if not ssot_file.is_file():
        report.add('fail', 'IA_MISSING', f'{SSOT_PATH} がない', SSOT_PATH)
        return None, None
    try:
        data = load_yaml(ssot_file)
    except YamlError as exc:
        report.add('fail', 'IA_YAML', f'YAML として読めない: {exc}', SSOT_PATH)
        return None, None
    legacy = isinstance(data, dict) and data.get('schema_version') == LEGACY_SCHEMA
    errs = errors_of(validator_for(f'info_arch.v{LEGACY_SCHEMA if legacy else SCHEMA}.schema.json'), data)
    for e in errs:
        report.add('fail', 'IA_SCHEMA', e.message, SSOT_PATH, pointer='/' + '/'.join(map(str, e.absolute_path)))
    if errs:
        return None, None
    if not legacy:
        return data, None
    report.add('warn', 'IA_SSOT_LEGACY_V4',
               f'{SSOT_PATH} はスキーマ版 4（固定値・正式名称・声を SSOT に持つ古い形）。固定値と正式名称は docs/dictionary/ の辞書へ移す'
               '（infoarch-author migrate-dictionary）。声は読まない（廃止）', SSOT_PATH)
    return lift_v4(data)


# ---- ids, ledger and references ------------------------------------------------------------------

def ssot_refs(block, atom):
    """Every id an SSOT atom references, as (field, id). One table, so no field is forgotten."""
    out = []
    single = {'proofs': ('segment',), **{b: ('segment',) for b in POSITIONING if b != 'segments'}}
    for fld in single.get(block, ()):
        if fld in atom:
            out.append((fld, atom[fld]))
    return out


def check_reference(m, report, src, aid, where, *, from_proposed, broken_code, empty_code):
    atom = m.atom(aid)
    if atom is None:
        t = m.tombstone(aid) if m.ssot else '__live__'
        hint = f'（墓標: {t} を使う）' if t not in ('__live__', None) else ''
        report.add('fail', broken_code, f'{src} が存在しない {aid} を参照している{hint}', where)
    elif from_proposed and atom['state'] == 'empty':
        report.add('fail', empty_code, f'{src} が空欄の {aid} を参照している', where)
    return atom


def check_ssot(m, report):
    d, path = m.ssot, str(SSOT_PATH)
    nxt, tomb = d['ledger']['next'], d['ledger']['tombstones']
    for key in [k for b in BLOCKS for k in d[b]] + list(tomb) + [v for v in tomb.values() if v]:
        if not ID_RE.fullmatch(key):
            report.add('fail', 'IA_SCHEMA', f'{key!r} は id の正規の書き方（接頭辞-0001〜9999、10000 以上は先頭 0 なし）でない', path)
    for block, prefix in BLOCKS.items():
        for aid, atom in d[block].items():
            if not ID_RE.fullmatch(aid):
                continue
            if int(ID_RE.fullmatch(aid).group(2)) >= nxt[prefix]:
                report.add('fail', 'IA_ID_NOT_ISSUED', f'{aid} は台帳の次番号 {nxt[prefix]} 以上で、台帳が発行していない', path)
            if aid in tomb:
                report.add('fail', 'IA_ID_TOMBSTONED', f'{aid} は墓標にある（id は再利用しない）', path)
            for fld, target in ssot_refs(block, atom):
                check_reference(m, report, f'{aid}.{fld}', target, path, from_proposed=atom['state'] == 'proposed',
                                broken_code='IA_REF_BROKEN', empty_code='IA_EMPTY_REFERENCED')
    for old, new in tomb.items():
        o = ID_RE.fullmatch(old)
        if o is None or (new is not None and not ID_RE.fullmatch(new)):
            continue
        if int(o.group(2)) >= nxt[o.group(1)]:
            report.add('fail', 'IA_ID_NOT_ISSUED', f'墓標 {old} は台帳が発行していない番号', path)
        if new is not None and m.atom(new) is None:
            report.add('fail', 'IA_TOMBSTONE_TARGET', f'墓標 {old} → {new} の先が無い', path)


# ---- the path of a value and its approval --------------------------------------------------------

def positioning_path(m, root_ids, segment):
    """The atoms and arrows the given values stand on in one segment, and the hash of their union. A value's path is
    its whole approval unit (requirement 3). -> (sorted keys, sha256 of their content); no value -> ([], None)."""
    ps = pos.paths(m)
    on = {}
    for v in root_ids:
        p = ps.get(v)
        if p is not None and p.segment == segment:
            on.update(p.roles)
    if not on:
        return [], None
    A, E = pos.atoms(m), pos.edges(m)
    ids = sorted(on)
    return ids, pos.sha256([[k, pos.content(A.get(k) or E[k])] for k in ids])


def element(m, key):
    """An atom or an arrow by its key (None when there is none)."""
    return m.atom(key) or pos.edges(m, proposed_only=False).get(key)


APPROVAL_PREFIX = pos.PATH_PREFIX

REVIEW_NAMES = {'seg': 'セグメント', 'job': 'ジョブ', 'pan': '痛み', 'gan': '得たいもの', 'alt': '代替',
                'cat': 'カテゴリー', 'str': '強み', 'val': '提供価値', 'prf': '証拠'}
REVIEW_FIELDS = {'text': '', 'qualifier': '見分ける条件', 'situation': '状況', 'kind': '種類', 'segment': 'セグメント',
                 'source': '出典', 'as_of': '時点', 'expires': '期限', 'copied_from': '写しの出どころ', 'provenance': '出所'}
EDGE_NAMES = {'has_job': 'ジョブを持つ', 'need': '痛み・得たいものを持つ', 'fulfills': '満たす度合い', 'recalls': '思い出す',
              'includes': '入れている', 'member_of': '所属', 'lacks': '代替に無い', 'produces': '生む',
              'satisfies': '満たす', 'supports': '支える'}
ROLE_NAMES = {'claim': '主張', 'basis': '差別化の根拠', 'premise': '前提', 'evidence': '証拠'}


def review_markdown(m, keys, hashes):
    """Markdown of the paths an approval request binds: each path whole, claim -> basis -> premises -> evidence, with
    the hash that binds it. What is shown is exactly what the hash covers (ADR-20261118-ia-approval-per-value-path)."""
    lines, ps = [], pos.paths(m)
    atoms, edges = pos.atoms(m, proposed_only=False), pos.edges(m, proposed_only=False)
    for key in sorted(keys):
        vid = key[len(APPROVAL_PREFIX):]
        p = ps.get(vid)
        if p is None:
            continue
        lines.append(f'## 提供価値 {vid}（{p.segment}・{p.kind}）　`{(hashes.get(key) or "")[:12]}`')
        lines.append(f'- 比べる集合：{"、".join(p.comparison) or "なし"}　満たす先：{"、".join(p.targets) or "なし"}')
        for code, message in p.problems:
            lines.append(f'- 失格：{code} {message}')
        for role in pos.ROLES:
            members = [k for k in p.members if p.roles[k] == role]
            if not members:
                continue
            lines.append(f'### {ROLE_NAMES[role]}（{len(members)}）')
            for k in members:
                if k in atoms:
                    a = atoms[k]
                    body = '／'.join(f'{label + "：" if label else ""}{"、".join(v) if isinstance(v, list) else v}'
                                    for fld, label in REVIEW_FIELDS.items() if (v := a.get(fld)) and fld != 'provenance')
                    lines.append(f'- {REVIEW_NAMES.get(pos._prefix(k), k)} {k}：{body}')
                else:
                    e = edges[k]
                    degree = f'（{e["degree"]}）' if 'degree' in e else ''
                    lines.append(f'- 矢印 {e["from"]} → {e["to"]}：{EDGE_NAMES.get(e["kind"], e["kind"])}{degree}')
        lines.append('')
    return '\n'.join(lines)


def read_approvals(root, report):
    """The owner's ia:path rows of design/approvals.yaml. Its full validation belongs to own-design-author."""
    return pos.read_approvals(Path(root), report, load_yaml, YamlError)
