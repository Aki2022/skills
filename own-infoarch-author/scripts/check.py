"""Deterministic checks of design/info_arch/ (token 0): the positioning SSOT, channel declarations, outlines (v6).

What is checked here is structure and references, never meaning. Meaning (term variants, emphasis, whether a
sentence is a claim) belongs to the Jev judgement ledger; rendered artifacts belong to the DOM checks.
"""
from __future__ import annotations

import datetime as dt
import hashlib
import json
import re
from dataclasses import dataclass, field
from pathlib import Path

import outline as ol
import positioning as pos
from common import INFO_ARCH_DIR, Report, YamlError, load_ref, load_yaml, validator_for

# Positioning (v6): segment, job, pain, gain, alternative, category, strength, value (+ evidence in proofs).
POSITIONING = dict(pos.ATOM_BLOCKS)
BLOCKS = {'proofs': 'prf', 'facts': 'fct', 'entities': 'ent', 'voices': 'vce', **POSITIONING}
PREFIX_BLOCK = {v: k for k, v in BLOCKS.items()}
ID_RE = re.compile(r'(' + '|'.join(BLOCKS.values()) + r')-((?!0000)[0-9]{4}|[1-9][0-9]{4,})')  # always fullmatch: $ would accept a trailing newline
PLACEHOLDER = re.compile(r'\{\{(fct-(?:(?!0000)[0-9]{4}|[1-9][0-9]{4,}))\}\}')  # exactly {{fct-NNNN}}, no padding
ANY_BRACES = re.compile(r'\{\{.*?\}\}', re.S)
SSOT_PATH = INFO_ARCH_DIR / 'info_arch.yaml'


@dataclass
class Model:
    root: Path
    ssot: dict | None = None
    channels: dict = field(default_factory=dict)  # media -> (rel path, data)
    outlines: dict = field(default_factory=dict)  # id -> (rel path, data)

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


def _errors(validator, data):
    return sorted(validator.iter_errors(data), key=lambda e: list(e.absolute_path))


def _load(path, rel, report):
    try:
        return load_yaml(path), True
    except YamlError as exc:
        report.add('fail', 'IA_YAML', f'YAML として読めない: {exc}', rel)
        return None, False


def load(root, report):
    root = Path(root)
    base = root / INFO_ARCH_DIR
    m = Model(root)
    if not (root / SSOT_PATH).is_file():
        report.add('fail', 'IA_MISSING', f'{SSOT_PATH} がない', SSOT_PATH)
        return m
    data, ok = _load(root / SSOT_PATH, SSOT_PATH, report)
    if not ok:
        return m
    errs = _errors(validator_for('info_arch.v4.schema.json'), data)
    for e in errs:
        report.add('fail', 'IA_SCHEMA', e.message, SSOT_PATH, pointer='/' + '/'.join(map(str, e.absolute_path)))
    if errs:
        return m
    m.ssot = data
    for kind, sub, schema, store in (('channel', 'channels', 'channel.v2.schema.json', m.channels),
                                     ('outline', 'outlines', 'outline.v3.schema.json', m.outlines)):
        d = base / sub
        validator = validator_for(schema)
        for f in sorted(d.iterdir()) if d.is_dir() else []:
            rel = f.relative_to(root)
            if f.name.startswith('.') and f.is_file():
                continue  # editor and OS dotfiles (.DS_Store, .gitkeep) are not declarations; dot directories still fail
            if f.is_dir() or f.suffix != '.yaml':
                report.add('fail', f'IA_{kind.upper()}_FILE', f'{f.name} は読まれない（{sub}/<name>.yaml だけ。黙って無視しない）', rel)
                continue
            data, ok = _load(f, rel, report)
            if not ok:
                continue
            errs = _errors(validator, data)
            for e in errs:
                report.add('fail', 'IA_SCHEMA', e.message, rel, pointer='/' + '/'.join(map(str, e.absolute_path)))
            if errs:
                continue
            key = data['media'] if kind == 'channel' else data['id']
            if key != f.stem:
                code = 'IA_CHANNEL_NAME' if kind == 'channel' else 'IA_OUTLINE_ID'
                report.add('fail', code, f'{key!r} がファイル名 {f.stem!r} と違う', rel)
                continue
            store[key] = (rel, data)
    return m


# ---- SSOT --------------------------------------------------------------------------------

def ssot_refs(block, atom):
    """Every id an SSOT atom references, as (field, id). One table, so no field is forgotten."""
    out = []
    single = {'proofs': ('segment',), **{b: ('segment',) for b in POSITIONING if b != 'segments'}}
    many = {}
    for fld in single.get(block, ()):
        if fld in atom:
            out.append((fld, atom[fld]))
    for fld in many.get(block, ()):
        out += [(fld, x) for x in atom.get(fld, [])]
    return out


def _ref(m, report, src, aid, where, *, from_proposed, broken_code, empty_code):
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
                _ref(m, report, f'{aid}.{fld}', target, path, from_proposed=atom['state'] == 'proposed',
                     broken_code='IA_REF_BROKEN', empty_code='IA_EMPTY_REFERENCED')
    for old, new in tomb.items():
        o = ID_RE.fullmatch(old)
        if o is None or (new is not None and not ID_RE.fullmatch(new)):
            continue
        if int(o.group(2)) >= nxt[o.group(1)]:
            report.add('fail', 'IA_ID_NOT_ISSUED', f'墓標 {old} は台帳が発行していない番号', path)
        if new is not None and m.atom(new) is None:
            report.add('fail', 'IA_TOMBSTONE_TARGET', f'墓標 {old} → {new} の先が無い', path)
    voices = [v for v, a in d['voices'].items() if a['state'] == 'proposed']
    if len(voices) > 1:
        report.add('fail', 'IA_VOICE_MULTIPLE', f'proposed の声が {len(voices)} 個ある（声はプロジェクトに 1 つ）: {voices}', path)


def check_lead_support(m, report, oid, rel, o, all_paths):
    """Warn when the lead message's value has nothing on its path that evidence supports.

    Shape cannot tell whether a claim is true, but it can tell that the one message a deliverable leads with stands
    on nothing evidenced. Only the lead: warning on every message would turn the warning into noise.
    """
    lead = [msg for msg in o['messages'].values() if msg['rank'] == 'lead']
    if len(lead) != 1 or lead[0]['type'] == 'cta':
        return
    p = all_paths.get(lead[0]['value'])
    if p is not None and p.segment == ol.segment_of(o) and 'evidence' not in p.roles.values():
        report.add('warn', 'IA_LEAD_UNSUPPORTED',
                   f'{oid}: 一番のメッセージの提供価値 {lead[0]["value"]} の経路に、証拠が支える原子・矢印が無い（証拠を足すか、一番を替える）',
                   rel, id=lead[0]['value'])


# ---- the paths an outline points at (v6 requirements 3, 8 and 15) ---------------------------------------------

def referenced_values(m, o):
    """The values an outline's messages point at, inside its target's segment (a value of another segment is
    IA_OUTLINE_VALUE_OTHER_POSITIONING, and points at no path)."""
    seg, values = ol.segment_of(o), pos.atoms(m)
    return sorted({msg['value'] for msg in o['messages'].values()
                   if pos._prefix(msg['value']) == 'val' and (values.get(msg['value']) or {}).get('segment') == seg})


def referenced_paths(m):
    """{value id: True when a shipped outline points at its path, False when only draft outlines do}."""
    out = {}
    for _, (_, o) in m.outlines.items():
        ch = m.channels.get(o['channel'])
        shipped = bool(ch and ch[1]['ship'])
        for v in referenced_values(m, o):
            out[v] = out.get(v, False) or shipped
    return out


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


# ---- owner approval per value path (v6 requirements 3 and 15) -------------------------------------------------

APPROVAL_PREFIX = pos.PATH_PREFIX


def positioning_targets(m):
    """-> ({ia:path:<value>: path hash} for every proposed value, sorted keys of the paths shipped outlines use)."""
    hashes = pos.approval_targets(m)
    must = {APPROVAL_PREFIX + v for v, shipped in referenced_paths(m).items() if shipped}
    return hashes, sorted(must & set(hashes))


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


def positioning_dependencies(m, o):
    """Every positioning element an outline stands on: the paths of the values its messages point at (their gaps
    and points of parity come from the alternatives and degree arrows on those paths), so evidence behind any of
    them, and a change to any of them, reaches the outline (expiry, reverse lookup, maintenance)."""
    seg = ol.segment_of(o)
    if not seg:
        return set()
    ps = pos.paths(m)
    out = {seg}
    for v in referenced_values(m, o):
        out |= set(ps[v].roles)
    for msg in o['messages'].values():
        out |= {a for a in msg.get('compares', []) if m.atom(a) is not None}
    return out


def outline_dependencies(m, o):
    """Direct references plus the positioning underneath (requirement 14: reverse lookup is transitive)."""
    return {x for _, x in outline_refs(o)} | positioning_dependencies(m, o)


def element(m, key):
    """An atom or an arrow by its key (None when there is none)."""
    return m.atom(key) or pos.edges(m, proposed_only=False).get(key)


def check_positioning_use(m, report, oid, rel, o, shipped, approved):
    """Every path an outline's messages point at must carry the owner's current approval (requirement 15)."""
    missing = [v for v in referenced_values(m, o) if v not in approved]
    if missing:
        report.add('fail' if shipped else 'warn', 'IA_POSITIONING_UNAPPROVED',
                   f'{oid}: 指している提供価値 {missing} の経路に、オーナーの承認が無いか外れている'
                   '（design-author approve --prepare で束を作り、オーナーの承認の言葉で記録する）', rel, values=missing)


# ---- outlines ----------------------------------------------------------------------------

def _glob_re(pattern):
    out, i = [], 0
    while i < len(pattern):
        if pattern.startswith('**/', i):
            out.append('(?:.*/)?'); i += 3
        elif pattern.startswith('**', i):
            out.append('.*'); i += 2
        elif pattern[i] == '*':
            out.append('[^/]*'); i += 1
        elif pattern[i] == '?':
            out.append('[^/]'); i += 1
        else:
            out.append(re.escape(pattern[i])); i += 1
    return re.compile(''.join(out) + r'\Z')


def outline_refs(o):
    """Every SSOT id an outline references, as (where, id)."""
    out = []
    for mid, msg in sorted(o['messages'].items()):
        out.append((f'messages.{mid}.value', msg['value']))
        out += [(f'messages.{mid}.compares', a) for a in msg.get('compares', [])]
    for where, text in outline_texts(o):
        out += [(where, f) for f in PLACEHOLDER.findall(text)]
    return out


def outline_texts(o):
    """(where, text) of every piece of copy and declared text an outline writes."""
    return ([(f'copies.{cid}', cp['text']) for cid, cp in sorted(o['copies'].items())]
            + [(f'texts.{k}', t['text']) for k, t in sorted((o.get('texts') or {}).items())])


def check_outline(m, report, oid, rel, o):
    """Placement (channel and outputs), references into the SSOT, and fact placeholders in the copy."""
    ch = m.channels.get(o['channel'])
    if ch is None:
        report.add('fail', 'IA_OUTLINE_CHANNEL', f'{oid}: チャネル {o["channel"]!r} の宣言（channels/{o["channel"]}.yaml）が無い', rel)
    else:
        globs = [_glob_re(g) for g in ch[1]['outputs']]
        for out in o['outputs']:
            if not any(g.match(out) for g in globs):
                report.add('fail', 'IA_OUTLINE_OUTPUT_UNDECLARED', f'{oid}: 成果物 {out} がチャネル {o["channel"]} の outputs {ch[1]["outputs"]} に一致しない', rel)
    for where, aid in outline_refs(o):
        _ref(m, report, f'{oid}.{where}', aid, rel, from_proposed=True,
             broken_code='IA_OUTLINE_REF_BROKEN', empty_code='IA_OUTLINE_REF_EMPTY')
    facts = {fid: f['value'] for fid, f in m.ssot['facts'].items() if f['state'] == 'proposed' and f.get('value')}
    for where, text in outline_texts(o):
        w = f'{oid}.{where}'
        for bad in [b for b in ANY_BRACES.findall(text) if not PLACEHOLDER.fullmatch(b)]:
            report.add('fail', 'IA_OUTLINE_BAD_PLACEHOLDER', f'{w}: {bad} は固定値の参照 {{{{fct-NNNN}}}} の形でない', rel)
        plain = PLACEHOLDER.sub(' ', text)
        for fid, value in facts.items():
            if value in plain:
                report.add('fail', 'IA_OUTLINE_FACT_LITERAL', f'{w}: 固定値 {fid} の値「{value}」を直接書いている（{{{{{fid}}}}} で参照する）', rel)


# ---- design page inventory (ADR-20261115-design-ia-two-skills-shared-base) -------------------

HTML_SUFFIXES = ('.html', '.htm')


def check_pages(m, report):
    """Shipped HTML outputs against design/validation.json pages: the two declarations are kept, but must agree."""
    path = m.root / 'design/validation.json'
    if not path.is_file():
        return
    try:
        pages = json.loads(path.read_text(encoding='utf-8')).get('pages')
    except (ValueError, AttributeError):
        return  # design's own check reports a malformed validation.json
    if not isinstance(pages, dict):
        return
    covered, excluded = pages.get('covered') or {}, pages.get('excluded') or {}
    for oid, (rel, o) in sorted(m.outlines.items()):
        ch = m.channels.get(o['channel'])
        if not (ch and ch[1]['ship']):
            continue
        for out in o['outputs']:
            if not out.lower().endswith(HTML_SUFFIXES):
                continue
            if out in excluded:
                report.add('info', 'IA_PAGES_OUTPUT_EXCLUDED_BY_DESIGN',
                           f'{oid}: {out} は design の棚卸しで対象外だが、IA の検査（check・dom・判定）は続ける', rel)
            elif out not in covered:
                report.add('warn', 'IA_PAGES_OUTPUT_NOT_INVENTORIED',
                           f'{oid}: 出荷される {out} が design/validation.json の pages（covered・excluded）に無い。design の検査が見ていない', rel)


# ---- expiry ------------------------------------------------------------------------------

def _expiry(atom, days):
    if atom.get('evergreen'):
        return None
    if atom.get('expires'):
        return dt.date.fromisoformat(atom['expires'])
    if atom.get('as_of'):
        return dt.date.fromisoformat(atom['as_of']) + dt.timedelta(days=days)
    return None


def check_expiry(m, report, date):
    today = dt.date.fromisoformat(date)
    days = {**load_ref('expiry.json')['expiry_days'], **((m.ssot.get('policy') or {}).get('expiry_days') or {})}
    shipped = {}
    for oid, (rel, o) in m.outlines.items():
        ch = m.channels.get(o['channel'])
        if ch and ch[1]['ship']:
            for aid in outline_dependencies(m, o):
                shipped.setdefault(aid, []).append(oid)
    for block, kind in (('proofs', 'proof'), ('facts', 'fact')):
        for aid, atom in m.ssot[block].items():
            if atom['state'] != 'proposed':
                continue
            exp = _expiry(atom, days[kind])
            if exp is None or exp >= today:
                continue
            if aid in shipped:
                report.add('fail', 'IA_EXPIRED_SHIPPED', f'{aid} の期限 {exp} が過ぎているのに、出荷される成果物 {sorted(shipped[aid])} が参照している',
                           SSOT_PATH, id=aid, outlines=sorted(shipped[aid]))
            else:
                report.add('info', 'IA_EXPIRED_UNSHIPPED', f'{aid} の期限 {exp} が過ぎている（出荷される成果物からは参照されていない）', SSOT_PATH, id=aid)


# ---- entry points ------------------------------------------------------------------------

def check(root, date):
    report = Report('check', date)
    m = load(root, report)
    if m.ssot is None:
        return report, m
    check_ssot(m, report)
    approved = pos.approved_keys(m, read_approvals(root, report))
    _, all_paths = pos.check_positioning(m, report, approved, referenced_paths(m), str(SSOT_PATH))
    if not m.channels:
        report.add('fail', 'IA_NO_CHANNELS', 'channels/*.yaml が 0 件（対象 0 件を合格にしない）', INFO_ARCH_DIR / 'channels')
    if not m.outlines:
        report.add('fail', 'IA_NO_OUTLINES', 'outlines/*.yaml が 0 件（対象 0 件を合格にしない）', INFO_ARCH_DIR / 'outlines')
    for oid, (rel, o) in sorted(m.outlines.items()):
        ch = m.channels.get(o['channel'])
        shipped = bool(ch and ch[1]['ship'])
        check_outline(m, report, oid, rel, o)
        ol.check(m, report, oid, rel, o, shipped, approved)
        check_positioning_use(m, report, oid, rel, o, shipped, approved)
        check_lead_support(m, report, oid, rel, o, all_paths)
    check_expiry(m, report, date)
    check_pages(m, report)
    report.targets = {'channels': len(m.channels), 'outlines': len(m.outlines),
                      'sections': sum(len(o['sections']) for _, o in m.outlines.values())}
    return report, m


def impact(root, aid):
    report = Report('impact', dt.date.today().isoformat())
    m = load(root, report)
    if m.ssot is not None and m.atom(aid) is None and m.tombstone(aid) == '__live__':
        report.add('fail', 'IA_IMPACT_UNKNOWN_ID', f'{aid} は SSOT にも墓標にも無い（綴りを確かめる）')
    hits = sorted(oid for oid, (_, o) in m.outlines.items() if m.ssot is not None and aid in outline_dependencies(m, o))
    outputs = sorted({out for oid in hits for out in m.outlines[oid][1]['outputs']})
    return report, {'id': aid, 'outlines': hits, 'outputs': outputs}
