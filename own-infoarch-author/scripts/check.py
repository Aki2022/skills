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


def supported_elements(m, approved):
    """Atoms and arrows that approved positioning evidence supports (through approved `supports` arrows)."""
    return {e['to'] for k, e in pos.edges(m).items()
            if e['kind'] == 'supports' and k in approved and e['from'] in approved and pos._prefix(e['from']) == 'prf'}


def check_lead_support(m, report, oid, rel, o, approved):
    """Warn when the lead message's value has nothing on its positioning path that evidence supports.

    Shape cannot tell whether a claim is true, but it can tell that the one message a deliverable leads with stands
    on nothing evidenced. Only the lead: warning on every message would turn the warning into noise.
    """
    lead = [msg for msg in o['messages'].values() if msg['rank'] == 'lead']
    if len(lead) != 1 or lead[0]['type'] == 'cta':
        return
    path, _ = positioning_path(m, [lead[0]['value']], ol.segment_of(o))
    if path and not set(path) & supported_elements(m, approved):
        report.add('warn', 'IA_LEAD_UNSUPPORTED',
                   f'{oid}: 一番のメッセージの提供価値 {lead[0]["value"]} の経路に、証拠が支える原子・矢印が無い（証拠を足すか、一番を替える）',
                   rel, id=lead[0]['value'])


# ---- positioning path (v6 requirements 2, 3 and 15) -----------------------------------------------------------

# How a message's positioning atoms pull in what they stand on: (atom prefix, arrow kind, direction, other end).
_PATH_RULES = (('val', 'satisfies', 'out'), ('val', 'produces', 'in'), ('str', 'lacks', 'out'),
               ('pan', 'need', 'in'), ('gan', 'need', 'in'), ('job', 'has_job', 'in'))


def positioning_path(m, root_ids, segment):
    """The positioning atoms and arrows a message's value stands on in one segment, and their hash.

    Derived from arrows alone, so a deliverable cannot leave anything out: a value pulls in the needs it satisfies
    and the strengths that produce it (with the alternatives they are compared against); a need pulls in its job;
    a job its segment; and the evidence that supports any of those comes with its `supports` arrow (a message
    inherits it, so it is approved and expires like the rest). Proposed elements only; approval is checked
    separately. Arrows into another segment are not followed. -> (sorted keys, sha256 of their content).
    An empty path hashes to None.
    """
    atoms, edges, segs = pos.atoms(m), pos.edges(m), pos.segment_map(m)
    def mine(aid):
        return aid in atoms and segs.get(aid) == segment
    roots = [r for r in root_ids if mine(r) and pos._prefix(r) in ('val', 'job', 'pan', 'gan')]
    if not roots:
        return [], None
    on, todo = set(), list(roots)
    while todo:
        cur = todo.pop()
        if cur in on:
            continue
        on.add(cur)
        for prefix, kind, direction in _PATH_RULES:
            if pos._prefix(cur) != prefix:
                continue
            for key, e in edges.items():
                end = e['to'] if direction == 'out' else e['from']
                if e['kind'] == kind and (e['from'] if direction == 'out' else e['to']) == cur and mine(end):
                    on.add(key)
                    todo.append(end)
    on.add(segment)
    for key, e in edges.items():
        if e['kind'] == 'supports' and e['to'] in on and mine(e['from']):
            on |= {key, e['from']}
    ids = sorted(on)
    body = [[k, pos.content(atoms.get(k) or edges[k])] for k in ids]
    return ids, pos.sha256(body)


# ---- owner approval per atom and arrow (v6 requirements 3 and 15) ---------------------------------------------

APPROVAL_PREFIX = pos.APPROVAL_PREFIX


def positioning_targets(m):
    """-> ({ia:positioning:<key>: hash} for every proposed atom and arrow, sorted keys the shipped paths stand on)."""
    hashes = pos.approval_targets(m)
    must = set()
    for _, (_, o) in m.outlines.items():
        ch = m.channels.get(o['channel'])
        if not (ch and ch[1]['ship']):
            continue
        for msg in o['messages'].values():
            must |= {APPROVAL_PREFIX + x for x in positioning_path(m, [msg['value']], ol.segment_of(o))[0]}
            must |= {APPROVAL_PREFIX + a for a in msg.get('compares', [])}
    return hashes, sorted(must & set(hashes))


REVIEW_NAMES = {'seg': 'セグメント', 'job': 'ジョブ', 'pan': '痛み', 'gan': '得たいもの', 'alt': '代替',
                'cat': 'カテゴリー', 'str': '強み', 'val': '提供価値', 'prf': '証拠'}
REVIEW_FIELDS = {'text': '', 'qualifier': '見分ける条件', 'situation': '状況', 'kind': '種類', 'segment': 'セグメント',
                 'source': '出典', 'as_of': '時点', 'expires': '期限', 'copied_from': '写しの出どころ', 'provenance': '出所'}
EDGE_NAMES = {'has_job': 'ジョブを持つ', 'need': '痛み・得たいものを持つ', 'fulfills': '満たす度合い', 'recalls': '思い出す',
              'includes': '入れている', 'member_of': '所属', 'lacks': '代替に無い', 'produces': '生む',
              'satisfies': '満たす', 'supports': '支える'}


def review_markdown(m, keys, hashes):
    """Markdown of the atoms and arrows an approval request binds, grouped by kind, with each binding hash."""
    lines, chosen = [], sorted(k[len(APPROVAL_PREFIX):] for k in keys)
    atoms, edges = pos.atoms(m, proposed_only=False), pos.edges(m, proposed_only=False)
    for prefix, name in REVIEW_NAMES.items():
        ids = [k for k in chosen if pos._prefix(k) == prefix and k in atoms]
        if not ids:
            continue
        lines.append(f'## {name}（{len(ids)}）')
        for aid in ids:
            lines.append(f'### {aid}　`{(hashes.get(APPROVAL_PREFIX + aid) or "")[:12]}`')
            for fld, label in REVIEW_FIELDS.items():
                v = atoms[aid].get(fld)
                if v:
                    lines.append(f'- {label + "：" if label else ""}{"、".join(v) if isinstance(v, list) else v}')
        lines.append('')
    arrows = [k for k in chosen if k in edges]
    if arrows:
        lines.append(f'## 矢印（{len(arrows)}）')
        for k in arrows:
            e = edges[k]
            degree = f'（{e["degree"]}）' if 'degree' in e else ''
            lines.append(f'- {e["from"]} → {e["to"]}：{EDGE_NAMES.get(e["kind"], e["kind"])}{degree}　'
                         f'`{(hashes.get(APPROVAL_PREFIX + k) or "")[:12]}`')
        lines.append('')
    return '\n'.join(lines)


def read_approvals(root, report):
    """The owner's ia:positioning rows of design/approvals.yaml. Its full validation belongs to own-design-author."""
    return pos.read_approvals(Path(root), report, load_yaml, YamlError)


def positioning_dependencies(m, o):
    """Every positioning element an outline stands on: the whole positioning of its target's segment. Its gaps and
    points of parity come from every alternative and fulfillment arrow there, so evidence behind any of them, and a
    change to any of them, reaches the outline (expiry, reverse lookup, maintenance)."""
    seg = ol.segment_of(o)
    if not seg:
        return set()
    segs = pos.segment_map(m)
    out = {k for k, s in segs.items() if s == seg and pos._prefix(k) != 'prf'} | {seg}
    # Evidence counts once it supports something there; an unconnected draft supports nothing yet.
    out |= {e['from'] for e in pos.edges(m).values() if e['kind'] == 'supports' and e['to'] in out}
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
    """Every element on the paths of an outline's values must carry the owner's current approval (requirement 15)."""
    missing = set()
    for msg in o['messages'].values():
        missing |= {x for x in positioning_path(m, [msg['value']], ol.segment_of(o))[0] if x not in approved}
    if missing:
        report.add('fail' if shipped else 'warn', 'IA_POSITIONING_UNAPPROVED',
                   f'{oid}: 立っている positioning に、オーナーの承認が無いか外れた原子・矢印がある {sorted(missing)}'
                   '（design-author approve --prepare で束を作り、オーナーが記録する）', rel, atoms=sorted(missing))


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
    pos.check_positioning(m, report, approved, str(SSOT_PATH))
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
        check_lead_support(m, report, oid, rel, o, approved)
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
