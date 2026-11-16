"""Deterministic checks of design/info_arch/ (token 0): the messaging SSOT, channel declarations, structured outlines.

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

import positioning as pos
from common import INFO_ARCH_DIR, Report, YamlError, load_ref, load_yaml, validator_for

# Positioning (v6): segment, job, pain, gain, alternative, category, strength, value (+ evidence in proofs).
POSITIONING = dict(pos.ATOM_BLOCKS)
BLOCKS = {'audiences': 'aud', 'messages': 'msg', 'proofs': 'prf', 'facts': 'fct',
          'entities': 'ent', 'voices': 'vce', **POSITIONING}
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
                                     ('outline', 'outlines', 'outline.v2.schema.json', m.outlines)):
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
    single = {'audiences': ('segment',), 'proofs': ('segment',), **{b: ('segment',) for b in POSITIONING if b != 'segments'}}
    many = {'audiences': ('priorities',), 'messages': ('addresses', 'serves'), 'proofs': ('supports',)}
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
    check_lead_support(m, report)
    voices = [v for v, a in d['voices'].items() if a['state'] == 'proposed']
    if len(voices) > 1:
        report.add('fail', 'IA_VOICE_MULTIPLE', f'proposed の声が {len(voices)} 個ある（声はプロジェクトに 1 つ）: {voices}', path)


NO_CLAIM_ROLES = {'cta', 'vision'}  # an invitation or an aspiration asserts nothing that needs proof


def check_lead_support(m, report):
    """Warn when an audience's lead message makes a claim that no proof supports.

    Shape and references cannot tell whether a message is true, but they can tell that the one message an audience
    hears first has nothing behind it. Only priorities[0]: later messages are often context, and warning on every
    message would turn the warning into noise.
    """
    d = m.ssot
    supported = {mid for p in d['proofs'].values() if p['state'] == 'proposed' for mid in p.get('supports', [])}
    for aid, aud in d['audiences'].items():
        if aud['state'] != 'proposed' or not aud.get('priorities'):
            continue
        lead = aud['priorities'][0]
        msg = m.atom(lead)
        if msg is None or msg['state'] != 'proposed' or msg.get('role') in NO_CLAIM_ROLES or lead in supported:
            continue
        report.add('warn', 'IA_PRIORITY_UNSUPPORTED',
                   f'{aid} の 1 位のメッセージ {lead}（{msg.get("role")}）を支える証拠が無い（証拠を足すか、主張を下げる）',
                   str(SSOT_PATH), id=lead)


# ---- positioning path (v6 requirements 2, 3 and 15; ISSUE-14 replaces the message side) ---------------------

# How a message's positioning atoms pull in what they stand on: (atom prefix, arrow kind, direction, other end).
_PATH_RULES = (('val', 'satisfies', 'out'), ('val', 'produces', 'in'), ('str', 'lacks', 'out'),
               ('pan', 'need', 'in'), ('gan', 'need', 'in'), ('job', 'has_job', 'in'))


def positioning_path(m, root_ids, segment):
    """The positioning atoms and arrows a message (or pain) stands on in one segment, and their hash.

    Derived from arrows alone, so a deliverable cannot leave anything out: a value pulls in the needs it satisfies
    and the strengths that produce it (with the alternatives they are compared against); a need pulls in its job;
    a job its segment. Proposed elements only; approval is checked separately. Arrows into another segment are not
    followed. -> (sorted keys, sha256 of their content). An empty path hashes to None.
    """
    if isinstance(root_ids, str):
        root_ids = _roots(m, root_ids)
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
    ids = sorted(on)
    body = [[k, pos.content(atoms.get(k) or edges[k])] for k in ids]
    return ids, pos.sha256(body)


# ---- owner approval per atom and arrow (v6 requirement 3; ISSUE-13 adds prepare, the contract and the ship report) --

APPROVAL_PREFIX = pos.APPROVAL_PREFIX


def positioning_targets(m):
    """-> ({ia:positioning:<key>: hash} for every proposed atom and arrow, sorted keys the shipped paths stand on)."""
    hashes = pos.approval_targets(m)
    must = set()
    for _, (_, o) in m.outlines.items():
        ch = m.channels.get(o['channel'])
        if not (ch and ch[1]['ship']):
            continue
        segment = (m.atom(o['audience']) or {}).get('segment')
        for p in o['pages']:
            for s in p['sections']:
                for key in ('message', 'pain'):
                    if s.get(key) and m.atom(s[key]) is not None:
                        must |= {APPROVAL_PREFIX + x for x in positioning_path(m, _roots(m, s[key]), segment)[0]}
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
    """Every positioning element an outline stands on: the paths of its messages and pains, and of its chosen atoms."""
    segment = (m.atom(o['audience']) or {}).get('segment')
    roots = [s[k] for p in o['pages'] for s in p['sections'] for k in ('message', 'pain') if s.get(k) and m.atom(s[k])]
    out = set()
    for r in roots:
        out |= set(positioning_path(m, _roots(m, r), segment)[0])
    for x in (o.get('strategy') or {}).get('choose', []):
        out |= set(positioning_path(m, [x], segment)[0])
    return out


def outline_dependencies(m, o):
    """Direct references plus the positioning underneath (requirement 14: reverse lookup is transitive)."""
    return {x for _, x in outline_refs(o)} | positioning_dependencies(m, o)


def _roots(m, aid):
    atom = m.atom(aid) or {}
    if aid.startswith('msg-'):
        return list(atom.get('serves', [])) + list(atom.get('addresses', []))
    return [aid]


def element(m, key):
    """An atom or an arrow by its key (None when there is none)."""
    return m.atom(key) or pos.edges(m, proposed_only=False).get(key)


def check_positioning_use(m, report, oid, rel, o, shipped, approved=None):
    """Every message and pain an outline uses must stand on the audience's segment, and every element of those
    paths must carry the owner's current approval (requirement 15)."""
    level = 'fail' if shipped else 'warn'
    segment = (m.atom(o['audience']) or {}).get('segment')
    missing = set()
    for p in o['pages']:
        for i, s in enumerate(p['sections']):
            w = f'{oid}.{p["id"]}[{i}]'
            for key in ('message', 'pain'):
                aid = s.get(key)
                if not aid or m.atom(aid) is None:
                    continue
                roots = _roots(m, aid)
                if not roots:
                    report.add(level, 'IA_POSITIONING_UNTRACED',
                               f'{w}: {aid} はどの positioning の原子にも立っていない（serves で提供価値・ジョブ・痛み・得たいものを結ぶ）', rel, id=aid)
                    continue
                ids, _ = positioning_path(m, roots, segment)
                if not ids:
                    report.add(level, 'IA_POSITIONING_SEGMENT_MISMATCH',
                               f'{w}: {aid} の positioning は対象者 {o["audience"]} のセグメント {segment} にたどり着かない', rel, id=aid)
                if approved is not None:
                    missing |= {x for x in ids if x not in approved}
    if missing:
        report.add(level, 'IA_POSITIONING_UNAPPROVED',
                   f'{oid}: 立っている positioning に、オーナーの承認が無いか外れた原子・矢印がある {sorted(missing)}'
                   '（design-author approve --prepare で束を作り、オーナーが記録する）', rel, atoms=sorted(missing))


# ---- composition (steps 2-4) -----------------------------------------------------------------------------------

def compose(m, audience, chosen):
    """Messages that stand on the chosen positioning atoms for the audience's segment, in the audience's priority
    order (then the rest, by id). The deterministic half of step 3; choosing and cutting stays with the author."""
    aud = m.atom(audience) or {}
    segment = aud.get('segment')
    chosen = set(chosen)
    picked = [mid for mid, a in sorted((m.ssot.get('messages') or {}).items())
              if a['state'] == 'proposed' and set(positioning_path(m, _roots(m, mid), segment)[0]) & chosen]
    order = {mid: i for i, mid in enumerate(aud.get('priorities') or [])}
    return sorted(picked, key=lambda mid: (order.get(mid, len(order)), mid))


def check_composition(m, report, oid, rel, o):
    """The recorded choice must stand on the audience's segment; messages should stand on it; dropping the
    audience's top message needs a recorded reason. (The outline field keeps the name `strategy` until ISSUE-14.)"""
    segment = (m.atom(o['audience']) or {}).get('segment')
    st = o['strategy']
    bad = [x for x in st['choose'] if m.atom(x) is None or not positioning_path(m, [x], segment)[0]]
    if bad:
        report.add('fail', 'IA_COMPOSITION_POSITIONING_INVALID',
                   f'{oid}: 選んだ positioning の原子 {bad} が無いか、対象者のセグメント {segment} の原子でない', rel)
        return
    chosen = set(st['choose'])
    for p in o['pages']:
        for i, s in enumerate(p['sections']):
            mid = s.get('message')
            if mid and m.atom(mid) is not None and _roots(m, mid):
                if not set(positioning_path(m, _roots(m, mid), segment)[0]) & chosen:
                    report.add('warn', 'IA_COMPOSITION_OFF_POSITIONING',
                               f'{oid}.{p["id"]}[{i}]: {mid} は選んだ positioning の原子 {sorted(chosen)} に立っていない', rel, id=mid)
    top = ((m.atom(o['audience']) or {}).get('priorities') or [None])[0]
    used = {s.get('message') for p in o['pages'] for s in p['sections']}
    cut = {m.resolve(x) for x in o['brief'].get('cut', [])} - {None}
    if top and top not in used and top not in cut and not st.get('deviation'):
        report.add('warn', 'IA_COMPOSITION_DEVIATION_UNEXPLAINED',
                   f'{oid}: 対象者の 1 位のメッセージ {top} を使っていない（理由を strategy.deviation に書くか、brief.cut で削ると決める）', rel)


# ---- frames (requirement 16) ------------------------------------------------------------

def frame_types(media):
    return (load_ref('media/frames.json')['media']).get(media)


def check_frames(m, report, oid, rel, o):
    """Declared frame instances: known media and type, unique ids, multiplicity, and element frames anchored once."""
    used = [(f'{oid}.{p["id"]}[{i}]', s['frame']) for p in o['pages'] for i, s in enumerate(p['sections']) if 'frame' in s]
    decl = o.get('frames')
    if decl is None:
        for w, fid in used:
            report.add('fail', 'IA_FRAME_UNDECLARED', f'{w}: 枠 {fid} を使っているが、アウトラインに frames が無い', rel)
        return
    types = frame_types(decl['media'])
    if types is None:
        report.add('fail', 'IA_FRAME_UNKNOWN_MEDIA', f'{oid}: 枠の媒体 {decl["media"]!r} はノウハウ（media/frames.json）に無い', rel)
        return
    kinds, seen = {}, set()
    for inst in decl['instances']:
        if inst['id'] in seen:
            report.add('fail', 'IA_FRAME_DUPLICATE_ID', f'{oid}: 枠の実体 id {inst["id"]} が重複している', rel)
            continue
        seen.add(inst['id'])
        if inst['type'] not in types:
            report.add('fail', 'IA_FRAME_UNKNOWN_TYPE', f'{oid}: 枠の種類 {inst["type"]!r} は {decl["media"]} のノウハウに無い', rel)
            continue
        kinds[inst['id']] = inst['type']
    for t, rule in types.items():
        n = sum(1 for v in kinds.values() if v == t)
        if n < rule['min'] or (rule['max'] is not None and n > rule['max']):
            want = f'{rule["min"]}〜{rule["max"] if rule["max"] is not None else ""}'
            report.add('fail', 'IA_FRAME_MULTIPLICITY', f'{oid}: 枠の種類 {t} の実体が {n} 個（{want} 個）', rel)
    counts = {}
    for w, fid in used:
        if fid not in kinds:
            report.add('fail', 'IA_FRAME_UNDECLARED', f'{w}: 枠 {fid} は frames.instances で宣言されていない', rel)
        elif types[kinds[fid]]['target'] == 'sequence':
            report.add('fail', 'IA_FRAME_SEQUENCE_ANCHORED', f'{w}: 並びの枠 {fid} はセクションに付けない（アウトラインの上だけで持つ）', rel)
        else:
            counts[fid] = counts.get(fid, 0) + 1
    for fid, t in kinds.items():
        if types[t]['target'] != 'element':
            continue
        if counts.get(fid, 0) > 1:
            report.add('fail', 'IA_FRAME_DUPLICATE_USE', f'{oid}: 枠の実体 {fid} を {counts[fid]} つのセクションに付けている（1 つだけ）', rel)
        elif not counts.get(fid):
            report.add('fail', 'IA_FRAME_UNUSED', f'{oid}: 宣言した枠の実体 {fid} がどのセクションにも付いていない', rel)


def mapping(m, oid):
    """The table an external experiment tool points at (requirement 17): version, frame instances, atoms and hashes."""
    rel, o = m.outlines[oid]
    def h(value):
        return hashlib.sha256(json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(',', ':')).encode()).hexdigest()
    def atoms_of(s):
        ids = [s[k] for k in ('pain', 'message') if k in s] + list(s.get('proofs', [])) + PLACEHOLDER.findall(s['copy'])
        return {aid: h([aid, m.atom(aid)]) for aid in ids}
    sections = [s for p in o['pages'] for s in p['sections']]
    decl = o.get('frames') or {'instances': []}
    rows = []
    for inst in decl['instances']:
        target = ((frame_types(decl.get('media')) or {}).get(inst['type']) or {}).get('target')
        if target == 'sequence':
            rows.append({'id': inst['id'], 'type': inst['type'],
                         'atoms': [s[k] for s in sections for k in ('pain', 'message') if k in s]})
        else:
            hit = [s for s in sections if s.get('frame') == inst['id']]
            rows.append({'id': inst['id'], 'type': inst['type'], 'atoms': atoms_of(hit[0]) if hit else {}})
    outputs = {}
    for out in o['outputs']:
        f = m.root / out
        outputs[out] = hashlib.sha256(f.read_bytes()).hexdigest() if f.is_file() else None
    standing = {k: element(m, k) for k in sorted(positioning_dependencies(m, o))}
    version = h({'outline': o, 'atoms': {aid: m.atom(aid) for s in sections for aid in atoms_of(s)},
                 'positioning': standing, 'outputs': outputs})
    return {'deliverable': oid, 'version_hash': version, 'outputs': outputs, 'strategy': o['strategy'], 'frames': rows}


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
    """Every id an outline references, as (where, id)."""
    out = [('audience', o['audience'])]
    out += [('brief.cut', x) for x in o['brief'].get('cut', [])]
    for p in o['pages']:
        for i, s in enumerate(p['sections']):
            w = f'{p["id"]}[{i}]'
            for key in ('message', 'pain'):
                if key in s:
                    out.append((f'{w}.{key}', s[key]))
            out += [(f'{w}.proofs', x) for x in s.get('proofs', [])]
            out += [(f'{w}.copy', x) for x in PLACEHOLDER.findall(s['copy'])]
    return out


TEXT_ROLES = {'navigation', 'other'}  # roles whose non-claim text is declared with a key instead of a message


def check_text_keys(report, oid, rel, o):
    """Declared non-claim text: a navigation / other section without a message or pain names a stable key, unique in
    the outline, so the factory can anchor it (data-ia-text) and dom can compare it; claims stay messages."""
    seen = {}
    for p in o['pages']:
        for i, s in enumerate(p['sections']):
            w = f'{oid}.{p["id"]}[{i}]'
            claim = 'message' in s or 'pain' in s
            if 'key' in s:
                if claim or s['role'] not in TEXT_ROLES:
                    report.add('fail', 'IA_OUTLINE_TEXT_KEY_MISPLACED',
                               f'{w}: key は message も pain も無い navigation・other のセクションだけに書く（主張はメッセージで書く）', rel)
                elif s['key'] in seen:
                    report.add('fail', 'IA_OUTLINE_TEXT_KEY_DUPLICATE', f'{w}: key {s["key"]!r} は {seen[s["key"]]} と重複している', rel)
                seen.setdefault(s['key'], w)
            elif not claim and s['role'] in TEXT_ROLES:
                report.add('fail', 'IA_OUTLINE_TEXT_KEY_MISSING',
                           f'{w}: 主張でない文言のセクションに key が無い（画面の印 data-ia-text を付けられない）', rel)


def check_outline(m, report, oid, rel, o):
    ch = m.channels.get(o['channel'])
    if ch is None:
        report.add('fail', 'IA_OUTLINE_CHANNEL', f'{oid}: チャネル {o["channel"]!r} の宣言（channels/{o["channel"]}.yaml）が無い', rel)
    else:
        globs = [_glob_re(g) for g in ch[1]['outputs']]
        for out in o['outputs']:
            if not any(g.match(out) for g in globs):
                report.add('fail', 'IA_OUTLINE_OUTPUT_UNDECLARED', f'{oid}: 成果物 {out} がチャネル {o["channel"]} の outputs {ch[1]["outputs"]} に一致しない', rel)
    check_text_keys(report, oid, rel, o)
    for where, aid in outline_refs(o):
        _ref(m, report, f'{oid}.{where}', aid, rel, from_proposed=True,
             broken_code='IA_OUTLINE_REF_BROKEN', empty_code='IA_OUTLINE_REF_EMPTY')
    cut = {m.resolve(x) for x in o['brief'].get('cut', [])} - {None}
    aud = m.atom(o['audience']) or {}
    priorities = set(aud.get('priorities', []))
    facts = {fid: f['value'] for fid, f in m.ssot['facts'].items() if f['state'] == 'proposed' and f.get('value')}
    for p in o['pages']:
        for i, s in enumerate(p['sections']):
            w = f'{oid}.{p["id"]}[{i}]'
            msg = s.get('message')
            if msg in cut:
                report.add('fail', 'IA_OUTLINE_CUT_USED', f'{w}: 削ると決めた {msg} を使っている', rel)
            elif msg and priorities and msg not in priorities:
                report.add('warn', 'IA_OUTLINE_OFF_PRIORITY', f'{w}: {msg} は対象者 {o["audience"]} の優先リストに無い', rel)
            for bad in [b for b in ANY_BRACES.findall(s['copy']) if not PLACEHOLDER.fullmatch(b)]:
                report.add('fail', 'IA_OUTLINE_BAD_PLACEHOLDER', f'{w}: {bad} は固定値の参照 {{{{fct-NNNN}}}} の形でない', rel)
            plain = PLACEHOLDER.sub(' ', s['copy'])
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
            for _, aid in outline_refs(o):
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
        check_outline(m, report, oid, rel, o)
        check_frames(m, report, oid, rel, o)
        check_composition(m, report, oid, rel, o)
        ch = m.channels.get(o['channel'])
        check_positioning_use(m, report, oid, rel, o, shipped=bool(ch and ch[1]['ship']), approved=approved)
    check_expiry(m, report, date)
    check_pages(m, report)
    report.targets = {'channels': len(m.channels), 'outlines': len(m.outlines),
                      'sections': sum(len(p['sections']) for _, o in m.outlines.values() for p in o['pages'])}
    return report, m


def impact(root, aid):
    report = Report('impact', dt.date.today().isoformat())
    m = load(root, report)
    if m.ssot is not None and m.atom(aid) is None and m.tombstone(aid) == '__live__':
        report.add('fail', 'IA_IMPACT_UNKNOWN_ID', f'{aid} は SSOT にも墓標にも無い（綴りを確かめる）')
    hits = sorted(oid for oid, (_, o) in m.outlines.items() if m.ssot is not None and aid in outline_dependencies(m, o))
    outputs = sorted({out for oid in hits for out in m.outlines[oid][1]['outputs']})
    return report, {'id': aid, 'outlines': hits, 'outputs': outputs}
