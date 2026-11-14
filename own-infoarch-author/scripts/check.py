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

from common import INFO_ARCH_DIR, Report, YamlError, load_ref, load_yaml, validator_for

# Strategy layer (v5): segment, job, pain, gain, value, alternative, strength, concept, promise.
STRATEGY = {'segments': 'seg', 'jobs': 'job', 'pains': 'pan', 'gains': 'gan', 'values': 'val',
            'alternatives': 'alt', 'strengths': 'str', 'concepts': 'cpt', 'promises': 'prm'}
BLOCKS = {'audiences': 'aud', 'messages': 'msg', 'proofs': 'prf', 'facts': 'fct',
          'entities': 'ent', 'voices': 'vce', **STRATEGY}
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
    errs = _errors(validator_for('info_arch.v3.schema.json'), data)
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
    single = {'audiences': ('segment',), 'jobs': ('segment',), 'pains': ('job',), 'gains': ('job',),
              'alternatives': ('segment',)}
    many = {'audiences': ('priorities',), 'messages': ('addresses', 'serves'), 'proofs': ('supports',),
            'values': ('relieves', 'creates'), 'strengths': ('versus', 'produces'), 'promises': ('values',)}
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


# ---- strategy path (requirements 13 and 15) ----------------------------------------------

def _live(m, block, aid):
    atom = (m.ssot.get(block) or {}).get(aid)
    return atom if atom and atom['state'] == 'proposed' else None


def strategy_path(m, root_ids, segment):
    """The strategy atoms a message (or pain) stands on for one segment, and their hash.

    Derived from references only, so a deliverable cannot leave anything out: a job is on the path when it
    belongs to the segment; a pain or gain when its job is; a value when it relieves or creates one of those;
    a promise when it states such a value; every strength that produces such a value, with the alternatives of
    the same segment it is compared against. Branches that lead to another segment are not followed.
    -> (sorted ids, sha256 of the canonical JSON of those atoms). An empty path hashes to None.
    """
    if isinstance(root_ids, str):
        root_ids = _roots(m, root_ids)
    d = m.ssot
    jobs = {j for j, a in (d.get('jobs') or {}).items() if a['state'] == 'proposed' and a.get('segment') == segment}
    needs = {b: {k for k, a in (d.get(b) or {}).items() if a['state'] == 'proposed' and a.get('job') in jobs}
             for b in ('pains', 'gains')}
    values = {v for v, a in (d.get('values') or {}).items() if a['state'] == 'proposed'
              and (set(a.get('relieves', [])) & needs['pains'] or set(a.get('creates', [])) & needs['gains'])}
    promises = {p for p, a in (d.get('promises') or {}).items() if a['state'] == 'proposed' and set(a.get('values', [])) & values}
    reachable = jobs | needs['pains'] | needs['gains'] | values | promises
    roots = [r for r in root_ids if r in reachable]
    if not roots:
        return [], None
    on = set()
    for r in roots:
        prefix = r.split('-')[0]
        if prefix == 'prm':
            vals = set(d['promises'][r].get('values', [])) & values
            on |= {r} | vals
        elif prefix == 'val':
            vals = {r}
            on |= vals
        else:
            vals = set()
            on.add(r)
        for v in vals:
            a = d['values'][v]
            on |= (set(a.get('relieves', [])) & needs['pains']) | (set(a.get('creates', [])) & needs['gains'])
    for k in [x for x in on if x.split('-')[0] in ('pan', 'gan')]:
        on.add(d['pains' if k.startswith('pan') else 'gains'][k]['job'])
    on_values = {x for x in on if x.startswith('val-')}
    for sid, a in (d.get('strengths') or {}).items():
        if a['state'] == 'proposed' and set(a.get('produces', [])) & on_values:
            on.add(sid)
            on |= {x for x in a.get('versus', []) if (_live(m, 'alternatives', x) or {}).get('segment') == segment}
    on.add(segment)
    ids = sorted(on)
    body = [[aid, m.atom(aid)] for aid in ids]
    digest = hashlib.sha256(json.dumps(body, ensure_ascii=False, sort_keys=True, separators=(',', ':')).encode()).hexdigest()
    return ids, digest


def _roots(m, aid):
    atom = m.atom(aid) or {}
    if aid.startswith('msg-'):
        return list(atom.get('serves', [])) + list(atom.get('addresses', []))
    return [aid]


def check_strategy(m, report, oid, rel, o, shipped):
    """Every message and pain an outline uses must stand on the audience's segment (requirement 15)."""
    level = 'fail' if shipped else 'warn'
    segment = (m.atom(o['audience']) or {}).get('segment')
    for p in o['pages']:
        for i, s in enumerate(p['sections']):
            w = f'{oid}.{p["id"]}[{i}]'
            for key in ('message', 'pain'):
                aid = s.get(key)
                if not aid or m.atom(aid) is None:
                    continue
                roots = _roots(m, aid)
                if not roots:
                    report.add(level, 'IA_STRATEGY_UNTRACED',
                               f'{w}: {aid} はどの戦略の原子も担っていない（serves で提供価値・約束・ジョブ・痛み・得たいものを結ぶ）', rel, id=aid)
                    continue
                ids, _ = strategy_path(m, roots, segment)
                if not ids:
                    report.add(level, 'IA_STRATEGY_SEGMENT_MISMATCH',
                               f'{w}: {aid} の戦略は対象者 {o["audience"]} のセグメント {segment} にたどり着かない', rel, id=aid)


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


def check_outline(m, report, oid, rel, o):
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
    if not m.channels:
        report.add('fail', 'IA_NO_CHANNELS', 'channels/*.yaml が 0 件（対象 0 件を合格にしない）', INFO_ARCH_DIR / 'channels')
    if not m.outlines:
        report.add('fail', 'IA_NO_OUTLINES', 'outlines/*.yaml が 0 件（対象 0 件を合格にしない）', INFO_ARCH_DIR / 'outlines')
    for oid, (rel, o) in sorted(m.outlines.items()):
        check_outline(m, report, oid, rel, o)
        ch = m.channels.get(o['channel'])
        check_strategy(m, report, oid, rel, o, shipped=bool(ch and ch[1]['ship']))
    check_expiry(m, report, date)
    report.targets = {'channels': len(m.channels), 'outlines': len(m.outlines),
                      'sections': sum(len(p['sections']) for _, o in m.outlines.values() for p in o['pages'])}
    return report, m


def impact(root, aid):
    report = Report('impact', dt.date.today().isoformat())
    m = load(root, report)
    if m.ssot is not None and m.atom(aid) is None and m.tombstone(aid) == '__live__':
        report.add('fail', 'IA_IMPACT_UNKNOWN_ID', f'{aid} は SSOT にも墓標にも無い（綴りを確かめる）')
    hits = sorted(oid for oid, (_, o) in m.outlines.items() if aid in {x for _, x in outline_refs(o)})
    outputs = sorted({out for oid in hits for out in m.outlines[oid][1]['outputs']})
    return report, {'id': aid, 'outlines': hits, 'outputs': outputs}
