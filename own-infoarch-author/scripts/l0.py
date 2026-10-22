"""L0: consistency of design/info_arch/ itself (no artifacts, no model, token 0).

Scope of this module (ISSUE-01a): schema, id ledger and tombstones, references, preset musts,
claim expiry and scope, raw numbers in renditions, terms.prh.yml generated section and specs.
Signed approvals, judgement ledger, manifests and ship gates belong to the shared runner.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

import numeric
import prh
from common import INFO_ARCH_DIR, Report, YamlError, load_ref, load_yaml, safe_load_text, validator_for

KINDS = {'audiences': 'aud', 'problems': 'prb', 'messages': 'msg', 'claims': 'clm', 'entities': 'ent'}
PLACEHOLDER = re.compile(r'\{\{\s*(clm-[0-9]{4,})\s*\}\}')
ID_RE = re.compile(r'^(aud|prb|msg|clm|trm|ent)-([0-9]{4,})$')


@dataclass
class Model:
    root: Path
    data: dict | None = None
    channels: dict = field(default_factory=dict)  # name -> (path, data)
    terms_text: str | None = None
    rules: list = field(default_factory=list)

    def atom(self, aid):
        for block, prefix in KINDS.items():
            if aid.startswith(prefix + '-'):
                return (self.data.get(block) or {}).get(aid)
        return None

    def tombstone(self, aid):
        return (self.data['ledger'].get('tombstones') or {}).get(aid, '__live__')


def _schema_errors(validator, data):
    return sorted(validator.iter_errors(data), key=lambda e: list(e.absolute_path))


def load(root, report):
    root = Path(root)
    base = root / INFO_ARCH_DIR
    m = Model(root)
    main = base / 'info_arch.yaml'
    if not main.is_file():
        report.add('fail', 'IA_MISSING', 'design/info_arch/info_arch.yaml がない', main.relative_to(root))
        return m
    try:
        data = load_yaml(main)
    except YamlError as exc:
        report.add('fail', 'IA_YAML', f'YAML として読めない: {exc}', 'design/info_arch/info_arch.yaml')
        return m
    errs = _schema_errors(validator_for('info_arch.v1.schema.json'), data)
    for e in errs:
        report.add('fail', 'IA_SCHEMA', e.message, 'design/info_arch/info_arch.yaml',
                   pointer='/' + '/'.join(map(str, e.absolute_path)))
    if errs:
        return m
    m.data = data
    chan_validator = validator_for('channel.v1.schema.json')
    chan_dir = base / 'channels'
    files = sorted(chan_dir.glob('*.yaml')) if chan_dir.is_dir() else []
    if not files:
        report.add('fail', 'IA_NO_CHANNELS', 'channels/*.yaml が 0 件（検査対象が無いことを合格にしない）', 'design/info_arch/channels')
    for f in files:
        rel = f.relative_to(root)
        try:
            cdata = load_yaml(f)
        except YamlError as exc:
            report.add('fail', 'IA_YAML', f'YAML として読めない: {exc}', rel)
            continue
        cerrs = _schema_errors(chan_validator, cdata)
        for e in cerrs:
            report.add('fail', 'IA_SCHEMA', e.message, rel, pointer='/' + '/'.join(map(str, e.absolute_path)))
        if not cerrs:
            if cdata['media'] != f.stem:
                report.add('fail', 'IA_CHANNEL_NAME', f'media {cdata["media"]!r} がファイル名と違う', rel)
            m.channels[f.stem] = (rel, cdata)
    terms = base / 'terms.prh.yml'
    if terms.is_file():
        try:
            m.terms_text = terms.read_bytes().decode('utf-8')
        except UnicodeDecodeError as exc:
            report.add('fail', 'IA_YAML', f'UTF-8 として読めない: {exc}', 'design/info_arch/terms.prh.yml')
    return m


def _check_ledger(m, report):
    nxt, tomb = m.data['ledger']['next'], m.data['ledger']['tombstones']
    for block, prefix in KINDS.items():
        for aid in m.data[block] or {}:
            n = int(ID_RE.match(aid).group(2))
            if n >= nxt[prefix]:
                report.add('fail', 'IA_ID_NOT_ISSUED', f'{aid} は台帳の次番号 {nxt[prefix]} 以上で、台帳が発行していない', 'design/info_arch/info_arch.yaml')
            if aid in tomb:
                report.add('fail', 'IA_ID_TOMBSTONED', f'{aid} は墓標にある（id は再利用しない）', 'design/info_arch/info_arch.yaml')
    for old, new in tomb.items():
        o = ID_RE.match(old)
        if int(o.group(2)) >= nxt[o.group(1)]:
            report.add('fail', 'IA_ID_NOT_ISSUED', f'墓標 {old} は台帳が発行していない番号', 'design/info_arch/info_arch.yaml')
        if new is None:
            continue
        n = ID_RE.match(new)
        if n.group(1) == 'trm':  # terms live in terms.prh.yml; the ledger can only vouch that the id was issued
            if int(n.group(2)) >= nxt['trm']:
                report.add('fail', 'IA_TOMBSTONE_TARGET', f'墓標 {old} → {new} の先は台帳が発行していない', 'design/info_arch/info_arch.yaml')
        elif m.atom(new) is None:
            report.add('fail', 'IA_TOMBSTONE_TARGET', f'墓標 {old} → {new} の先が無い', 'design/info_arch/info_arch.yaml')


def _ref(m, report, src, aid, where, *, from_proposed=False):
    atom = m.atom(aid)
    if atom is not None:
        if from_proposed and atom.get('state') == 'empty':
            report.add('fail', 'IA_EMPTY_REFERENCED', f'proposed の {src} が空欄の {aid} を参照している', where)
        return atom
    t = m.tombstone(aid)
    hint = f'（墓標: {t} を使う）' if t not in ('__live__', None) else ''
    report.add('fail', 'IA_REF_BROKEN', f'{src} が存在しない {aid} を参照している{hint}', where)
    return None


def atom_refs(block, atom):
    """Every id an atom references, as (field, id). One table for all reference fields, so a new field
    cannot be checked for broken references but forgotten for empty-target references."""
    out = []
    if block in ('problems', 'messages') and 'audience' in atom:
        out.append(('audience', atom['audience']))
    if block == 'problems':
        out += [('evidence_claims', c) for c in atom.get('evidence_claims', [])]
    if block == 'messages':
        out += [('addresses', x) for x in atom.get('addresses', [])]
        out += [('supports', c) for c in atom.get('supports', [])]
        for media, text in (atom.get('renditions') or {}).items():
            out += [(f'renditions.{media}', c) for c in PLACEHOLDER.findall(text)]
    if block == 'claims':
        out += [('scope.audiences', a) for a in (atom.get('scope') or {}).get('audiences', [])]
    return out


def _check_refs(m, report, date):
    path = 'design/info_arch/info_arch.yaml'
    d = m.data
    for block in ('problems', 'messages', 'claims'):
        for aid, atom in d[block].items():
            proposed = atom['state'] == 'proposed'
            for fld, target in atom_refs(block, atom):
                ref = _ref(m, report, f'{aid}.{fld}', target, path, from_proposed=proposed)
                if proposed and fld in ('supports', 'evidence_claims'):
                    _scope(report, aid, atom.get('audience'), target, ref, path)
    for mid, msg in d['messages'].items():
        for media, text in (msg.get('renditions') or {}).items():
            raw = PLACEHOLDER.sub(' ', text)
            toks = numeric.candidate_tokens(raw)
            if toks:
                report.add('fail', 'IA_RENDITION_RAW_NUMBER',
                           f'{mid}.renditions.{media} に生の数値 {[t.text for t in toks]}（主張は {{{{clm-…}}}} で参照する）', path)
    for cid, c in d['claims'].items():
        if c.get('expires') and c['expires'] < date:
            report.add('fail', 'IA_CLAIM_EXPIRED', f'{cid} の期限 {c["expires"]} が評価日 {date} を過ぎている', path)
        if 'value' in c and c['value']['unit'] not in _canonical_units():
            report.add('fail', 'IA_CLAIM_UNIT', f'{cid} の単位 {c["value"]["unit"]!r} が正規化表の正規単位に無い', path)
    for name, (rel, ch) in m.channels.items():
        for x in ch['messages']:
            atom = _ref(m, report, f'channels/{name}', x, rel)
            if atom is not None and atom['state'] == 'empty':
                report.add('fail', 'IA_EMPTY_REFERENCED', f'channels/{name} が空欄の {x} を使っている', rel)
        for x in ch.get('required_claims', []):
            atom = _ref(m, report, f'channels/{name}', x, rel)
            if atom is not None and atom['state'] == 'empty':
                report.add('fail', 'IA_EMPTY_REFERENCED', f'channels/{name} が空欄の {x} を必須にしている', rel)


def _canonical_units():
    return {''} | {u['unit'] for u in load_ref('numeric.json')['units']}


def _scope(report, src, audience, cid, claim, path):
    if claim is None or audience is None or claim.get('state') != 'proposed':
        return
    scope = (claim.get('scope') or {}).get('audiences', [])
    if audience not in scope:
        report.add('fail', 'IA_CLAIM_SCOPE', f'{src}（対象者 {audience}）が適用範囲外の {cid}（{scope}）を使っている', path)


def _check_presets(m, report):
    presets = load_ref('presets.json')['presets']
    for name, (rel, ch) in m.channels.items():
        preset = presets[ch['preset']]
        for block in preset['required_blocks']:
            if not any(a['state'] == 'proposed' for a in m.data[block].values()):
                report.add('fail', 'IA_PRESET_BLOCK', f'プリセット {ch["preset"]} の必須ブロック {block} に中身が無い', rel)
        roles = {m.data['messages'][x].get('role') for x in ch['messages']
                 if x in m.data['messages'] and m.data['messages'][x]['state'] == 'proposed'}
        for role in preset['required_roles']:
            if role not in roles:
                report.add('fail', 'IA_PRESET_ROLE', f'channels/{name} に必須役割 {role} のメッセージが無い', rel)


def _check_terms(m, report):
    rel = 'design/info_arch/terms.prh.yml'
    if m.terms_text is None:
        report.add('fail', 'IA_TERMS_MISSING', 'terms.prh.yml がない', rel)
        return
    block = prh.generated_block(m.terms_text)
    want = prh.render_generated(m.data['entities'])
    if block is None:
        report.add('fail', 'IA_TERMS_MARKERS', '生成区間のマーカーが無いか重複している（gen-terms で作る）', rel)
    elif block != want:
        report.add('fail', 'IA_TERMS_GENERATED_DRIFT', '生成区間が entity から再生成した内容と違う（手編集か entity の変更。gen-terms で直す）', rel)
    try:
        m.rules = prh.load_rules(safe_load_text(m.terms_text), where=rel)
    except YamlError as exc:
        report.add('fail', 'IA_YAML', f'YAML として読めない: {exc}', rel)
        return
    except prh.PrhError as exc:
        report.add('fail', 'IA_TERMS_INVALID', str(exc), rel)
        return
    for f in prh.run_specs(m.rules):
        report.add('fail', 'IA_TERMS_SPEC', f'specs が通らない: {f}', rel)


def _info_unreferenced(m, report):
    used = set()
    for msg in m.data['messages'].values():
        used.update(msg.get('supports', []))
        for text in (msg.get('renditions') or {}).values():
            used.update(PLACEHOLDER.findall(text))
    for p in m.data['problems'].values():
        used.update(p.get('evidence_claims', []))
    for _, ch in m.channels.values():
        used.update(ch.get('required_claims', []))
    idle = sorted(c for c, v in m.data['claims'].items() if v['state'] == 'proposed' and c not in used)
    if idle:
        report.add('info', 'IA_CLAIMS_UNREFERENCED', f'どこからも使われていない proposed の主張 {len(idle)} 件', ids=idle)


def check(root, date):
    report = Report('l0', date)
    m = load(root, report)
    if m.data is None:
        return report, m
    _check_ledger(m, report)
    _check_refs(m, report, date)
    if m.channels:
        _check_presets(m, report)
    _check_terms(m, report)
    _info_unreferenced(m, report)
    return report, m
