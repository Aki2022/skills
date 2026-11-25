"""Deterministic checks of ux/infoarch/ (token 0): the positioning SSOT, channel declarations, scripts (台本).

What is checked here is structure and references, never meaning. The one meaning judgment (copy x the value its message
points to) belongs to the Jev judgement ledger; rendered artifacts belong to the DOM checks.

The positioning SSOT is the positioning layer's (own-positioning-author): this module reads it through that skill's
loader and adds what stands on it: channels, scripts, the dictionary's values and names (docs/dictionary/, read
through own-doc-update's reader), approvals per path, expiry and the design page inventory.

The script (scripts/<id>.yaml, script.v1) replaced the v6 outline. The store keeps its v6 name `outlines` and the
helpers their `outline_*` names; the retired outlines/ directory is never read, and any file in it fails.
"""
from __future__ import annotations

import datetime as dt
import hashlib
import json
import re
from dataclasses import dataclass, field
from pathlib import Path

import common  # noqa: F401  (puts the sibling skills' scripts on sys.path)
import layer_paths
import blueprint
import dictionary
import outline as ol
import positioning as pos
import positioning_model as pm
from common import Report, YamlError, load_ref, load_yaml, validator_for
from positioning_model import (  # noqa: F401  (the positioning layer's, re-exported for the rest of this skill)
    APPROVAL_PREFIX, BLOCKS, ID_RE, POSITIONING, PREFIX_BLOCK, SSOT_PATH, check_ssot, element, positioning_path,
    review_markdown, ssot_refs)

# A fixed value is the dictionary's: a copy of a `value` entry is written {{<dictionary id>}}. A schema 4 tree still
# writes {{fct-NNNN}} for the facts its SSOT carries (read with a warning until it moves to the dictionary).
PLACEHOLDER = re.compile(r'\{\{(value-[a-z0-9]+(?:-[a-z0-9]+)*|fct-(?:(?!0000)[0-9]{4}|[1-9][0-9]{4,}))\}\}')  # no padding
ANY_BRACES = re.compile(r'\{\{.*?\}\}', re.S)
SCRIPTS_DIR = 'scripts'  # ux/infoarch/scripts/<deliverable id>.yaml: the script (台本)
SCRIPT_SCHEMA = 'script.v1.schema.json'
RETIRED_DIR = 'outlines'  # the v6 outline's place: never read; a file left there fails (IA_OUTLINE_RETIRED)


@dataclass
class Model(pm.Model):
    channels: dict = field(default_factory=dict)  # media -> (rel path, data)
    blueprint: object = None  # the product UI's infoarch (blueprint.Blueprint), when ux/infoarch/blueprint/ exists
    outlines: dict = field(default_factory=dict)  # scripts (台本): id -> (rel path, data); v6 name kept
    # The dictionary as the rest of the checks read it: facts (kind value) and entities (kind name), keyed by the
    # dictionary id. Built at load time, so dom, judge, skeleton and the answer key read one shape wherever the
    # values came from (docs/dictionary/, or a schema 4 SSOT that still carries them).
    facts: dict = field(default_factory=dict)
    entities: dict = field(default_factory=dict)

    def atom(self, aid):
        return self.facts.get(aid) or self.entities.get(aid) or super().atom(aid)


def _errors(validator, data):
    return pm.errors_of(validator, data)


def _load(path, rel, report):
    try:
        return load_yaml(path), True
    except YamlError as exc:
        report.add('fail', 'IA_YAML', f'YAML として読めない: {exc}', rel)
        return None, False


def terms_of(entries):
    """-> (facts, entities) of dictionary entries: kind value becomes a fact, kind name an entity, keyed by dictionary id."""
    facts = {e['id']: {'state': 'proposed', 'label': e['meaning'], 'value': e['text'],
                       **({'expires': e['expires']} if e.get('expires') else {}),
                       **({'evergreen': True} if e.get('evergreen') else {}),
                       **({'as_of': e['as_of']} if e.get('as_of') else {}),
                       'from': 'dictionary'} for e in entries if e['kind'] == 'value'}
    entities = {e['id']: {'state': 'proposed', 'type': 'name', 'name': e['text'], 'from': 'dictionary'}
                for e in entries if e['kind'] == 'name'}
    return facts, entities


def _dictionary_view(root, legacy, report):
    """-> (facts, entities). From docs/dictionary/, or from a schema 4 SSOT's own blocks."""
    entries, nfiles, errors = dictionary.load(root)
    for e in errors:
        report.add('fail', 'IA_DICTIONARY_INVALID', f'docs/dictionary/ が辞書の形でない: {e}（辞書の値を読めない）', 'docs/dictionary')
    facts, entities = terms_of(entries)
    if legacy:  # a schema 4 SSOT keeps its own, under its own ids; the dictionary's win when both exist
        facts = {**{k: {**v, 'from': 'ssot-v4'} for k, v in legacy['facts'].items()}, **facts}
        entities = {**{k: {**v, 'from': 'ssot-v4'} for k, v in legacy['entities'].items()}, **entities}
    return facts, entities


def load(root, report):
    root = Path(root)
    base = layer_paths.path(root, 'infoarch')  # a repository that moved stops here, never "0 scripts"
    m = Model(root)
    data, legacy = pm.load_ssot(root, report)
    if data is None:
        return m
    m.ssot = data
    m.facts, m.entities = _dictionary_view(root, legacy, report)
    # downstream readers (dom, judge, skeleton, answer key) take facts and entities from the SSOT dict
    m.ssot['facts'], m.ssot['entities'] = m.facts, m.entities
    retired = base / RETIRED_DIR
    left = (sorted(p.name for p in retired.iterdir()) if retired.is_dir() and not retired.is_symlink()
            else [retired.name] if retired.is_symlink() or retired.exists() else [])
    if left:  # an empty directory holds nothing; anything in it (dotfiles too) is a leftover
        report.add('fail', 'IA_OUTLINE_RETIRED',
                   f'{RETIRED_DIR}/ は廃止した（台本は {SCRIPTS_DIR}/<成果物 id>.yaml）。残っている {left} を読まずに 0 件にはしない。'
                   'outline は機械で写さず、positioning から台本を作り直してから消す', layer_paths.rel('infoarch', RETIRED_DIR))
    for kind, sub, schema, store in (('channel', 'channels', 'channel.v2.schema.json', m.channels),
                                     ('outline', SCRIPTS_DIR, SCRIPT_SCHEMA, m.outlines)):
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
    m.blueprint = blueprint.load(root, report)
    return m


def check_blueprint(m, report, approved_values, rows):
    """The product UI's infoarch (blueprint), when the repository has one: product items and screens."""
    if m.blueprint is not None and m.blueprint.present:
        blueprint.check(m.blueprint, report, pos.atoms(m), approved_values, blueprint.approved_items(m.blueprint, rows))


def _ref(m, report, src, aid, where, **kw):
    return pm.check_reference(m, report, src, aid, where, **kw)


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


# ---- owner approval per value path (v6 requirements 3 and 15) -------------------------------------------------

APPROVAL_PREFIX = pos.PATH_PREFIX


def positioning_targets(m):
    """-> ({approval key: hash}, sorted keys shipping needs). Keys: ia:path:<value> (the hash of a value's whole path) for
    every proposed value, and ia:item:<item id> (the hash of that item alone) for every proposed product item. Shipping
    needs the paths shipped outlines use and the items shipped screens cite."""
    hashes = pos.approval_targets(m)
    must = {APPROVAL_PREFIX + v for v, shipped in referenced_paths(m).items() if shipped}
    must &= set(hashes)
    bp = m.blueprint
    if bp is not None and bp.present and bp.product:
        hashes = {**hashes, **blueprint.targets(bp)}
        must |= set(blueprint.required_keys(bp))
    return hashes, sorted(must)


def read_approvals(root, report):
    return pm.read_approvals(root, report)


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
    for cid, cp in sorted(o['copies'].items()):
        out += [(f'copies.{cid}.refs', p) for p in cp.get('refs', [])]
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
            report.add('fail', 'IA_OUTLINE_BAD_PLACEHOLDER', f'{w}: {bad} は固定値の参照 {{{{value-<辞書の id>}}}} の形でない', rel)
        plain = PLACEHOLDER.sub(' ', text)
        for fid, value in facts.items():
            if value in plain:
                report.add('fail', 'IA_OUTLINE_FACT_LITERAL', f'{w}: 固定値 {fid} の値「{value}」を直接書いている（{{{{{fid}}}}} で参照する）', rel)


# ---- design page inventory (ADR-20261115-design-ia-two-skills-shared-base) -------------------

HTML_SUFFIXES = ('.html', '.htm')


def check_pages(m, report):
    """Shipped HTML outputs against ux/design/validation.json pages: the two declarations are kept, but must agree."""
    path = layer_paths.path(m.root, 'validation')
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
                           f'{oid}: 出荷される {out} が {layer_paths.rel("validation")} の pages（covered・excluded）に無い。design の検査が見ていない', rel)


# ---- expiry ------------------------------------------------------------------------------

# INFO_ARCH.md is prose that refers to atoms by id and holds no value (SPEC-information-architecture-standardization,
# requirement 1). The ids it may cite: a positioning atom or a product item (the prefixes the two layers issue), a screen
# `scr-<slug>`, and a script written `script:<id>` (or `outline:<id>`, the v6 form, still read). v5 ids (msg-…, aud-…) are the migration's business, not read here.
_PROSE_PREFIXES = '|'.join(sorted(set(pm.BLOCKS.values()) | set(blueprint.ITEM_BLOCKS.values())))
PROSE_ID = re.compile(r'(?<![A-Za-z0-9_-])((?:' + _PROSE_PREFIXES + r')-(?:(?!0000)[0-9]{4}|[1-9][0-9]{4,})'
                      r'|scr-[a-z0-9]+(?:-[a-z0-9]+)*|(?:outline|script):[a-z0-9][a-z0-9_-]*)(?![A-Za-z0-9_-])')


def _tombstoned(m, aid):
    bp = m.blueprint.product if m.blueprint is not None and m.blueprint.present and m.blueprint.product else None
    return any(aid in ((src or {}).get('ledger') or {}).get('tombstones', {}) for src in (m.ssot, bp))


def check_info_arch_refs(m, report):
    """Every id INFO_ARCH.md cites exists (fail) or was retired (warn). A repository without the file has nothing to check."""
    path = layer_paths.path(m.root, 'infoarch', 'INFO_ARCH.md')
    if not path.is_file():
        return
    rel = path.relative_to(Path(m.root).resolve()) if path.is_absolute() else path
    bp = m.blueprint if m.blueprint is not None and m.blueprint.present else None
    seen = set()
    for lineno, line in enumerate(path.read_text(encoding='utf-8').splitlines(), 1):
        for mention in PROSE_ID.findall(line):
            if mention in seen:
                continue
            seen.add(mention)
            if mention.startswith(('outline:', 'script:')):
                aid = mention.split(':', 1)[1]
                live = aid in m.outlines
            elif mention.startswith('scr-'):
                aid, live = mention, bool(bp and mention in bp.screens)
            elif blueprint.ITEM_ID.fullmatch(mention):
                aid, live = mention, bool(bp and bp.item(mention))
            else:
                aid, live = mention, m.atom(mention) is not None
            if live:
                continue
            if _tombstoned(m, aid):
                report.add('warn', 'IA_INFO_ARCH_REF_TOMBSTONED', f'INFO_ARCH.md の {lineno} 行が、取り下げた id {aid} を引いている', rel, lineno, id=aid)
            else:
                report.add('fail', 'IA_INFO_ARCH_REF_BROKEN', f'INFO_ARCH.md の {lineno} 行が、無い id {aid} を引いている（綴りか、消した id を確かめる）', rel, lineno, id=aid)


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
    rows = read_approvals(root, report)
    approved = pos.approved_keys(m, rows)
    _, all_paths = pos.check_positioning(m, report, approved, referenced_paths(m), str(SSOT_PATH))
    product_ui = bool(m.blueprint and m.blueprint.present)  # a product UI is checked by its blueprint, not by outlines
    if not m.channels and (m.outlines or not product_ui):
        report.add('fail', 'IA_NO_CHANNELS', 'channels/*.yaml が 0 件（対象 0 件を合格にしない）', layer_paths.rel('infoarch', 'channels'))
    if not m.outlines and not product_ui:
        report.add('fail', 'IA_NO_OUTLINES', f'{SCRIPTS_DIR}/*.yaml（台本）が 0 件（対象 0 件を合格にしない）', layer_paths.rel('infoarch', SCRIPTS_DIR))
    for oid, (rel, o) in sorted(m.outlines.items()):
        ch = m.channels.get(o['channel'])
        shipped = bool(ch and ch[1]['ship'])
        check_outline(m, report, oid, rel, o)
        ol.check(m, report, oid, rel, o, shipped, approved)
        check_positioning_use(m, report, oid, rel, o, shipped, approved)
        check_lead_support(m, report, oid, rel, o, all_paths)
    check_blueprint(m, report, approved, rows)
    check_expiry(m, report, date)
    check_pages(m, report)
    check_info_arch_refs(m, report)
    # The script side (channels, outlines = scripts, blocks) and the blueprint side (screens, items) are counted apart: a product
    # UI has screens and items and no outline, a deliverable has outlines and no blueprint. The bridge (ia_module) refuses
    # a check where neither side has anything, and a side that is half empty.
    bp = m.blueprint if m.blueprint is not None and m.blueprint.present else None
    report.targets = {'channels': len(m.channels), 'outlines': len(m.outlines),
                      'blocks': sum(len(o['blocks']) for _, o in m.outlines.values()),
                      'screens': len(bp.screens) if bp else 0, 'items': len(blueprint.items(bp)) if bp else 0}
    return report, m


# What `check` and `blueprint` never run, whoever owns it: named in `skipped`, so that a green report is not read as a
# green rendered page or a green judgment (SPEC-information-architecture-standardization, Quality 2).
NOT_RUN_HERE = (
    {'id': 'ia-dom', 'reason': '描画した画面は見ない（infoarch-author dom／ship が、出荷する HTML の成果物を見る）'},
    {'id': 'ia-semantic-judgments', 'reason': 'Jev の判定は読まない（infoarch-author judge／evaluate が、共有の判定台帳を読む）'},
)


def categorize(report, m, command):
    """Give a `check` (or `blueprint`) report its five categories (ia_common.CATEGORIES). Call it last, after the report
    holds every finding: `missing` also lists the inputs that a failure says are absent. -> report"""
    report.categorized = True
    report.skipped = [dict(e) for e in NOT_RUN_HERE]
    bp = m.blueprint if m.blueprint is not None and m.blueprint.present else None
    if m.ssot is not None:
        report.executed.append({'id': 'ia-positioning', 'targets': {'values': len(pos.paths(m))}})
    if command == 'check':
        if m.outlines:
            report.executed.append({'id': 'ia-outlines', 'targets': {
                'channels': len(m.channels), 'outlines': len(m.outlines),
                'blocks': sum(len(o['blocks']) for _, o in m.outlines.values())}})
        elif bp is not None:
            report.skipped.append({'id': 'ia-outlines', 'reason': '台本が無い（プロダクト UI は blueprint で検査する）'})
    else:
        report.skipped.append({'id': 'ia-outlines', 'reason': 'blueprint は台本を見ない（infoarch-author check）'})
    if bp is not None:
        report.executed.append({'id': 'ia-blueprint', 'targets': {'screens': len(bp.screens), 'items': len(blueprint.items(bp))}})
    elif command == 'check':
        report.skipped.append({'id': 'ia-blueprint', 'reason': 'blueprint/ が無い（台本の成果物）'})
    return report


def impact(root, aid):
    report = Report('impact', dt.date.today().isoformat())
    m = load(root, report)
    if m.ssot is not None and m.atom(aid) is None and m.tombstone(aid) == '__live__':
        report.add('fail', 'IA_IMPACT_UNKNOWN_ID', f'{aid} は SSOT にも墓標にも無い（綴りを確かめる）')
    hits = sorted(oid for oid, (_, o) in m.outlines.items() if m.ssot is not None and aid in outline_dependencies(m, o))
    outputs = sorted({out for oid in hits for out in m.outlines[oid][1]['outputs']})
    bp = m.blueprint
    if bp is not None and bp.present and bp.item(aid) is not None:
        report.findings = [f for f in report.findings if f.code != 'IA_IMPACT_UNKNOWN_ID']
    screens = sorted(sid for sid, (_, s) in bp.screens.items() if aid in blueprint.dependencies(bp, s)) if bp is not None else []
    return report, {'id': aid, 'outlines': hits, 'outputs': outputs, 'screens': screens}
