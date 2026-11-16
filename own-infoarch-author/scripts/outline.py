"""outline: the AI-made layer that stands on one positioning (token 0).

SPEC-information-architecture-standardization v6, requirements 6 (atoms and arrows), 7 (disqualification),
8 (outline and set ids), 9 (difference between two sets) and 13 (one positioning per outline).

- An outline stands on the positioning of its target's segment. Its constants (target, touchpoint, tone) and its
  objective are held fixed; messages (what), sections (where and in which order) and copy (in which words) vary.
- Order is the `order` field, never the order things are written in. The deliverable id, channel and outputs say
  where the outline is written and shipped, not what it says: they are not part of the outline id.
- The set id is the pair of the positioning id and the outline id.
"""
from __future__ import annotations

import json

import positioning as pos

CONTENT_KEYS = ('target', 'touchpoint', 'tone', 'objective', 'no_cta', 'messages', 'sections', 'copies', 'texts')
CONSTANT_KEYS = ('target', 'touchpoint', 'tone', 'objective')  # a variant that changes one is another experiment
NO_CTA_NOTE = 'CTA なし（意図）'


def _canonical_message(msg):
    out = dict(msg)
    if 'compares' in out:
        out['compares'] = sorted(out['compares'])
    return out


def _canonical_section(sec):
    return {**sec, 'places': sorted(sec.get('places', []))}


def content(o):
    """What an outline says, in a canonical form (sets sorted, placement fields left out)."""
    out = {k: o[k] for k in CONTENT_KEYS if k in o}
    if not out.get('no_cta'):
        out.pop('no_cta', None)  # `no_cta: false` and no declaration say the same thing
    out['messages'] = {k: _canonical_message(v) for k, v in (o.get('messages') or {}).items()}
    out['sections'] = {k: _canonical_section(v) for k, v in (o.get('sections') or {}).items()}
    return out


def outline_id(o):
    return 'out-' + pos.sha256(content(o))


def segment_of(o):
    target = o.get('target')
    return target.get('segment') if isinstance(target, dict) else None


def set_ids(m, oid, approved):
    """{'positioning', 'outline', 'set'} of one deliverable. No approved positioning -> positioning and set are None."""
    _, o = m.outlines[oid]
    p = pos.derive(m, approved).get(segment_of(o))
    pid, out = (p.id if p else None), outline_id(o)
    return {'positioning': pid, 'outline': out, 'set': f'{pid}+{out}' if pid else None}


def ordered_sections(o):
    """[(section id, section, [(copy id, copy), ...])] in `order`, copies in their own `order` (ties by id)."""
    secs = sorted((o.get('sections') or {}).items(), key=lambda kv: (kv[1].get('order', 0), kv[0]))
    copies = o.get('copies') or {}
    return [(sid, s, sorted(((cid, c) for cid, c in copies.items() if c.get('section') == sid),
                            key=lambda kv: (kv[1].get('order', 0), kv[0])))
            for sid, s in secs]


def elements(o):
    """{element key: content} of an outline, for the difference between two sets."""
    c = content(o)
    out = {k: c[k] for k in ('target', 'touchpoint', 'tone', 'objective') if k in c}
    out['no_cta'] = bool(c.get('no_cta'))
    for block in ('messages', 'sections', 'copies', 'texts'):
        for k, v in (c.get(block) or {}).items():
            out[f'{block}/{k}'] = v
    return out


def positioning_elements(m, approved, segment):
    p = pos.derive(m, approved).get(segment)
    if p is None:
        return {}
    atoms, edges = pos.atoms(m), pos.edges(m)
    return {k: pos.content(atoms.get(k) or edges[k]) for k in p.atoms + p.edges}


def _changes(layer, a, b):
    out = []
    for key in sorted(set(a) | set(b)):
        if key not in b:
            out.append({'layer': layer, 'element': key, 'change': 'removed'})
        elif key not in a:
            out.append({'layer': layer, 'element': key, 'change': 'added'})
        elif a[key] != b[key]:
            va, vb = a[key], b[key]
            attrs = (sorted(k for k in set(va) | set(vb) if va.get(k) != vb.get(k))
                     if isinstance(va, dict) and isinstance(vb, dict) else [])
            out.append({'layer': layer, 'element': key, 'change': 'changed', 'attributes': attrs})
    return out


def _unanchored(dom):
    """{viewport: sorted texts outside every anchor} of one rendered variant (a dom report's `unanchored`). Sorted, so
    the order of the outputs (and of the files a variant writes to) never reads as a difference."""
    out = {}
    for row in (dom or {}).get('unanchored') or []:
        out.setdefault(row.get('viewport'), []).append(' '.join(str(row.get('text', '')).split()))
    return {vp: sorted(t for t in texts if t) for vp, texts in out.items()}


def _fact_values(m, o):
    """{fact id: value} of the facts an outline's copy and declared text refer to."""
    import check
    facts = m.ssot.get('facts') or {}
    texts = [c['text'] for c in o['copies'].values()] + [t['text'] for t in (o.get('texts') or {}).values()]
    return {fid: (facts.get(fid) or {}).get('value') for t in texts for fid in check.PLACEHOLDER.findall(t)}


def diff(ma, oid_a, approved_a, mb, oid_b, approved_b, dom_a=None, dom_b=None):
    """Requirement 9: every changed atom, arrow and attribute between two sets, which layer moved, and whether the
    variants are the same experiment. A difference no element owns (text outside every anchor) disqualifies."""
    _, a = ma.outlines[oid_a]
    _, b = mb.outlines[oid_b]
    ids = {'a': set_ids(ma, oid_a, approved_a), 'b': set_ids(mb, oid_b, approved_b)}
    changes = _changes('outline', elements(a), elements(b))
    changes += _changes('positioning', positioning_elements(ma, approved_a, segment_of(a)),
                        positioning_elements(mb, approved_b, segment_of(b)))
    layers = sorted({c['layer'] for c in changes})
    differs = []
    for key in CONSTANT_KEYS:
        va, vb = a.get(key), b.get(key)
        if isinstance(va, dict) and isinstance(vb, dict):
            differs += [f'{key}/{k}' for k in sorted(set(va) | set(vb)) if va.get(k) != vb.get(k)]
        elif va != vb:
            differs.append(key)
    findings, not_checked = [], []
    fa, fb = _fact_values(ma, a), _fact_values(mb, b)
    moved_facts = sorted(f for f in set(fa) & set(fb) if fa[f] != fb[f])
    if moved_facts:  # the values are not echoed: only which facts differ
        findings.append({'level': 'fail', 'code': 'IA_DIFF_UNASSIGNED',
                         'message': f'別案どうしで固定値の値が違う {moved_facts}（outline にも positioning にも割り当てられない差分）'})
    if dom_a is None or dom_b is None:
        not_checked.append('印の外の文の差分：両方の dom の報告（--dom-a・--dom-b）が無いので比べていない')
    else:
        ua, ub = _unanchored(dom_a), _unanchored(dom_b)
        moved = sorted(v for v in set(ua) | set(ub) if ua.get(v, []) != ub.get(v, []))
        if moved:
            findings.append({'level': 'fail', 'code': 'IA_DIFF_UNASSIGNED',
                             'message': f'印の外の文が別案どうしで違う（どの要素にも割り当てられない差分。画面幅 {moved}）'})
    return {'a': oid_a, 'b': oid_b, 'ids': ids, 'changes': changes, 'changed_layers': layers,
            'same_experiment': not differs, 'experiment_differs': differs, 'findings': findings,
            'not_checked': not_checked, 'exit_code': 2 if findings else 0}


# ---- requirement 7 (and the shape of requirement 6) ------------------------------------------------------------

def check(m, report, oid, rel, o, shipped, approved):
    """Disqualification of one outline. Value approval follows requirement 15: fail when shipped, warn in a draft."""
    seg = segment_of(o)
    if seg is not None and (m.atom(seg) or {}).get('state') != 'proposed':
        report.add('fail', 'IA_OUTLINE_TARGET', f'{oid}: target のセグメント {seg} が無いか空欄', rel)
        seg = None
    msgs, secs, copies = o['messages'], o['sections'], o['copies']
    values = pos.atoms(m)
    for mid, msg in sorted(msgs.items()):
        v = values.get(msg['value'])
        if v is not None and seg is not None:
            if v.get('segment') != seg:
                report.add('fail', 'IA_OUTLINE_VALUE_OTHER_POSITIONING',
                           f'{oid}.{mid}: 提供価値 {msg["value"]} は立っている positioning（{seg}）のものでない', rel, id=mid)
            elif msg['value'] not in approved:
                report.add('fail' if shipped else 'warn', 'IA_OUTLINE_VALUE_UNAPPROVED',
                           f'{oid}.{mid}: 提供価値 {msg["value"]} にオーナーの承認が無い', rel, id=mid)
        for alt in msg.get('compares', []):
            a = values.get(alt)
            if a is None or seg is None:
                continue  # a missing alternative is IA_OUTLINE_REF_BROKEN; a missing segment IA_OUTLINE_TARGET
            if pos._prefix(alt) != 'alt' or a.get('segment') != seg:
                report.add('fail', 'IA_OUTLINE_COMPARE_INVALID',
                           f'{oid}.{mid}: 比べる相手 {alt} は立っている positioning（{seg}）の代替でない', rel, id=mid)
            elif alt not in approved:
                report.add('fail' if shipped else 'warn', 'IA_OUTLINE_COMPARE_UNAPPROVED',
                           f'{oid}.{mid}: 比べる相手 {alt} にオーナーの承認が無い', rel, id=mid)
        if msg['type'] == 'cta' and msg.get('achieves') != o['objective']['id']:
            report.add('fail', 'IA_OUTLINE_CTA_NO_ACTION',
                       f'{oid}.{mid}: CTA が取ってほしい行動 {o["objective"]["id"]} を指さない（achieves）', rel, id=mid)
    placed = {}
    for sid, sec in sorted(secs.items()):
        for mid in sec['places']:
            if mid not in msgs:
                report.add('fail', 'IA_OUTLINE_REF_BROKEN', f'{oid}.{sid}: 存在しないメッセージ {mid} を置いている', rel)
            else:
                placed.setdefault(mid, []).append(sid)
        kinds = [msgs[x]['type'] for x in sec['places'] if x in msgs]
        if kinds.count('normal') == 0:
            report.add('fail', 'IA_OUTLINE_SECTION_NO_MESSAGE', f'{oid}.{sid}: 通常のメッセージを置いていない', rel)
        if kinds.count('normal') > 1 or kinds.count('cta') > 1:
            report.add('fail', 'IA_OUTLINE_SECTION_OVERLOADED',
                       f'{oid}.{sid}: 通常のメッセージ {kinds.count("normal")} つ・CTA {kinds.count("cta")} つ（通常 1 つと CTA 0〜1 つ）', rel)
    if not secs:
        report.add('fail', 'IA_OUTLINE_NO_SECTIONS', f'{oid}: セクションが 0 件', rel)
    orders = {}
    for sid, sec in secs.items():
        orders.setdefault(sec['order'], []).append(sid)
    for n, sids in sorted(orders.items()):
        if len(sids) > 1:
            report.add('fail', 'IA_OUTLINE_ORDER_DUPLICATE', f'{oid}: セクション {sorted(sids)} の順序がどれも {n}', rel)
    leads = [mid for mid, msg in msgs.items() if msg['rank'] == 'lead']
    if len(leads) != 1:
        report.add('fail', 'IA_OUTLINE_LEAD_COUNT', f'{oid}: 一番のメッセージが {len(leads)} つ {sorted(leads)}（ちょうど 1 つ）', rel)
    for mid in leads:
        if mid not in placed:
            report.add('fail', 'IA_OUTLINE_LEAD_UNPLACED', f'{oid}: 一番のメッセージ {mid} がどのセクションにも置かれていない', rel, id=mid)
    ctas = [mid for mid, msg in msgs.items() if msg['type'] == 'cta']
    if o.get('no_cta'):
        if ctas:
            report.add('fail', 'IA_OUTLINE_CTA_CONTRADICTS', f'{oid}: 「CTA を置かない」と書いたのに CTA {sorted(ctas)} がある', rel)
        report.add('info', 'IA_OUTLINE_NO_CTA_INTENDED', f'{oid}: {NO_CTA_NOTE}', rel)
    elif not any(mid in placed for mid in ctas):
        report.add('fail', 'IA_OUTLINE_CTA_MISSING', f'{oid}: CTA をどのセクションにも置いていない（置かないなら no_cta: true と書く）', rel)
    for mid in sorted(set(msgs) - set(placed) - set(leads)):
        report.add('warn', 'IA_OUTLINE_MESSAGE_UNPLACED', f'{oid}: メッセージ {mid} がどのセクションにも置かれていない', rel, id=mid)
    conveyed, copy_orders = set(), {}
    for cid, cp in sorted(copies.items()):
        if cp['section'] not in secs:
            report.add('fail', 'IA_OUTLINE_COPY_SECTION', f'{oid}.{cid}: 属するセクション {cp["section"]} が無い', rel, id=cid)
            continue
        copy_orders.setdefault((cp['section'], cp['order']), []).append(cid)
        if cp['message'] not in secs[cp['section']]['places']:
            report.add('fail', 'IA_OUTLINE_COPY_MESSAGE',
                       f'{oid}.{cid}: 伝えるメッセージ {cp["message"]} が、属するセクション {cp["section"]} に置かれていない', rel, id=cid)
        else:
            conveyed.add((cp['section'], cp['message']))
    for (sid, n), cids in sorted(copy_orders.items()):
        if len(cids) > 1:
            report.add('fail', 'IA_OUTLINE_ORDER_DUPLICATE', f'{oid}.{sid}: 文案 {sorted(cids)} の順番がどれも {n}', rel)
    for sid, sec in sorted(secs.items()):
        for mid in sec['places']:
            if mid in msgs and (sid, mid) not in conveyed:
                report.add('fail', 'IA_OUTLINE_MESSAGE_UNCONVEYED', f'{oid}.{sid}: 置かれたメッセージ {mid} を伝える文案が無い', rel, id=mid)
    for key, t in sorted((o.get('texts') or {}).items()):
        if t['kind'] == 'navigation' and 'section' in t:
            report.add('fail', 'IA_OUTLINE_TEXT_MISPLACED', f'{oid}.texts.{key}: ナビの文言はセクションに置かない', rel)
        elif t['kind'] == 'label' and 'section' not in t:
            report.add('fail', 'IA_OUTLINE_TEXT_MISPLACED', f'{oid}.texts.{key}: ラベルは置くセクション（section）を書く', rel)
        elif 'section' in t and t['section'] not in secs:
            report.add('fail', 'IA_OUTLINE_REF_BROKEN', f'{oid}.texts.{key}: セクション {t["section"]} が無い', rel)


def load_dom_report(path):
    return json.loads(open(path, encoding='utf-8').read()) if path else None
