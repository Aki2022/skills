"""script (台本): the AI-made layer that stands on one positioning (token 0). It replaced the v6 outline.

SPEC-information-architecture-standardization, requirements 6 (atoms and arrows), 7 (disqualification),
8 (script and set ids), 9 (difference between two sets) and 13 (one positioning per script);
SPEC-deliverable-production-layers, requirements 16 (content functions), 17 (proposals) and 20 (script replaces outline).

- A script stands on the positioning of its target's segment. Its constants (target, touchpoint, tone) and its
  objective are held fixed; messages (what), blocks (where, in which order, with which content function) and copy
  (in which words) vary. Only a claim block places a normal message (exactly one) and only an action block a CTA (0-1).
- A copy has a reference (the message it conveys, positioning evidence in `refs`, a dictionary value) or the proposal
  mark. Proposals warn while producing; the ship gate refuses a shipped script that still has one (`proposals`).
- Order is the `order` field, never the order things are written in. The deliverable id, channel and outputs say
  where the script is written and shipped, not what it says: they are not part of the script id.
- The set id is the script id with the ids of the value paths its messages point at (requirement 8). The
  positioning id is reported, not part of the set id: editing a path nobody points at never moves it.

Module and store names (`outline`, `m.outlines`) and the codes of unchanged rules (IA_OUTLINE_*) are kept from v6.
"""
from __future__ import annotations

import json

import positioning as pos

CONTENT_KEYS = ('target', 'touchpoint', 'tone', 'objective', 'no_cta', 'messages', 'blocks', 'copies', 'texts')
CONSTANT_KEYS = ('target', 'touchpoint', 'tone', 'objective')  # a variant that changes one is another experiment
# A variant declares its base (`variant_of`) and the elements it means to change (`overrides`, element keys as
# `elements` gives them). The file is still a whole script; the declaration is placement, like channel and outputs:
# it is in neither the script id nor the set id.
CONFOUNDED_LABEL = '交絡あり'
PLAN_NOT_CHECKED = '組み方の計画は比べていない（形が未定）'
NO_CTA_NOTE = 'CTA なし（意図）'
SCRIPT_PREFIX = 'spt-'  # `scr-` is a screen's (blueprint)
# The content functions (docs/dictionary/deliverables.json term-function-*): which may place messages, and which copy
# must carry a reference or the proposal mark.
FUNCTIONS = ('identify', 'claim', 'fact', 'evidence', 'offer', 'action', 'source')
CITING = ('claim', 'fact', 'evidence')


def _canonical_message(msg):
    out = dict(msg)
    if 'compares' in out:
        out['compares'] = sorted(out['compares'])
    return out


def _canonical_block(blk):
    out = {**blk, 'places': sorted(blk.get('places', []))}
    if not out['places']:
        out.pop('places')  # `places: []` and no places say the same thing
    return out


def _canonical_copy(cp):
    out = dict(cp)
    if 'refs' in out:
        out['refs'] = sorted(out['refs'])
    return out


def content(o):
    """What a script says, in a canonical form (sets sorted, placement fields left out)."""
    out = {k: o[k] for k in CONTENT_KEYS if k in o}
    if not out.get('no_cta'):
        out.pop('no_cta', None)  # `no_cta: false` and no declaration say the same thing
    out['messages'] = {k: _canonical_message(v) for k, v in (o.get('messages') or {}).items()}
    out['blocks'] = {k: _canonical_block(v) for k, v in (o.get('blocks') or {}).items()}
    out['copies'] = {k: _canonical_copy(v) for k, v in (o.get('copies') or {}).items()}
    return out


def outline_id(o):
    """The script id (requirement 8): `spt-` + the hash of its content."""
    return SCRIPT_PREFIX + pos.sha256(content(o))


def segment_of(o):
    target = o.get('target')
    return target.get('segment') if isinstance(target, dict) else None


def referenced_paths(m, o, all_paths=None):
    """{value id: ValuePath} of the values the outline's messages point at, inside its target's segment."""
    import check
    ps = pos.paths(m) if all_paths is None else all_paths
    return {v: ps[v] for v in check.referenced_values(m, o)}


def set_ids(m, oid, approved):
    """{'positioning', 'outline', 'paths', 'set'} of one deliverable (requirement 8). The set id is the outline id
    with the path ids of the values it points at, sorted by value id; it hashes content, so it is the same whether or
    not the owner has approved yet (approval is the ship gate's, not the id's). No path pointed at -> set is None.
    The positioning id (every approved path of the segment) is for reporting only."""
    _, o = m.outlines[oid]
    ps = pos.paths(m)
    p = pos.derive(m, approved, ps).get(segment_of(o))
    out = outline_id(o)
    refs = {v: q.id for v, q in referenced_paths(m, o, ps).items()}
    return {'positioning': p.id if p else None, 'outline': out, 'paths': refs,
            'set': out + ''.join('+' + refs[v] for v in sorted(refs)) if refs else None}


def ordered_blocks(o):
    """[(block id, block, [(copy id, copy), ...])] in `order`, copies in their own `order` (ties by id)."""
    blks = sorted((o.get('blocks') or {}).items(), key=lambda kv: (kv[1].get('order', 0), kv[0]))
    copies = o.get('copies') or {}
    return [(bid, b, sorted(((cid, c) for cid, c in copies.items() if c.get('block') == bid),
                            key=lambda kv: (kv[1].get('order', 0), kv[0])))
            for bid, b in blks]


def proposals(o):
    """Sorted ids of the copies that still carry the proposal mark. Production warns on each; the ship gate refuses a
    shipped script with any (SPEC-deliverable-production-layers requirement 17, IA requirement 15)."""
    return sorted(cid for cid, c in (o.get('copies') or {}).items() if c.get('proposal'))


def elements(o):
    """{element key: content} of a script, for the difference between two sets."""
    c = content(o)
    out = {k: c[k] for k in ('target', 'touchpoint', 'tone', 'objective') if k in c}
    out['no_cta'] = bool(c.get('no_cta'))
    for part in ('messages', 'blocks', 'copies', 'texts'):
        for k, v in (c.get(part) or {}).items():
            out[f'{part}/{k}'] = v
    return out


def path_elements(m, p):
    """{element key: content} of one path, for the difference between two sets."""
    A, E = pos.atoms(m), pos.edges(m)
    return {k: pos.content(A.get(k) or E[k]) for k in p.members}


def path_changes(ma, a, mb, b):
    """Requirement 9: paths are matched by value id. The same value with another hash is a positioning edit (listed
    element by element); a value on one side only is an outline selection."""
    pa, pb = referenced_paths(ma, a), referenced_paths(mb, b)
    out = []
    for v in sorted(set(pa) | set(pb)):
        if v in pa and v in pb:
            if pa[v].hash != pb[v].hash:
                out += [{**c, 'path': v} for c in _changes('positioning', path_elements(ma, pa[v]),
                                                            path_elements(mb, pb[v]))]
        else:
            out.append({'layer': 'outline', 'element': f'paths/{v}', 'change': 'added' if v in pb else 'removed',
                        'kind': 'selection'})
    return out


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
    changes += path_changes(ma, a, mb, b)
    layers = sorted({c['layer'] for c in changes})
    differs = experiment_differs(a, b)
    declaring = b if b.get('variant_of') == oid_a else a if a.get('variant_of') == oid_b else None
    over = overrides(a if declaring is b else b, declaring) if declaring is not None else overrides(a, b, declared=())
    findings, not_checked = list(over['findings']), [PLAN_NOT_CHECKED]
    fa, fb = _fact_values(ma, a), _fact_values(mb, b)
    moved_facts = sorted(f for f in set(fa) & set(fb) if fa[f] != fb[f])
    if moved_facts:  # the values are not echoed: only which facts differ
        findings.append({'level': 'fail', 'code': 'IA_DIFF_UNASSIGNED',
                         'message': f'別案どうしで固定値の値が違う {moved_facts}（台本にも positioning にも割り当てられない差分）'})
    if dom_a is None or dom_b is None:
        not_checked.append('印の外の文の差分：両方の dom の報告（--dom-a・--dom-b）が無いので比べていない')
    else:
        ua, ub = _unanchored(dom_a), _unanchored(dom_b)
        moved = sorted(v for v in set(ua) | set(ub) if ua.get(v, []) != ub.get(v, []))
        if moved:
            findings.append({'level': 'fail', 'code': 'IA_DIFF_UNASSIGNED',
                             'message': f'印の外の文が別案どうしで違う（どの要素にも割り当てられない差分。画面幅 {moved}）'})
    levels = {f['level'] for f in findings}
    return {'a': oid_a, 'b': oid_b, 'ids': ids, 'changes': changes, 'changed_layers': layers,
            'same_experiment': not differs, 'experiment_differs': differs,
            # which side of a variant differs: the script, the layout plan, or both. The plan has no file yet, so
            # only the script side is compared (the plan is listed in not_checked).
            'differs_in': 'script' if any(c['layer'] == 'outline' for c in changes) else 'none',
            'variant_of': (declaring or {}).get('variant_of'), 'confounded': bool(over['undeclared']),
            'confounded_label': CONFOUNDED_LABEL if over['undeclared'] else None,
            'undeclared': over['undeclared'], 'unchanged_overrides': over['unchanged'], 'findings': findings,
            'not_checked': not_checked, 'exit_code': 2 if 'fail' in levels else 1 if 'warn' in levels else 0}


def experiment_differs(a, b):
    """The constants and objective keys that differ between two scripts: any one makes them two experiments."""
    differs = []
    for key in CONSTANT_KEYS:
        va, vb = a.get(key), b.get(key)
        if isinstance(va, dict) and isinstance(vb, dict):
            differs += [f'{key}/{k}' for k in sorted(set(va) | set(vb)) if va.get(k) != vb.get(k)]
        elif va != vb:
            differs.append(key)
    return differs


_ABSENT = object()


def overrides(base, variant, declared=None):
    """Requirement 9 / production layers 18: the elements of `variant` that differ from `base` and are not declared
    (`undeclared`: a confounded comparison) and the declared ones that do not differ (`unchanged`), with a warning
    each. The constants and objective are left out: a change there is another experiment, not a confound.
    `declared` defaults to the variant's own `overrides`."""
    ea, eb = elements(base), elements(variant)
    keys = set(ea) | set(eb)
    changed = {k for k in keys if ea.get(k, _ABSENT) != eb.get(k, _ABSENT)}
    declared = set(variant.get('overrides', []) if declared is None else declared)
    undeclared = sorted(k for k in changed - declared if k not in CONSTANT_KEYS)
    unchanged = sorted(k for k in declared & keys if k not in changed)
    findings = []
    if undeclared:
        findings.append({'level': 'warn', 'code': 'IA_SCRIPT_UNDECLARED_DIFFERENCE',
                         'message': f'{CONFOUNDED_LABEL}：基本の台本と違うのに上書きとして宣言していない要素 {undeclared}'})
    if unchanged:
        findings.append({'level': 'warn', 'code': 'IA_SCRIPT_OVERRIDE_UNCHANGED',
                         'message': f'上書きとして宣言したのに基本の台本と違わない要素 {unchanged}'})
    return {'undeclared': undeclared, 'unchanged': unchanged, 'findings': findings}


def check_variant(m, report, oid, rel, o):
    """A variant's declaration (requirement 6, last bullet; 9): its base exists and is another script, the constants
    and objective are the base's, every override names an element of either script; then the comparison above."""
    base_id = o.get('variant_of')
    if base_id is None:
        if 'overrides' in o:
            report.add('fail', 'IA_SCRIPT_OVERRIDES_WITHOUT_BASE', f'{oid}: overrides を書いたのに variant_of（基本の台本）が無い', rel)
        return
    if base_id == oid or base_id not in m.outlines:
        report.add('fail', 'IA_SCRIPT_VARIANT_BASE_MISSING', f'{oid}: 基本の台本 {base_id} が無い（自分自身は基本にできない）', rel)
        return
    _, base = m.outlines[base_id]
    differs = experiment_differs(base, o)
    if differs:
        report.add('fail', 'IA_SCRIPT_VARIANT_OTHER_EXPERIMENT',
                   f'{oid}: 定数か目的変数 {differs} が基本の台本 {base_id} と違う（変種でなく別の試験。variant_of を外す）', rel)
    known = set(elements(base)) | set(elements(o))
    unknown = sorted(k for k in o.get('overrides', []) if k not in known)
    if unknown:
        report.add('fail', 'IA_SCRIPT_OVERRIDE_UNKNOWN', f'{oid}: 上書きの宣言 {unknown} が、{base_id} にもこの台本にも無い要素', rel)
    for f in overrides(base, o)['findings']:
        report.add(f['level'], f['code'], f'{oid}（基本は {base_id}）: {f["message"]}', rel)


# ---- requirement 7 (and the shape of requirement 6) ------------------------------------------------------------

def check(m, report, oid, rel, o, shipped, approved):
    """Disqualification of one script (requirement 7). Value approval follows requirement 15: fail when shipped, warn in a
    draft. A proposal always warns here: refusing it is the ship gate's (ia_ship, IA_SCRIPT_PROPOSAL_AT_SHIP)."""
    seg = segment_of(o)
    if seg is not None and (m.atom(seg) or {}).get('state') != 'proposed':
        report.add('fail', 'IA_OUTLINE_TARGET', f'{oid}: target のセグメント {seg} が無いか空欄', rel)
        seg = None
    msgs, blks, copies = o['messages'], o['blocks'], o['copies']
    values, ps = pos.atoms(m), pos.paths(m)
    for mid, msg in sorted(msgs.items()):
        v = values.get(msg['value'])
        p = ps.get(msg['value'])
        if v is not None and seg is not None:
            if v.get('segment') != seg:
                report.add('fail', 'IA_OUTLINE_VALUE_OTHER_POSITIONING',
                           f'{oid}.{mid}: 提供価値 {msg["value"]} の経路のセグメントは {v.get("segment")}、target は {seg}', rel, id=mid)
                p = None
            elif msg['value'] not in approved:
                report.add('fail' if shipped else 'warn', 'IA_OUTLINE_VALUE_UNAPPROVED',
                           f'{oid}.{mid}: 提供価値 {msg["value"]} の経路にオーナーの承認が無い', rel, id=mid)
        for alt in msg.get('compares', []):
            a = values.get(alt)
            if a is None or seg is None:
                continue  # a missing alternative is IA_OUTLINE_REF_BROKEN; a missing segment IA_OUTLINE_TARGET
            if pos._prefix(alt) != 'alt' or a.get('segment') != seg:
                report.add('fail', 'IA_OUTLINE_COMPARE_INVALID',
                           f'{oid}.{mid}: 比べる相手 {alt} は立っている positioning（{seg}）の代替でない', rel, id=mid)
            elif p is not None and alt not in p.comparison:
                report.add('fail', 'IA_OUTLINE_COMPARE_OUTSIDE_SET',
                           f'{oid}.{mid}: 比べる相手 {alt} は提供価値 {msg["value"]} の経路の比べる集合 {p.comparison} に無い', rel, id=mid)
        if msg['type'] == 'cta' and msg.get('achieves') != o['objective']['id']:
            report.add('fail', 'IA_OUTLINE_CTA_NO_ACTION',
                       f'{oid}.{mid}: CTA が取ってほしい行動 {o["objective"]["id"]} を指さない（achieves）', rel, id=mid)
    placed = {}  # message id -> [block ids]
    for bid, blk in sorted(blks.items()):
        places = blk.get('places', [])
        for mid in places:
            if mid not in msgs:
                report.add('fail', 'IA_OUTLINE_REF_BROKEN', f'{oid}.{bid}: 存在しないメッセージ {mid} を置いている', rel)
            else:
                placed.setdefault(mid, []).append(bid)
        kinds = [msgs[x]['type'] for x in places if x in msgs]
        fn = blk['function']
        allowed = {'claim': 'normal', 'action': 'cta'}.get(fn)
        misplaced = sorted(x for x in places if x in msgs and msgs[x]['type'] != allowed)
        if misplaced:
            report.add('fail', 'IA_SCRIPT_MESSAGE_MISPLACED',
                       f'{oid}.{bid}: 働き {fn} のまとまりに置けないメッセージ {misplaced}（主張は通常のメッセージ、行動は CTA だけを置く）', rel)
        if fn == 'claim' and kinds.count('normal') == 0:
            report.add('fail', 'IA_SCRIPT_CLAIM_NO_MESSAGE', f'{oid}.{bid}: 主張のまとまりに通常のメッセージを置いていない', rel)
        if (fn == 'claim' and kinds.count('normal') > 1) or (fn == 'action' and kinds.count('cta') > 1):
            report.add('fail', 'IA_OUTLINE_SECTION_OVERLOADED',
                       f'{oid}.{bid}: 通常のメッセージ {kinds.count("normal")} つ・CTA {kinds.count("cta")} つ'
                       '（主張は通常のメッセージ 1 つ、行動は CTA 0〜1 つ）', rel)
    if not blks:
        report.add('fail', 'IA_OUTLINE_NO_SECTIONS', f'{oid}: まとまりが 0 件', rel)
    orders = {}
    for bid, blk in blks.items():
        orders.setdefault(blk['order'], []).append(bid)
    for n, bids in sorted(orders.items()):
        if len(bids) > 1:
            report.add('fail', 'IA_OUTLINE_ORDER_DUPLICATE', f'{oid}: まとまり {sorted(bids)} の順序がどれも {n}', rel)
    leads = [mid for mid, msg in msgs.items() if msg['rank'] == 'lead']
    if len(leads) != 1:
        report.add('fail', 'IA_OUTLINE_LEAD_COUNT', f'{oid}: 一番のメッセージが {len(leads)} つ {sorted(leads)}（ちょうど 1 つ）', rel)
    for mid in leads:
        if not any(blks[b]['function'] == 'claim' for b in placed.get(mid, [])):
            report.add('fail', 'IA_OUTLINE_LEAD_UNPLACED', f'{oid}: 一番のメッセージ {mid} がどの主張のまとまりにも置かれていない', rel, id=mid)
    ctas = [mid for mid, msg in msgs.items() if msg['type'] == 'cta']
    if o.get('no_cta'):
        if ctas:
            report.add('fail', 'IA_OUTLINE_CTA_CONTRADICTS', f'{oid}: 「CTA を置かない」と書いたのに CTA {sorted(ctas)} がある', rel)
        report.add('info', 'IA_OUTLINE_NO_CTA_INTENDED', f'{oid}: {NO_CTA_NOTE}', rel)
    elif not any(mid in placed for mid in ctas):
        report.add('fail', 'IA_OUTLINE_CTA_MISSING', f'{oid}: CTA をどのまとまりにも置いていない（置かないなら no_cta: true と書く）', rel)
    for mid in sorted(set(msgs) - set(placed) - set(leads)):
        report.add('warn', 'IA_OUTLINE_MESSAGE_UNPLACED', f'{oid}: メッセージ {mid} がどのまとまりにも置かれていない', rel, id=mid)
    proofs = (m.ssot or {}).get('proofs') or {}
    conveyed, copy_orders = set(), {}
    for cid, cp in sorted(copies.items()):
        bid = cp['block']
        if bid not in blks:
            report.add('fail', 'IA_OUTLINE_COPY_SECTION', f'{oid}.{cid}: 属するまとまり {bid} が無い', rel, id=cid)
            continue
        fn = blks[bid]['function']
        copy_orders.setdefault((bid, cp['order']), []).append(cid)
        mid = cp.get('message')
        if mid is not None and mid not in blks[bid].get('places', []):
            report.add('fail', 'IA_OUTLINE_COPY_MESSAGE',
                       f'{oid}.{cid}: 伝えるメッセージ {mid} が、属するまとまり {bid} に置かれていない', rel, id=cid)
        elif mid is not None:
            conveyed.add((bid, mid))
        elif fn == 'claim':
            report.add('fail', 'IA_OUTLINE_COPY_MESSAGE', f'{oid}.{cid}: 主張のまとまりの文案が、伝えるメッセージを指さない', rel, id=cid)
        for prf in cp.get('refs', []):
            other = (proofs.get(prf) or {}).get('segment')
            if prf in proofs and seg is not None and other != seg:
                report.add('fail', 'IA_SCRIPT_PROOF_OTHER_POSITIONING',
                           f'{oid}.{cid}: 引く証拠 {prf} のセグメントは {other}、target は {seg}', rel, id=cid)
        cited = bool(mid or cp.get('refs') or check_mod().PLACEHOLDER.search(cp['text']))
        if cited and cp.get('proposal'):
            report.add('fail', 'IA_SCRIPT_PROPOSAL_WITH_REFERENCE',
                       f'{oid}.{cid}: 引き先（メッセージ・証拠・辞書の値）があるのに提案の印を持つ', rel, id=cid)
        elif fn in CITING and not cited and not cp.get('proposal'):
            report.add('fail', 'IA_SCRIPT_COPY_UNREFERENCED',
                       f'{oid}.{cid}: 働き {fn} の文案に引き先も提案の印も無い（証拠を引くか、辞書の値を引くか、proposal: true を付ける）', rel, id=cid)
        if cp.get('proposal'):
            report.add('warn', 'IA_SCRIPT_PROPOSAL',
                       f'{oid}.{cid}: 提案の印が残っている（制作中は警告。出荷では止まる。証拠を positioning に足すか、文を消す）', rel, id=cid)
    for (bid, n), cids in sorted(copy_orders.items()):
        if len(cids) > 1:
            report.add('fail', 'IA_OUTLINE_ORDER_DUPLICATE', f'{oid}.{bid}: 文案 {sorted(cids)} の順番がどれも {n}', rel)
    for bid, blk in sorted(blks.items()):
        for mid in blk.get('places', []):
            if mid in msgs and (bid, mid) not in conveyed:
                report.add('fail', 'IA_OUTLINE_MESSAGE_UNCONVEYED', f'{oid}.{bid}: 置かれたメッセージ {mid} を伝える文案が無い', rel, id=mid)
    for key, t in sorted((o.get('texts') or {}).items()):
        if t['kind'] == 'navigation' and 'block' in t:
            report.add('fail', 'IA_OUTLINE_TEXT_MISPLACED', f'{oid}.texts.{key}: ナビの文言はまとまりに置かない', rel)
        elif t['kind'] == 'label' and 'block' not in t:
            report.add('fail', 'IA_OUTLINE_TEXT_MISPLACED', f'{oid}.texts.{key}: ラベルは置くまとまり（block）を書く', rel)
        elif 'block' in t and t['block'] not in blks:
            report.add('fail', 'IA_OUTLINE_REF_BROKEN', f'{oid}.texts.{key}: まとまり {t["block"]} が無い', rel)


def check_mod():
    import check as _check  # check imports this module: imported late
    return _check


def load_dom_report(path):
    return json.loads(open(path, encoding='utf-8').read()) if path else None
