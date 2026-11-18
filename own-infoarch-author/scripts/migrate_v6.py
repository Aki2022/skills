"""One batch migration of design/info_arch/ from v5 to v6 (SPEC-information-architecture-standardization v6, requirement 17).

v5 is info_arch.yaml schema 3 with outlines schema 2; v6 is schema 4 with outlines schema 3. Mechanical moves only:
- promises (prm) and concepts (cpt) are dropped; the four forces of a job are dropped (v6 does not hold them);
- pains, gains, values and strengths get the segment their v5 links lead to, and those links become edges
  (has_job, need, satisfies, lacks, produces); proofs become positioning evidence (`supports` edges to what the
  messages they supported serve), with a copy per segment (`copied_from`) when they reached two segments;
- each outline's audience becomes its constant `target`, each v5 message it uses becomes an outline message pointing
  at the one value of the target segment it served, each v5 section becomes a section with one copy, frames become
  section ids (the sequence frame is the section order), navigation/other text becomes `texts`;
- the rendered v5 marks (data-ia-role / data-ia-msg / data-ia-pain / data-ia-frame) become data-ia-section and
  data-ia-copy, and design.yaml's modules.ia moves from 0.3.0 to the v6 contract version.
What needs judgment (value kinds, categories, fulfills degrees, recalls/includes, a normal message beside a CTA) is
listed under drafting_required and never invented. Anything that has no mechanical destination is unmigrated: one
item is a failure (exit 2) and nothing is written (one batch). design/approvals.yaml and design/judgments.json are
only read, to list what stops applying. Nothing is committed. A v6 tree is a no-op, so the tool is rerunnable.
"""
from __future__ import annotations

import json
import re
from collections import defaultdict
from html.parser import HTMLParser
from pathlib import Path

import yaml

import check
import common
import ia_module
import judge
import positioning as pos

V5_PIN = '0.3.0'
APPROVALS = 'design/approvals.yaml'
PROV = 'extracted:infoarch-author migrate-v6'
# v5 section role -> v6 section role (aware / pain / compare / trust / act). A fixed table, printed in the report.
ROLE = {'pain': 'pain', 'benefit': 'aware', 'vision': 'aware', 'proof': 'trust', 'how': 'trust',
        'objection': 'compare', 'cta': 'act', 'navigation': 'aware', 'other': 'aware'}
TEXT_ROLES = {'navigation', 'other'}
LEGACY_ATTRS = ('data-ia-problem', 'data-ia-msg', 'data-ia-pain', 'data-ia-frame')
LEGACY_RE = re.compile(r'\b(?:' + '|'.join(LEGACY_ATTRS) + r')\s*=')
V4_PREFIXES = ('prf', 'fct', 'ent', 'vce', 'seg', 'job', 'pan', 'gan', 'val', 'alt', 'str', 'cat')
FORCES = ('push', 'pull', 'anxiety', 'habit')
HEADER = ('# design/info_arch/info_arch.yaml — 伝え方の SSOT（スキーマ版 4）\n'
          '# v5 から infoarch-author migrate-v6 で移した。提供価値の kind・カテゴリー・満たす度合いなどは下書きが要る。\n')
OUTLINE_HEADER = '# v5 の outline から infoarch-author migrate-v6 で移した（版 3）。\n'


def _dump(data):
    return yaml.safe_dump(data, allow_unicode=True, sort_keys=False)


def _prefix(aid):
    return str(aid).split('-')[0]


class _Run:
    def __init__(self):
        self.id_map = defaultdict(list)
        self.unmigrated, self.dropped, self.drafting = [], [], []

    def fail(self, item, reason, **detail):
        self.unmigrated.append({'item': item, 'reason': reason, **detail})

    def drop(self, item, reason):
        self.dropped.append({'item': item, 'reason': reason})


# ---- SSOT ------------------------------------------------------------------------------------------------------------

def _atom(old, **fields):
    out = {'state': old.get('state', 'empty'), 'provenance': old.get('provenance', 'proposal')}
    out.update({k: v for k, v in fields.items() if v is not None})
    return out


def _edge(kind, frm, to, source):
    return {'kind': kind, 'from': frm, 'to': to, 'state': 'proposed', 'provenance': f'{PROV} ({source})'}


def _values_of(ssot, mid):
    """Positioning atoms a v5 message served: values (direct, or through a promise) and jobs, pains and gains."""
    msg = (ssot.get('messages') or {}).get(mid) or {}
    out = []
    for x in msg.get('serves') or []:
        if _prefix(x) == 'prm':
            out += ((ssot.get('promises') or {}).get(x) or {}).get('values') or []
        elif _prefix(x) in ('val', 'job', 'pan', 'gan'):
            out.append(x)
    return list(dict.fromkeys(out))


def migrate_ssot(ssot, r):
    """-> (v4 SSOT, {atom id: segment})."""
    new = {'schema_version': 4, 'ledger': None}
    if 'policy' in ssot:
        new['policy'] = ssot['policy']
    edges, seg_of = [], {}
    proposed = lambda a: a.get('state') == 'proposed'  # noqa: E731
    new['segments'] = {sid: _atom(s, qualifier=s.get('qualifier'), situation=s.get('situation'))
                       for sid, s in (ssot.get('segments') or {}).items()}
    seg_of.update({sid: sid for sid in new['segments']})

    def one_segment(aid, segs, block):
        segs = sorted({s for s in segs if s})
        if len(segs) == 1:
            return segs[0]
        r.fail(aid, 'atom-segment', block=block, segments=segs)
        return None

    new['jobs'] = {}
    for jid, j in (ssot.get('jobs') or {}).items():
        seg = j.get('segment')
        if proposed(j):
            seg = one_segment(jid, [seg], 'jobs')
            if seg:
                edges.append(_edge('has_job', seg, jid, f'{jid}.segment'))
        seg_of[jid] = seg
        new['jobs'][jid] = _atom(j, segment=seg, text=j.get('text'))
        if any(j.get(f) for f in FORCES):
            r.drop(f'{jid}.{"/".join(f for f in FORCES if j.get(f))}', 'v6 はジョブの 4 つの力を持たない（Deferred Q13）')
    for block in ('pains', 'gains'):
        new[block] = {}
        for nid, n in (ssot.get(block) or {}).items():
            job = n.get('job')
            seg = seg_of.get(job)
            if proposed(n):
                seg = one_segment(nid, [seg], block)
                if seg:
                    edges.append(_edge('need', job, nid, f'{nid}.job'))
            seg_of[nid] = seg
            new[block][nid] = _atom(n, segment=seg, text=n.get('text'))
    new['alternatives'] = {}
    for aid, a in (ssot.get('alternatives') or {}).items():
        seg_of[aid] = a.get('segment')
        new['alternatives'][aid] = _atom(a, segment=a.get('segment'), text=a.get('text'))
    new['values'] = {}
    for vid, v in (ssot.get('values') or {}).items():
        needs = (v.get('relieves') or []) + (v.get('creates') or [])
        seg = one_segment(vid, [seg_of.get(n) for n in needs], 'values') if proposed(v) else None
        if seg:
            edges += [_edge('satisfies', vid, n, f'{vid}.relieves/creates') for n in needs]
            r.drafting.append(f'values:{vid}.kind（difference か parity。v5 に無いので作らない）')
        seg_of[vid] = seg
        new['values'][vid] = _atom(v, segment=seg, text=v.get('text'))
    new['strengths'] = {}
    for sid, s in (ssot.get('strengths') or {}).items():
        links = [seg_of.get(a) for a in s.get('versus') or []] + [seg_of.get(v) for v in s.get('produces') or []]
        seg = one_segment(sid, links, 'strengths') if proposed(s) else None
        if seg:
            edges += [_edge('lacks', sid, a, f'{sid}.versus') for a in s.get('versus') or []]
            edges += [_edge('produces', sid, v, f'{sid}.produces') for v in s.get('produces') or []]
        new['strengths'][sid] = _atom(s, segment=seg, text=s.get('text'))
    new['categories'] = {}
    for cid, c in (ssot.get('concepts') or {}).items():
        r.id_map[cid] = []
        r.drop(cid, '約束とコンセプトは v6 に無い。カテゴリーの材料として drafting_required に出す')
        if c.get('category'):
            r.drafting.append(f'categories: {cid} のカテゴリー「{c["category"]}」を材料に、セグメントごとに cat と member_of・recalls・includes を下書きする')
    for pid in ssot.get('promises') or {}:
        r.id_map[pid] = []
        r.drop(pid, '約束は v6 で無くした（要件 17）')
    segs = sorted(sid for sid, s in new['segments'].items() if proposed(s))
    for seg in segs:
        r.drafting.append(f'{seg}: 所属カテゴリー（cat・member_of ちょうど 1 本）、ジョブの recalls、カテゴリーの includes、'
                          '代替 × 痛み・得たいものの fulfills（度合い full / partial / none。欠けると導けない）')
    nxt = dict((ssot.get('ledger') or {}).get('next') or {})
    proofs = {}
    for fid, f in (ssot.get('proofs') or {}).items():
        base = dict(text=f.get('text'), source=f.get('source'), as_of=f.get('as_of'), expires=f.get('expires'))
        if not proposed(f):
            proofs[fid] = _atom(f, **base)
            continue
        by_seg = defaultdict(list)
        for mid in f.get('supports') or []:
            for t in _values_of(ssot, mid):
                if seg_of.get(t):
                    by_seg[seg_of[t]].append(t)
        if not by_seg:
            r.fail(fid, 'proof-without-target', supports=f.get('supports') or [])
            continue
        for i, seg in enumerate(sorted(by_seg)):
            pid = fid
            if i:
                pid = f'prf-{nxt.get("prf", 1):04d}'
                nxt['prf'] = nxt.get('prf', 1) + 1
            proofs[pid] = _atom(f, **base, segment=seg)
            if i:  # the proof schema has no copied_from: the copy names its source in provenance
                proofs[pid]['provenance'] = f'{PROV} (copy of {fid})'
            edges += [_edge('supports', pid, t, f'{fid}.supports') for t in dict.fromkeys(by_seg[seg])]
            r.id_map[fid].append(pid)
    new['proofs'] = proofs
    for block in ('facts', 'entities', 'voices'):
        new[block] = ssot.get(block) or {}
    tombs = {}
    for old, target in ((ssot.get('ledger') or {}).get('tombstones') or {}).items():
        if _prefix(old) in V4_PREFIXES and (target is None or _prefix(target) in V4_PREFIXES):
            tombs[old] = target
        else:
            r.drop(f'ledger.tombstones.{old}', 'v6 に無い種類の墓標')
    new['ledger'] = {'next': {p: int(nxt.get(p, 1)) for p in V4_PREFIXES}, 'tombstones': tombs}
    new['edges'] = list({(e['kind'], e['from'], e['to']): e for e in edges}.values())
    order = ['schema_version', 'ledger', 'policy', 'segments', 'jobs', 'pains', 'gains', 'alternatives', 'categories',
             'strengths', 'values', 'edges', 'proofs', 'facts', 'entities', 'voices']
    return {k: new[k] for k in order if k in new}, seg_of


# ---- outlines --------------------------------------------------------------------------------------------------------

def _resolve_pain(ssot, seg_of, seg, pain, candidates):
    """A pain-only v5 section: the one message of this outline that addressed the pain, else whose value meets it."""
    msgs = ssot.get('messages') or {}
    hit = [m for m in candidates if pain in ((msgs.get(m) or {}).get('addresses') or [])]
    if not hit:
        vals = ssot.get('values') or {}
        hit = [m for m in candidates
               if any(pain in ((vals.get(v) or {}).get('relieves') or []) + ((vals.get(v) or {}).get('creates') or [])
                      for v in _values_of(ssot, m) if seg_of.get(v) == seg)]
    return hit[0] if len(hit) == 1 else None


def migrate_outline(ssot, seg_of, o, r, used):
    """-> [(new outline, [v5 sections with their new ids])] — one per page."""
    oid = o['id']
    aud = (ssot.get('audiences') or {}).get(o.get('audience')) or {}
    seg = aud.get('segment')
    if not seg or seg not in (ssot.get('segments') or {}):
        r.fail(f'{oid}.audience', 'audience-without-segment', audience=o.get('audience'))
        return []
    pages, outputs = o.get('pages') or [], o.get('outputs') or []
    if len(pages) == 1:
        parts = [(oid, pages[0], outputs)]
    elif len(pages) == len(outputs):
        parts = [(f'{oid}-{p["id"]}', p, [out]) for p, out in zip(pages, outputs)]
    else:
        r.fail(oid, 'pages-outputs-mismatch', pages=len(pages), outputs=len(outputs))
        return []
    brief = o.get('brief') or {}
    for k in ('strategy', 'brief'):
        if o.get(k):
            r.drop(f'{oid}.{k}', 'outline v3 に無い（取ってほしい行動だけを objective に移した）' if k == 'brief'
                   else 'outline v3 に無い（選んだ理由は set の id と差分が持つ）')
    msgs = ssot.get('messages') or {}
    result = []
    for nid, page, outs in parts:
        page_msgs = [s['message'] for s in page['sections'] if s.get('message')]
        messages, sections, copies, texts, placed = {}, {}, {}, {}, []
        n, last_sec = 0, None
        for i, s in enumerate(page['sections']):
            key = f'{oid}:{page["id"]}[{i}]'
            if s.get('role') in TEXT_ROLES and s.get('key') and not s.get('message') and not s.get('pain'):
                kind = 'navigation' if s['role'] == 'navigation' else 'label'
                texts[s['key']] = {'kind': kind, 'text': s['copy'], **({'section': last_sec} if kind == 'label' else {})}
                if kind == 'label':
                    r.drafting.append(f'{nid}.texts.{s["key"]}: ラベルを置くセクション（直前の {last_sec} にした）を確かめる')
                r.id_map[key] = [f'{nid}:texts.{s["key"]}']
                placed.append((s, None, None, None))
                continue
            if s.get('role') not in ROLE:
                r.fail(key, 'section-without-message', role=s.get('role'))
                continue
            mid = s.get('message')
            if not mid and s.get('pain'):
                mid = _resolve_pain(ssot, seg_of, seg, s['pain'], list(dict.fromkeys(page_msgs)))
            if not mid:
                r.fail(key, 'section-without-message', pain=s.get('pain'))
                continue
            if mid not in messages:
                vals = [v for v in _values_of(ssot, mid) if _prefix(v) == 'val' and seg_of.get(v) == seg]
                if len(vals) != 1:
                    r.fail(f'{nid}:{mid}', 'message-value-none' if not vals else 'message-value-several',
                           target_segment=seg, values=vals)
                    continue
                cta = (msgs.get(mid) or {}).get('role') == 'cta'
                messages[mid] = {'value': vals[0], 'rank': 'support', 'type': 'cta' if cta else 'normal',
                                 **({'achieves': 'obj-0001'} if cta else {})}
                r.id_map[mid].append(f'{nid}:{mid}')
                used.add(mid)
            n += 1
            sid, cid = f'sec-{n:04d}', f'cpy-{n:04d}'
            last_sec = sid
            sections[sid] = {'order': n, 'role': ROLE[s['role']], 'places': [mid]}
            copies[cid] = {'section': sid, 'order': 1, 'message': mid, 'text': s['copy']}
            r.id_map[key] = [f'{nid}:{sid}', f'{nid}:{cid}']
            if s.get('frame'):
                r.id_map[f'{oid}#{s["frame"]}'] = [f'{nid}:{sid}']
            if s.get('proofs'):
                r.drop(f'{key}.proofs', '証拠は提供価値の経路が持つ（outline は指さない）')
            if s['role'] in ('navigation', 'other', 'how', 'vision', 'objection'):
                r.drafting.append(f'{nid}.{sid}: v5 の役割 {s["role"]} を {ROLE[s["role"]]} にした。役割を確かめる')
            if messages[mid]['type'] == 'cta':
                r.drafting.append(f'{nid}.{sid}: CTA だけのセクション。通常のメッセージを 1 つ置き、それを伝える文案を足すか、'
                                  '前のセクションに CTA を移す')
            placed.append((s, sid, cid, mid))
        for inst in (o.get('frames') or {}).get('instances') or []:
            if f'{oid}#{inst["id"]}' not in r.id_map:
                r.id_map[f'{oid}#{inst["id"]}'] = [f'{nid}:sections.order'] if inst.get('type') == 'section-order' else []
        normal = [m for m, v in messages.items() if v['type'] == 'normal']
        lead = next((m for m in aud.get('priorities') or [] if m in normal), normal[0] if normal else None)
        if lead:
            messages[lead]['rank'] = 'lead'
        else:
            r.drafting.append(f'{nid}: 通常のメッセージが無いので、一番のメッセージを決める')
        if not any(v['type'] == 'cta' for v in messages.values()):
            r.drafting.append(f'{nid}: CTA が無い。置くか、置かないなら no_cta: true と書く')
        medium = str(brief.get('medium', '')).lower()
        touchpoint = medium if re.fullmatch(r'[a-z][a-z0-9_-]*', medium) else o['channel']
        if touchpoint != medium:
            r.drafting.append(f'{nid}.touchpoint: v5 の媒体「{brief.get("medium")}」を形にできないのでチャネル名 {touchpoint} にした。確かめる')
        new = {'schema_version': 3, 'id': nid, 'channel': o['channel'], 'outputs': outs,
               'target': {'segment': seg, 'condition': aud.get('label') or aud.get('situation') or seg},
               'touchpoint': touchpoint}
        if aud.get('tone'):
            new['tone'] = aud['tone']
        else:
            r.drafting.append(f'{nid}.tone: v5 の対象者にトーンが無い。書く')
        new['objective'] = {'id': 'obj-0001', 'action': brief.get('action') or o['channel']}
        new.update(messages=messages, sections=sections, copies=copies)
        if texts:
            new['texts'] = texts
        r.drafting.append(f'{nid}.target.condition: v5 の対象者の名前を写した。セグメントに足す条件として確かめる')
        r.id_map[o['audience']].append(f'{nid}:target')
        result.append((new, placed))
    return result


# ---- rendered HTML ---------------------------------------------------------------------------------------------------

# A wrapper is a div only around block content inside a container that may hold one; otherwise a span. No inline
# style: `display: contents` would leave the copy without a box, and dom orders copies by where they are drawn.
BLOCK = {'address', 'article', 'aside', 'blockquote', 'details', 'div', 'dl', 'fieldset', 'figure', 'footer', 'form',
         'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'header', 'hr', 'li', 'main', 'nav', 'ol', 'p', 'pre', 'section', 'table',
         'ul'}
PHRASING = {'a', 'b', 'button', 'em', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'i', 'label', 'p', 'small', 'span', 'strong'}
VOID = {'area', 'base', 'br', 'col', 'embed', 'hr', 'img', 'input', 'link', 'meta', 'source', 'track', 'wbr'}


class _Tree(HTMLParser):
    """Elements with their source offsets: start, end of the start tag, start and end of the end tag, parent."""

    def __init__(self, text):
        super().__init__(convert_charrefs=True)
        self.lines = [0]
        for line in text.splitlines(keepends=True):
            self.lines.append(self.lines[-1] + len(line))
        self.text, self.elems, self.stack = text, [], []
        self.feed(text)
        self.close()

    def _offset(self):
        line, col = self.getpos()
        return self.lines[line - 1] + col

    def _open(self, tag, attrs, closed):
        start = self._offset()
        raw = self.get_starttag_text()
        e = {'tag': tag, 'attrs': dict(attrs), 'start': start, 'start_end': start + len(raw),
             'end': None, 'end_end': None, 'parent': self.stack[-1] if self.stack else None}
        self.elems.append(e)
        if not closed and tag not in VOID:
            self.stack.append(len(self.elems) - 1)

    def handle_starttag(self, tag, attrs):
        self._open(tag, attrs, False)

    def handle_startendtag(self, tag, attrs):
        self._open(tag, attrs, True)

    def handle_endtag(self, tag):
        for depth in range(len(self.stack) - 1, -1, -1):
            if self.elems[self.stack[depth]]['tag'] == tag:
                e = self.elems[self.stack[depth]]
                e['end'] = self._offset()
                e['end_end'] = self.text.index('>', e['end']) + 1
                del self.stack[depth:]
                return


def _start_tag(text, e, new_attrs):
    tag = text[e['start']:e['start_end']]
    tag = re.sub(r'\s+data-ia-(?:msg|pain|problem|frame|role)\s*=\s*(?:"[^"]*"|\'[^\']*\'|[^\s>]+)', '', tag)
    close = '/>' if tag.endswith('/>') else '>'
    return tag[:-len(close)].rstrip() + ''.join(f' {k}="{v}"' for k, v in new_attrs) + close


def rewrite_html(text, placed):
    """-> (new text, error or None). `placed`: the v5 sections of the outline in order, with their new ids."""
    tree = _Tree(text)
    marked = [i for i, e in enumerate(tree.elems) if 'data-ia-role' in e['attrs']]
    if len(marked) != len(placed):
        return None, f'v5 のセクションの印が {len(marked)} 個、outline のセクションは {len(placed)} 個'
    edits = []  # (start, end, replacement)
    for i, (s, sid, cid, _) in zip(marked, placed):
        e = tree.elems[i]
        a = e['attrs']
        same = (a.get('data-ia-msg'), a.get('data-ia-pain')) == (s.get('message'), s.get('pain'))
        if not same or (sid is None and a.get('data-ia-text') != s.get('key')):
            return None, f'{e["tag"]}（{a}）が outline のセクション（{s.get("message") or s.get("pain") or s.get("key")}）と合わない'
        if sid is None:
            edits.append((e['start'], e['start_end'], _start_tag(text, e, [])))
            continue
        if e['end'] is None:
            return None, f'{e["tag"]} の閉じタグが無い（{sid}）'
        role = ROLE[s['role']]
        edits.append((e['start'], e['start_end'], _start_tag(text, e, [('data-ia-section', sid), ('data-ia-role', role)])))
        kids = [c for c in tree.elems if c['parent'] == i]
        rest = text[e['start_end']:e['end']]
        for c in sorted(kids, key=lambda c: -c['start']):
            stop = (c['end_end'] or c['start_end']) - e['start_end']
            rest = rest[:c['start'] - e['start_end']] + rest[stop:]
        rest = re.sub(r'<!--.*?-->', '', rest, flags=re.S)
        if len(kids) == 1 and kids[0]['end'] is not None and not rest.strip():
            k = kids[0]
            edits.append((k['start'], k['start_end'], _start_tag(text, k, [('data-ia-copy', cid)])))
        else:  # bare text or several elements: wrap them in an element with a real box, where dom reads its position
            tag = 'div' if e['tag'] not in PHRASING and any(c['tag'] in BLOCK for c in kids) else 'span'
            edits.append((e['start_end'], e['start_end'], f'<{tag} data-ia-copy="{cid}">'))
            edits.append((e['end'], e['end'], f'</{tag}>'))
    for start, end, repl in sorted(edits, key=lambda x: (-x[0], -x[1])):
        text = text[:start] + repl + text[end:]
    if LEGACY_RE.search(text):
        return None, '書き換えた後も旧い印が残る（セクションの外の data-ia-msg など）'
    return text, None


# ---- design.yaml, ledgers --------------------------------------------------------------------------------------------

def _design_pin(root, r, writes):
    path = Path(root) / 'design/design.yaml'
    if not path.is_file():
        return None
    raw = path.read_text(encoding='utf-8')
    cfg = json.loads(raw)
    pin = (cfg.get('modules') or {}).get('ia')
    if pin is None or pin == ia_module.VERSION:
        return {'from': pin, 'to': pin}
    if pin != V5_PIN:
        r.fail('design/design.yaml modules.ia', 'module-pin-unknown', pin=pin, expected=[V5_PIN, ia_module.VERSION])
        return {'from': pin, 'to': None}
    cfg['modules']['ia'] = ia_module.VERSION
    writes[path] = json.dumps(cfg, ensure_ascii=False, indent=2 if raw.startswith('{\n') else None) + ('\n' if raw.endswith('\n') else '')
    return {'from': pin, 'to': ia_module.VERSION}


def stops_applying(root):
    """Owner approvals and judgments of the v5 tree that no v6 key reads (listed, never written)."""
    out = {'approvals': [], 'judgments': []}
    path = Path(root) / APPROVALS
    if path.is_file():
        for row in common.load_yaml(path) or []:
            key = str((row or {}).get('key', ''))
            if key.startswith('ia:') and not key.startswith(pos.PATH_PREFIX):
                legacy = key.startswith(pos.LEGACY_PREFIX)
                out['approvals'].append({'key': key, 'approver': row.get('approver'), 'date': row.get('date'),
                                         'why': ('原子と矢印ごとの行。infoarch-author migrate-approvals で経路の行へ移せるか確かめる' if legacy
                                                 else 'v5 の戦略の層の確定。v6 は提供価値ごとの経路（ia:path:）をオーナーが新たに承認する')})
    ledger = judge.read_ledger(root)
    for row in (ledger or {}).get('rows', []):
        cand = (row or {}).get('candidate') or {}
        kind = str((cand.get('input') or {}).get('kind', ''))
        if kind.startswith(judge.IA_PREFIX):
            out['judgments'].append({'id': cand.get('id'), 'kind': kind,
                                     'why': 'v6 の判定は v6 の outline の id と文案・提供価値の中身から鍵を作るので、v5 の IA の判定は読まれない'})
    out['approvals'].sort(key=lambda x: x['key'])
    out['judgments'].sort(key=lambda x: str(x['id']))
    return out


# ---- the batch -------------------------------------------------------------------------------------------------------

def plan(root):
    """-> (report, {path: new text}, [paths to delete]) without writing anything."""
    root = Path(root)
    base = root / check.INFO_ARCH_DIR
    ssot = common.load_yaml(base / 'info_arch.yaml')
    version = (ssot or {}).get('schema_version')
    if version == 4:
        return {'status': 'already-v6', 'unmigrated': [], 'unmigrated_count': 0, 'exit_code': 0, 'written': False}, {}, []
    if version != 3:
        raise ValueError(f'info_arch.yaml のスキーマ版 {version!r} は v5（3）でも v6（4）でもない')
    r = _Run()
    new_ssot, seg_of = migrate_ssot(ssot, r)
    writes, deletes, used, html_parts = {base / 'info_arch.yaml': HEADER + _dump(new_ssot)}, [], set(), {}
    for f in sorted((base / 'outlines').glob('*.yaml')):
        o = common.load_yaml(f)
        if o.get('schema_version') != 2:
            raise ValueError(f'{f.name}: outline の版 {o.get("schema_version")!r} は v5（2）でない')
        parts = migrate_outline(ssot, seg_of, o, r, used)
        for new, placed in parts:
            writes[base / 'outlines' / f'{new["id"]}.yaml'] = OUTLINE_HEADER + _dump(new)
            for out in new['outputs']:
                html_parts[out] = placed
        if parts and all(new['id'] != o['id'] for new, _ in parts):
            deletes.append(f)
    for mid, msg in (ssot.get('messages') or {}).items():
        if mid not in used:
            if msg.get('state') == 'proposed':
                r.fail(mid, 'message-unused', text=msg.get('text'))
            else:
                r.id_map[mid] = []
                r.drop(mid, '空欄のメッセージ（outline の原子にしない）')
    for aid, a in (ssot.get('audiences') or {}).items():
        if aid not in r.id_map:
            r.id_map[aid] = []
            r.drop(aid, 'どの outline も使わない対象者（セグメントはそのまま残る）')
    globs = []
    for f in sorted((base / 'channels').glob('*.yaml')):
        globs += [check._glob_re(g) for g in (common.load_yaml(f) or {}).get('outputs', [])]
    pages = sorted(p for p in root.rglob('*') if p.is_file() and p.suffix.lower() in check.HTML_SUFFIXES
                   and '.git' not in p.parts and 'node_modules' not in p.parts)
    for p in pages:
        rel = p.relative_to(root).as_posix()
        text = p.read_text(encoding='utf-8')
        if rel in html_parts:
            new, error = rewrite_html(text, html_parts[rel])
            if error:
                r.fail(rel, 'html-unmatched', detail=error)
            elif new != text:
                writes[p] = new
        elif (any(g.match(rel) for g in globs) or 'data-ia-' in text) and LEGACY_RE.search(text):
            r.fail(rel, 'html-unmatched', detail='どの outline の成果物でもない描画物に旧い印がある')
    pin = _design_pin(root, r, writes)
    prose = base / 'INFO_ARCH.md'
    mentions = sorted(set(re.findall(r'\b(?:aud|msg|prm|cpt)-\d{4,}\b', prose.read_text(encoding='utf-8')))) if prose.is_file() else []
    created = sorted(p.relative_to(root).as_posix() for p in writes if not p.exists())
    report = {
        'status': 'preview',
        'id_map': dict(sorted(r.id_map.items())),
        'id_map_note': '載っていない id（セグメント・ジョブ・痛み・代替・固定値など）は同じ id のまま。[] は無くしたもの',
        'unmigrated': r.unmigrated, 'unmigrated_count': len(r.unmigrated),
        'dropped': r.dropped, 'drafting_required': r.drafting, 'role_map': ROLE,
        'stops_applying': stops_applying(root),
        'module_pin': pin, 'prose_mentions': {'design/info_arch/INFO_ARCH.md': mentions} if mentions else {},
        'set_ids': 'drafting_required を済ませて check が緑になってから infoarch-author set --outline <id> で出す（v5 の対応表の置き換え）',
        'files': {'write': sorted(p.relative_to(root).as_posix() for p in writes if p.exists()), 'create': created,
                  'delete': sorted(p.relative_to(root).as_posix() for p in deletes)},
        'rollback': [
            'この道具は commit しない。--write の後、commit する前なら git restore -- <files.write と files.delete> で戻し、'
            'files.create のファイルを 1 つずつ消す',
            'commit した後なら git revert <移行の commit>（info_arch.yaml・outlines・描画物・design.yaml が v5 に戻る）',
            'design/approvals.yaml と design/judgments.json は書いていないので戻すものは無い。v6 で足した ia:path: の行と判定は、'
            'v5 の道具では数えられない',
        ],
        'written': False,
    }
    report['exit_code'] = 2 if r.unmigrated else 0
    return report, writes, deletes


def run(root, write=False):
    report, writes, deletes = plan(root)
    if report['status'] == 'already-v6' or report['exit_code'] or not write:
        return report
    for p, text in writes.items():
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(text, encoding='utf-8')
    for p in deletes:
        p.unlink()
    report.update(status='migrated', written=True)
    return report
