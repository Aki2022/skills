"""Rendered DOM vs script (台本): screen conformance (SPEC-information-architecture-standardization, requirement 10).

The browser does the parsing and rendering; this module only compares its snapshot with the script: every block
and every piece of copy has exactly one visible mark (data-ia-section / data-ia-copy; the anchor names are v6's and
are kept until the factories rename them), the blocks appear exactly in the script's order, each piece of copy is
drawn inside its own block and in the block's order (order is the DOM's reading order; CSS that draws the marks in
another order fails, and moves that cannot be judged are reported as uncovered), each copy mark
shows the outline's copy word for word (facts expanded; width and whitespace folded), every fact the copy refers to
is shown with the SSOT value, and two runs give the same snapshot. A repeat is another piece of copy in the outline.
Declared non-claim text (data-ia-text) must show its text on every occurrence, and an official name (data-dict="name-…")
must show exactly the SSOT name. A v4/v5 attribute (data-ia-problem / -msg / -pain / -frame) left on the page fails.
Visible text outside every mark is text nobody manages: it fails (requirement 10) without echoing the text.
dom exit 0 means the page shows what the outline declares; it does not judge whether the claims are true.
"""
from __future__ import annotations

import json
import subprocess
import unicodedata
from pathlib import Path

import check
import outline as ol

DEFAULT_VIEWPORTS = [{'name': 'mobile', 'width': 375, 'height': 812}, {'name': 'desktop', 'width': 1280, 'height': 800}]
HTML_SUFFIXES = {'.html', '.htm'}
LEGACY = {'legacy-problem': 'data-ia-problem', 'legacy-msg': 'data-ia-msg', 'legacy-pain': 'data-ia-pain',
          'legacy-frame': 'data-ia-frame'}


def expected(outline):
    """-> ([(kind, id)] in the outline's order, [fact ids the copy refers to])."""
    seq, facts = [], []
    for sid, _, copies in ol.ordered_blocks(outline):
        seq.append(('section', sid))  # a block's anchor is data-ia-section (kind 'section')
        for cid, c in copies:
            seq.append(('copy', cid))
            facts += [f for f in check.PLACEHOLDER.findall(c['text']) if f not in facts]
    for t in (outline.get('texts') or {}).values():
        facts += [f for f in check.PLACEHOLDER.findall(t['text']) if f not in facts]
    return seq, facts


def _expand(text, ssot):
    values = {fid: str(f.get('value', '')) for fid, f in ssot['facts'].items()}
    return _copy_norm(check.PLACEHOLDER.sub(lambda m: values.get(m.group(1), m.group(0)), text))


def expected_copy(outline, ssot):
    """{copy id: normalized copy} — the text each copy anchor must show, facts expanded to their SSOT value."""
    return {cid: _expand(c['text'], ssot) for cid, c in outline['copies'].items()}


def declared_text(outline, ssot):
    """{key: normalized text} — declared non-claim text (navigation and labels)."""
    return {k: _expand(t['text'], ssot) for k, t in (outline.get('texts') or {}).items()}


def navigation_keys(outline):
    return {k for k, t in (outline.get('texts') or {}).items() if t['kind'] == 'navigation'}


def _norm(text):
    return ' '.join((text or '').split())


def _copy_norm(text):
    """Copy as read: width-folded (NFKC) and with every whitespace removed (line breaks are layout, not copy)."""
    return ''.join(unicodedata.normalize('NFKC', text or '').split())


def compare(outline, ssot, snap, output=None):
    """-> [(code, level, message)] for one snapshot (one file, one viewport)."""
    out = []
    vp = snap['viewport']['name']
    if not snap.get('text_chars'):
        return [('IA_DOM_NO_TARGETS', 'fail', f'{vp}: 描画されたテキストが 0 件（対象 0 件を合格にしない）')]
    seq, facts = expected(outline)
    anchors = snap['anchors']
    wanted = set(seq)
    for kind, attr in LEGACY.items():
        legacy = sorted({a['id'] for a in anchors if a['kind'] == kind})
        if legacy:
            out.append(('IA_DOM_LEGACY_ATTRIBUTE', 'fail', f'{vp}: 旧属性 {attr} が残っている {legacy}（v6 では data-ia-section・data-ia-copy）'))
    texts = declared_text(outline, ssot)
    for a in anchors:
        if a['kind'] in ('section', 'copy') and (a['kind'], a['id']) not in wanted:
            out.append(('IA_DOM_ANCHOR_UNKNOWN', 'fail', f'{vp}: アンカー {a["id"]} はアウトラインに無い'))
        elif a['kind'] == 'text' and a['id'] not in texts:
            out.append(('IA_DOM_ANCHOR_UNKNOWN', 'fail', f'{vp}: 宣言した文言の印 {a["id"]} はこのアウトラインに無い'))
        if a.get('nested'):
            out.append(('IA_DOM_ANCHOR_NESTED', 'fail',
                        f'{vp}: {a["id"]} の印（{a["kind"]}）の中にほかの IA の印がある（文案・宣言した文言・正式名称は、セクションや文案を包めない）'))
    # Order is the document order of the visible anchors (the reading order a screen reader follows), not the drawn
    # position: cards side by side in the script's order pass. CSS that draws them in another order fails below.
    rendered = [(a['kind'], a['id']) for a in sorted(anchors, key=lambda a: a['dom_index'])
                if a['visible'] and a['kind'] in ('section', 'copy')]
    nav = navigation_keys(outline)
    for key in dict.fromkeys(seq + [('text', k) for k in texts]):
        found = [a for a in anchors if (a['kind'], a['id']) == key]
        if not found:
            out.append(('IA_DOM_ANCHOR_MISSING', 'fail', f'{vp}: アウトラインの {key[1]} のアンカーが描画物に無い'))
        elif key[0] == 'text' and key[1] in nav:
            continue  # navigation may be collapsed at a viewport (a menu); it must exist, and what shows must match
        elif not any(a['visible'] for a in found):
            out.append(('IA_DOM_ANCHOR_HIDDEN', 'fail', f'{vp}: {key[1]} のアンカーが見えていない（display / visibility / opacity / 大きさ 0）'))
    have = [k[1] for k in rendered if k[0] == 'section']
    want = [sid for kind, sid in seq if kind == 'section' and sid in have]
    if have != want:
        out.append(('IA_DOM_ORDER', 'fail', f'{vp}: セクションの DOM の順 {have} が outline の順 {want} と違う'))
    owner = {cid: c['block'] for cid, c in outline['copies'].items()}
    for cid in sorted(owner):
        marks = [a for a in anchors if (a['kind'], a['id']) == ('copy', cid)]
        if len(marks) > 1:
            out.append(('IA_DOM_COPY_REPEATED', 'fail', f'{vp}: 文案 {cid} の印が {len(marks)} つある（繰り返すなら outline に別の文案を置く）'))
        for a in marks:
            if a.get('section') != owner[cid]:
                out.append(('IA_DOM_COPY_SECTION', 'fail', f'{vp}: 文案 {cid} が {a.get("section")} の中にある（outline では {owner[cid]}）'))
    for sid, _, copies in ol.ordered_blocks(outline):
        drawn = [a['id'] for a in sorted(anchors, key=lambda a: a['dom_index'])
                 if a['kind'] == 'copy' and a['visible'] and a.get('section') == sid and owner.get(a['id']) == sid]
        expected_order = [cid for cid, _ in copies if cid in drawn]
        if list(dict.fromkeys(drawn)) != expected_order:
            out.append(('IA_DOM_COPY_ORDER', 'fail', f'{vp}: {sid} の文案の DOM の順 {drawn} が outline の順 {expected_order} と違う'))
    for a in anchors:
        if a['kind'] not in ('section', 'copy') or not a['visible'] or (a['kind'], a['id']) not in wanted:
            continue
        if a['reorder']:
            out.append(('IA_DOM_CSS_REORDER', 'fail', f'{vp}: {a["id"]} を CSS が DOM と違う順に描いている {a["reorder"]}'
                        '（順は DOM の順で判定する。横に並べるのはよいが、見た目の順を入れ替えない）'))
        if a['unverified']:
            out.append(('IA_DOM_ORDER_UNVERIFIED', 'warn', f'{vp}: {a["id"]} の見た目の順を確かめられない {a["unverified"]}（未カバー）'))
    for cid, copy_ in expected_copy(outline, ssot).items():
        shown = [a for a in anchors if (a['kind'], a['id']) == ('copy', cid) and a['visible']]
        # The text itself is not echoed: it is unpublished copy and may be long.
        if any(_copy_norm(a['text']) != copy_ for a in shown):
            out.append(('IA_DOM_COPY_DRIFT', 'fail', f'{vp}: {cid} の表示が outline の文案と一字一句同じでない（文言を変えずにスタイリングする）'))
    for k, copy_ in texts.items():
        shown = [a for a in anchors if a['kind'] == 'text' and a['id'] == k and a['visible']]
        # Declared text has no repeat leniency: every visible occurrence is the declared text.
        if any(_copy_norm(a['text']) != copy_ for a in shown):
            out.append(('IA_DOM_COPY_DRIFT', 'fail', f'{vp}: 宣言した文言 {k} の表示がアウトラインの文案と違う（全出現で一致させる）'))
    entities = ssot.get('entities') or {}
    for a in anchors:
        if a['kind'] == 'term':
            out.append(('IA_DOM_ANCHOR_UNKNOWN', 'fail', f'{vp}: 辞書の印 {a["id"]} は用語（term）で、印にできるのは固定値（value）と正式名称（name）だけ'))
        if a['kind'] != 'entity':
            continue
        name = (entities.get(a['id']) or {}).get('name')
        if not name:
            out.append(('IA_DOM_ANCHOR_UNKNOWN', 'fail', f'{vp}: 正式名称の印 {a["id"]} は辞書に無い'))
        elif a['visible'] and _copy_norm(a['text']) != _copy_norm(name):
            out.append(('IA_DOM_ENTITY_NAME', 'fail', f'{vp}: 正式名称 {a["id"]} の表示が SSOT の名前「{name}」と違う（名前だけを包む）'))
    unanchored = _norm(snap.get('unanchored'))
    if unanchored:
        out.append(('IA_DOM_UNANCHORED_TEXT', 'fail',
                    f'{vp}: どの印にも属さない文言が {len(unanchored)} 文字ある（主張なら outline のメッセージと文案に、主張でなければ'
                    ' texts に宣言する）'))
    # A figure (data-ia-figure="<copy id>") repeats part of a copy outside it, in the copy's own block: the big number
    # above its sentence. It is read against the script's copy (values filled in), so it adds no word (2026-11-27).
    copies, owner_of = expected_copy(outline, ssot), {cid: c['block'] for cid, c in outline['copies'].items()}
    for a in anchors:
        if a['kind'] != 'figure':
            continue
        if a['id'] not in copies:
            out.append(('IA_DOM_FIGURE_UNKNOWN', 'fail', f'{vp}: 抜き出しの印 {a["id"]} の文案がアウトラインに無い'))
        elif a.get('section') != owner_of[a['id']]:
            out.append(('IA_DOM_FIGURE_SECTION', 'fail',
                        f'{vp}: 文案 {a["id"]} の抜き出しが {a.get("section")} の中にある（文案と同じ {owner_of[a["id"]]} の中に置く）'))
        elif a['visible'] and (not _copy_norm(a['text']) or _copy_norm(a['text']) not in _copy_norm(copies[a['id']])):
            out.append(('IA_DOM_FIGURE_TEXT', 'fail', f'{vp}: 文案 {a["id"]} の抜き出しが、その文案の一部でない（文案に無い文字は抜き出せない）'))
    values = {fid: f.get('value') for fid, f in ssot['facts'].items()}
    for a in anchors:
        # A fact anchor this page does not refer to may only show a registered value; otherwise it would wrap
        # arbitrary text that no other check reads.
        if a['kind'] != 'fact' or a['id'] in facts:
            continue
        if a['id'] not in values:
            out.append(('IA_DOM_ANCHOR_UNKNOWN', 'fail', f'{vp}: 固定値のアンカー {a["id"]} は辞書に無い'))
        elif _copy_norm(a['text']) != _copy_norm(values.get(a['id'])):
            out.append(('IA_DOM_FACT_VALUE', 'fail', f'{vp}: 固定値 {a["id"]} の表示が SSOT の値「{values.get(a["id"])}」と違う'))
    for fid in facts:
        shown = [a for a in anchors if a['kind'] == 'fact' and a['id'] == fid]
        if not shown:
            out.append(('IA_DOM_FACT_MISSING', 'fail', f'{vp}: 固定値 {fid} の表示（data-dict）が無い'))
        for a in shown:
            # Read like copy (whitespace removed): a value split into child elements to style part of it is read with a
            # space at each split, and a space is layout, not a different value (2026-11-26).
            if _copy_norm(a['text']) != _copy_norm(values.get(fid)):
                out.append(('IA_DOM_FACT_VALUE', 'fail', f'{vp}: 固定値 {fid} の表示「{_norm(a["text"])}」が SSOT の値「{values.get(fid)}」と違う'))
    return out


def unjudged_text(outline, ssot, snap, output=None):
    """Visible text no deterministic check vouches for: text outside every anchor, and repeats of an anchor whose
    text is not the outline's copy. Both go to the judgment ledger as claims with nothing behind them."""
    parts = [_norm(snap.get('unanchored'))]
    for cid, copy_ in expected_copy(outline, ssot).items():
        parts += [_norm(a['text']) for a in snap['anchors']
                  if (a['kind'], a['id']) == ('copy', cid) and a['visible'] and _copy_norm(a['text']) != copy_]
    return ' '.join(p for p in parts if p)


def report_dict(rep):
    """A dom report as `judge` reads it (evaluate and the ship gate both hand it on)."""
    data = rep.to_dict()
    data['first_screens'] = getattr(rep, 'first_screens', [])
    data['unanchored'] = getattr(rep, 'unanchored', [])
    return data


def compare_runs(snaps):
    first = json.dumps(snaps[0], sort_keys=True, ensure_ascii=False)
    if any(json.dumps(s, sort_keys=True, ensure_ascii=False) != first for s in snaps[1:]):
        return [('IA_DOM_UNSTABLE', 'fail', f'{snaps[0]["viewport"]["name"]}: 同じ入力で 2 回描画した結果が違う（判定に使わない）')]
    return []


def snapshot(root, html, viewports):
    proc = subprocess.run(['node', str(Path(__file__).resolve().parent / 'dom_snapshot.mjs'), str(html), json.dumps(viewports)],
                          capture_output=True, text=True, cwd=str(root), check=False, timeout=180)
    if proc.returncode != 0:
        raise RuntimeError(f'dom_snapshot が失敗した（exit {proc.returncode}）: {proc.stderr.strip()[:300]}')
    return json.loads(proc.stdout)


def run(root, outline_id, date, runs=2):
    report, m = check.check(root, date)
    report.level = 'dom'
    if m.ssot is None:
        return report
    if outline_id not in m.outlines:
        report.add('fail', 'IA_DOM_OUTLINE', f'アウトライン {outline_id} が無い')
        return report
    rel, o = m.outlines[outline_id]
    ch = m.channels.get(o['channel'])
    viewports = (ch[1].get('viewports') if ch else None) or DEFAULT_VIEWPORTS
    first_screens, unanchored, files = [], [], 0
    for out in o['outputs']:
        path = Path(root) / out
        if path.suffix.lower() not in HTML_SUFFIXES:
            report.uncovered.append({'path': out, 'reason': 'HTML 以外の成果物は DOM の検査の対象外'})
            report.add('warn', 'IA_DOM_UNCOVERED', f'{out}: HTML 以外なので DOM を確かめられない（未カバー）', out)
            continue
        if not path.is_file():
            report.add('fail', 'IA_DOM_OUTPUT_MISSING', f'{out}: 成果物のファイルが無い', out)
            continue
        try:
            series = [snapshot(root, path, viewports) for _ in range(runs)]
        except RuntimeError as exc:
            report.add('fail', 'IA_DOM_PLAYWRIGHT', str(exc), out)
            continue
        files += 1
        for i, vp in enumerate(viewports):
            per_vp = [s[i] for s in series]
            for code, level, msg in compare_runs(per_vp) or compare(o, m.ssot, per_vp[0], out):
                report.add(level, code, f'{out} {msg}', out)
                if code == 'IA_DOM_ORDER_UNVERIFIED':
                    report.uncovered.append({'path': out, 'reason': msg})
            first_screens.append({'outline': outline_id, 'file': out, 'viewport': vp['name'], 'text': per_vp[0]['first_screen']})
            loose = unjudged_text(o, m.ssot, per_vp[0], out)
            if loose:
                unanchored.append({'outline': outline_id, 'file': out, 'viewport': vp['name'], 'text': loose})
    report.targets = {'files': files, 'viewports': len(viewports) if files else 0}
    report.first_screens = first_screens
    report.unanchored = unanchored
    return report
