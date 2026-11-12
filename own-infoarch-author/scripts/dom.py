"""Rendered DOM vs structured outline (ISSUE-03).

The browser does the parsing and rendering; this module only compares its snapshot with the outline:
every message / problem the outline uses has a visible anchor, the anchors appear in the outline's order
(as a subsequence of the rendered order, so a repeated call to action is fine), each section's visible text is
the outline's copy (facts expanded), every fact the copy refers to is shown with the SSOT value, and two runs give
the same snapshot. Visible text outside every anchor is reported, and passed on to the judgement ledger as a claim. The first-screen text is passed on to
the Jev judgement ledger (ISSUE-04); nothing here judges meaning.
"""
from __future__ import annotations

import json
import subprocess
import unicodedata
from pathlib import Path

import check

DEFAULT_VIEWPORTS = [{'name': 'mobile', 'width': 375, 'height': 812}, {'name': 'desktop', 'width': 1280, 'height': 800}]
HTML_SUFFIXES = {'.html', '.htm'}


def expected_for(outline, output):
    """Anchors and facts one output file must show: its own page when pages and outputs pair up 1:1."""
    pages = outline['pages']
    if len(pages) == len(outline['outputs']):
        pages = [pages[outline['outputs'].index(output)]]
    return expected({'pages': pages})


def expected(outline):
    seq, facts = [], []
    for p in outline['pages']:
        for s in p['sections']:
            if 'problem' in s:
                seq.append(('problem', s['problem']))
            if 'message' in s:
                seq.append(('msg', s['message']))
            facts += [f for f in check.PLACEHOLDER.findall(s['copy']) if f not in facts]
    return seq, facts


def _section_pages(outline, output):
    pages = outline['pages']
    if output and len(pages) == len(outline['outputs']):
        pages = [pages[outline['outputs'].index(output)]]
    return pages


def expected_copy(outline, ssot, output=None):
    """{(kind, id): {normalized copy}} — the text each anchor must show, facts expanded to their SSOT value."""
    values = {fid: str(f.get('value', '')) for fid, f in ssot['facts'].items()}
    out = {}
    for p in _section_pages(outline, output):
        for s in p['sections']:
            text = _copy_norm(check.PLACEHOLDER.sub(lambda m: values.get(m.group(1), m.group(0)), s['copy']))
            # Every anchor the section carries shows the same copy (skeleton puts both attributes on one element).
            for kind, field in (('problem', 'problem'), ('msg', 'message')):
                if field in s:
                    out.setdefault((kind, s[field]), set()).add(text)
    return out


def _norm(text):
    return ' '.join((text or '').split())


def _copy_norm(text):
    """Copy as read: width-folded (NFKC) and with every whitespace removed (line breaks are layout, not copy)."""
    return ''.join(unicodedata.normalize('NFKC', text or '').split())


def _is_subsequence(want, have):
    it = iter(have)
    return all(any(w == h for h in it) for w in want)


def compare(outline, ssot, snap, output=None):
    """-> [(code, level, message)] for one snapshot (one file, one viewport)."""
    out = []
    vp = snap['viewport']['name']
    if not snap.get('text_chars'):
        return [('IA_DOM_NO_TARGETS', 'fail', f'{vp}: 描画されたテキストが 0 件（対象 0 件を合格にしない）')]
    seq, facts = expected_for(outline, output) if output else expected(outline)
    anchors = snap['anchors']
    wanted = set(seq)
    for a in anchors:
        if a['kind'] in ('msg', 'problem') and (a['kind'], a['id']) not in wanted:
            out.append(('IA_DOM_ANCHOR_UNKNOWN', 'fail', f'{vp}: アンカー {a["id"]} はアウトラインに無い'))
    rendered = [(a['kind'], a['id']) for a in sorted(anchors, key=lambda a: (a['top'], a['left']))
                if a['visible'] and a['kind'] in ('msg', 'problem')]
    for key in dict.fromkeys(seq):
        found = [a for a in anchors if (a['kind'], a['id']) == key]
        if not found:
            out.append(('IA_DOM_ANCHOR_MISSING', 'fail', f'{vp}: アウトラインの {key[1]} のアンカーが描画物に無い'))
        elif not any(a['visible'] for a in found):
            out.append(('IA_DOM_ANCHOR_HIDDEN', 'fail', f'{vp}: {key[1]} のアンカーが見えていない（display / visibility / opacity / 大きさ 0）'))
    present = [k for k in seq if k in set(rendered)]
    if not _is_subsequence(present, rendered):
        out.append(('IA_DOM_ORDER', 'fail', f'{vp}: 描画の順 {[k[1] for k in rendered]} がアウトラインの順 {[k[1] for k in seq]} と違う'))
    for key, copies in expected_copy(outline, ssot, output).items():
        shown = [a for a in anchors if (a['kind'], a['id']) == key and a['visible']]
        if not shown:
            continue  # missing or hidden is reported above
        matches = [a for a in shown if _copy_norm(a['text']) in copies]
        # The text itself is not echoed: it is unpublished copy and may be long.
        if not matches:
            out.append(('IA_DOM_COPY_DRIFT', 'fail', f'{vp}: {key[1]} の表示がアウトラインの文案と違う（文言を変えずにスタイリングする）'))
        elif len(matches) < len(shown):
            out.append(('IA_DOM_COPY_EXTRA', 'warn', f'{vp}: {key[1]} の繰り返し {len(shown) - len(matches)} か所が文案と違う'))
    unanchored = _norm(snap.get('unanchored'))
    if unanchored:
        out.append(('IA_DOM_UNANCHORED_TEXT', 'warn',
                    f'{vp}: どのアンカーにも属さない表示の文言が {len(unanchored)} 文字ある（判定台帳で主張として判定する）'))
    values = {fid: f.get('value') for fid, f in ssot['facts'].items()}
    for a in anchors:
        # A fact anchor this page does not refer to may only show a registered value; otherwise it would wrap
        # arbitrary text that no other check reads.
        if a['kind'] != 'fact' or a['id'] in facts:
            continue
        if a['id'] not in values:
            out.append(('IA_DOM_ANCHOR_UNKNOWN', 'fail', f'{vp}: 固定値のアンカー {a["id"]} は SSOT に無い'))
        elif _norm(a['text']) != _norm(values.get(a['id'])):
            out.append(('IA_DOM_FACT_VALUE', 'fail', f'{vp}: 固定値 {a["id"]} の表示が SSOT の値「{values.get(a["id"])}」と違う'))
    for fid in facts:
        shown = [a for a in anchors if a['kind'] == 'fact' and a['id'] == fid]
        if not shown:
            out.append(('IA_DOM_FACT_MISSING', 'fail', f'{vp}: 固定値 {fid} の表示（data-ia-fact）が無い'))
        for a in shown:
            if _norm(a['text']) != _norm(values.get(fid)):
                out.append(('IA_DOM_FACT_VALUE', 'fail', f'{vp}: 固定値 {fid} の表示「{_norm(a["text"])}」が SSOT の値「{values.get(fid)}」と違う'))
    return out


def unjudged_text(outline, ssot, snap, output=None):
    """Visible text no deterministic check vouches for: text outside every anchor, and repeats of an anchor whose
    text is not the outline's copy. Both go to the judgment ledger as claims with nothing behind them."""
    parts = [_norm(snap.get('unanchored'))]
    for key, copies in expected_copy(outline, ssot, output).items():
        parts += [_norm(a['text']) for a in snap['anchors']
                  if (a['kind'], a['id']) == key and a['visible'] and _copy_norm(a['text']) not in copies]
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
            first_screens.append({'outline': outline_id, 'file': out, 'viewport': vp['name'], 'text': per_vp[0]['first_screen']})
            loose = unjudged_text(o, m.ssot, per_vp[0], out)
            if loose:
                unanchored.append({'outline': outline_id, 'file': out, 'viewport': vp['name'], 'text': loose})
    report.targets = {'files': files, 'viewports': len(viewports) if files else 0}
    report.first_screens = first_screens
    report.unanchored = unanchored
    return report
