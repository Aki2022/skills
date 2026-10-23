"""Rendered DOM vs structured outline (ISSUE-03).

The browser does the parsing and rendering; this module only compares its snapshot with the outline:
every message / problem the outline uses has a visible anchor, the anchors appear in the outline's order
(as a subsequence of the rendered order, so a repeated call to action is fine), every fact the copy refers
to is shown with the SSOT value, and two runs give the same snapshot. The first-screen text is passed on to
the Jev judgement ledger (ISSUE-04); nothing here judges meaning.
"""
from __future__ import annotations

import json
import subprocess
from pathlib import Path

import check

DEFAULT_VIEWPORTS = [{'name': 'mobile', 'width': 375, 'height': 812}, {'name': 'desktop', 'width': 1280, 'height': 800}]
HTML_SUFFIXES = {'.html', '.htm'}


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


def _norm(text):
    return ' '.join((text or '').split())


def _is_subsequence(want, have):
    it = iter(have)
    return all(any(w == h for h in it) for w in want)


def compare(outline, ssot, snap):
    """-> [(code, level, message)] for one snapshot (one file, one viewport)."""
    out = []
    vp = snap['viewport']['name']
    if not snap.get('text_chars'):
        return [('IA_DOM_NO_TARGETS', 'fail', f'{vp}: 描画されたテキストが 0 件（対象 0 件を合格にしない）')]
    seq, facts = expected(outline)
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
    values = {fid: f.get('value') for fid, f in ssot['facts'].items()}
    for fid in facts:
        shown = [a for a in anchors if a['kind'] == 'fact' and a['id'] == fid]
        if not shown:
            out.append(('IA_DOM_FACT_MISSING', 'fail', f'{vp}: 固定値 {fid} の表示（data-ia-fact）が無い'))
        for a in shown:
            if _norm(a['text']) != _norm(values.get(fid)):
                out.append(('IA_DOM_FACT_VALUE', 'fail', f'{vp}: 固定値 {fid} の表示「{_norm(a["text"])}」が SSOT の値「{values.get(fid)}」と違う'))
    return out


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
    first_screens, files = [], 0
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
            for code, level, msg in compare_runs(per_vp) or compare(o, m.ssot, per_vp[0]):
                report.add(level, code, f'{out} {msg}', out)
            first_screens.append({'file': out, 'viewport': vp['name'], 'text': per_vp[0]['first_screen']})
    report.targets = {'files': files, 'viewports': len(viewports) if files else 0}
    report.first_screens = first_screens
    return report
