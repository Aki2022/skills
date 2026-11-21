#!/usr/bin/env python3
"""Conformance of a rendered product UI screen to its blueprint (infoarch) and design's allocation rules.

The page is drawn in a real browser at two widths (desktop and phone) and measured (`measure_screen.mjs`). Checked, per width:
- every declared region is marked and drawn (visible, with an area); no marked region is undeclared;
- priority and position: a region of higher priority (smaller number) is never drawn *below* a region of lower priority
  (more than the tolerance lower). Side by side is fine. The order of the priorities holds at both widths;
- exactly one action is drawn as primary (filled), and it is the declared one (none when the screen declares `no_primary`);
- every declared state exists on the page with its declared text (so it can be drawn).
A page with no marks at all is NOT CHECKED (exit 2, SCREEN_NOT_CHECKED), never a pass. An inversion the owner decided on is
declared in design/layout-overrides.yaml with a reason; it is reported (info) and does not fail.

Usage: audit_screen.py --screen <screen yaml> --html <page> [--overrides design/layout-overrides.yaml] [--rules allocation-rules.json]
Exit: 0 conforming, 1 warnings, 2 failing or not checked.
"""
from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
from pathlib import Path

SKILLS = Path(__file__).resolve().parents[2]
MEASURE = Path(__file__).resolve().parent / 'measure_screen.mjs'
RULES = SKILLS / 'own-design-author/references/allocation-rules.json'
VIEWPORTS = [{'name': 'desktop', 'width': 1280, 'height': 800}, {'name': 'mobile', 'width': 375, 'height': 812}]


def load_yaml(path):
    try:
        import yaml
    except ImportError as exc:
        raise RuntimeError('PyYAML is required: run this with the L0 dependency interpreter (the one DESIGN_AUTHOR_PYTHON names; the interpreter of the skill tests has it)') from exc
    return yaml.safe_load(Path(path).read_text(encoding='utf-8'))


def measure(html, viewports=VIEWPORTS, repo=None):
    """Draw the page at each width and measure it. The Playwright version is the one the repository pins in
    design/render-environment.json (`repo`, default: the current directory): the same fixed environment the other checks use."""
    env = {**os.environ, **({'SCREEN_PLAYWRIGHT_DIR': str(repo)} if repo else {})}
    proc = subprocess.run(['node', str(MEASURE), str(html), json.dumps(viewports)], capture_output=True, text=True, timeout=120, env=env)
    if proc.returncode != 0:
        raise RuntimeError(f'measure_screen が失敗した（exit {proc.returncode}）: {proc.stderr.strip()[:300]}')
    return json.loads(proc.stdout)


def overrides_for(path, screen_id):
    if not path or not Path(path).is_file():
        return []
    rows = load_yaml(path) or []
    out = []
    for r in rows:
        if r.get('screen') == screen_id:
            if not str(r.get('reason', '')).strip():
                raise ValueError('layout-overrides の行に理由（reason）が無い: ' + json.dumps(r, ensure_ascii=False))
            out.append(r)
    return out


def audit(screen, snapshots, overrides=(), tolerance=4):
    """-> (findings [(level, code, message)], stats). `screen`: the blueprint screen; `snapshots`: measure() output."""
    f, sid = [], screen['id']
    declared = screen['regions']
    marked_anywhere = any(s['regions'] or s['actions'] or s['states'] for s in snapshots)
    if not marked_anywhere:
        return [('fail', 'SCREEN_NOT_CHECKED', f'{sid}: ページに印（data-ia-region・data-ia-action・data-ia-state）が 1 つも無い。確かめていない（合格にしない）')], {'regions': 0}
    for snap in snapshots:
        vp = snap['viewport']['name']
        if sid not in snap['screens']:
            f.append(('fail', 'SCREEN_MARK_MISSING', f'{vp}: 画面の印 data-ia-screen="{sid}" が無い'))
        drawn = {r['id']: r for r in snap['regions']}
        for rid in sorted(set(declared) - set(drawn)):
            f.append(('fail', 'SCREEN_REGION_MISSING', f'{vp}: 宣言した領域 {rid} の印がページに無い'))
        for rid in sorted(set(drawn) - set(declared)):
            f.append(('fail', 'SCREEN_REGION_UNDECLARED', f'{vp}: 印のある領域 {rid} が blueprint に無い'))
        for rid, r in sorted(drawn.items()):
            if rid in declared and (not r['visible'] or r['width'] <= 0 or r['height'] <= 0):
                f.append(('fail', 'SCREEN_REGION_HIDDEN', f'{vp}: 領域 {rid} が描かれていない（見えない、または面積が 0）'))
        ids = [rid for rid in declared if rid in drawn and 'priority' in declared[rid]]
        for a in ids:
            for b in ids:
                if declared[a]['priority'] < declared[b]['priority'] and drawn[a]['top'] > drawn[b]['top'] + tolerance:
                    covered = [o for o in overrides if o['above'] == b and o['below'] == a and vp in o.get('viewports', [vp])]
                    if covered:
                        f.append(('info', 'SCREEN_ORDER_OVERRIDDEN', f'{vp}: {b} が {a} より上にある（規則では下）。人が決めた：{covered[0]["reason"]}'))
                    else:
                        f.append(('fail', 'SCREEN_ORDER', f'{vp}: 優先度 {declared[a]["priority"]} の {a} が、優先度 {declared[b]["priority"]} の {b} より下に描かれている'))
        actions = {a['id']: a for a in snap['actions']}
        for aid in sorted(set(screen['actions']) - set(actions)):
            f.append(('fail', 'SCREEN_ACTION_MISSING', f'{vp}: 宣言した操作 {aid} の印がページに無い'))
        want = {aid for aid, a in screen['actions'].items() if a.get('primary')}
        painted = {aid for aid, a in actions.items() if a['visible'] and a['painted']}
        marked = {aid for aid, a in actions.items() if a['visible'] and a['marked_primary']}
        if len(painted | marked) > 1 or (screen.get('no_primary') and (painted | marked)):
            f.append(('fail', 'SCREEN_PRIMARY_MANY', f'{vp}: 主操作として描かれている操作が {sorted(painted | marked)}（1 つだけ。主操作なしの画面は 0）'))
        elif want and painted != want:
            f.append(('fail', 'SCREEN_PRIMARY_WRONG', f'{vp}: 塗られている操作 {sorted(painted)} が、宣言した主操作 {sorted(want)} と違う'))
        for name, decl in sorted(screen['states'].items()):
            shown = [s for s in snap['states'] if s['id'] == name]
            if not shown or any(' '.join(s['text'].split()) != ' '.join(decl['text'].split()) for s in shown):
                f.append(('fail', 'SCREEN_STATE_MISSING', f'{vp}: 宣言した状態 {name} が、宣言した文言でページに無い（描けない）'))
    return f, {'regions': len(declared), 'viewports': [s['viewport']['name'] for s in snapshots]}


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument('--screen', required=True)
    ap.add_argument('--html', required=True)
    ap.add_argument('--repo', help='repository root: its design/render-environment.json pins the browser (default: the current directory)')
    ap.add_argument('--overrides')
    ap.add_argument('--rules', default=str(RULES))
    args = ap.parse_args(argv)
    try:
        screen = load_yaml(args.screen)
        rules = json.loads(Path(args.rules).read_text(encoding='utf-8'))
        findings, stats = audit(screen, measure(args.html, repo=args.repo), overrides_for(args.overrides, screen['id']), rules['order']['tolerance_px'])
    except (ValueError, OSError, RuntimeError) as exc:
        print(json.dumps({'exit_code': 2, 'error': str(exc)}, ensure_ascii=False))
        return 2
    code = max([{'fail': 2, 'warn': 1, 'info': 0}[lvl] for lvl, _, _ in findings], default=0)
    print(json.dumps({'tool': 'audit_screen', 'screen': screen['id'], 'exit_code': code, 'stats': stats,
                      'findings': [{'level': l, 'code': c, 'message': m} for l, c, m in findings]}, ensure_ascii=False, indent=2))
    return code


if __name__ == '__main__':
    sys.exit(main())
