#!/usr/bin/env python3
"""Conformance of a rendered product UI screen to its blueprint (infoarch) and design's allocation rules.

The page is drawn in a real browser at two widths (desktop and phone) and measured (`measure_screen.mjs`). Checked, per width:
- every declared region is marked and drawn (visible, with an area); no marked region is undeclared;
- priority and position: a region of higher priority (smaller number) is never drawn *below* a region of lower priority
  (more than the tolerance lower). Side by side is fine. The order of the priorities holds at both widths;
- exactly one action is drawn as primary (filled), and it is the declared one (none when the screen declares `no_primary`);
- every declared state exists on the page with its declared text (so it can be drawn).
A page with no marks at all is NOT CHECKED (exit 2, SCREEN_NOT_CHECKED), never a pass. An inversion the owner decided on is
declared in ux/design/layout-overrides.yaml, one row per (screen, pair of regions, widths), and each row holds the owner's approval
in the ledger (own-design-author's layout_overrides.py says what binds it). An approved row allows its screen, its pair and its widths
only, and is reported (info). A row with no approval, a reason alone, no width, no screen or region, or a copy of another row fails.

Usage: audit_screen.py --screen <screen yaml> --html <page> [--repo <repo>] [--overrides ux/design/layout-overrides.yaml] [--rules allocation-rules.json]
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
sys.path.insert(0, str(SKILLS / 'own-design-author/scripts'))
import layer_paths  # noqa: E402  (the one rule for where a repository keeps its layers)
import layout_overrides  # noqa: E402  (what an override row stands on, and whether the owner approved it)

MEASURE = Path(__file__).resolve().parent / 'measure_screen.mjs'
RULES = SKILLS / 'own-design-author/references/allocation-rules.json'
VIEWPORTS = [{'name': 'desktop', 'width': 1280, 'height': 800}, {'name': 'mobile', 'width': 375, 'height': 812}]


REEXEC_GUARD = 'AUDIT_SCREEN_REEXEC'


def start_again_under_the_dependency_interpreter():
    """The python3 on PATH may lack PyYAML. The L0 dependency interpreter (DESIGN_AUTHOR_PYTHON, the one build_screen already
    uses) has it, so start this script again under it, once. Without PyYAML and without that variable, load_yaml stops (exit 2)."""
    try:
        import yaml  # noqa: F401
        return
    except ImportError:
        pass
    target = os.environ.get('DESIGN_AUTHOR_PYTHON')
    if not target or os.environ.get(REEXEC_GUARD):
        return
    os.environ[REEXEC_GUARD] = '1'
    try:
        os.execv(target, [target, str(Path(__file__).resolve()), *sys.argv[1:]])
    except OSError:
        pass  # fall through: load_yaml reports the missing PyYAML


def load_yaml(path):
    try:
        import yaml
    except ImportError as exc:
        raise RuntimeError('PyYAML is required: run this with the L0 dependency interpreter (the one DESIGN_AUTHOR_PYTHON names; the interpreter of the skill tests has it)') from exc
    return yaml.safe_load(Path(path).read_text(encoding='utf-8'))


def measure(html, viewports=VIEWPORTS, repo=None):
    """Draw the page at each width and measure it. The Playwright version is the one the repository pins in
    its render environment file (`repo`, default: the current directory): the same fixed environment the other checks use."""
    env = {**os.environ, **({'SCREEN_PLAYWRIGHT_DIR': str(repo)} if repo else {})}
    proc = subprocess.run(['node', str(MEASURE), str(html), json.dumps(viewports)], capture_output=True, text=True, timeout=120, env=env)
    if proc.returncode != 0:
        raise RuntimeError(f'measure_screen が失敗した（exit {proc.returncode}）: {proc.stderr.strip()[:300]}')
    return json.loads(proc.stdout)


def repository_roots(repo, overrides=None):
    """The repository `repo`, else the one the overrides file sits in (when it sits where the rule puts it), else the current
    directory. -> [root, ...]: the first is the one the ledger and the blueprint are read from; none of them is read around."""
    roots = [Path(repo) if repo else Path.cwd()]
    if overrides:
        want = Path(layer_paths.rel('layout_overrides')).parts
        parts = Path(overrides).resolve().parts
        if len(parts) > len(want) and parts[len(parts) - len(want):] == want:
            derived = Path(*parts[:len(parts) - len(want)])
            roots = [derived] + roots if not repo else roots + [derived]
    return roots


def guard_repository(repo, overrides=None):
    """Stop on a repository that moved to a layout this version does not read. The repository is `repo`, else the one the
    overrides file sits in (when it sits where the rule puts it), else the current directory; none of them is read around."""
    for root in repository_roots(repo, overrides):
        layer_paths.guard(root)


def override_state(repo, overrides, screen_id):
    """-> (rows that hold the owner's current approval for this screen, [(code, message)] this audit must fail on).

    The file is `overrides` (the rows are read there); the ledger and the blueprint are the repository's. Nothing here widens a row:
    a row that is not complete, not approved or no longer standing on its screen and regions is a failure and allows nothing."""
    if not overrides or not Path(overrides).is_file():
        return [], []
    root = repository_roots(repo, overrides)[0]
    state = layout_overrides.evaluate(root, layout_overrides.read_rows(overrides), layout_overrides.read_approvals(root))
    return state.approved_for(screen_id), state.problems_for(screen_id)


def audit(screen, snapshots, overrides=(), tolerance=4):
    """-> (findings [(level, code, message)], stats). `screen`: the blueprint screen; `snapshots`: measure() output; `overrides`: the
    rows that hold the owner's current approval for this screen (`override_state`)."""
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
                    # a row allows the pair and the widths it names: a width left out allows none (it used to allow every width)
                    covered = [o for o in overrides if o.get('above') == b and o.get('below') == a and vp in (o.get('viewports') or [])]
                    if covered:
                        f.append(('info', 'SCREEN_ORDER_OVERRIDDEN', f'{vp}: {b} が {a} より上にある（規則では下）。オーナーが承認した例外 {covered[0].get("id", "?")}：{covered[0]["reason"]}'))
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
    ap.add_argument('--repo', help=f'repository root: its {layer_paths.rel("render_environment")} pins the browser (default: the current directory)')
    ap.add_argument('--overrides', help=f'the overrides file ({layer_paths.rel("layout_overrides")} of the repository); none when not given')
    ap.add_argument('--rules', default=str(RULES))
    args = ap.parse_args(argv)
    try:
        guard_repository(args.repo, args.overrides)
        screen = load_yaml(args.screen)
        rules = json.loads(Path(args.rules).read_text(encoding='utf-8'))
        rows, problems = override_state(args.repo, args.overrides, screen['id'])
        findings, stats = audit(screen, measure(args.html, repo=args.repo), rows, rules['order']['tolerance_px'])
        findings = [('fail', code, message) for code, message in problems] + findings
    except (ValueError, OSError, RuntimeError) as exc:  # LayerLayoutError is a ValueError: its message is the layout one
        print(json.dumps({'exit_code': 2, 'error': str(exc)}, ensure_ascii=False))
        return 2
    code = max([{'fail': 2, 'warn': 1, 'info': 0}[lvl] for lvl, _, _ in findings], default=0)
    print(json.dumps({'tool': 'audit_screen', 'screen': screen['id'], 'exit_code': code, 'stats': stats,
                      'findings': [{'level': l, 'code': c, 'message': m} for l, c, m in findings]}, ensure_ascii=False, indent=2))
    return code


if __name__ == '__main__':
    start_again_under_the_dependency_interpreter()
    sys.exit(main())
