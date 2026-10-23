"""own-infoarch-author CLI: check (deterministic), impact (reverse lookup). Never writes approvals."""
from __future__ import annotations

import argparse
import datetime as dt
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import check  # noqa: E402
import dom  # noqa: E402
import skeleton  # noqa: E402


def _print(obj):
    print(json.dumps(obj, ensure_ascii=False, indent=2, default=str))


def cmd_check(args):
    date = args.date or dt.date.today().isoformat()
    dt.date.fromisoformat(date)  # ValueError -> exit 2 in main
    report, _ = check.check(args.root, date)
    _print(report.to_dict())
    return report.exit_code


def cmd_dom(args):
    date = args.date or dt.date.today().isoformat()
    dt.date.fromisoformat(date)
    report = dom.run(args.root, args.outline, date)
    out = report.to_dict()
    out['first_screens'] = getattr(report, 'first_screens', [])
    _print(out)
    return report.exit_code


def cmd_skeleton(args):
    """Print (or write under --out, never overwriting) the HTML skeleton of a checked outline."""
    report, m = check.check(args.root, dt.date.today().isoformat())
    if m.ssot is None or args.outline not in m.outlines:
        report.add('fail', 'IA_SKELETON_OUTLINE', f'アウトライン {args.outline} が無い')
    elif report.exit_code == 2:
        report.add('fail', 'IA_SKELETON_REFUSED', 'check が失格しているので骨組みを出さない。先に直す')
    if report.exit_code == 2:
        _print(report.to_dict())
        return 2
    pages = skeleton.render(m.outlines[args.outline][1], m.ssot)
    if not args.out:
        print('\n'.join(pages.values()), end='')
        return 0
    targets = {Path(args.out) / rel: html for rel, html in pages.items()}
    existing = [str(t) for t in targets if t.exists()]
    if existing:
        _print({'exit_code': 2, 'error': f'上書きしない: {existing}'})
        return 2
    for t, html in targets.items():
        t.parent.mkdir(parents=True, exist_ok=True)
        t.write_text(html, encoding='utf-8')
    _print({'written': sorted(str(t) for t in targets)})
    return 0


def cmd_impact(args):
    report, result = check.impact(args.root, args.id)
    if report.exit_code == 2:
        _print(report.to_dict())
        return 2
    _print(result)
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(prog='infoarch-author')
    sub = ap.add_subparsers(dest='cmd', required=True)
    c = sub.add_parser('check', help='deterministic checks of the SSOT, channels and outlines')
    c.add_argument('--root', default='.')
    c.add_argument('--date', help='evaluation date YYYY-MM-DD (default: today)')
    c.set_defaults(fn=cmd_check)
    d = sub.add_parser('dom', help='render the outline outputs in a browser and compare anchors, order and facts')
    d.add_argument('--root', default='.')
    d.add_argument('--outline', required=True)
    d.add_argument('--date', help='evaluation date YYYY-MM-DD (default: today)')
    d.set_defaults(fn=cmd_dom)
    k = sub.add_parser('skeleton', help='print or write the HTML skeleton (anchors, order, facts) of an outline for a factory to style')
    k.add_argument('--root', default='.')
    k.add_argument('--outline', required=True)
    k.add_argument('--out', help='directory to write into (files are never overwritten)')
    k.set_defaults(fn=cmd_skeleton)
    i = sub.add_parser('impact', help='list outlines and outputs that reference an atom id')
    i.add_argument('--root', default='.')
    i.add_argument('--id', required=True)
    i.set_defaults(fn=cmd_impact)
    args = ap.parse_args(argv)
    try:
        return args.fn(args)
    except Exception as exc:  # noqa: BLE001 - any crash must be a failure (2), never the warning code (1)
        print(json.dumps({'exit_code': 2, 'error': f'{type(exc).__name__}: {exc}'}, ensure_ascii=False))
        return 2


if __name__ == '__main__':
    sys.exit(main())
