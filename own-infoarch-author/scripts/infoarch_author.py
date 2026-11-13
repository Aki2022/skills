"""own-infoarch-author CLI: check (deterministic), dom, judge (judgment ledger reader), evaluate (all three), ship (gate), maintain, skeleton, impact. Never writes approvals."""
from __future__ import annotations

import argparse
import datetime as dt
import json
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import check  # noqa: E402
import dom  # noqa: E402
import evaluate  # noqa: E402
import ia_ship  # noqa: E402
import judge  # noqa: E402
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
    out['unanchored'] = getattr(report, 'unanchored', [])
    _print(out)
    return report.exit_code


def cmd_judge(args):
    date = args.date or dt.date.today().isoformat()
    dt.date.fromisoformat(date)
    if bool(args.answers) != bool(args.replacement_out):
        raise ValueError('--answers と --replacement-out は一緒に使う')
    report, rows, summary = judge.run(args.root, date, args.dom_report or [], args.answers, args.replacement_out)
    out = report.to_dict()
    out['screened'] = getattr(report, 'screened', [])
    out['missing'] = getattr(report, 'missing', [])
    if args.include_candidates:
        out['judgment_candidates'] = rows
    if summary is not None:
        out['replacement'] = summary
    _print(out)
    return report.exit_code


def cmd_evaluate(args):
    """check -> dom -> judge in one report; --acquire fetches missing Jev judgments (preview unless --execute)."""
    date = args.date or dt.date.today().isoformat()
    dt.date.fromisoformat(date)
    acquire = None
    if args.acquire:
        if not args.answers_dir or not args.limit:
            raise ValueError('--acquire には --answers-dir と --limit が要る')
        acquire = {'limit': args.limit, 'answers_dir': args.answers_dir, 'execute': args.execute}
    elif args.execute:
        raise ValueError('--execute は --acquire と一緒に使う')
    result = evaluate.run(args.root, date, args.outline, acquire=acquire)
    result.pop('candidates', None)  # candidate rows carry unpublished copy; judge --include-candidates prints them on purpose
    _print(result)
    return result['exit_code']


def _line(obj):
    print(json.dumps(obj, ensure_ascii=False, sort_keys=True, default=str))


def cmd_ship(args):
    """Ship gate: one JSON line; any refusal is a failure (2), never a pass."""
    date = args.date or dt.date.today().isoformat()
    dt.date.fromisoformat(date)
    try:
        if args.doctor:
            result = ia_ship.doctor(args.root)
        elif args.pre_push:
            result = ia_ship.pre_push(args.root, args.refs_file, date)
        elif args.channel:
            result = ia_ship.ship_channel(args.root, args.channel, date)
        else:
            result, _ = ia_ship.gate(args.root, date)
    except (ValueError, OSError, KeyError, TypeError, subprocess.SubprocessError) as exc:
        result = {'exit_code': 2, 'errors': [str(exc) or type(exc).__name__]}
    _line(result)
    return result['exit_code']


def cmd_maintain(args):
    """own-session-close Step 2.5: one JSON line; skips when design/info_arch/ is absent."""
    date = args.date or dt.date.today().isoformat()
    dt.date.fromisoformat(date)
    try:
        result = ia_ship.maintain(args.root, date, args.within_days)
    except (ValueError, OSError, KeyError, TypeError, subprocess.SubprocessError) as exc:
        result = {'exit_code': 2, 'status': 'failed', 'errors': [str(exc) or type(exc).__name__]}
    _line(result)
    return result['exit_code']


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
    j = sub.add_parser('judge', help='Jev judgment candidates and a read-only check against the shared judgment ledger')
    j.add_argument('--root', default='.')
    j.add_argument('--date', help='evaluation date YYYY-MM-DD (default: today)')
    j.add_argument('--dom-report', action='append', help='JSON output of `dom` whose first_screens feed the emphasis judgment (repeatable)')
    j.add_argument('--include-candidates', action='store_true', help='print judgment_candidates (contains copy text) for judgment_run.py')
    j.add_argument('--answers', help='judgment-answers.json from judgment_run.py')
    j.add_argument('--replacement-out', help='write the shared-ledger replacement here (never overwrites)')
    j.set_defaults(fn=cmd_judge)
    e = sub.add_parser('evaluate', help='evaluate generated creatives: check, then dom, then judge, one report and one exit code')
    e.add_argument('--root', default='.')
    e.add_argument('--date', help='evaluation date YYYY-MM-DD (default: today)')
    e.add_argument('--outline', action='append', help='outline id to evaluate (repeatable; default: every outline)')
    e.add_argument('--acquire', action='store_true', help='also fetch missing Jev judgments (preview only unless --execute)')
    e.add_argument('--answers-dir', help='directory that keeps judgment-answers.json between runs (outside the repo)')
    e.add_argument('--limit', type=int, help='most judgments to request in this run')
    e.add_argument('--execute', action='store_true', help='with --acquire: call Jev (metered) and record design/judgments.json')
    e.set_defaults(fn=cmd_evaluate)
    s = sub.add_parser('ship', help='ship gate on the committed SHA: check, dom of shipped HTML, judge; --channel runs ship_command only when green')
    s.add_argument('--root', default='.')
    s.add_argument('--date', help='evaluation date YYYY-MM-DD (default: today)')
    mode = s.add_mutually_exclusive_group()
    mode.add_argument('--channel', help='the formal entry: gate, then run this channel\'s ship_command on the same commit')
    mode.add_argument('--pre-push', action='store_true', help='gate the predicted merge of each pushed commit into merge_base')
    mode.add_argument('--doctor', action='store_true', help='diagnose ship_command, merge_base, marker and the pre-push extension')
    s.add_argument('--refs-file', help='pre-push refs preserved by the dispatcher')
    s.set_defaults(fn=cmd_ship)
    mt = sub.add_parser('maintain', help='own-session-close Step 2.5: changed SSOT/outlines, evidence falling due, uncovered judgments')
    mt.add_argument('--root', default='.')
    mt.add_argument('--date', help='evaluation date YYYY-MM-DD (default: today)')
    mt.add_argument('--within-days', type=int, help='due window in days (default: references/expiry.json due_within_days)')
    mt.set_defaults(fn=cmd_maintain)
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
