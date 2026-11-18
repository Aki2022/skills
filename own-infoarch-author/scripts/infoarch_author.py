"""own-infoarch-author CLI: check (deterministic), dom, judge (judgment ledger reader), evaluate (all three), ship (gate), maintain, skeleton, impact, positioning, set, diff. Never writes approvals."""
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


def cmd_review(args):
    """The owner reads what an approval request binds (requirement 3), and a stale request is refused."""
    report = check.Report('review', dt.date.today().isoformat())
    m = check.load(args.root, report)
    if m.ssot is None or report.exit_code == 2:
        _print(report.to_dict())
        return 2
    request = json.loads(Path(args.request).read_text(encoding='utf-8'))
    hashes, must = check.positioning_targets(m)
    rows = [r for r in request.get('rows', []) if str(r.get('key', '')).startswith(check.APPROVAL_PREFIX)]
    stale = sorted(r['key'] for r in rows if hashes.get(r['key']) != r.get('value_hash'))
    _print({'rows': len(rows), 'stale': stale, 'missing_required': sorted(set(must) - {r['key'] for r in rows}),
            'markdown': check.review_markdown(m, [r['key'] for r in rows], hashes)})
    return 2 if stale or not rows else 0


def _loaded(args, level):
    report = check.Report(level, dt.date.today().isoformat())
    m = check.load(args.root, report)
    if m.ssot is None or report.exit_code == 2:
        return report, None, None
    import positioning
    approved = positioning.approved_keys(m, check.read_approvals(args.root, report))
    return report, m, (None if report.exit_code == 2 else approved)


def cmd_set(args):
    """The set id (requirement 8) of one deliverable, with the content hash of each of its outline elements."""
    import outline
    report, m, approved = _loaded(args, 'set')
    if approved is not None and args.outline not in m.outlines:
        report.add('fail', 'IA_SET_UNKNOWN_OUTLINE', f'アウトライン {args.outline} が無い')
    if approved is None or report.exit_code == 2:
        _print(report.to_dict())
        return 2
    _, o = m.outlines[args.outline]
    ids = outline.set_ids(m, args.outline, approved)
    elems = {}
    for key, value in outline.elements(o).items():
        block, _, name = key.partition('/')
        if name:
            elems.setdefault(block, {})[name] = outline.pos.sha256(value)
    _print({'deliverable': args.outline, **ids, 'elements': elems})
    return 0


def cmd_diff(args):
    """Requirement 9: the difference between two sets (two deliverables, optionally of two checkouts)."""
    import outline
    root_b = args.root_b or args.root
    ra, ma, aa = _loaded(argparse.Namespace(root=args.root), 'diff')
    rb, mb, ab = _loaded(argparse.Namespace(root=root_b), 'diff')
    for rep_, m, appr, oid in ((ra, ma, aa, args.a), (rb, mb, ab, args.b)):
        if appr is not None and oid not in m.outlines:
            rep_.add('fail', 'IA_SET_UNKNOWN_OUTLINE', f'アウトライン {oid} が無い')
        if appr is None or rep_.exit_code == 2:
            _print(rep_.to_dict())
            return 2
    out = outline.diff(ma, args.a, aa, mb, args.b, ab,
                       outline.load_dom_report(args.dom_a), outline.load_dom_report(args.dom_b))
    _print(out)
    return out['exit_code']


def cmd_positioning(args):
    """Every approved positioning: its id, derived gaps and points of parity, and what waits for the owner."""
    import positioning
    report = check.Report('positioning', dt.date.today().isoformat())
    m = check.load(args.root, report)
    if m.ssot is None or report.exit_code == 2:
        _print(report.to_dict())
        return 2
    approved = positioning.approved_keys(m, check.read_approvals(args.root, report))
    if report.exit_code == 2:
        _print(report.to_dict())
        return 2
    ps = positioning.paths(m)
    _print({'positionings': {seg: p.to_dict() for seg, p in positioning.derive(m, approved, ps).items()},
            'paths': {v: {**p.to_dict(), 'approved': v in approved} for v, p in ps.items()}})
    return 0


def cmd_answer_key(args):
    """Requirement 12: write the answer key (pairs, candidates for Jev, the owner's labeling page), or measure it."""
    import answer_key
    report, m, approved = _loaded(args, 'answer-key')
    if approved is None or report.exit_code == 2:
        _print(report.to_dict())
        return 2
    jl = judge._jl()
    policy = jl.load_policy()
    if args.measure:
        if not args.key:
            _print({'exit_code': 2, 'error': '--measure には、作った時の --key（answer-key.json）が要る'})
            return 2
        items = answer_key.load_items(m, approved, policy, jl, args.key)
        labels = json.loads(Path(args.measure).read_text(encoding='utf-8'))
        result = answer_key.measure(items, labels, judge.read_ledger(args.root), policy['theta'])
        _print(result)
        return 0
    items = answer_key.build(m, approved, policy, jl)
    if not items:
        _print({'exit_code': 2, 'error': '正解表の組が 0 件（承認済みの提供価値に立つ文案が無い）'})
        return 2
    out = Path(args.out)
    files = {out / 'answer-key.json': json.dumps(answer_key.key_file(items), ensure_ascii=False, indent=2),
             out / 'candidates.json': json.dumps(answer_key.candidates_file(items), ensure_ascii=False, indent=2),
             out / 'answer-key-page.html': answer_key.page(items)}
    existing = [str(p) for p in files if p.exists() or p.is_symlink()]
    if existing:
        _print({'exit_code': 2, 'error': f'上書きしない: {existing}'})
        return 2
    out.mkdir(parents=True, exist_ok=True)
    for p, text in files.items():
        p.write_text(text + ('' if text.endswith('\n') else '\n'), encoding='utf-8')
    _print({'answer_key': answer_key.key_hash(items), 'pairs': len(items),
            'own': sum(i['kind'] == 'own' for i in items), 'swap': sum(i['kind'] == 'swap' for i in items),
            'written': sorted(str(p) for p in files)})
    return 0


def cmd_migrate_approvals(args):
    """Requirement 17: replace the per-element ia:positioning rows by ia:path rows, once (dry run without --write)."""
    import approval_migration
    import positioning
    report = check.Report('migrate-approvals', dt.date.today().isoformat())
    m = check.load(args.root, report)
    if m.ssot is None or report.exit_code == 2:
        _print(report.to_dict())
        return 2
    path = Path(args.root) / approval_migration.LEDGER
    if path.is_symlink() or not path.is_file():
        _print({'exit_code': 2, 'written': False, 'error': f'{approval_migration.LEDGER} が無いか symlink'})
        return 2
    rows = check.load_yaml(path) or []
    try:
        positioning._verify_rows()(rows, Path(args.root))
    except ValueError as exc:
        _print({'exit_code': 2, 'written': False, 'error': f'承認台帳の形でない: {exc}'})
        return 2
    out = approval_migration.migrate(args.root, m, rows, write=args.write)
    _print(out)
    return out['exit_code']


def cmd_targets(args):
    """Owner-approval targets for own-design-author: one row per value (the hash of its whole path), and the keys
    shipping needs."""
    report = check.Report('targets', dt.date.today().isoformat())
    m = check.load(args.root, report)
    if m.ssot is None or report.exit_code == 2:
        _print(report.to_dict())
        return 2
    hashes, must = check.positioning_targets(m)
    _print({'approval_hashes': hashes, 'must_approval_keys': must})
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
    rv = sub.add_parser('review', help='readable content of an owner-approval request (ia:positioning rows); a stale request fails')
    rv.add_argument('--root', default='.')
    rv.add_argument('--request', required=True)
    rv.set_defaults(fn=cmd_review)
    st = sub.add_parser('set', help='set id of one deliverable: positioning id + outline id, and its element hashes')
    st.add_argument('--root', default='.')
    st.add_argument('--outline', required=True)
    st.set_defaults(fn=cmd_set)
    df = sub.add_parser('diff', help='difference between two sets: changed elements, changed layers, same experiment or not')
    df.add_argument('--root', default='.')
    df.add_argument('--root-b', help='checkout of the second set (default: --root)')
    df.add_argument('--a', required=True, help='first deliverable (outline id)')
    df.add_argument('--b', required=True, help='second deliverable (outline id)')
    df.add_argument('--dom-a', help='dom report of the first deliverable (compares text outside anchors)')
    df.add_argument('--dom-b', help='dom report of the second deliverable')
    df.set_defaults(fn=cmd_diff)
    t = sub.add_parser('targets', help='owner-approval targets (ia:path keys and path hashes) for own-design-author approvals')
    t.add_argument('--root', default='.')
    t.set_defaults(fn=cmd_targets)
    mg = sub.add_parser('migrate-approvals', help='replace per-element ia:positioning approval rows by ia:path rows once')
    mg.add_argument('--root', default='.')
    mg.add_argument('--write', action='store_true', help='write the ledger (without it: dry run)')
    mg.set_defaults(fn=cmd_migrate_approvals)
    ak = sub.add_parser('answer-key', help='answer key for the meaning judgment: pairs, candidates and the labeling page, or --measure')
    ak.add_argument('--root', default='.')
    ak.add_argument('--out', default='answer-key', help='directory to write into (files are never overwritten)')
    ak.add_argument('--measure', help="the owner's labels.json: measure agreement against design/judgments.json")
    ak.add_argument('--key', help='answer-key.json written with the page (required with --measure)')
    ak.set_defaults(fn=cmd_answer_key)
    po = sub.add_parser('positioning', help='approved positionings per segment: id, gaps, points of parity, member category, drafts')
    po.add_argument('--root', default='.')
    po.set_defaults(fn=cmd_positioning)
    args = ap.parse_args(argv)
    try:
        return args.fn(args)
    except Exception as exc:  # noqa: BLE001 - any crash must be a failure (2), never the warning code (1)
        print(json.dumps({'exit_code': 2, 'error': f'{type(exc).__name__}: {exc}'}, ensure_ascii=False))
        return 2


if __name__ == '__main__':
    sys.exit(main())
