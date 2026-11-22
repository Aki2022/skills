"""own-positioning-author CLI: check, positioning, targets, migrate-approvals, review, request.

The positioning layer is the owner-approved core of what a repository says: segments, jobs, pains, gains, alternatives,
categories, strengths, values and evidence, with the arrows between them (design/info_arch/info_arch.yaml, schema 5).
This CLI reads and checks it alone. It never writes an approval: the owner approves per value path through
own-design-author's ledger, and the request it shows the owner is built here from the same data the hash binds.
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import approval_migration  # noqa: E402
import ia_common  # noqa: E402
import layer_paths  # noqa: E402  (on sys.path through ia_common)
import positioning as pos  # noqa: E402
import positioning_model as pm  # noqa: E402


def _print(obj):
    print(json.dumps(obj, ensure_ascii=False, indent=2, default=str))


def _report(level):
    return ia_common.Report(level, dt.date.today().isoformat(), tool='own-positioning-author')


def _loaded(args, level):
    """-> (report, model or None). A model is returned only when the SSOT loads."""
    report = _report(level)
    ssot, _legacy = pm.load_ssot(args.root, report)
    if ssot is None:
        return report, None
    return report, pm.Model(Path(args.root), ssot)


def cmd_check(args):
    """Shape and references of every element, the path of every value, ids and the ledger. `--strict` judges every
    value's path as if a shipped outline pointed at it (what the infoarch check does for the paths shipped outlines use)."""
    date = args.date or dt.date.today().isoformat()
    dt.date.fromisoformat(date)
    report = ia_common.Report('check', date, tool='own-positioning-author')
    ssot, _ = pm.load_ssot(args.root, report)
    if ssot is None:
        _print(report.to_dict())
        return 2
    m = pm.Model(Path(args.root), ssot)
    pm.check_ssot(m, report)
    approved = pos.approved_keys(m, pm.read_approvals(args.root, report))
    ps = pos.paths(m)
    pos.check_positioning(m, report, approved, {v: True for v in ps} if args.strict else None, str(pm.SSOT_PATH))
    report.targets = {'values': len(ps), 'approved': sum(1 for v in ps if v in approved)}
    if not ps:
        report.add('fail', 'IA_NO_VALUES', 'positioning に提供価値が 0 件（対象 0 件を合格にしない）', pm.SSOT_PATH)
    _print(report.to_dict())
    return report.exit_code


def cmd_positioning(args):
    """Every approved positioning: its id, derived gaps and points of parity, and what waits for the owner."""
    report, m = _loaded(args, 'positioning')
    if m is None or report.exit_code == 2:
        _print(report.to_dict())
        return 2
    approved = pos.approved_keys(m, pm.read_approvals(args.root, report))
    if report.exit_code == 2:
        _print(report.to_dict())
        return 2
    ps = pos.paths(m)
    _print({'positionings': {seg: p.to_dict() for seg, p in pos.derive(m, approved, ps).items()},
            'paths': {v: {**p.to_dict(), 'approved': v in approved} for v, p in ps.items()}})
    return 0


def cmd_targets(args):
    """The approval targets of the positioning alone: one row per value, the hash of its whole path. Which of them
    shipping needs is the infoarch layer's (`infoarch-author targets`: it knows which outlines point at which path)."""
    report, m = _loaded(args, 'targets')
    if m is None or report.exit_code == 2:
        _print(report.to_dict())
        return 2
    _print({'approval_hashes': pos.approval_targets(m)})
    return 0


def cmd_migrate_approvals(args):
    """Requirement 17: replace the per-element ia:positioning rows by ia:path rows, once (dry run without --write)."""
    report, m = _loaded(args, 'migrate-approvals')
    if m is None or report.exit_code == 2:
        _print(report.to_dict())
        return 2
    path = layer_paths.path(Path(args.root), 'ledger')
    if path.is_symlink() or not path.is_file():
        _print({'exit_code': 2, 'written': False, 'error': f'{approval_migration.LEDGER} が無いか symlink'})
        return 2
    rows = ia_common.load_yaml(path) or []
    try:
        pos._verify_rows()(rows, Path(args.root))
    except ValueError as exc:
        _print({'exit_code': 2, 'written': False, 'error': f'承認台帳の形でない: {exc}'})
        return 2
    out = approval_migration.migrate(args.root, m, rows, write=args.write)
    _print(out)
    return out['exit_code']


def cmd_review(args):
    """The owner reads what an approval request binds (requirement 3), and a stale request is refused."""
    report, m = _loaded(args, 'review')
    if m is None or report.exit_code == 2:
        _print(report.to_dict())
        return 2
    request = json.loads(Path(args.request).read_text(encoding='utf-8'))
    hashes = pos.approval_targets(m)
    rows = [r for r in request.get('rows', []) if str(r.get('key', '')).startswith(pm.APPROVAL_PREFIX)]
    stale = sorted(r['key'] for r in rows if hashes.get(r['key']) != r.get('value_hash'))
    _print({'rows': len(rows), 'stale': stale,
            'markdown': pm.review_markdown(m, [r['key'] for r in rows], hashes)})
    return 2 if stale or not rows else 0


def cmd_request(args):
    """The approval request the owner reads and answers (○ / ×), made from the data the approval hash binds. With
    `--read-decisions` it instead reads the owner's saved answers back against the current paths."""
    import approval_request
    report, m = _loaded(args, 'request')
    if m is None or report.exit_code == 2:
        _print(report.to_dict())
        return 2
    approved = pos.approved_keys(m, pm.read_approvals(args.root, report))
    if report.exit_code == 2:
        _print(report.to_dict())
        return 2
    must = []
    if args.must_file:
        must = json.loads(Path(args.must_file).read_text(encoding='utf-8')).get('must_approval_keys') or []
    items = approval_request.build(m, approved, must)
    if not items:
        _print({'exit_code': 2, 'error': '承認の対象（提供価値の経路）が 0 件'})
        return 2
    if args.read_decisions:
        out = approval_request.decisions_from_rows(items, approval_request.load_rows(args.read_decisions))
        _print(out)
        return 0 if not (out['stale'] or out['unknown']) else 1
    out_dir = Path(args.out)
    files = {out_dir / 'approval-request.html': approval_request.page(items, args.title),
             out_dir / 'approval-request.json': json.dumps({'rows': approval_request.request_rows(items)},
                                                          ensure_ascii=False, indent=2)}
    existing = [str(p) for p in files if p.exists() or p.is_symlink()]
    if existing:
        _print({'exit_code': 2, 'error': f'上書きしない: {existing}'})
        return 2
    out_dir.mkdir(parents=True, exist_ok=True)
    for p, text in files.items():
        p.write_text(text + ('' if text.endswith('\n') else '\n'), encoding='utf-8')
    _print({'cards': len(items), 'approved': sum(i['approved'] for i in items), 'written': sorted(str(p) for p in files),
            'publish': 'この html を非公開の Artifact として出す（capabilities: {db: {}}）。保存先は decisions/<提供価値 id>'})
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(prog='positioning-author')
    sub = ap.add_subparsers(dest='cmd', required=True)
    c = sub.add_parser('check', help='shape, ids and the path of every value of the positioning SSOT')
    c.add_argument('--root', default='.')
    c.add_argument('--date', help='evaluation date YYYY-MM-DD (default: today)')
    c.add_argument('--strict', action='store_true', help="fail on every value's disqualified path (default: only when an outline points at it, which infoarch-author check knows)")
    c.set_defaults(fn=cmd_check)
    po = sub.add_parser('positioning', help='approved positionings per segment: id, gaps, points of parity, member category, drafts')
    po.add_argument('--root', default='.')
    po.set_defaults(fn=cmd_positioning)
    t = sub.add_parser('targets', help='approval targets (ia:path keys and path hashes) of the positioning alone')
    t.add_argument('--root', default='.')
    t.set_defaults(fn=cmd_targets)
    mg = sub.add_parser('migrate-approvals', help='replace per-element ia:positioning approval rows by ia:path rows once')
    mg.add_argument('--root', default='.')
    mg.add_argument('--write', action='store_true', help='write the ledger (without it: dry run)')
    mg.set_defaults(fn=cmd_migrate_approvals)
    rv = sub.add_parser('review', help='readable content of an owner-approval request (ia:path rows); a stale request fails')
    rv.add_argument('--root', default='.')
    rv.add_argument('--request', required=True)
    rv.set_defaults(fn=cmd_review)
    rq = sub.add_parser('request', help='the approval request page (○ / ×, saved to the Artifact db) built from the data the hash binds; or read the answers back')
    rq.add_argument('--root', default='.')
    rq.add_argument('--out', default='approval-request', help='directory to write into (files are never overwritten)')
    rq.add_argument('--title', default='承認依頼')
    rq.add_argument('--must-file', help='`infoarch-author targets` output: marks the paths shipping needs')
    rq.add_argument('--read-decisions', help="the owner's saved answers (the directory `ArtifactData` list wrote with out_dir, or the page's copied JSON): ok / ng / stale / undecided")
    rq.set_defaults(fn=cmd_request)
    args = ap.parse_args(argv)
    try:
        layer_paths.guard(args.root)  # every subcommand reads the layers: a repository that moved stops before any of them
        return args.fn(args)
    except Exception as exc:  # noqa: BLE001 - any crash must be a failure (2), never the warning code (1)
        print(json.dumps({'exit_code': 2, 'error': f'{type(exc).__name__}: {exc}'}, ensure_ascii=False))
        return 2


if __name__ == '__main__':
    sys.exit(main())
