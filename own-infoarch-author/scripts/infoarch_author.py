"""own-infoarch-author CLI: check (deterministic), dom, judge (judgment ledger reader), evaluate (all three), ship (gate), maintain, skeleton, impact, targets, set, diff, answer-key, migrate-v6, migrate-dictionary. Never writes approvals. The positioning layer (paths, approval requests, the approval ledger migration) is own-positioning-author's CLI."""
from __future__ import annotations

import argparse
import datetime as dt
import json
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import blueprint  # noqa: E402
import check  # noqa: E402
import dom  # noqa: E402
import evaluate  # noqa: E402
import ia_ship  # noqa: E402
import judge  # noqa: E402
import layer_paths  # noqa: E402  (on sys.path through check -> common)
import skeleton  # noqa: E402

# These two print one JSON line whatever happens (their own contract), so they guard inside their own try block.
GUARD_ITSELF = ('ship', 'maintain')


# What only the script side of a repository reports (channels, scripts, their use of a path, expiry of what they ship,
# the design page inventory). `blueprint` leaves these to `check`.
OUTLINE_SIDE = ('IA_OUTLINE', 'IA_SCRIPT_', 'IA_CHANNEL', 'IA_NO_OUTLINES', 'IA_NO_CHANNELS', 'IA_LEAD_UNSUPPORTED', 'IA_PAGES_',
                'IA_EXPIRED_', 'IA_POSITIONING_UNAPPROVED')


def _print(obj):
    print(json.dumps(obj, ensure_ascii=False, indent=2, default=str))


def cmd_check(args):
    date = args.date or dt.date.today().isoformat()
    dt.date.fromisoformat(date)  # ValueError -> exit 2 in main
    report, m = check.check(args.root, date)
    _print(check.categorize(report, m, 'check').to_dict())
    return report.exit_code


def cmd_blueprint(args):
    """Deterministic checks of the product UI's infoarch alone (product items, screens, approvals per item)."""
    date = args.date or dt.date.today().isoformat()
    dt.date.fromisoformat(date)
    report, m = check.check(args.root, date)
    bp = m.blueprint
    if m.ssot is not None:
        # Everything `check` reports except what only an outline repository has. A product UI stands on the positioning,
        # so its failures (IA_POSITIONING_*, IA_REF_BROKEN, IA_ID_*, ...) count here as in `check`. Drop by the outline side's
        # names, never keep by a list of the blueprint's: a code that a list does not name would vanish without a trace.
        report.findings = [f for f in report.findings if not f.code.startswith(OUTLINE_SIDE)]
        if bp is None or not bp.present or not bp.screens:
            if not [f for f in report.findings if f.code in ('IA_BLUEPRINT_NO_SCREENS', 'IA_BLUEPRINT_PRODUCT_MISSING')]:
                report.add('fail', 'IA_BLUEPRINT_NO_SCREENS', 'blueprint/screens/*.yaml が 0 件（対象 0 件を合格にしない）', layer_paths.rel('infoarch', 'blueprint', 'screens'))
        report.targets = {'screens': len(bp.screens) if bp else 0, 'items': len(blueprint.items(bp)) if bp else 0}
    _print(check.categorize(report, m, 'blueprint').to_dict())
    return report.exit_code


def cmd_request(args):
    """The approval request for the product items (○ / ×, saved to the Artifact db); or read the owner's answers back."""
    import blueprint_request
    date = dt.date.today().isoformat()
    report, m = check.check(args.root, date)
    bp = m.blueprint
    if m.ssot is None or bp is None or not bp.present or bp.product is None or report.exit_code == 2:
        _print({'exit_code': 2, 'error': 'blueprint の検査が失格しているか、product.yaml が無い（先に blueprint を緑にする）',
                'findings': [f.code for f in report.findings if f.level == 'fail']})
        return 2
    rows = check.read_approvals(args.root, check.Report('request', date))
    items = blueprint_request.build(bp, blueprint.approved_items(bp, rows), blueprint.required_keys(bp))
    if not items:
        _print({'exit_code': 2, 'error': '承認の対象（proposed の項目）が 0 件'})
        return 2
    if args.read_decisions:
        import approval_request
        out = approval_request.decisions_from_rows(items, json.loads(Path(args.read_decisions).read_text(encoding='utf-8')))
        _print(out)
        return 0 if not (out['stale'] or out['unknown']) else 1
    import positioning
    values = {k: (a.get('text') or k) for k, a in positioning.atoms(m, proposed_only=False).items() if k.startswith('val-')}
    out_dir = Path(args.out)
    files = {out_dir / 'approval-request.html': blueprint_request.page(items, args.title, blueprint_request.overview(bp, values)),
             out_dir / 'approval-request.json': json.dumps({'rows': [{'key': i['key'], 'value_hash': i['value_hash']} for i in items]},
                                                          ensure_ascii=False, indent=2)}
    existing = [str(p) for p in files if p.exists() or p.is_symlink()]
    if existing:
        _print({'exit_code': 2, 'error': f'上書きしない: {existing}'})
        return 2
    out_dir.mkdir(parents=True, exist_ok=True)
    for p, text in files.items():
        p.write_text(text + ('' if text.endswith('\n') else '\n'), encoding='utf-8')
    _print({'cards': len(items), 'approved': sum(i['approved'] for i in items), 'written': sorted(str(p) for p in files),
            'publish': 'この html を非公開の Artifact として出す（capabilities: {db: {}}）。保存先は decisions/<項目 id>'})
    return 0


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
    report, rows, summary = judge.run(args.root, date, args.answers, args.replacement_out)
    out = report.to_dict()
    out['screened'] = getattr(report, 'screened', [])
    out['missing'] = getattr(report, 'missing', [])
    if args.include_candidates:
        out['judgment_candidates'] = rows
    if summary is not None:
        out['replacement'] = summary
    _print(out)
    return report.exit_code


def cmd_accept_reviews(args):
    date = args.date or dt.date.today().isoformat()
    dt.date.fromisoformat(date)
    _print(judge.accept_reviews(args.root, date, args.id or [], args.owner_words,
                                args.approved_on, write=args.write))
    return 0


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
        layer_paths.guard(args.root)
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
    """own-session-close Step 2.5: one JSON line; skips when the repository has no infoarch tree, fails when it moved to a layout this version does not read."""
    date = args.date or dt.date.today().isoformat()
    dt.date.fromisoformat(date)
    try:
        layer_paths.guard(args.root)
        result = ia_ship.maintain(args.root, date, args.within_days)
    except (ValueError, OSError, KeyError, TypeError, subprocess.SubprocessError) as exc:
        result = {'exit_code': 2, 'status': 'failed', 'errors': [str(exc) or type(exc).__name__]}
    _line(result)
    return result['exit_code']


def cmd_skeleton_screen(args):
    """The HTML skeleton of a blueprint screen (marks for regions, groups, actions and states). A failing blueprint gives none."""
    import screen_skeleton
    report, m = check.check(args.root, dt.date.today().isoformat())
    bp = m.blueprint
    if m.ssot is None or bp is None or args.screen not in bp.screens:
        report.add('fail', 'IA_SKELETON_OUTLINE', f'画面 {args.screen} が無い')
    elif report.exit_code == 2:
        report.add('fail', 'IA_SKELETON_REFUSED', 'check が失格しているので骨組みを出さない。先に直す')
    if report.exit_code == 2:
        _print(report.to_dict())
        return 2
    page = screen_skeleton.render(bp.screens[args.screen][1], css=Path(args.css).read_text(encoding='utf-8') if args.css else '')
    if not args.out:
        print(page, end='')
        return 0
    target = Path(args.out)
    if target.exists():
        _print({'exit_code': 2, 'error': f'上書きしない: {target}'})
        return 2
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(page, encoding='utf-8')
    _print({'written': [str(target)]})
    return 0


def cmd_skeleton(args):
    """Print (or write under --out, never overwriting) the HTML skeleton of a checked outline (or, with --screen, of a blueprint screen)."""
    if args.screen:
        return cmd_skeleton_screen(args)
    report, m = check.check(args.root, dt.date.today().isoformat())
    if m.ssot is None or args.outline not in m.outlines:
        report.add('fail', 'IA_SKELETON_OUTLINE', f'台本 {args.outline} が無い')
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


def _loaded(args, level):
    report = check.Report(level, dt.date.today().isoformat())
    m = check.load(args.root, report)
    if m.ssot is None or report.exit_code == 2:
        return report, None, None
    import positioning
    approved = positioning.approved_keys(m, check.read_approvals(args.root, report))
    return report, m, (None if report.exit_code == 2 else approved)


def cmd_set(args):
    """The set id (requirement 8) of one deliverable, with the content hash of each of its script elements."""
    import outline
    report, m, approved = _loaded(args, 'set')
    if approved is not None and args.outline not in m.outlines:
        report.add('fail', 'IA_SET_UNKNOWN_OUTLINE', f'台本 {args.outline} が無い')
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
        result = answer_key.measure(items, labels, judge.read_ledger(args.root), policy['theta'],
                                    answer_key.read_measurement(Path(args.key).parent))
        _print(result)
        return 0
    if args.record:
        if not args.key or not args.limit:
            _print({'exit_code': 2, 'error': '--record には、作った時の --key（answer-key.json）と --limit が要る'})
            return 2
        items = answer_key.load_items(m, approved, policy, jl, args.key)
        _print(answer_key.record(items, Path(args.key).parent, args.limit, execute=args.execute))
        return 0
    if args.execute or args.limit:
        _print({'exit_code': 2, 'error': '--limit と --execute は --record と一緒に使う'})
        return 2
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


def cmd_migrate_v6(args):
    """Requirement 17: one batch v5 -> v6 migration of ux/infoarch/ (dry run without --write)."""
    import migrate_v6
    out = migrate_v6.run(args.root, write=args.write)
    _print(out)
    return out['exit_code']


def cmd_migrate_dictionary(args):
    """One migration of a schema 4 tree to schema 5 and the docs dictionary (dry run without --write)."""
    import migrate_dictionary
    mapping = json.loads(Path(args.map).read_text(encoding='utf-8'))
    out = migrate_dictionary.run(args.root, mapping, domain=args.domain, write=args.write, date=args.date)
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
    c = sub.add_parser('check', help='deterministic checks of the SSOT, channels and scripts (台本)')
    c.add_argument('--root', default='.')
    c.add_argument('--date', help='evaluation date YYYY-MM-DD (default: today)')
    c.set_defaults(fn=cmd_check)
    bpc = sub.add_parser('blueprint', help='deterministic checks of the product UI\'s infoarch: product items, screens, approvals per item')
    bpc.add_argument('--root', default='.')
    bpc.add_argument('--date', help='evaluation date YYYY-MM-DD (default: today)')
    bpc.set_defaults(fn=cmd_blueprint)
    rq = sub.add_parser('request', help='the approval request page for the product items of a blueprint (○ / ×), or read the answers back')
    rq.add_argument('--root', default='.')
    rq.add_argument('--out', default='approval-request', help='directory to write into (files are never overwritten)')
    rq.add_argument('--title', default='プロダクト UI の承認依頼')
    rq.add_argument('--read-decisions', help="the owner's saved answers (the Artifact db `decisions` rows, or the page's copied JSON)")
    rq.set_defaults(fn=cmd_request)
    d = sub.add_parser('dom', help='render the script outputs in a browser and compare anchors, order and facts')
    d.add_argument('--root', default='.')
    d.add_argument('--outline', '--script', dest='outline', required=True, help='script id (--outline is the v6 name, kept)')
    d.add_argument('--date', help='evaluation date YYYY-MM-DD (default: today)')
    d.set_defaults(fn=cmd_dom)
    j = sub.add_parser('judge', help='Jev judgment candidates and a read-only check against the shared judgment ledger')
    j.add_argument('--root', default='.')
    j.add_argument('--date', help='evaluation date YYYY-MM-DD (default: today)')
    j.add_argument('--include-candidates', action='store_true', help='print judgment_candidates (contains copy text) for judgment_run.py')
    j.add_argument('--answers', help='judgment-answers.json from judgment_run.py')
    j.add_argument('--replacement-out', help='write the shared-ledger replacement here (never overwrites)')
    j.set_defaults(fn=cmd_judge)
    ar = sub.add_parser('accept-reviews', help="record reviews (IA_JUDGE_REVIEW) the owner read and accepted, with the owner's words")
    ar.add_argument('--root', default='.')
    ar.add_argument('--date', help='evaluation date YYYY-MM-DD (default: today)')
    ar.add_argument('--id', action='append', help='judgment id of a current review (repeatable)')
    ar.add_argument('--owner-words', help="the owner's own words accepting them (quoted exactly)")
    ar.add_argument('--approved-on', help='the ISO date the owner said them')
    ar.add_argument('--write', action='store_true', help=f'write {layer_paths.rel("accepted_reviews")} (without it: preview)')
    ar.set_defaults(fn=cmd_accept_reviews)
    e = sub.add_parser('evaluate', help='evaluate generated creatives: check, then dom, then judge, one report and one exit code')
    e.add_argument('--root', default='.')
    e.add_argument('--date', help='evaluation date YYYY-MM-DD (default: today)')
    e.add_argument('--outline', '--script', dest='outline', action='append', help='script id to evaluate (repeatable; default: every script)')
    e.add_argument('--acquire', action='store_true', help='also fetch missing Jev judgments (preview only unless --execute)')
    e.add_argument('--answers-dir', help='directory that keeps judgment-answers.json between runs (outside the repo)')
    e.add_argument('--limit', type=int, help='most judgments to request in this run')
    e.add_argument('--execute', action='store_true', help=f'with --acquire: call Jev (metered) and record {layer_paths.rel("judgments")}')
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
    mt = sub.add_parser('maintain', help='own-session-close Step 2.5: changed SSOT/scripts, evidence falling due, uncovered judgments')
    mt.add_argument('--root', default='.')
    mt.add_argument('--date', help='evaluation date YYYY-MM-DD (default: today)')
    mt.add_argument('--within-days', type=int, help='due window in days (default: references/expiry.json due_within_days)')
    mt.set_defaults(fn=cmd_maintain)
    k = sub.add_parser('skeleton', help='print or write the HTML skeleton (anchors, order, facts) of a script for a factory to style')
    k.add_argument('--root', default='.')
    k.add_argument('--outline', '--script', dest='outline', help='script id (--outline is the v6 name, kept)')
    k.add_argument('--screen', help='a blueprint screen id: print or write its HTML skeleton (marks for regions, groups, actions, states)')
    k.add_argument('--css', help='with --screen: a stylesheet to embed (design\'s allocation CSS)')
    k.add_argument('--out', help='script: a directory; screen: the HTML file to write (never overwritten)')
    k.set_defaults(fn=cmd_skeleton)
    i = sub.add_parser('impact', help='list scripts and outputs that reference an atom id')
    i.add_argument('--root', default='.')
    i.add_argument('--id', required=True)
    i.set_defaults(fn=cmd_impact)
    st = sub.add_parser('set', help='set id of one deliverable: script id + the ids of the paths it points at, and its element hashes')
    st.add_argument('--root', default='.')
    st.add_argument('--outline', '--script', dest='outline', required=True, help='script id (--outline is the v6 name, kept)')
    st.set_defaults(fn=cmd_set)
    df = sub.add_parser('diff', help='difference between two sets: changed elements, changed layers, same experiment or not')
    df.add_argument('--root', default='.')
    df.add_argument('--root-b', help='checkout of the second set (default: --root)')
    df.add_argument('--a', required=True, help='first deliverable (script id)')
    df.add_argument('--b', required=True, help='second deliverable (script id)')
    df.add_argument('--dom-a', help='dom report of the first deliverable (compares text outside anchors)')
    df.add_argument('--dom-b', help='dom report of the second deliverable')
    df.set_defaults(fn=cmd_diff)
    t = sub.add_parser('targets', help='owner-approval targets (ia:path keys and path hashes) for own-design-author approvals')
    t.add_argument('--root', default='.')
    t.set_defaults(fn=cmd_targets)
    m6 = sub.add_parser('migrate-v6', help='one batch v5 migration: SSOT, scripts (from v5 outlines), rendered marks, modules.ia pin')
    m6.add_argument('--root', default='.')
    m6.add_argument('--write', action='store_true', help='write the migrated files (without it: dry run)')
    m6.set_defaults(fn=cmd_migrate_v6)
    md = sub.add_parser('migrate-dictionary', help='one migration of a schema 4 tree: facts and entities to docs/dictionary/, the voice dropped, scripts and rendered marks rewritten')
    md.add_argument('--root', default='.')
    md.add_argument('--map', required=True, help='JSON: {old id: {"id": "<kind>-<slug>", "expires": "YYYY-MM-DD" | "evergreen": true, "variants": [..]}}')
    md.add_argument('--domain', default='product', help='dictionary file docs/dictionary/<domain>.json (default: product)')
    md.add_argument('--date', help='evaluation date YYYY-MM-DD (default: today)')
    md.add_argument('--write', action='store_true', help='write the files (without it: dry run, with the check of the migrated tree)')
    md.set_defaults(fn=cmd_migrate_dictionary)
    ak = sub.add_parser('answer-key', help='answer key for the meaning judgment: pairs, candidates and the labeling page, or --measure')
    ak.add_argument('--root', default='.')
    ak.add_argument('--out', default='answer-key', help='directory to write into (files are never overwritten)')
    ak.add_argument('--measure', help=f'the owner\'s labels.json: measure agreement (own pairs against {layer_paths.rel("judgments")}, swap pairs against the measurement ledger beside --key)')
    ak.add_argument('--key', help='answer-key.json written with the page (required with --measure and --record)')
    ak.add_argument('--record', action='store_true', help='ask Jev the swap pairs not yet answered and record them in the measurement ledger beside --key (a preview unless --execute)')
    ak.add_argument('--limit', type=int, help='most pairs to ask Jev in one --record run')
    ak.add_argument('--execute', action='store_true', help='with --record: really ask Jev (without it: the count only, no key read, no network, no writes)')
    ak.set_defaults(fn=cmd_answer_key)
    args = ap.parse_args(argv)
    try:
        if args.cmd not in GUARD_ITSELF:
            for name in ('root', 'root_b'):  # every other subcommand reads the layers of the checkout(s) it is given
                if getattr(args, name, None):
                    layer_paths.guard(getattr(args, name))
        return args.fn(args)
    except Exception as exc:  # noqa: BLE001 - any crash must be a failure (2), never the warning code (1)
        print(json.dumps({'exit_code': 2, 'error': f'{type(exc).__name__}: {exc}'}, ensure_ascii=False))
        return 2


if __name__ == '__main__':
    sys.exit(main())
