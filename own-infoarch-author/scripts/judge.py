"""IA semantic judgments (ISSUE-04): candidates for Jev and a read-only check against the shared ledger.

Four kinds share own-design-author's judgment policy and `design/judgments.json`
(ADR-20261111-ia-judgments-share-design-policy-and-ledger, biz_ops):
- ia-name-variant: a section's copy writes an official name in another notation.
- ia-unregistered-claim: a section's copy claims something its message / proofs / facts do not support.
- ia-audience-fit: the outline's order and emphasis do not suit its audience and brief.
- ia-emphasis: a rendered first screen does not carry the audience's top-priority message.

The gate only reads the ledger. It never calls a model: missing judgments are "uncovered", stored ones are
"screened" (never "clean"), and a score at or above theta goes to a human as IA_JUDGE_REVIEW.
"""
from __future__ import annotations

import json
import os
import sys
from pathlib import Path

import check
import dom
from common import Report

DESIGN_SCRIPTS = Path(__file__).resolve().parents[2] / 'own-design-author/scripts'
LEDGER = Path('design/judgments.json')
IA_PREFIX = 'ia-'


def _design_path():
    if not (DESIGN_SCRIPTS / 'judgment_ledger.py').is_file():
        raise RuntimeError('own-design-author/scripts/judgment_ledger.py が無い（判定台帳は design と共有）')
    if str(DESIGN_SCRIPTS) not in sys.path:
        sys.path.insert(0, str(DESIGN_SCRIPTS))


def _jl():
    _design_path()
    import judgment_ledger
    return judgment_ledger


def _jr():
    # The answers file is validated and the output path guarded by the same reader that wrote them.
    _design_path()
    import judgment_run
    return judgment_run


def _expand(copy, facts):
    return check.PLACEHOLDER.sub(lambda m: str(facts.get(m.group(1), {}).get('value', m.group(0))), copy)


def outline_candidates(oid, outline, ssot, policy, jl):
    facts = ssot.get('facts', {})
    entities = [{'id': eid, 'name': e['name']} for eid, e in sorted(ssot.get('entities', {}).items()) if e.get('name')]
    rows, lines = [], []
    for page in outline['pages']:
        for i, s in enumerate(page['sections']):
            text = _expand(s.get('copy', ''), facts)
            lines.append(f"{s['role']}: {text}")
            if not text.strip():
                continue
            ident = f"{oid}/{page['id']}/{i}"
            if entities:
                rows.append(jl.candidate('ia-name-variant', ident, text, {'outline': oid, 'entities': entities}, policy=policy))
            msg = ssot.get('messages', {}).get(s.get('message'), {})
            ctx = {'outline': oid, 'role': s['role'],
                   'message': {'id': s['message'], 'text': msg.get('text', '')} if 'message' in s else None,
                   'problem': ssot.get('problems', {}).get(s.get('problem'), {}).get('text') if 'problem' in s else None,
                   'proofs': [{'id': p, 'text': ssot.get('proofs', {}).get(p, {}).get('text', '')} for p in s.get('proofs', [])],
                   'facts': [{'id': f, 'label': facts.get(f, {}).get('label', ''), 'value': facts.get(f, {}).get('value', '')}
                             for f in dict.fromkeys(check.PLACEHOLDER.findall(s.get('copy', '')))]}
            rows.append(jl.candidate('ia-unregistered-claim', ident, text, ctx, policy=policy))
    aud = ssot.get('audiences', {}).get(outline['audience'], {})
    # The priority order is what the fit is judged against, so reordering it must invalidate the judgment.
    priorities = [{'id': p, 'text': ssot.get('messages', {}).get(p, {}).get('text', '')} for p in aud.get('priorities') or []]
    fit_ctx = {'outline': oid, 'audience': {**{k: aud.get(k) for k in ('label', 'situation', 'desired_action')},
                                            'priorities': priorities},
               'brief': outline.get('brief', {})}
    rows.append(jl.candidate('ia-audience-fit', oid, '\n'.join(lines), fit_ctx, policy=policy))
    return rows


def emphasis_candidates(oid, outline, ssot, screens, policy, jl):
    aud = ssot.get('audiences', {}).get(outline['audience'], {})
    top = (aud.get('priorities') or [None])[0]
    top_msg = {'id': top, 'text': ssot.get('messages', {}).get(top, {}).get('text', '')}
    return [jl.candidate('ia-emphasis', f"{oid}/{s['file']}@{s['viewport']}", s['text'],
                         {'outline': oid, 'file': s['file'], 'viewport': s['viewport'],
                          'audience': aud.get('label'), 'top_message': top_msg}, policy=policy)
            for s in screens]


def expected_screens(outline, channels):
    """(file, viewport) pairs `dom` renders for an outline: its HTML outputs times the channel's viewports."""
    ch = channels.get(outline['channel'])
    viewports = (ch[1].get('viewports') if ch else None) or dom.DEFAULT_VIEWPORTS
    return [(o, v['name']) for o in outline['outputs'] if Path(o).suffix.lower() in dom.HTML_SUFFIXES for v in viewports]


def read_dom_reports(paths, expected, report):
    """-> {outline id: [first screens]}; a failed or malformed render is never judged.

    `expected` maps outline id -> the (file, viewport) pairs dom renders. A row outside them, or a duplicate,
    fails; row text is never echoed into findings (it is unpublished copy)."""
    screens, seen = {}, set()
    for path in paths:
        data = json.loads(Path(path).read_text(encoding='utf-8'))
        if data.get('exit_code') != 0:
            report.add('fail', 'IA_JUDGE_DOM_FAILED', f'{path}: dom が合格していない描画物は判定しない（exit {data.get("exit_code")}）')
            continue
        rows = data.get('first_screens')
        if not isinstance(rows, list) or not rows:
            report.add('fail', 'IA_JUDGE_DOM_REPORT', f'{path}: first_screens が空（対象 0 件を合格にしない）')
            continue
        for i, s in enumerate(rows):
            ok = (isinstance(s, dict) and all(isinstance(s.get(k), str) and s[k].strip() for k in ('outline', 'file', 'viewport', 'text'))
                  and (s['file'], s['viewport']) in expected.get(s['outline'], ()))
            key = (s['outline'], s['file'], s['viewport']) if ok else None
            if not ok or key in seen:
                report.add('fail', 'IA_JUDGE_DOM_REPORT',
                           f'{path}: first_screens[{i}] が読めない（型・未知のアウトライン/成果物/画面幅・重複のどれか）')
                continue
            seen.add(key)
            screens.setdefault(s['outline'], []).append(s)
    return screens


def read_ledger(root):
    path = Path(root) / LEDGER
    if any(p.is_symlink() for p in (path, path.parent)):
        raise ValueError('判定台帳が symlink')
    if not path.exists():
        return None
    return json.loads(path.read_text(encoding='utf-8'))


def evaluate(rows, ledger, *, missing, policy=None, jl=None):
    report = Report(level='judge', evaluated_on='')
    report.screened, report.missing = [], list(missing)
    if not rows:
        report.add('fail', 'IA_JUDGE_NO_TARGETS', '判定の候補が 0 件（対象 0 件を合格にしない）')
        return report
    jl = jl or _jl()
    result = jl.evaluate(rows, ledger, policy=policy)
    report.uncovered = result['uncovered']
    report.screened = result['screened']
    if result['uncovered']:
        report.add('warn', 'IA_JUDGE_UNCOVERED', f'{len(result["uncovered"])} 件の判定が台帳に無い（未カバー。Jev で取得して記録する）')
    for w in result['warnings']:
        report.add('warn', 'IA_JUDGE_REVIEW', f'要確認: {w["id"]} の score {w["score"]} が θ {w["theta"]} 以上。原文を人が確かめる',
                   id=w['id'], judgment_hash=w['judgment_hash'], score=w['score'])
    for m in report.missing:
        # A shipped HTML outline without its first screen is a gap (warn); other media are reported as missing only.
        where = f'{m["outline"]}/{m["file"]}@{m["viewport"]}' if 'viewport' in m else m['outline']
        report.add('warn' if m.get('shipped_html') else 'info', 'IA_JUDGE_EMPHASIS_MISSING',
                   f'{where}: 最初の画面の入力（dom の報告）が無いので強調を判定していない')
    return report


def read_answers(path, policy):
    """judgment_run.py's own reader: shape, policy, candidate hashes, and score == raw_answer.noul."""
    jr = _jr()
    if not jr._safe_path(Path(path)):
        raise ValueError(f'回答ファイルが symlink: {path}')
    try:
        answers, _ = jr._answers(Path(path), policy)
    except jr.api.ClassificationError as exc:
        raise ValueError(f'回答ファイルが読めない・policy と合わない・score と raw_answer が違う（{exc}）') from None
    return answers


def write_new(target, text):
    """Create `target` exclusively: never overwrite, never follow a symlink (the path or any parent)."""
    target = Path(target)
    if not _jr()._safe_path(target):
        raise ValueError(f'出力先が symlink を通る: {target}')
    try:
        fd = os.open(target, os.O_WRONLY | os.O_CREAT | os.O_EXCL | os.O_NOFOLLOW, 0o644)
    except FileExistsError:
        raise ValueError(f'上書きしない: {target}') from None
    with os.fdopen(fd, 'w', encoding='utf-8') as f:
        f.write(text)


def replacement(ledger, rows, answers, policy, jl):
    """Shared-ledger replacement: keep every non-IA row, take answered current IA candidates, drop stale IA rows."""
    if answers.get('policy') != policy or answers.get('status') != 'model-answers-unlabeled':
        raise ValueError('回答の policy が現在の共有 policy と違う（migrate が要る）')
    current = {r['judgment_hash']: r for r in rows}
    answered = {}
    for a in answers['rows']:
        c = a['candidate']
        if c['judgment_hash'] in current and c == current[c['judgment_hash']]:
            answered[c['judgment_hash']] = {'candidate': c, 'score': jl.probability(a['score'])}
    old = (ledger or {}).get('rows', [])
    design = [r for r in old if not r['candidate']['input']['kind'].startswith(IA_PREFIX)]
    dropped = sum(1 for r in old if r['candidate']['input']['kind'].startswith(IA_PREFIX)
                  and r['candidate']['judgment_hash'] not in answered)
    out = {'schema_version': 2, 'policy': policy, 'rows': design + list(answered.values()),
           'holdout': (ledger or {}).get('holdout', {'source': 'unmeasured'})}
    jl.validate_ledger(out, policy)
    summary = {'design_rows': len(design), 'ia_rows': len(answered),
               'unanswered': len(current) - len(answered), 'dropped_ia_rows': dropped}
    return out, summary


def run(root, date, dom_reports=(), answers=None, replacement_out=None):
    """-> (report, candidates, replacement summary or None)."""
    base, m = check.check(root, date)
    if m.ssot is None or base.exit_code == 2:
        rep = Report(level='judge', evaluated_on=date)
        rep.screened, rep.missing = [], []
        rep.add('fail', 'IA_JUDGE_REFUSED', 'check が失格しているので判定しない。先に直す')
        return rep, [], None
    jl = _jl()
    policy = jl.load_policy()
    pre = Report(level='judge', evaluated_on=date)
    expected = {oid: expected_screens(outline, m.channels) for oid, (_, outline) in m.outlines.items()}
    screens = read_dom_reports(dom_reports, expected, pre)
    rows, missing = [], []
    for oid, (_, outline) in sorted(m.outlines.items()):
        rows += outline_candidates(oid, outline, m.ssot, policy, jl)
        ch = m.channels.get(outline['channel'])
        shipped = bool(ch and ch[1].get('ship'))
        if oid in screens:
            rows += emphasis_candidates(oid, outline, m.ssot, screens[oid], policy, jl)
            # A partial report (some files or viewports only) must not read as covered.
            got = {(s['file'], s['viewport']) for s in screens[oid]}
            missing += [{'kind': 'ia-emphasis', 'outline': oid, 'file': f, 'viewport': v, 'shipped_html': shipped,
                         'reason': 'dom の報告（first_screens）にこの画面幅が無い'}
                        for f, v in expected[oid] if (f, v) not in got]
        else:
            missing.append({'kind': 'ia-emphasis', 'outline': oid, 'shipped_html': shipped and bool(expected[oid]),
                            'reason': 'dom の報告（first_screens）が渡されていない'})
    try:
        ledger = read_ledger(root)
        report = evaluate(rows, ledger, missing=missing, policy=policy, jl=jl)
    except (ValueError, KeyError, TypeError) as exc:
        report = Report(level='judge', evaluated_on=date)
        report.screened, report.missing, ledger = [], missing, None
        report.add('fail', 'IA_JUDGE_LEDGER', f'判定台帳 {LEDGER} が読めない・policy と合わない: {exc}')
    report.evaluated_on = date
    report.findings = pre.findings + report.findings
    report.targets = {'candidates': len(rows), 'outlines': len(m.outlines),
                      'by_kind': {k: sum(r['input']['kind'] == k for r in rows) for k in sorted({r['input']['kind'] for r in rows})}}
    summary = None
    if answers is not None and report.exit_code < 2:
        out, summary = replacement(ledger, rows, read_answers(answers, policy), policy, jl)
        write_new(replacement_out, json.dumps(out, ensure_ascii=False, indent=2) + '\n')
    return report, rows, summary
