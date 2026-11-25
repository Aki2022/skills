"""IA semantic judgment: candidates for Jev and a read-only check against the shared ledger.

There is one kind, ia-value-fit: a piece of copy does not convey the one value its message points to (information
architecture requirement 11). It shares own-design-author's judgment policy and `ux/judgments.json`
(ADR-20261111-ia-judgments-share-design-policy-and-ledger, biz_ops). The input carries the copy and the value's
content, so changing either never reuses an old judgment. It is advisory until its accuracy is measured with the
owner's answer key (answer_key.py); it replaced v5's ia-strategy-fit.

The four older kinds (name variant, unregistered claim, audience fit, emphasis) are not made, not read and never stop
a shipment (ADR-20261123-ia-skill-completion-judgments-and-measurement). Rows of those kinds may remain in a
repository's ledger: they are ignored, never rewritten. A name variant is the dictionary check's work.

The gate only reads the ledger. It never calls a model: missing judgments are "uncovered", stored ones are
"screened" (never "clean"), and a score at or above theta goes to a human as IA_JUDGE_REVIEW. Neither blocks.

A review the owner read and accepted is recorded, with the owner's words, in `ux/infoarch/accepted_reviews.yaml`
bound to the judgment hash (ADR-20261118-ai-records-owner-approvals, biz_ops). It is then reported as
IA_JUDGE_ACCEPTED (info) instead of a warning. Changing the copy or its context changes the hash, so the record stops
applying (IA_JUDGE_ACCEPTED_STALE, info) and the review comes back. A broken or symlinked record fails closed.
"""
from __future__ import annotations

import datetime as dt
import json
import os
import sys
from pathlib import Path

import check
import outline as ol
from common import Report
import layer_paths  # on sys.path through common

DESIGN_SCRIPTS = Path(__file__).resolve().parents[2] / 'own-design-author/scripts'
# Repository-relative names, derived from the layer rule; every read and write goes through layer_paths.path (the guard).
LEDGER = Path(layer_paths.rel('judgments'))
ACCEPTED = Path(layer_paths.rel('accepted_reviews'))
# Only the copy of a claim block conveys a value (requirement 11). Facts, evidence (with testimonials), offers and
# the call to action are not there to convey one: in the I16 example and the officeohana pilot the owner labeled every
# such copy "does not convey" (v6's trust and act), so asking only produced reviews. They are not asked, nor keyed.
VALUE_FIT_FUNCTIONS = ('claim',)
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


def _value(ssot, vid):
    return {'id': vid, 'text': (ssot.get('values', {}).get(vid) or {}).get('text', '')}


def outline_candidates(oid, outline, ssot, policy, jl):
    """One candidate per piece of copy of a claim block whose message points to a value with text."""
    facts = ssot.get('facts', {})
    rows = []
    for sid, blk, copies in ol.ordered_blocks(outline):
        if blk['function'] not in VALUE_FIT_FUNCTIONS:
            continue
        for cid, c in copies:
            text = _expand(c['text'], facts)
            msg = outline['messages'].get(c.get('message'))
            if text.strip() and msg and (ssot.get('values', {}).get(msg['value']) or {}).get('text'):
                rows.append(value_fit(jl, policy, f'{oid}/{sid}/{cid}', text, oid, c['message'], _value(ssot, msg['value'])))
    return rows


def value_fit(jl, policy, ident, text, oid, message, value):
    """The one question Jev is asked about meaning: does this copy convey this value (v6 requirement 11)."""
    return jl.candidate('ia-value-fit', ident, text, {'outline': oid, 'message': message, 'value': value}, policy=policy)


def read_ledger(root):
    path = layer_paths.path(Path(root), 'judgments')
    if any(p.is_symlink() for p in (path, path.parent)):
        raise ValueError('判定台帳が symlink')
    if not path.exists():
        return None
    return json.loads(path.read_text(encoding='utf-8'))


def read_accepted(root):
    """-> {judgment hash: row} of reviews the owner accepted; {} when there is no record. Raises on any damage."""
    path = layer_paths.path(Path(root), 'accepted_reviews')
    if any(p.is_symlink() for p in (path, path.parent)):
        raise ValueError(f'{ACCEPTED} が symlink')
    if not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding='utf-8'))
    except json.JSONDecodeError as exc:
        raise ValueError(f'{ACCEPTED} が読めない: {exc}') from None
    if not isinstance(data, dict) or set(data) != {'schema_version', 'rows'} or data['schema_version'] != 1 \
            or not isinstance(data['rows'], list):
        raise ValueError(f'{ACCEPTED} の形が違う（schema_version 1 と rows だけ）')
    out = {}
    for row in data['rows']:
        if not isinstance(row, dict) or set(row) != {'judgment_hash', 'id', 'approver', 'date'} \
                or not all(isinstance(row[k], str) and row[k].strip() for k in row):
            raise ValueError(f'{ACCEPTED} の行は judgment_hash・id・approver・date だけ')
        try:
            ok = dt.date.fromisoformat(row['date']).isoformat() == row['date']
        except ValueError:
            ok = False
        if not ok or row['judgment_hash'] in out:
            raise ValueError(f'{ACCEPTED} の行の日付が ISO でない・同じ判定が 2 行ある')
        out[row['judgment_hash']] = row
    return out


def evaluate(rows, ledger, *, policy=None, jl=None, accepted=None):
    report = Report(level='judge', evaluated_on='')
    report.screened, report.missing = [], []
    if not rows:
        # An outline of only fact and call-to-action copy has nothing to judge. That is a visible gap, never a green
        # "judged" and never a stop: no judgment stops a shipment.
        report.add('warn', 'IA_JUDGE_NO_TARGETS', '判定の候補が 0 件（文案 × 提供価値の対象が無い。判定していない）')
        return report
    jl = jl or _jl()
    result = jl.evaluate(rows, ledger, policy=policy)
    report.uncovered = result['uncovered']
    report.screened = result['screened']
    if result['uncovered']:
        report.add('warn', 'IA_JUDGE_UNCOVERED', f'{len(result["uncovered"])} 件の判定が台帳に無い（未カバー。Jev で取得して記録する）')
    accepted = accepted or {}
    taken = []
    for w in result['warnings']:
        if w['judgment_hash'] in accepted:
            taken.append(w)
            continue
        report.add('warn', 'IA_JUDGE_REVIEW', f'要確認: {w["id"]} の score {w["score"]} が θ {w["theta"]} 以上。原文を人が確かめる',
                   id=w['id'], judgment_hash=w['judgment_hash'], score=w['score'])
    if taken:
        report.add('info', 'IA_JUDGE_ACCEPTED', f'{len(taken)} 件の要確認は、オーナーが原文を見て認めた（{ACCEPTED}）',
                   ids=[w['id'] for w in taken], judgment_hashes=[w['judgment_hash'] for w in taken])
    stale = sorted(r['id'] for h, r in accepted.items() if h not in {w['judgment_hash'] for w in result['warnings']})
    if stale:
        report.add('info', 'IA_JUDGE_ACCEPTED_STALE', f'{len(stale)} 件の「認めた」記録は今の判定に当たらない（文案・文脈が変わったか、θ 未満になった）',
                   ids=stale)
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


def replacement(ledger, rows, answers, policy, jl, outlines=None):
    """Shared-ledger replacement: keep every non-IA row, take answered current IA candidates, keep IA rows that are
    still valid, drop only superseded ones and those of outlines that no longer exist.

    A run may cover part of the IA (one deliverable's render): a row whose candidate id this run did not produce
    belongs to work outside the run and stays while its outline exists. A row whose id the run produced with a
    different hash is superseded."""
    if answers.get('policy') != policy or answers.get('status') != 'model-answers-unlabeled':
        raise ValueError('回答の policy が現在の共有 policy と違う（migrate が要る）')
    if ledger is not None and ledger.get('schema_version') != jl.CURRENT_SCHEME:
        raise ValueError('共有の台帳が古い鍵の形（schema 2 以前）。先に judgment_ledger.py rekey で移す')
    current = {r['judgment_hash']: r for r in rows}
    answered = {}
    for a in answers['rows']:
        c = a['candidate']
        if c['judgment_hash'] in current and c == current[c['judgment_hash']]:
            answered[c['judgment_hash']] = {'candidate': c, 'score': jl.probability(a['score'])}
    old = (ledger or {}).get('rows', [])
    design = [r for r in old if not r['candidate']['input']['kind'].startswith(IA_PREFIX)]
    current_ids = {r['id']: r['judgment_hash'] for r in rows}
    kept, dropped = [], 0
    for r in old:
        c = r['candidate']
        if not c['input']['kind'].startswith(IA_PREFIX) or c['judgment_hash'] in answered:
            continue
        if c['id'] in current_ids:
            valid = current_ids[c['id']] == c['judgment_hash']
        else:
            outline = c['id'].split(':', 1)[1].split('/', 1)[0].split('@', 1)[0]
            valid = outlines is not None and outline in outlines
        if valid:
            kept.append(r)
        else:
            dropped += 1
    out = {'schema_version': jl.CURRENT_SCHEME, 'policy': policy, 'rows': design + kept + list(answered.values()),
           'holdout': (ledger or {}).get('holdout', {'source': 'unmeasured'})}
    jl.validate_ledger(out, policy)
    summary = {'design_rows': len(design), 'ia_rows': len(answered), 'kept_ia_rows': len(kept),
               'unanswered': len(current) - len(answered), 'dropped_ia_rows': dropped}
    return out, summary


def run(root, date, answers=None, replacement_out=None):
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
    rows = []
    for oid, (_, outline) in sorted(m.outlines.items()):
        rows += outline_candidates(oid, outline, m.ssot, policy, jl)
    try:
        accepted = read_accepted(root)
    except layer_paths.LayerLayoutError:
        raise  # a moved repository stops the command; it is not a damaged record
    except (ValueError, OSError) as exc:
        accepted = {}  # a broken record never hides a review
        pre.add('fail', 'IA_JUDGE_ACCEPTED_FILE', f'認めた要確認の記録が読めない: {exc}', ACCEPTED)
    try:
        ledger = read_ledger(root)
        report = evaluate(rows, ledger, policy=policy, jl=jl, accepted=accepted)
    except layer_paths.LayerLayoutError:
        raise
    except (ValueError, KeyError, TypeError) as exc:
        report = Report(level='judge', evaluated_on=date)
        report.screened, report.missing, ledger = [], [], None
        report.add('fail', 'IA_JUDGE_LEDGER', f'判定台帳 {LEDGER} が読めない・policy と合わない: {exc}')
    report.evaluated_on = date
    report.findings = pre.findings + report.findings
    report.targets = {'candidates': len(rows), 'outlines': len(m.outlines),
                      'by_kind': {k: sum(r['input']['kind'] == k for r in rows) for k in sorted({r['input']['kind'] for r in rows})}}
    summary = None
    if answers is not None and report.exit_code < 2:
        out, summary = replacement(ledger, rows, read_answers(answers, policy), policy, jl, outlines=set(m.outlines))
        write_new(replacement_out, json.dumps(out, ensure_ascii=False, indent=2) + '\n')
    return report, rows, summary


def accept_reviews(root, date, ids, owner_words, approved_on, write=False):
    """Record reviews the owner accepted. Only a review judge reports now can be accepted; nothing is written without
    the owner's words and the ISO date they were given, nor without `write`."""
    _design_path()
    from owner_approval import _owner_approval
    owner = _owner_approval(owner_words, approved_on)
    if not ids:
        raise ValueError('認める要確認の id を 1 つ以上渡す')
    report, _, _ = run(root, date)
    if report.exit_code == 2:
        raise ValueError('judge が失格しているので記録しない（先に直す）: ' + ', '.join(report.codes('fail')))
    current = {}
    for f in report.findings:
        if f.code == 'IA_JUDGE_REVIEW':
            current[f.detail['id']] = f.detail['judgment_hash']
        elif f.code == 'IA_JUDGE_ACCEPTED':
            current.update(zip(f.detail['ids'], f.detail['judgment_hashes']))
    unknown = sorted(set(ids) - set(current))
    if unknown:
        raise ValueError(f'今の要確認に無い id は認められない: {unknown}')
    rows = read_accepted(root)
    new = [{'judgment_hash': current[i], 'id': i, 'approver': 'owner', 'date': owner['date']}
           for i in dict.fromkeys(ids) if current[i] not in rows]
    out = {'exit_code': 0, 'status': 'recorded' if write else 'preview', 'recorded': len(new) if write else 0,
           'ids': list(dict.fromkeys(ids)), 'owner_approval': owner,
           'commit_trailer': f'Owner-Approval: 「{owner["words"]}」 ({owner["date"]})'}
    if write and new:
        path = layer_paths.path(Path(root), 'accepted_reviews')
        data = {'schema_version': 1, 'rows': sorted(list(rows.values()) + new, key=lambda r: (r['id'], r['judgment_hash']))}
        tmp = path.with_name(path.name + '.tmp')
        if tmp.is_symlink() or tmp.exists():
            raise ValueError(f'一時ファイルが既にある: {tmp}')
        tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
        os.replace(tmp, path)
    return out
