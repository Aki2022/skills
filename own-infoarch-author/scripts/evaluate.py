"""One evaluation entry for a generated creative: check -> dom -> judge, one report, one exit code.

A factory runs this after rendering. It stops at the first stage that fails (a failed check is not rendered,
a failed render is not judged), merges the findings without repeating them, and reports each stage's exit code.

With `acquire`, it also fetches the missing Jev judgments, bounded by `limit` and the shared usage ledger:
without `execute` it only previews (no key, no network, no writes); with `execute` it asks Jev through
own-design-author's `judgment_run`, records the answers in `design/judgments.json` through
`judgment_ledger.py` (record when no ledger exists, refresh otherwise), and evaluates again.
"""
from __future__ import annotations

import json
import subprocess
import sys
import tempfile
from pathlib import Path

import check
import dom
import judge
from common import Report


def _merge(target, findings, seen):
    for f in findings:
        key = (f['level'], f['code'], f['message'], f['path'])
        if key not in seen:
            seen.add(key)
            target.append(f)


def _stages(root, date, outlines, work):
    """-> (stages, findings, judge report or None, candidates, dom report paths)."""
    findings, seen = [], set()
    base, m = check.check(root, date)
    _merge(findings, base.to_dict()['findings'], seen)
    stages = {'check': base.exit_code, 'dom': 'skipped', 'judge': 'skipped'}
    if base.exit_code == 2 or m.ssot is None:
        return stages, findings, None, [], []
    targets = list(outlines) if outlines else sorted(m.outlines)
    unknown = [o for o in targets if o not in m.outlines]
    if unknown:
        _merge(findings, [{'level': 'fail', 'code': 'IA_EVALUATE_OUTLINE', 'message': f'アウトライン {unknown} が無い',
                           'path': str(check.INFO_ARCH_DIR / 'outlines')}], seen)
        return stages, findings, None, [], []
    stages['dom'], paths = {}, []
    for oid in targets:
        rep = dom.run(root, oid, date)
        data = rep.to_dict()
        data['first_screens'] = getattr(rep, 'first_screens', [])
        data['unanchored'] = getattr(rep, 'unanchored', [])
        stages['dom'][oid] = rep.exit_code
        _merge(findings, data['findings'], seen)
        if data['first_screens']:  # an outline with no HTML output has nothing for judge to read
            path = Path(work) / f'dom-{len(paths)}.json'
            path.write_text(json.dumps(data, ensure_ascii=False, default=str), encoding='utf-8')
            paths.append(path)
    if any(code == 2 for code in stages['dom'].values()):
        return stages, findings, None, [], paths
    jrep, rows, _ = judge.run(root, date, [str(p) for p in paths])
    stages['judge'] = jrep.exit_code
    _merge(findings, jrep.to_dict()['findings'], seen)
    return stages, findings, jrep, rows, paths


def _record(root, replacement):
    """Write the shared ledger through own-design-author's recorder (its lock and checks), never directly."""
    command = 'refresh' if (Path(root) / judge.LEDGER).exists() else 'record'
    proc = subprocess.run([sys.executable, str(judge.DESIGN_SCRIPTS / 'judgment_ledger.py'), command, '--repo', str(root),
                           '--replacement', str(replacement), '--write'], capture_output=True, text=True, check=False)
    if proc.returncode != 0:
        raise ValueError(f'判定台帳に書けない（judgment_ledger.py {command}: {proc.stdout.strip()[:200]}）')
    return {'record': 'recorded', 'refresh': 'refreshed'}[command]


def _acquire(root, date, rows, paths, acquire, work):
    jr = judge._jr()
    answers_dir = Path(acquire['answers_dir'])
    answers = answers_dir / 'judgment-answers.json'
    source = Path(work) / 'candidates.json'
    source.write_text(json.dumps({'judgment_candidates': rows}, ensure_ascii=False), encoding='utf-8')
    result = jr.run(source, answers, acquire['limit'], dry_run=not acquire.get('execute'),
                    _transport=acquire.get('_transport'), _usage_path=acquire.get('_usage_path', jr.USAGE))
    if acquire.get('execute') and answers.exists():
        replacement = Path(work) / 'replacement.json'
        rep, _, _ = judge.run(root, date, [str(p) for p in paths], str(answers), str(replacement))
        if rep.exit_code == 2:
            raise ValueError('判定台帳の置き換えを作れない（judge が失格）')
        result['ledger'] = _record(root, replacement)
    return result


def run(root, date, outlines=None, *, acquire=None):
    """-> result dict (exit_code, stages, findings, uncovered, screened, missing, candidates[, acquisition])."""
    with tempfile.TemporaryDirectory(prefix='infoarch-evaluate-') as tmp:
        work = Path(tmp).resolve()  # the shared reader refuses paths through a symlink (/var on macOS)
        stages, findings, jrep, rows, paths = _stages(root, date, outlines, work)
        acquisition = None
        if acquire is not None:
            if jrep is None or jrep.exit_code == 2:
                acquisition = {'status': 'refused', 'reason': '評価が失格しているので判定を取らない（先に直す）'}
            else:
                try:
                    acquisition = _acquire(root, date, rows, paths, acquire, work)
                except Exception as exc:  # noqa: BLE001 - the acquisition error names itself; never a pass
                    acquisition = {'status': 'failed', 'error': str(exc) or type(exc).__name__}
                    findings.append({'level': 'fail', 'code': 'IA_EVALUATE_ACQUIRE',
                                     'message': f'Jev の判定を取れない: {acquisition["error"]}', 'path': str(judge.LEDGER)})
                else:
                    if acquisition.get('ledger'):
                        stages, findings, jrep, rows, paths = _stages(root, date, outlines, work)
    report = Report(level='evaluate', evaluated_on=date)
    for f in findings:
        report.add(f['level'], f['code'], f['message'], f['path'], f.get('line', 0), **f.get('detail', {}))
    result = report.to_dict()
    result.update(stages=stages, candidates=rows,
                  uncovered=getattr(jrep, 'uncovered', []), screened=getattr(jrep, 'screened', []),
                  missing=getattr(jrep, 'missing', []))
    if acquisition is not None:
        result['acquisition'] = acquisition
    return result
