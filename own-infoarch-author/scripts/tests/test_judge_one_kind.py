"""WS-20261123 ISSUE-01: the only meaning judgment is "copy x the value its message points to" (ia-value-fit).

Spec requirement 11 / ADR-20261123: the four older kinds (name variant, unregistered claim, audience fit, emphasis)
are not made, not read and never stop a shipment. Old rows of those kinds may remain in a repository's ledger; they
are ignored without being rewritten.
"""
import json
import re
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))
from test_judge import DATE, cli, candidates, jl, layer_paths, repo, write_ledger  # noqa: E402,F401  (repo is a fixture)

SCRIPTS = Path(__file__).resolve().parents[1]
REMOVED = ('ia-name-variant', 'ia-unregistered-claim', 'ia-audience-fit', 'ia-emphasis')
PATTERN = re.compile('|'.join(REMOVED))


def test_judge_makes_only_value_fit_candidates(repo):
    code, out, rows = candidates(repo)
    assert rows, 'the fixture must have judgeable copy (a vacuous pass is not a pass)'
    assert {r['input']['kind'] for r in rows} == {'ia-value-fit'}
    assert set(out['targets']['by_kind']) == {'ia-value-fit'}
    assert not PATTERN.search(json.dumps(out, ensure_ascii=False))


def test_no_script_makes_or_reads_a_removed_kind():
    """The guard that turns red when a removed kind is put back: the scripts (not the tests) never name one."""
    files = sorted(p for p in SCRIPTS.glob('*.py'))
    assert len(files) >= 10, files  # the scan covers the skill's scripts, not an empty directory
    named = {p.name: PATTERN.findall(p.read_text(encoding='utf-8')) for p in files}
    assert {k: v for k, v in named.items() if v} == {}, named


def test_old_rows_in_the_ledger_are_ignored_not_screened_or_reported(repo):
    _, _, rows = candidates(repo)
    old = [jl.candidate(kind, 'lp-municipal/blk-0001/cpy-0001', '運用の手間が半分になります', {'outline': 'lp-municipal'})
           for kind in REMOVED[:3]]
    old.append(jl.candidate('ia-emphasis', 'lp-municipal/site/index.html@mobile', '運用の手間が半分になります', {'outline': 'lp-municipal'}))
    write_ledger(repo, [{'candidate': r, 'score': 0.99} for r in old] + [{'candidate': r, 'score': 0.0} for r in rows])
    code, out = cli('judge', '--root', str(repo), '--date', DATE)
    assert code == 0, out['findings']  # 0.99 old rows would be review warnings if they were read
    assert out['uncovered'] == [] and len(out['screened']) == len(rows)
    assert not PATTERN.search(json.dumps(out, ensure_ascii=False))


def test_missing_value_fit_answers_are_uncovered_but_never_a_failure(repo):
    code, out = cli('judge', '--root', str(repo), '--date', DATE)
    assert code == 1 and out['uncovered'], out
    assert all(u['id'].startswith('ia-value-fit:') for u in out['uncovered'])
    assert 'IA_JUDGE_EMPHASIS_MISSING' not in [f['code'] for f in out['findings']]
    assert out['missing'] == []


def test_judge_takes_no_dom_report(repo, tmp_path):
    """Nothing a render produces feeds a judgment any more; a dom report must not be accepted and silently dropped."""
    report = tmp_path / 'dom.json'
    report.write_text(json.dumps({'level': 'dom', 'exit_code': 0, 'first_screens': []}), encoding='utf-8')
    proc = __import__('subprocess').run([sys.executable, str(SCRIPTS / 'infoarch_author.py'), 'judge', '--root', str(repo),
                                         '--date', DATE, '--dom-report', str(report)], capture_output=True, text=True)
    assert proc.returncode != 0 and 'unrecognized arguments' in proc.stderr, proc.stderr


def test_copy_without_a_value_to_judge_is_reported_not_a_failure(repo):
    """Outlines whose copy is all trust/act have no candidate; that is a visible gap (warn), not a stop."""
    from judge import evaluate
    report = evaluate([], None)
    assert report.exit_code == 1 and 'IA_JUDGE_NO_TARGETS' in report.codes('warn')
