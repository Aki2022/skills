"""`evaluate`: one entry that runs check -> dom -> judge and returns one exit code; optional bounded Jev acquisition."""
import copy
import json
import shutil
import subprocess
import sys
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SCRIPTS))

import dom  # noqa: E402
import evaluate  # noqa: E402
from test_check import DATE, write_repo  # noqa: E402
from test_dom import green_snapshot  # noqa: E402
from test_judge import write_ledger  # noqa: E402

DESIGN = SCRIPTS.parents[1] / 'own-design-author'
sys.path.insert(0, str(DESIGN / 'scripts'))
import judgment_ledger as jl  # noqa: E402


def fake_snapshot(mutate=None):
    def snap(root, html, viewports):
        out = []
        for v in viewports:
            s = green_snapshot()
            s['viewport'] = v
            s['unanchored'] = ''
            if mutate:
                mutate(s)
            out.append(s)
        return out
    return snap


@pytest.fixture
def repo(tmp_path, monkeypatch):
    root = tmp_path / 'repo'
    write_repo(root)
    (root / 'site').mkdir()
    (root / 'site/index.html').write_text('<p>rendered by the fake snapshot</p>', encoding='utf-8')
    monkeypatch.setattr(dom, 'snapshot', fake_snapshot())
    return root


def codes(result, level=None):
    return [f['code'] for f in result['findings'] if level is None or f['level'] == level]


def cover(root, result, score=0.1):
    """Record every current candidate in the shared ledger with `score`."""
    rows = [{'candidate': c, 'score': score} for c in result['candidates']]
    write_ledger(root, rows)


def test_uncovered_judgments_make_it_a_warning(repo):
    result = evaluate.run(repo, DATE, ['lp-municipal'])
    assert result['exit_code'] == 1 and 'IA_JUDGE_UNCOVERED' in codes(result, 'warn'), result['findings']
    assert result['stages'] == {'check': 0, 'dom': {'lp-municipal': 0}, 'judge': 1}


def test_covered_and_below_theta_is_green(repo):
    cover(repo, evaluate.run(repo, DATE, ['lp-municipal']))
    result = evaluate.run(repo, DATE, ['lp-municipal'])
    assert result['exit_code'] == 0, result['findings']
    assert result['screened'] and not result['uncovered']


def test_score_at_theta_goes_to_review(repo):
    cover(repo, evaluate.run(repo, DATE, ['lp-municipal']), score=0.9)
    result = evaluate.run(repo, DATE, ['lp-municipal'])
    assert result['exit_code'] == 1 and 'IA_JUDGE_REVIEW' in codes(result, 'warn')


def test_rewritten_copy_fails_and_is_not_judged(repo, monkeypatch):
    monkeypatch.setattr(dom, 'snapshot', fake_snapshot(lambda s: s['anchors'][1].update(text='運用の手間がゼロになります')))
    result = evaluate.run(repo, DATE, ['lp-municipal'])
    assert result['exit_code'] == 2 and 'IA_DOM_COPY_DRIFT' in codes(result, 'fail'), result['findings']
    assert result['stages']['judge'] == 'skipped'


def test_unanchored_text_is_warned_and_sent_to_judgment(repo, monkeypatch):
    monkeypatch.setattr(dom, 'snapshot', fake_snapshot(lambda s: s.update(unanchored='業界No.1 の実績')))
    result = evaluate.run(repo, DATE, ['lp-municipal'])
    assert 'IA_DOM_UNANCHORED_TEXT' in codes(result, 'warn')
    assert any(c['input']['catalog_id'].endswith('#unanchored') for c in result['candidates'])


def test_failing_check_stops_before_rendering(repo, monkeypatch):
    monkeypatch.setattr(dom, 'snapshot', lambda *a: pytest.fail('rendered although check failed'))
    outline = repo / 'design/info_arch/outlines/lp-municipal.yaml'
    outline.write_text(outline.read_text(encoding='utf-8').replace('msg-0001', 'msg-0099', 1), encoding='utf-8')
    result = evaluate.run(repo, DATE, ['lp-municipal'])
    assert result['exit_code'] == 2 and result['stages'] == {'check': 2, 'dom': 'skipped', 'judge': 'skipped'}


def test_default_evaluates_every_outline(repo):
    result = evaluate.run(repo, DATE)
    assert set(result['stages']['dom']) == {'lp-municipal', 'deck-smb'}
    assert 'IA_DOM_UNCOVERED' in codes(result, 'warn')  # the deck is not HTML: reported, never silently passed


def test_unknown_outline_fails(repo):
    result = evaluate.run(repo, DATE, ['no-such'])
    assert result['exit_code'] == 2 and 'IA_EVALUATE_OUTLINE' in codes(result, 'fail')


def test_findings_are_not_repeated_per_stage(repo):
    # A check warning is also carried by every dom report: it must still appear once.
    import yaml
    path = repo / 'design/info_arch/info_arch.yaml'
    ssot = yaml.safe_load(path.read_text(encoding='utf-8'))
    del ssot['proofs']['prf-0002']
    path.write_text(yaml.safe_dump(ssot, allow_unicode=True, sort_keys=False), encoding='utf-8')
    result = evaluate.run(repo, DATE)
    assert codes(result).count('IA_PRIORITY_UNSUPPORTED') == 1, result['findings']
    keys = [(f['level'], f['code'], f['message'], f['path']) for f in result['findings']]
    assert len(keys) == len(set(keys))


def test_cli_prints_one_report_without_candidate_text(tmp_path):
    root = tmp_path / 'empty'
    shutil.copytree(SCRIPTS.parent / 'templates/design', root / 'design')
    proc = subprocess.run([sys.executable, str(SCRIPTS / 'infoarch_author.py'), 'evaluate', '--root', str(root), '--date', DATE],
                          capture_output=True, text=True, check=False)
    out = json.loads(proc.stdout)
    assert proc.returncode == 2 == out['exit_code'] and 'IA_NO_OUTLINES' in codes(out)
    assert 'candidates' not in out


# --- bounded Jev acquisition ---

def usage_ledger(tmp_path, attempted=0):
    path = tmp_path / 'usage/classification-usage.json'
    path.parent.mkdir()
    path.write_text(json.dumps({'schema_version': 1, 'attempted_judgments': attempted, 'batches': []}), encoding='utf-8')
    return path


def responder(calls, score=0.1):
    policy = jl.load_policy()

    def respond(payload, key, timeout):
        calls.append(len(payload['questions']))
        return {'model': policy['model'], 'usage': {'input_tokens': 10, 'output_tokens': 1},
                'answers': {q: {'type': 'noul', 'noul': score} for q in payload['questions']}}
    return respond


def git_init(root):
    subprocess.run(['git', 'init', '-q', str(root)], check=True)


def test_acquire_preview_calls_nothing_and_writes_nothing(repo, tmp_path, monkeypatch):
    import judgment_run
    monkeypatch.setattr(judgment_run.api, 'read_api_key', lambda: pytest.fail('auth read on preview'))
    answers = tmp_path / 'answers'
    answers.mkdir()
    result = evaluate.run(repo, DATE, ['lp-municipal'], acquire={
        'limit': 5, 'answers_dir': answers, 'execute': False, '_usage_path': usage_ledger(tmp_path),
        '_transport': lambda *a: pytest.fail('network on preview')})
    assert result['acquisition']['status'] == 'dry-run' and result['acquisition']['selected'] == 5
    assert not (answers / 'judgment-answers.json').exists() and not (repo / 'design/judgments.json').exists()


def test_acquire_execute_records_the_ledger_and_re_evaluates(repo, tmp_path, monkeypatch):
    git_init(repo)
    monkeypatch.setenv('TYPESAFE_API_KEY', 'test-key')
    answers, calls, usage = tmp_path / 'answers', [], usage_ledger(tmp_path)
    answers.mkdir()
    result = evaluate.run(repo, DATE, ['lp-municipal'], acquire={
        'limit': 500, 'answers_dir': answers, 'execute': True, '_usage_path': usage, '_transport': responder(calls)})
    assert calls and result['acquisition']['status'] == 'answered', result.get('acquisition')
    assert result['acquisition']['ledger'] == 'recorded'
    assert result['exit_code'] == 0 and not result['uncovered'], result['findings']
    assert json.loads(usage.read_text())['attempted_judgments'] == sum(calls)


def test_acquire_respects_the_limit_and_refreshes_an_existing_ledger(repo, tmp_path, monkeypatch):
    git_init(repo)
    monkeypatch.setenv('TYPESAFE_API_KEY', 'test-key')
    answers, calls, usage = tmp_path / 'answers', [], usage_ledger(tmp_path)
    answers.mkdir()
    acquire = {'limit': 3, 'answers_dir': answers, 'execute': True, '_usage_path': usage, '_transport': responder(calls)}
    first = evaluate.run(repo, DATE, ['lp-municipal'], acquire=acquire)
    assert sum(calls) == 3 and first['exit_code'] == 1 and first['acquisition']['ledger'] == 'recorded'
    second = evaluate.run(repo, DATE, ['lp-municipal'], acquire=dict(acquire, limit=500))
    assert second['acquisition']['ledger'] == 'refreshed' and second['exit_code'] == 0, second['findings']


def test_acquire_is_refused_when_the_evaluation_failed(repo, tmp_path, monkeypatch):
    monkeypatch.setattr(dom, 'snapshot', fake_snapshot(lambda s: s['anchors'][1].update(text='別の文')))
    answers = tmp_path / 'answers'
    answers.mkdir()
    result = evaluate.run(repo, DATE, ['lp-municipal'], acquire={
        'limit': 5, 'answers_dir': answers, 'execute': True, '_usage_path': usage_ledger(tmp_path),
        '_transport': lambda *a: pytest.fail('judged a failed render')})
    assert result['exit_code'] == 2 and result['acquisition']['status'] == 'refused'


def test_acquire_is_refused_when_the_judge_failed(repo, tmp_path):
    (repo / 'design/judgments.json').write_text('{"not": "a ledger"}', encoding='utf-8')
    answers = tmp_path / 'answers'
    answers.mkdir()
    result = evaluate.run(repo, DATE, ['lp-municipal'], acquire={
        'limit': 5, 'answers_dir': answers, 'execute': True, '_usage_path': usage_ledger(tmp_path),
        '_transport': lambda *a: pytest.fail('judged against a broken ledger')})
    assert result['stages']['judge'] == 2 and result['acquisition']['status'] == 'refused', result['findings']


def test_acquire_over_budget_fails_without_calling(repo, tmp_path):
    answers = tmp_path / 'answers'
    answers.mkdir()
    result = evaluate.run(repo, DATE, ['lp-municipal'], acquire={
        'limit': 5, 'answers_dir': answers, 'execute': True, '_usage_path': usage_ledger(tmp_path, attempted=1998),
        '_transport': lambda *a: pytest.fail('over budget call')})
    assert result['exit_code'] == 2 and 'IA_EVALUATE_ACQUIRE' in codes(result, 'fail')
    assert 'budget-exceeded' in result['acquisition']['error']


import os  # noqa: E402

from test_dom import DRIFT, HTML, PW_DIR  # noqa: E402


@pytest.mark.skipif(not PW_DIR, reason='INFOARCH_TEST_PLAYWRIGHT_DIR is not set: real-browser integration not run (reported as skipped, not passed)')
@pytest.mark.parametrize('page,expect,dom_code', [(HTML, 1, 0), (DRIFT, 2, 2)])
def test_real_browser_evaluate_end_to_end(tmp_path, page, expect, dom_code):
    root = tmp_path / 'repo'
    write_repo(root)
    (root / 'site').mkdir()
    (root / 'site/index.html').write_text(page.format(hidden=''), encoding='utf-8')
    env = dict(os.environ, INFOARCH_PLAYWRIGHT_DIR=PW_DIR)
    proc = subprocess.run([sys.executable, str(SCRIPTS / 'infoarch_author.py'), 'evaluate', '--root', str(root), '--date', DATE,
                           '--outline', 'lp-municipal'], capture_output=True, text=True, env=env, check=False)
    out = json.loads(proc.stdout)
    assert proc.returncode == expect == out['exit_code'] and out['stages']['dom'] == {'lp-municipal': dom_code}, out['findings']
    if expect == 1:  # rendered as designed; only the judgments are not recorded yet
        assert out['stages']['judge'] == 1 and 'IA_JUDGE_UNCOVERED' in codes(out, 'warn')
