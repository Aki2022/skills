"""ISSUE-04: IA semantic judgments are pinned in the shared judgment ledger; the gate never calls a model."""
import copy
import json
import subprocess
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import dom  # noqa: E402
import judge  # noqa: E402
from test_check import DATE, DECK, OUTLINE, SSOT, dump, write_repo  # noqa: E402

SCRIPTS = Path(__file__).resolve().parents[1]
DESIGN_SCRIPTS = SCRIPTS.parents[1] / 'own-design-author/scripts'
sys.path.insert(0, str(DESIGN_SCRIPTS))
import judgment_ledger as jl  # noqa: E402

IA_KINDS = ('ia-name-variant', 'ia-emphasis', 'ia-unregistered-claim', 'ia-audience-fit')
# OUTLINE has 4 sections and DECK 2; SSOT has one official name, so every section gets a name and a claim candidate.
SECTIONS = len(OUTLINE['pages'][0]['sections']) + len(DECK['pages'][0]['sections'])


def cli(*args):
    proc = subprocess.run([sys.executable, str(SCRIPTS / 'infoarch_author.py'), *args],
                          capture_output=True, text=True, check=False)
    return proc.returncode, json.loads(proc.stdout)


def dom_report(path, outline='lp-municipal', exit_code=0, screens=None):
    screens = screens if screens is not None else [
        {'outline': outline, 'file': 'site/index.html', 'viewport': 'mobile', 'text': '既存システムの運用、重くありませんか 運用の手間が半分になります'},
        {'outline': outline, 'file': 'site/index.html', 'viewport': 'desktop', 'text': '既存システムの運用、重くありませんか 運用の手間が半分になります 導入は100社以上'}]
    path.write_text(json.dumps({'level': 'dom', 'exit_code': exit_code, 'first_screens': screens}, ensure_ascii=False), encoding='utf-8')
    return path


def candidates(root, *extra):
    code, out = cli('judge', '--root', str(root), '--date', DATE, '--include-candidates', *extra)
    return code, out, out.get('judgment_candidates', [])


def write_ledger(root, rows, holdout=None):
    policy = jl.load_policy()
    ledger = {'schema_version': 2, 'policy': policy, 'rows': rows, 'holdout': holdout or {'source': 'unmeasured'}}
    (root / 'design').mkdir(exist_ok=True)
    (root / 'design/judgments.json').write_text(json.dumps(ledger, ensure_ascii=False), encoding='utf-8')
    return ledger


@pytest.fixture
def repo(tmp_path):
    write_repo(tmp_path)
    return tmp_path


def test_shared_policy_carries_the_four_ia_kinds_next_to_design():
    policy = jl.load_policy()
    assert set(IA_KINDS) <= set(policy['prompts'])
    assert {'doc-relevance', 'na-rationale'} <= set(policy['prompts'])
    for kind in IA_KINDS:
        assert jl.candidate(kind, 'x', 'text', {}, policy=policy)['id'] == f'{kind}:x'


def test_unknown_kind_is_still_rejected_by_the_shared_policy():
    policy = jl.load_policy()
    bad = copy.deepcopy(policy)
    bad['prompts']['ia-made-up'] = 'x'
    with pytest.raises(ValueError):
        jl.validate_policy(bad)


def test_no_ledger_reports_every_candidate_uncovered_and_emphasis_missing(repo):
    code, out = cli('judge', '--root', str(repo), '--date', DATE)
    assert code == 1
    kinds = [u['id'].split(':')[0] for u in out['uncovered']]
    assert kinds.count('ia-name-variant') == SECTIONS
    assert kinds.count('ia-unregistered-claim') == SECTIONS
    assert kinds.count('ia-audience-fit') == 2
    assert 'IA_JUDGE_UNCOVERED' in [f['code'] for f in out['findings']]
    assert 'IA_JUDGE_EMPHASIS_MISSING' in [f['code'] for f in out['findings']]
    assert any(m['kind'] == 'ia-emphasis' for m in out['missing'])
    assert out['screened'] == []
    # Text is withheld unless the candidates are explicitly requested.
    assert 'judgment_candidates' not in out and '運用の手間' not in json.dumps(out, ensure_ascii=False)


def test_candidates_are_reproducible_by_the_shared_reader_and_expand_facts(repo):
    code, out, rows = candidates(repo)
    assert code == 1 and rows
    policy = jl.load_policy()
    for row in rows:
        data = row['input']
        assert row == jl.candidate(data['kind'], data['catalog_id'], data['text'], data['context'], policy=policy)
    claim = next(r for r in rows if r['id'] == 'ia-unregistered-claim:lp-municipal/top/2')
    assert claim['input']['text'] == '導入は100社以上'
    assert claim['input']['context']['proofs'] == [{'id': 'prf-0001', 'text': '導入自治体の数'}]
    names = next(r for r in rows if r['id'] == 'ia-name-variant:lp-municipal/top/0')
    assert names['input']['context']['entities'] == [{'id': 'ent-0001', 'name': 'Webサーバー'}]


def test_emphasis_candidates_come_from_the_rendered_first_screen(repo, tmp_path):
    report = dom_report(tmp_path / 'dom.json')
    code, out, rows = candidates(repo, '--dom-report', str(report))
    emphasis = [r for r in rows if r['input']['kind'] == 'ia-emphasis']
    assert [r['id'] for r in emphasis] == ['ia-emphasis:lp-municipal/site/index.html@mobile',
                                          'ia-emphasis:lp-municipal/site/index.html@desktop']
    ctx = emphasis[0]['input']['context']
    assert ctx['top_message'] == {'id': 'msg-0001', 'text': '運用の手間が半分になる'}
    assert ctx['audience'] == '自治体の情報部門'
    assert not any(m['kind'] == 'ia-emphasis' and m.get('outline') == 'lp-municipal' for m in out['missing'])


def test_failed_render_is_not_judged(repo, tmp_path):
    code, out = cli('judge', '--root', str(repo), '--date', DATE, '--dom-report', str(dom_report(tmp_path / 'd.json', exit_code=2)))
    assert code == 2 and 'IA_JUDGE_DOM_FAILED' in [f['code'] for f in out['findings']]


def test_dom_report_for_unknown_outline_or_empty_screens_is_rejected(repo, tmp_path):
    code, out = cli('judge', '--root', str(repo), '--date', DATE, '--dom-report', str(dom_report(tmp_path / 'a.json', outline='nope')))
    assert code == 2 and 'IA_JUDGE_DOM_REPORT' in [f['code'] for f in out['findings']]
    code, out = cli('judge', '--root', str(repo), '--date', DATE, '--dom-report', str(dom_report(tmp_path / 'b.json', screens=[])))
    assert code == 2 and 'IA_JUDGE_DOM_REPORT' in [f['code'] for f in out['findings']]


def test_full_ledger_below_theta_is_screened_and_green_without_any_model_call(repo, tmp_path, monkeypatch):
    report = dom_report(tmp_path / 'dom.json')
    _, _, rows = candidates(repo, '--dom-report', str(report))
    theta = jl.load_policy()['theta']
    write_ledger(repo, [{'candidate': r, 'score': theta / 2} for r in rows])
    code, out = cli('judge', '--root', str(repo), '--date', DATE, '--dom-report', str(report))
    assert code == 0, out['findings']
    assert out['uncovered'] == [] and len(out['screened']) == len(rows)
    assert all(s['semantic_guarantee'] is False and s['evaluation'] == {'status': 'unmeasured'} for s in out['screened'])
    # Emphasis is judged only for outlines with a dom report; the deck has none and stays visibly missing.
    assert [m['outline'] for m in out['missing']] == ['deck-smb']


def test_at_or_above_theta_goes_to_human_review(repo, tmp_path):
    report = dom_report(tmp_path / 'dom.json')
    _, _, rows = candidates(repo, '--dom-report', str(report))
    theta = jl.load_policy()['theta']
    flagged = 'ia-name-variant:lp-municipal/top/1'
    write_ledger(repo, [{'candidate': r, 'score': theta if r['id'] == flagged else 0.0} for r in rows])
    code, out = cli('judge', '--root', str(repo), '--date', DATE, '--dom-report', str(report))
    assert code == 1
    review = [f for f in out['findings'] if f['code'] == 'IA_JUDGE_REVIEW']
    assert [f['detail']['id'] for f in review] == [flagged]


def test_changed_copy_invalidates_only_its_own_judgments(repo, tmp_path):
    report = dom_report(tmp_path / 'dom.json')
    _, _, rows = candidates(repo, '--dom-report', str(report))
    write_ledger(repo, [{'candidate': r, 'score': 0.0} for r in rows])
    changed = copy.deepcopy(OUTLINE)
    changed['pages'][0]['sections'][1]['copy'] = '運用の手間がかなり減ります'
    dump(repo / 'design/info_arch/outlines/lp-municipal.yaml', changed)
    code, out = cli('judge', '--root', str(repo), '--date', DATE, '--dom-report', str(report))
    assert code == 1
    # The section itself, plus the outline-level fit that reads every section.
    assert sorted(u['id'] for u in out['uncovered']) == ['ia-audience-fit:lp-municipal',
                                                          'ia-name-variant:lp-municipal/top/1',
                                                          'ia-unregistered-claim:lp-municipal/top/1']


def test_design_rows_in_the_shared_ledger_do_not_disturb_ia(repo):
    _, _, rows = candidates(repo)
    design = jl.candidate('doc-relevance', 'dads-000001', 'A clear policy.', {'title': 'Brand'})
    write_ledger(repo, [{'candidate': design, 'score': 0.9}] + [{'candidate': r, 'score': 0.0} for r in rows])
    code, out = cli('judge', '--root', str(repo), '--date', DATE)
    assert out['uncovered'] == [] and not any(s['id'].startswith('doc-relevance') for s in out['screened'])
    assert 'IA_JUDGE_REVIEW' not in [f['code'] for f in out['findings']]


def test_broken_or_symlinked_ledger_fails(repo, tmp_path):
    (repo / 'design/judgments.json').write_text('{"schema_version": 2}', encoding='utf-8')
    code, out = cli('judge', '--root', str(repo), '--date', DATE)
    assert code == 2 and 'IA_JUDGE_LEDGER' in [f['code'] for f in out['findings']]
    (repo / 'design/judgments.json').unlink()
    # A symlink to an otherwise valid, fully covering ledger: only the symlink guard can stop it.
    _, _, rows = candidates(repo)
    write_ledger(repo, [{'candidate': r, 'score': 0.0} for r in rows])
    target = tmp_path / 'elsewhere.json'
    (repo / 'design/judgments.json').rename(target)
    (repo / 'design/judgments.json').symlink_to(target)
    code, out = cli('judge', '--root', str(repo), '--date', DATE)
    assert code == 2 and 'IA_JUDGE_LEDGER' in [f['code'] for f in out['findings']]


def test_shipped_html_without_first_screen_is_a_gap_other_media_are_only_reported(repo):
    _, _, rows = candidates(repo)
    write_ledger(repo, [{'candidate': r, 'score': 0.0} for r in rows])
    code, out = cli('judge', '--root', str(repo), '--date', DATE)
    assert code == 1 and out['uncovered'] == []
    levels = {f['message'].split(':')[0]: f['level'] for f in out['findings'] if f['code'] == 'IA_JUDGE_EMPHASIS_MISSING'}
    assert levels == {'lp-municipal': 'warn', 'deck-smb': 'info'}


def test_failing_check_refuses_to_judge(repo):
    broken = copy.deepcopy(OUTLINE)
    broken['pages'][0]['sections'][1]['message'] = 'msg-0099'
    dump(repo / 'design/info_arch/outlines/lp-municipal.yaml', broken)
    code, out = cli('judge', '--root', str(repo), '--date', DATE)
    assert code == 2 and 'IA_JUDGE_REFUSED' in [f['code'] for f in out['findings']]


def test_zero_candidates_is_never_green():
    report = judge.evaluate([], None, missing=[])
    assert report.exit_code == 2 and 'IA_JUDGE_NO_TARGETS' in report.codes()


def test_replacement_keeps_design_rows_and_takes_only_current_ia_answers(repo, tmp_path):
    report = dom_report(tmp_path / 'dom.json')
    _, _, rows = candidates(repo, '--dom-report', str(report))
    policy = jl.load_policy()
    design = jl.candidate('na-rationale', 'dads-000002', 'Not a web product.', {'item': 'x'})
    stale = jl.candidate('ia-audience-fit', 'gone', 'old', {})
    holdout = {'source': 'unmeasured'}
    write_ledger(repo, [{'candidate': design, 'score': 0.2}, {'candidate': stale, 'score': 0.1}], holdout)
    answers = {'schema_version': 1, 'status': 'model-answers-unlabeled', 'policy': policy,
               'rows': [{'candidate': r, 'score': 0.25, 'raw_answer': {'type': 'noul', 'noul': 0.25}} for r in rows[:-1]]}
    (tmp_path / 'judgment-answers.json').write_text(json.dumps(answers, ensure_ascii=False), encoding='utf-8')
    out_path = tmp_path / 'replacement.json'
    code, out = cli('judge', '--root', str(repo), '--date', DATE, '--dom-report', str(report),
                    '--answers', str(tmp_path / 'judgment-answers.json'), '--replacement-out', str(out_path))
    # Judged against the ledger as it is now: the IA candidates are still uncovered there.
    assert code == 1, out
    assert out['replacement'] == {'design_rows': 1, 'ia_rows': len(rows) - 1, 'unanswered': 1, 'dropped_ia_rows': 1}
    replacement = json.loads(out_path.read_text(encoding='utf-8'))
    ids = [r['candidate']['id'] for r in replacement['rows']]
    assert design['id'] in ids and stale['id'] not in ids
    assert replacement['holdout'] == holdout
    jl.validate_ledger(replacement, policy)
    # The shared recorder accepts it as a refresh of the existing ledger.
    jl.refresh_ledger(json.loads((repo / 'design/judgments.json').read_text(encoding='utf-8')), replacement, policy)


def test_replacement_rejects_answers_under_another_policy(repo, tmp_path):
    policy = {**jl.load_policy(), 'theta': 0.99}
    answers = {'schema_version': 1, 'status': 'model-answers-unlabeled', 'policy': policy, 'rows': []}
    (tmp_path / 'judgment-answers.json').write_text(json.dumps(answers), encoding='utf-8')
    code, out = cli('judge', '--root', str(repo), '--date', DATE,
                    '--answers', str(tmp_path / 'judgment-answers.json'), '--replacement-out', str(tmp_path / 'r.json'))
    assert code == 2 and not (tmp_path / 'r.json').exists()


def test_dom_first_screens_name_their_outline(repo, monkeypatch):
    (repo / 'site').mkdir()
    (repo / 'site/index.html').write_text('<p>x</p>', encoding='utf-8')
    snap = {'viewport': {'name': 'mobile'}, 'text_chars': 1, 'anchors': [], 'first_screen': 'x'}
    monkeypatch.setattr(dom, 'snapshot', lambda root, html, vps: [dict(snap, viewport=v) for v in vps])
    report = dom.run(repo, 'lp-municipal', DATE)
    assert report.first_screens and all(s['outline'] == 'lp-municipal' for s in report.first_screens)
