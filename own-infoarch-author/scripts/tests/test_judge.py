"""ISSUE-04: the IA semantic judgment (copy x value) is pinned in the shared judgment ledger; the gate never calls a model.

WS-20261123 ISSUE-01 cut the four older kinds; test_judge_one_kind.py pins that."""
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
import layer_paths  # noqa: E402  (own-design-author's layer rule: where each layer lives)

# The design policy still lists the older kinds so that rows already in a repository's ledger stay valid.
OLD_KINDS = ('ia-name-variant', 'ia-emphasis', 'ia-unregistered-claim', 'ia-audience-fit')
# Of OUTLINE's 5 pieces of copy only the 2 in its pain section are judged (trust and act copy is not asked); DECK has
# 1 in its aware section.
JUDGED = ['ia-value-fit:deck-smb/sec-0001/cpy-0001', 'ia-value-fit:lp-municipal/sec-0001/cpy-0001',
          'ia-value-fit:lp-municipal/sec-0001/cpy-0002']
FLAGGED = 'ia-value-fit:lp-municipal/sec-0001/cpy-0002'


def cli(*args):
    proc = subprocess.run([sys.executable, str(SCRIPTS / 'infoarch_author.py'), *args],
                          capture_output=True, text=True, check=False)
    return proc.returncode, json.loads(proc.stdout)


def candidates(root, *extra):
    code, out = cli('judge', '--root', str(root), '--date', DATE, '--include-candidates', *extra)
    return code, out, out.get('judgment_candidates', [])


def write_ledger(root, rows, holdout=None):
    policy = jl.load_policy()
    ledger = {'schema_version': 3, 'policy': policy, 'rows': rows, 'holdout': holdout or {'source': 'unmeasured'}}
    path = layer_paths.path(root, 'judgments')
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(ledger, ensure_ascii=False), encoding='utf-8')
    return ledger


@pytest.fixture
def repo(tmp_path):
    write_repo(tmp_path)
    return tmp_path


def test_shared_policy_carries_value_fit_next_to_design_and_still_lists_the_old_kinds():
    policy = jl.load_policy()
    assert {'ia-value-fit', 'doc-relevance', 'na-rationale'} <= set(policy['prompts'])
    for kind in ('ia-value-fit', *OLD_KINDS):  # old rows stay valid in a ledger; the IA skill just never makes them
        assert jl.candidate(kind, 'x', 'text', {}, policy=policy)['id'] == f'{kind}:x'


def test_unknown_kind_is_still_rejected_by_the_shared_policy():
    policy = jl.load_policy()
    bad = copy.deepcopy(policy)
    bad['prompts']['ia-made-up'] = 'x'
    with pytest.raises(ValueError):
        jl.validate_policy(bad)


def test_no_ledger_reports_every_candidate_uncovered(repo):
    code, out = cli('judge', '--root', str(repo), '--date', DATE)
    assert code == 1
    assert sorted(u['id'] for u in out['uncovered']) == JUDGED
    assert 'IA_JUDGE_UNCOVERED' in [f['code'] for f in out['findings']]
    assert out['screened'] == [] and out['missing'] == []
    # Text is withheld unless the candidates are explicitly requested.
    assert 'judgment_candidates' not in out and '運用の手間' not in json.dumps(out, ensure_ascii=False)


def test_candidates_are_reproducible_by_the_shared_reader_and_expand_facts(repo):
    code, out, rows = candidates(repo)
    assert code == 1 and rows
    policy = jl.load_policy()
    for row in rows:
        data = row['input']
        assert row == jl.candidate(data['kind'], data['catalog_id'], data['text'], data['context'], policy=policy)
    deck = next(r for r in rows if r['id'] == 'ia-value-fit:deck-smb/sec-0001/cpy-0001')
    assert deck['input']['text'] == '100社以上が導入、少人数でも回ります'  # the dictionary's fixed value is expanded
    assert deck['input']['context']['message'] == 'msg-0001'
    assert deck['input']['context']['value']['id'] == 'val-0002'


def test_full_ledger_below_theta_is_screened_and_green_without_any_model_call(repo):
    _, _, rows = candidates(repo)
    theta = jl.load_policy()['theta']
    write_ledger(repo, [{'candidate': r, 'score': theta / 2} for r in rows])
    code, out = cli('judge', '--root', str(repo), '--date', DATE)
    assert code == 0, out['findings']
    assert out['uncovered'] == [] and len(out['screened']) == len(rows) == len(JUDGED)
    assert all(s['semantic_guarantee'] is False and s['evaluation'] == {'status': 'unmeasured'} for s in out['screened'])
    assert out['missing'] == []


def test_at_or_above_theta_goes_to_human_review(repo):
    _, _, rows = candidates(repo)
    theta = jl.load_policy()['theta']
    write_ledger(repo, [{'candidate': r, 'score': theta if r['id'] == FLAGGED else 0.0} for r in rows])
    code, out = cli('judge', '--root', str(repo), '--date', DATE)
    assert code == 1
    review = [f for f in out['findings'] if f['code'] == 'IA_JUDGE_REVIEW']
    assert [f['detail']['id'] for f in review] == [FLAGGED]


def test_changed_copy_invalidates_only_its_own_judgments(repo):
    _, _, rows = candidates(repo)
    write_ledger(repo, [{'candidate': r, 'score': 0.0} for r in rows])
    changed = copy.deepcopy(OUTLINE)
    changed['copies']['cpy-0002']['text'] = '運用の手間がかなり減ります'
    dump(layer_paths.path(repo, 'infoarch', 'outlines', 'lp-municipal.yaml'), changed)
    code, out = cli('judge', '--root', str(repo), '--date', DATE)
    assert code == 1
    assert sorted(u['id'] for u in out['uncovered']) == [FLAGGED]  # only the copy that changed


def test_design_rows_in_the_shared_ledger_do_not_disturb_ia(repo):
    _, _, rows = candidates(repo)
    design = jl.candidate('doc-relevance', 'dads-000001', 'A clear policy.', {'title': 'Brand'})
    write_ledger(repo, [{'candidate': design, 'score': 0.9}] + [{'candidate': r, 'score': 0.0} for r in rows])
    code, out = cli('judge', '--root', str(repo), '--date', DATE)
    assert out['uncovered'] == [] and not any(s['id'].startswith('doc-relevance') for s in out['screened'])
    assert 'IA_JUDGE_REVIEW' not in [f['code'] for f in out['findings']]


def test_broken_or_symlinked_ledger_fails(repo, tmp_path):
    layer_paths.path(repo, 'judgments').write_text('{"schema_version": 2}', encoding='utf-8')
    code, out = cli('judge', '--root', str(repo), '--date', DATE)
    assert code == 2 and 'IA_JUDGE_LEDGER' in [f['code'] for f in out['findings']]
    layer_paths.path(repo, 'judgments').unlink()
    # A symlink to an otherwise valid, fully covering ledger: only the symlink guard can stop it.
    _, _, rows = candidates(repo)
    write_ledger(repo, [{'candidate': r, 'score': 0.0} for r in rows])
    target = tmp_path / 'elsewhere.json'
    layer_paths.path(repo, 'judgments').rename(target)
    layer_paths.path(repo, 'judgments').symlink_to(target)
    code, out = cli('judge', '--root', str(repo), '--date', DATE)
    assert code == 2 and 'IA_JUDGE_LEDGER' in [f['code'] for f in out['findings']]


def test_failing_check_refuses_to_judge(repo):
    broken = copy.deepcopy(OUTLINE)
    broken['sections']['sec-0002']['places'] = ['msg-0099']
    dump(layer_paths.path(repo, 'infoarch', 'outlines', 'lp-municipal.yaml'), broken)
    code, out = cli('judge', '--root', str(repo), '--date', DATE)
    assert code == 2 and 'IA_JUDGE_REFUSED' in [f['code'] for f in out['findings']]


def test_zero_candidates_is_never_green():
    report = judge.evaluate([], None)
    assert report.exit_code == 1 and 'IA_JUDGE_NO_TARGETS' in report.codes('warn')  # a visible gap, never a stop


def test_replacement_keeps_design_rows_and_takes_only_current_ia_answers(repo, tmp_path):
    _, _, rows = candidates(repo)
    policy = jl.load_policy()
    design = jl.candidate('na-rationale', 'dads-000002', 'Not a web product.', {'item': 'x'})
    stale = jl.candidate('ia-value-fit', 'gone/sec-0001/cpy-0001', 'old', {})
    holdout = {'source': 'unmeasured'}
    write_ledger(repo, [{'candidate': design, 'score': 0.2}, {'candidate': stale, 'score': 0.1}], holdout)
    answers = {'schema_version': 1, 'status': 'model-answers-unlabeled', 'policy': policy,
               'rows': [{'candidate': r, 'score': 0.25, 'raw_answer': {'type': 'noul', 'noul': 0.25}} for r in rows[:-1]]}
    (tmp_path / 'judgment-answers.json').write_text(json.dumps(answers, ensure_ascii=False), encoding='utf-8')
    out_path = tmp_path / 'replacement.json'
    code, out = cli('judge', '--root', str(repo), '--date', DATE,
                    '--answers', str(tmp_path / 'judgment-answers.json'), '--replacement-out', str(out_path))
    # Judged against the ledger as it is now: the IA candidates are still uncovered there.
    assert code == 1, out
    assert out['replacement'] == {'design_rows': 1, 'ia_rows': len(rows) - 1, 'kept_ia_rows': 0, 'unanswered': 1, 'dropped_ia_rows': 1}
    replacement = json.loads(out_path.read_text(encoding='utf-8'))
    ids = [r['candidate']['id'] for r in replacement['rows']]
    assert design['id'] in ids and stale['id'] not in ids
    assert replacement['holdout'] == holdout
    jl.validate_ledger(replacement, policy)
    # The shared recorder accepts it as a refresh of the existing ledger.
    jl.refresh_ledger(json.loads(layer_paths.path(repo, 'judgments').read_text(encoding='utf-8')), replacement, policy)


def test_replacement_rejects_answers_under_another_policy(repo, tmp_path):
    policy = {**jl.load_policy(), 'theta': 0.99}
    answers = {'schema_version': 1, 'status': 'model-answers-unlabeled', 'policy': policy, 'rows': []}
    (tmp_path / 'judgment-answers.json').write_text(json.dumps(answers), encoding='utf-8')
    code, out = cli('judge', '--root', str(repo), '--date', DATE,
                    '--answers', str(tmp_path / 'judgment-answers.json'), '--replacement-out', str(tmp_path / 'r.json'))
    assert code == 2 and not (tmp_path / 'r.json').exists()


# --- external review 2026-11-11 (WS-20261022 ISSUE-04) ---

def _answers_file(tmp_path, rows, patch=None):
    policy = jl.load_policy()
    data = {'schema_version': 1, 'status': 'model-answers-unlabeled', 'policy': policy,
            'rows': [{'candidate': r, 'score': 0.25, 'raw_answer': {'type': 'noul', 'noul': 0.25}} for r in rows]}
    for i, row in (patch or {}).items():
        data['rows'][i].update(row)
    path = tmp_path / 'judgment-answers.json'
    path.write_text(json.dumps(data, ensure_ascii=False), encoding='utf-8')
    return path


def test_replacement_rejects_a_score_that_differs_from_the_raw_answer(repo, tmp_path):
    _, _, rows = candidates(repo)
    answers = _answers_file(tmp_path, rows, patch={0: {'score': 0.9}})
    code, out = cli('judge', '--root', str(repo), '--date', DATE, '--answers', str(answers), '--replacement-out', str(tmp_path / 'r.json'))
    assert code == 2 and not (tmp_path / 'r.json').exists()


def test_replacement_out_never_follows_a_symlink(repo, tmp_path):
    _, _, rows = candidates(repo)
    answers = _answers_file(tmp_path, rows)
    victim = tmp_path / 'victim.json'
    link = tmp_path / 'r.json'
    link.symlink_to(victim)
    code, _ = cli('judge', '--root', str(repo), '--date', DATE, '--answers', str(answers), '--replacement-out', str(link))
    assert code == 2 and not victim.exists()
    linked_dir = tmp_path / 'linked'
    linked_dir.symlink_to(tmp_path / 'real', target_is_directory=True)
    (tmp_path / 'real').mkdir()
    code, _ = cli('judge', '--root', str(repo), '--date', DATE, '--answers', str(answers), '--replacement-out', str(linked_dir / 'r.json'))
    assert code == 2 and not (tmp_path / 'real/r.json').exists()


def test_replacement_out_is_created_exclusively(repo, tmp_path, monkeypatch):
    _, _, rows = candidates(repo)
    answers = _answers_file(tmp_path, rows)
    target = tmp_path / 'r.json'
    target.write_text('someone else', encoding='utf-8')
    # Simulate the file appearing after any existence check: the write itself must refuse.
    monkeypatch.setattr(Path, 'exists', lambda self, **k: False)
    with pytest.raises((FileExistsError, ValueError)):
        judge.run(repo, DATE, answers, target)
    monkeypatch.undo()
    assert target.read_text(encoding='utf-8') == 'someone else'


def test_dom_first_screens_name_their_outline(repo, monkeypatch):
    """dom still reports the first screens (a rendered-page fact); no judgment reads them any more."""
    (repo / 'site').mkdir()
    (repo / 'site/index.html').write_text('<p>x</p>', encoding='utf-8')
    snap = {'viewport': {'name': 'mobile'}, 'text_chars': 1, 'anchors': [], 'first_screen': 'x'}
    monkeypatch.setattr(dom, 'snapshot', lambda root, html, vps: [dict(snap, viewport=v) for v in vps])
    report = dom.run(repo, 'lp-municipal', DATE)
    assert report.first_screens and all(s['outline'] == 'lp-municipal' for s in report.first_screens)
