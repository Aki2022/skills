"""The layout plan check (SPEC-deliverable-production-layers requirements 14, 15, 18, 21, 22).

plan_check.py --repo <repo> --deliverable <id>: exit 0 pass, 1 warnings, 2 failing or nothing checked. The plan covers
every block of its script exactly once, names only patterns of the resolved vocabulary (the L0 manifest), says the
narrow-screen shape or takes the pattern's default (reported), and keeps the comp images out of git."""
import copy
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))
from plan_testkit import PLAN, SCRIPT, Repo, codes, plan, run  # noqa: E402


@pytest.fixture
def repo(tmp_path):
    r = Repo(tmp_path)
    r.l0()
    return r


def check(repo, deliverable='lp-sample'):
    rc, out, raw = run('plan_check.py', '--repo', repo.root, '--deliverable', deliverable)
    assert out is not None, raw.stdout + raw.stderr
    assert out['exit_code'] == rc
    return rc, out


def test_a_whole_plan_passes_and_reports_the_narrow_defaults_it_takes(repo):
    rc, out = check(repo)
    assert rc == 0, out['findings']
    assert set(out['report']) == {'executed', 'skipped', 'missing', 'uncovered', 'screened'}
    assert out['blocks'] == 5
    # blk-0003 says its narrow shape; the others take the pattern's default and the report says which
    assert out['narrow']['blk-0003'] == {'pattern': 'stack-1', 'source': 'plan'}
    assert out['narrow']['blk-0002'] == {'pattern': 'stack-1', 'source': 'default'}
    assert out['narrow']['blk-0001'] == {'pattern': 'header-bar', 'source': 'default'}
    assert 'PLAN_NARROW_DEFAULT' in codes(out, 'info')
    assert 'PLAN_COMP_NOT_REVIEWED' in codes(out, 'info')
    assert any(row['id'] == 'comp-ignore' for row in out['report']['skipped'])  # no comp dir yet


def test_no_plan_is_not_checked(repo):
    rc, out = check(repo, 'lp-other')
    assert rc == 2 and codes(out, 'fail') == ['PLAN_MISSING']


def test_an_unreadable_plan_fails(repo):
    repo.plan_text('schema_version: 1\nblocks:\n  blk-0001: {pattern: band}\n  blk-0001: {pattern: band}\n')
    rc, out = check(repo)
    assert rc == 2 and 'PLAN_UNREADABLE' in codes(out, 'fail')


@pytest.mark.parametrize('change', [{'extra': 1}, {'comp': {}}, {'blocks': {'blk-0001': {'pattern': 'band', 'width': 2}}}])
def test_a_plan_outside_the_schema_fails(repo, change):
    repo.plan(plan(**change))
    rc, out = check(repo)
    assert rc == 2 and 'PLAN_SCHEMA' in codes(out, 'fail')


def test_the_deliverable_must_be_the_directory_name(repo):
    repo.plan(plan(deliverable='lp-elsewhere'), 'lp-sample')
    rc, out = check(repo)
    assert rc == 2 and 'PLAN_DELIVERABLE_MISMATCH' in codes(out, 'fail')


def test_the_script_must_exist(repo):
    repo.plan(plan(script='no-such-script'))
    rc, out = check(repo)
    assert rc == 2 and 'PLAN_SCRIPT_MISSING' in codes(out, 'fail')


def test_a_script_without_blocks_checks_nothing(repo):
    empty = copy.deepcopy(SCRIPT)
    empty['blocks'], empty['copies'] = {}, {}
    repo.script(empty)
    rc, out = check(repo)
    assert rc == 2 and 'PLAN_NOTHING_CHECKED' in codes(out, 'fail')


def test_without_a_fresh_manifest_there_is_no_vocabulary(tmp_path):
    repo = Repo(tmp_path)  # L0 never ran
    rc, out = check(repo)
    assert rc == 2 and 'PLAN_MANIFEST_UNAVAILABLE' in codes(out, 'fail')


def test_a_stale_manifest_is_refused(repo):
    repo.design_config(spec='docs/specs/sample.md')  # a design input moved after L0
    rc, out = check(repo)
    assert rc == 2 and 'PLAN_MANIFEST_UNAVAILABLE' in codes(out, 'fail')


def test_a_missing_block_fails(repo):
    blocks = copy.deepcopy(PLAN['blocks'])
    del blocks['blk-0004']
    repo.plan(plan(blocks=blocks))
    rc, out = check(repo)
    assert rc == 2 and 'PLAN_BLOCK_MISSING' in codes(out, 'fail')
    assert 'blk-0004' in ' '.join(f['message'] for f in out['findings'])


def test_an_extra_block_fails(repo):
    blocks = copy.deepcopy(PLAN['blocks'])
    blocks['blk-0009'] = {'pattern': 'stack-1'}
    repo.plan(plan(blocks=blocks))
    rc, out = check(repo)
    assert rc == 2 and 'PLAN_BLOCK_UNKNOWN' in codes(out, 'fail')


def test_a_pattern_outside_the_vocabulary_fails(repo):
    blocks = copy.deepcopy(PLAN['blocks'])
    blocks['blk-0002']['pattern'] = 'hero-magic'
    repo.plan(plan(blocks=blocks))
    rc, out = check(repo)
    assert rc == 2 and 'PLAN_PATTERN_UNKNOWN' in codes(out, 'fail')


def test_a_pattern_the_repository_disabled_is_outside_its_vocabulary(repo):
    repo.design_config(layout_patterns={'disabled': ['stats-row']})
    repo.l0()
    rc, out = check(repo)
    assert rc == 2 and 'PLAN_PATTERN_UNKNOWN' in codes(out, 'fail')


def test_a_narrow_shape_outside_the_vocabulary_fails(repo):
    blocks = copy.deepcopy(PLAN['blocks'])
    blocks['blk-0002']['narrow'] = 'squeeze-2'
    repo.plan(plan(blocks=blocks))
    rc, out = check(repo)
    assert rc == 2 and 'PLAN_NARROW_UNKNOWN' in codes(out, 'fail')


def test_a_pattern_not_suited_to_the_block_function_warns(repo):
    blocks = copy.deepcopy(PLAN['blocks'])
    blocks['blk-0005']['pattern'] = 'quotes'  # source block in a quotes pattern
    repo.plan(plan(blocks=blocks))
    rc, out = check(repo)
    assert rc == 1 and codes(out, 'warn') == ['PLAN_PATTERN_UNSUITED']


def test_a_comp_dir_that_git_would_commit_warns(repo):
    repo.comp_dir().mkdir(parents=True)
    rc, out = check(repo)
    assert rc == 1 and codes(out, 'warn') == ['PLAN_COMP_NOT_IGNORED']
    repo.ignore_comps()
    rc, out = check(repo)
    assert rc == 0, out['findings']
    assert any(row['id'] == 'comp-ignore' for row in out['report']['executed'])


def test_outside_git_the_comp_ignore_is_reported_not_checked(tmp_path):
    repo = Repo(tmp_path, git=False)
    repo.l0()
    repo.comp_dir().mkdir(parents=True)
    rc, out = check(repo)
    assert rc == 1 and codes(out, 'warn') == ['PLAN_COMP_IGNORE_NOT_CHECKED']
    assert any(row['id'] == 'comp-ignore' for row in out['report']['uncovered'])


def test_a_reviewed_comp_is_not_reported(repo):
    repo.plan(plan(comp={'reviewed': True}))
    rc, out = check(repo)
    assert rc == 0 and 'PLAN_COMP_NOT_REVIEWED' not in codes(out)


# --- variants (requirement 18) ---------------------------------------------------------------------------------------

def variant(repo, overrides, change=None, base='lp-sample'):
    data = plan(deliverable='lp-sample-b', variant_of=base, overrides=overrides)
    if change:
        change(data)
    repo.plan(data)
    return data


def _move_claim(data):
    data['blocks']['blk-0002'] = {'pattern': 'cover'}


def test_a_declared_variant_passes(repo):
    variant(repo, ['blocks/blk-0002'], _move_claim)
    rc, out = check(repo, 'lp-sample-b')
    assert rc == 0, out['findings']
    assert out['variant']['undeclared'] == [] and out['variant']['changed'] == ['blocks/blk-0002']


def test_an_undeclared_difference_warns_as_confounded(repo):
    variant(repo, [], _move_claim)
    rc, out = check(repo, 'lp-sample-b')
    assert rc == 1 and codes(out, 'warn') == ['PLAN_UNDECLARED_DIFFERENCE']
    assert out['variant']['confounded'] is True and out['variant']['undeclared'] == ['blocks/blk-0002']


def test_a_declared_override_that_does_not_differ_warns(repo):
    variant(repo, ['blocks/blk-0003'])
    rc, out = check(repo, 'lp-sample-b')
    assert rc == 1 and codes(out, 'warn') == ['PLAN_OVERRIDE_UNCHANGED']


def test_overrides_without_a_base_fail(repo):
    data = plan(deliverable='lp-sample-b', overrides=['blocks/blk-0002'])
    repo.plan(data)
    rc, out = check(repo, 'lp-sample-b')
    assert rc == 2 and 'PLAN_OVERRIDES_WITHOUT_BASE' in codes(out, 'fail')


def test_a_missing_base_plan_fails(repo):
    variant(repo, ['blocks/blk-0002'], _move_claim, base='lp-nowhere')
    rc, out = check(repo, 'lp-sample-b')
    assert rc == 2 and 'PLAN_VARIANT_BASE_MISSING' in codes(out, 'fail')


def test_an_override_naming_no_block_fails(repo):
    variant(repo, ['blocks/blk-0009'])
    rc, out = check(repo, 'lp-sample-b')
    assert rc == 2 and 'PLAN_OVERRIDE_UNKNOWN' in codes(out, 'fail')


def test_compare_is_the_one_comparison_infoarch_reads():
    """The infoarch diff calls this function: the element keys and the changed set are the plan side of differs_in."""
    import plan_check
    a = copy.deepcopy(PLAN)
    b = plan(deliverable='lp-sample-b', variant_of='lp-sample', overrides=['blocks/blk-0002'])
    _move_claim(b)
    out = plan_check.compare(a, b, declared=b['overrides'])
    assert out['changed'] == ['blocks/blk-0002'] and out['undeclared'] == [] and out['findings'] == []
    assert out['changes'] == [{'layer': 'plan', 'element': 'blocks/blk-0002', 'change': 'changed',
                               'attributes': ['details', 'pattern']}]
    assert plan_check.compare(a, copy.deepcopy(a), declared=())['changed'] == []


def test_an_unexpected_error_is_a_failure_never_a_warning(repo):
    """An uncaught exception would exit 1, which reads as a warning: it must be 2."""
    import json as _json
    from plan_testkit import lp
    path = lp.path(repo.root, 'resolved', 'web.json')
    data = _json.loads(path.read_text())
    data['layout_patterns']['patterns']['band'] = 'not a pattern'
    path.write_text(_json.dumps(data))
    rc, out = check(repo)
    assert rc == 2 and 'PLAN_CHECK_FAILED' in codes(out, 'fail')
