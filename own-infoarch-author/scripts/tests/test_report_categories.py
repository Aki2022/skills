"""The report's five categories (SPEC-information-architecture-standardization, Quality 2): executed, skipped, missing,
uncovered, screened. `check`, `blueprint` and the positioning layer's `check` always carry the five keys, empty or not;
judge, evaluate and ship keep the shape they already had.

Every test goes through the CLI (the real road), never through Report or a mock of it.
"""
import copy
import json
import subprocess
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from blueprint_fixtures import write_blueprint  # noqa: E402
from ia_fixtures import SSOT, dump  # noqa: E402
from test_check import DATE, write_repo  # noqa: E402
import layer_paths  # noqa: E402  (own-design-author's layer rule: where each layer lives)

SCRIPTS = Path(__file__).resolve().parents[1]
POSITIONING = SCRIPTS.parents[1] / 'own-positioning-author/scripts/positioning_author.py'
CATEGORIES = ('executed', 'skipped', 'missing', 'uncovered', 'screened')


def run(script, *args):
    proc = subprocess.run([sys.executable, str(script), *args], capture_output=True, text=True, check=False)
    assert proc.stdout, proc.stderr
    return proc.returncode, json.loads(proc.stdout)


def infoarch(command, root):
    return run(SCRIPTS / 'infoarch_author.py', command, '--root', str(root), '--date', DATE)


def positioning(root):
    return run(POSITIONING, 'check', '--root', str(root), '--date', DATE)


def ids(entries):
    return sorted(e['id'] for e in entries)


def assert_five(report):
    for key in CATEGORIES:
        assert isinstance(report[key], list), (key, report.keys())


# ---- the five keys are always there ---------------------------------------------------------------------------

@pytest.mark.parametrize('command', ['check', 'blueprint'])
def test_green_outline_repository_reports_all_five(tmp_path, command):
    write_repo(tmp_path)
    write_blueprint(tmp_path)  # one tree with both sides
    rc, report = infoarch(command, tmp_path)
    assert rc == 0, report['findings']
    assert_five(report)
    assert 'ia-positioning' in ids(report['executed'])
    assert report['missing'] == [] and report['uncovered'] == [] and report['screened'] == []
    # what the command does not run is named, with a reason, never left silent
    assert {'ia-dom', 'ia-semantic-judgments'} <= set(ids(report['skipped']))
    assert all(e['reason'] for e in report['skipped'])


def test_check_counts_what_it_executed_per_layer(tmp_path):
    write_repo(tmp_path)
    rc, report = infoarch('check', tmp_path)
    executed = {e['id']: e for e in report['executed']}
    assert executed['ia-outlines']['targets'] == {'channels': 2, 'outlines': 2, 'sections': report['targets']['sections']}
    assert 'ia-blueprint' in ids(report['skipped'])  # no blueprint in this tree: said, not silent
    assert 'ia-blueprint' not in executed


def test_product_ui_executes_the_blueprint_and_skips_outlines(tmp_path):
    write_blueprint(tmp_path)
    for command in ('check', 'blueprint'):
        rc, report = infoarch(command, tmp_path)
        executed = {e['id']: e for e in report['executed']}
        assert rc == 0 and executed['ia-blueprint']['targets'] == {'screens': 2, 'items': 10}, report
        assert 'ia-outlines' in ids(report['skipped']) and 'ia-outlines' not in executed


@pytest.mark.parametrize('command', ['check', 'blueprint'])
def test_empty_inputs_are_missing_never_an_empty_pass(tmp_path, command):
    """A positioning with nothing standing on it: failed, and the absent inputs are named in `missing`."""
    dump(layer_paths.path(tmp_path, 'positioning_file'), SSOT)
    rc, report = infoarch(command, tmp_path)
    assert rc == 2
    assert_five(report)
    absent = ids(report['missing'])
    if command == 'check':
        assert {'ia-channels', 'ia-outlines'} <= set(absent), report['missing']
    else:
        assert 'ia-blueprint-screens' in absent, report['missing']
    assert all(e['reason'] for e in report['missing'])


def test_no_positioning_file_is_missing_and_still_has_the_five_keys(tmp_path):
    rc, report = infoarch('check', tmp_path)
    assert rc == 2
    assert_five(report)
    assert 'ia-positioning' in ids(report['missing']) and report['executed'] == []


def test_positioning_check_reports_all_five(tmp_path):
    write_repo(tmp_path)
    rc, report = positioning(tmp_path)
    assert rc == 0, report['findings']
    assert_five(report)
    assert ids(report['executed']) == ['ia-positioning']
    assert report['executed'][0]['targets'] == report['targets']
    assert {'ia-outlines', 'ia-semantic-judgments'} <= set(ids(report['skipped']))  # the infoarch layer's, named
    assert report['missing'] == [] and report['screened'] == [] and report['uncovered'] == []


def test_positioning_check_with_no_file_is_missing_with_five_keys(tmp_path):
    rc, report = positioning(tmp_path)
    assert rc == 2
    assert_five(report)
    assert 'ia-positioning' in ids(report['missing'])


def test_a_failed_check_still_lists_what_it_executed(tmp_path):
    ssot = copy.deepcopy(SSOT)
    ssot['edges'] = [e for e in ssot['edges'] if f'{e["kind"]}:{e["from"]}:{e["to"]}' != 'fulfills:alt-0002:pan-0001']
    write_repo(tmp_path, ssot)
    rc, report = infoarch('check', tmp_path)
    assert rc == 2
    assert_five(report)
    assert 'ia-positioning' in ids(report['executed'])


# ---- the consumers keep working ---------------------------------------------------------------------------------

def test_design_l0_reads_the_check_with_the_added_keys(tmp_path):
    """ia_module validates the check response by its exact set of keys; the five categories must be part of that set."""
    import ia_module
    write_repo(tmp_path)
    write_blueprint(tmp_path)
    subprocess.run(['git', 'init', '-q', '-b', 'main', str(tmp_path)], check=True)
    out = ia_module.resolve(root=tmp_path, approvals={}, profiles={'web': {}})
    assert out['targets']['outlines'] == 2 and set(out['report']) == set(CATEGORIES)


# ---- judge, evaluate and ship keep the shape they had ------------------------------------------------------------

JUDGE_KEYS = {'tool', 'version', 'level', 'evaluated_on', 'exit_code', 'environment', 'counts', 'targets', 'uncovered',
              'findings', 'screened', 'missing'}


def test_judge_keeps_its_own_shape(tmp_path):
    write_repo(tmp_path)
    rc, report = infoarch('judge', tmp_path)
    assert set(report) == JUDGE_KEYS, set(report) ^ JUDGE_KEYS
    assert 'executed' not in report and 'skipped' not in report


def test_evaluate_keeps_its_own_shape(tmp_path):
    write_repo(tmp_path)
    rc, report = infoarch('evaluate', tmp_path)
    assert set(report) == JUDGE_KEYS | {'stages'}, set(report) ^ (JUDGE_KEYS | {'stages'})
    assert 'executed' not in report and 'skipped' not in report


def test_ship_stages_keep_their_own_shape():
    """Every stage of `ship` is built by `_stage`, which names its keys: the categories are not among them."""
    import ia_ship
    from ia_common import Report
    stage = ia_ship._stage('check', Report('check', DATE))
    assert set(stage) == {'stage', 'exit_code', 'counts', 'targets', 'findings'}
