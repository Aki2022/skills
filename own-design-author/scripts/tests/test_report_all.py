"""Report-all mode: render every page and state, and list every failure instead of stopping at the first.

The gate stops at the first failing page so a broken page never ships, but then a repository sees one page's
problems at a time and the skill itself cannot be checked on the rest without fixing the site first.
`design-check --report-all` keeps going: every page and state is rendered and judged, every failure is listed, and
the run still exits 2 when anything failed. It is never a ship result.
"""
import copy
import json
import subprocess

import pytest

from test_core import project, write  # noqa: F401
from test_item_coverage import catalog_item, use_catalog_item, set_items, rendered, stages  # noqa: F401
from test_interaction_states import ship_fixture, OPEN
from ship_contract import validate_web_result

DECISION = ('components/accordion', 'ユースケース')
RULE = {'id': 'few', 'kind': 'count', 'targets': ['details'], 'maximum': 1}
TWO = '<details><summary style="display:block;padding:16px">Q1</summary>A</details>' * 2


def setup(run_fixture, project, urls):
    run, root, _ = run_fixture
    use_catalog_item(project, catalog_item(*DECISION, grade='nice-to-have'))
    write(root / 'design/checks/web.yaml', {'rules': [RULE]})
    set_items(root, {'dads-000001': {'state': 'defined', 'evidence': ['rule:few']}})
    validation = json.loads((root / 'design/validation.json').read_text())
    validation['urls'] = urls
    write(root / 'design/validation.json', validation)
    return run, root


def items(root):
    return json.loads((root / 'design/.resolved/items.json').read_text())


TWO_PAGES = [{'env': 'DESIGN_FIXTURE_URL'}, {'env': 'DESIGN_FIXTURE_URL_2'}]


def test_the_gate_still_stops_at_the_first_failing_page(rendered, project):
    run, root = setup(rendered, project, TWO_PAGES)
    code, result = run(TWO)
    assert code == 2 and stages(result) == ['L0', 'L1', 'L2', 'ITEMS'], result['stages']
    assert result.get('mode') != 'report-all'


def test_report_all_renders_every_page_and_lists_every_failure(rendered, project):
    run, root = setup(rendered, project, TWO_PAGES)
    code, result = run(TWO, report_all=True)
    assert code == 2, result
    assert result['mode'] == 'report-all'
    assert stages(result) == ['L0', 'L1', 'L2', 'L2', 'ITEMS'], result['stages']
    errors = ' '.join(result['errors'])
    assert 'DESIGN_FIXTURE_URL: rule:few' in errors and 'DESIGN_FIXTURE_URL_2: rule:few' in errors, result['errors']
    assert items(root)['machine']['dads-000001'] == {'DESIGN_FIXTURE_URL': 'fail', 'DESIGN_FIXTURE_URL_2': 'fail'}
    assert items(root)['mode'] == 'report-all'


def test_report_all_keeps_going_past_a_page_that_cannot_render(rendered, project):
    broken = [{'env': 'DESIGN_FIXTURE_URL', 'states': {'open': [{'click': '#missing'}]}}, {'env': 'DESIGN_FIXTURE_URL_2'}]
    run, root = setup(rendered, project, broken)
    code, result = run(TWO, report_all=True)
    assert code == 2, result
    assert [s.get('target_env') for s in result['stages'] if s['stage'] == 'L2'] == [
        'DESIGN_FIXTURE_URL', 'DESIGN_FIXTURE_URL#open', 'DESIGN_FIXTURE_URL_2'], result['stages']
    assert any('could not reach state DESIGN_FIXTURE_URL#open' in e for e in result['errors']), result['errors']
    # The unreachable state is never judged, never recorded as passing; the pages around it are judged.
    machine = items(root)['machine']['dads-000001']
    assert 'DESIGN_FIXTURE_URL#open' not in machine
    assert machine == {'DESIGN_FIXTURE_URL': 'fail', 'DESIGN_FIXTURE_URL_2': 'fail'}
    assert 'DESIGN_FIXTURE_URL#open' in items(root)['unrendered']


def test_report_all_passes_clean_pages_but_exits_like_the_gate(rendered, project):
    run, root = setup(rendered, project, TWO_PAGES)
    code, result = run('<p>no accordion</p>', report_all=True)
    gate_code, gate = run('<p>no accordion</p>')
    assert code == gate_code and code != 2, (result['errors'], gate['errors'])


def test_a_report_all_result_is_never_ship_evidence(tmp_path):
    result = ship_fixture(tmp_path, {'open': OPEN}, [{'name': 'open', 'steps': OPEN}])
    urls = {'DESIGN_WEB_URL': 'http://127.0.0.1:12345/'}
    assert validate_web_result(tmp_path, result, urls)
    with pytest.raises(ValueError):
        validate_web_result(tmp_path, {**copy.deepcopy(result), 'mode': 'report-all'}, urls)
    record = json.loads((tmp_path / 'design/.resolved/items.json').read_text())
    write(tmp_path / 'design/.resolved/items.json', {**record, 'mode': 'report-all'})
    with pytest.raises(ValueError):
        validate_web_result(tmp_path, result, urls)


def test_report_all_judges_a_page_whose_l2_found_problems_and_names_them(rendered, project):
    # A colour outside the palette fails L2 (exit 1) but the page rendered: report-all still judges its items
    # and says which L2 checks failed, instead of hiding the page behind a generic L2 error.
    run, root = setup(rendered, project, TWO_PAGES)
    code, result = run(TWO + '<p style="color:rgb(255,0,0)">off palette</p>', report_all=True)
    assert code == 2, result
    assert [s['exit_code'] for s in result['stages'] if s['stage'] == 'L2'] == [1, 1], result['stages']
    assert any(e.startswith('L2 DESIGN_FIXTURE_URL found: ') and 'design-palette' in e for e in result['errors']), result['errors']
    assert items(root)['machine']['dads-000001'] == {'DESIGN_FIXTURE_URL': 'fail', 'DESIGN_FIXTURE_URL_2': 'fail'}
    assert items(root)['unrendered'] == []
