"""A design's prohibitions (parts and ways of building it does not allow) are declared in a form code can read, and a
page that breaks one is disqualified like any other measured failure (SPEC-deliverable-production-layers requirement 27).

design.yaml `prohibitions` rows pair a statement with the page rule (ux/design/checks) that checks it. A prohibition with
no rule is uncovered, never a pass. The owner approves each row by hash, so editing a row drops its approval. The prose in
DESIGN.md is not touched, so the approvals that bind it stay valid. The browser cases need DESIGN_BROWSER_TEST=1."""
import json
import os
import subprocess
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))
from test_core import project, write, check, approve_rows  # noqa: E402,F401
import layer_paths as lp  # noqa: E402
from test_measurement_profiles import prepare, measure, NORMAL, BODY, browser  # noqa: E402,F401

NO_NAV = {'id': 'no-nav', 'kind': 'count', 'targets': ['nav', '[role=navigation]'], 'maximum': 0}
ROW = {'id': 'no-nav', 'statement': 'ナビゲーションの帯を置かない', 'rule': 'no-nav'}


def declare(root, rows, *rules):
    write(lp.path(root, 'checks', 'web.yaml'), {'rules': list(rules)})
    config = json.loads(lp.path(root, 'design_config').read_text())
    config['prohibitions'] = rows
    write(lp.path(root, 'design_config'), config)


def manifest(project):
    return check(project)['manifest']['web']


def test_a_prohibition_with_a_page_rule_is_collected_on_every_page_and_is_not_uncovered(project):
    root, _ = project
    declare(root, [ROW], NO_NAV)
    result = check(project)
    assert result['exit_code'] != 2, result['errors']
    web = result['manifest']['web']
    assert web['prohibitions'] == [ROW]
    assert 'rule:no-nav' in web['item_checks']['collect']
    assert 'prohibition:no-nav' not in web['coverage']['uncovered']


def test_a_design_without_prohibitions_has_none_in_its_manifest(project):
    assert manifest(project)['prohibitions'] == []


def test_a_prohibition_without_a_rule_is_uncovered_and_stops_ship(project):
    root, _ = project
    declare(root, [{'id': 'no-quote', 'statement': '引用の枠を置かない'}])
    web = manifest(project)
    assert web['coverage']['uncovered'].count('prohibition:no-quote') == 1
    assert web['prohibitions'] == [{'id': 'no-quote', 'statement': '引用の枠を置かない', 'rule': None}]
    assert any('prohibitions without a declared check 1' in e for e in check(project, ship=True)['errors'])


@pytest.mark.parametrize('rows, rules, word', [
    ([{**ROW, 'rule': 'ghost'}], [NO_NAV], 'unknown rule'),
    ([ROW], [{'id': 'no-nav', 'kind': 'contrast', 'targets': ['nav']}], 'page rule'),
    ([ROW, ROW], [NO_NAV], 'duplicate'),
    ([{'id': 'no-nav'}], [NO_NAV], 'statement'),
    ([{**ROW, 'extra': 1}], [NO_NAV], 'unknown field'),
    ([{**ROW, 'statement': ' '}], [NO_NAV], 'statement'),
    ('no-nav', [NO_NAV], 'list'),
])
def test_a_malformed_prohibition_stops_l0_and_names_the_fault(project, rows, rules, word):
    root, _ = project
    declare(root, rows, *rules)
    result = check(project)
    assert result['exit_code'] == 2
    assert any('prohibition' in e and word in e for e in result['errors']), result['errors']


def key(row):
    return 'ds:web:prohibition:' + row['id']


def test_each_prohibition_needs_the_owners_approval_by_hash_and_an_edit_drops_it(project):
    root, _ = project
    declare(root, [ROW], NO_NAV)
    assert any('unapproved design prohibitions 1' in e for e in check(project, ship=True)['errors'])
    assert any('prohibitions awaiting owner approval: 1' in w for w in check(project)['warnings'])
    digest = lambda: manifest(project)['prohibition_hashes']['no-nav']
    approve_rows(project, [{'key': key(ROW), 'value_hash': digest()}])
    assert not any('prohibitions' in e for e in check(project, ship=True)['errors'])
    # a loosened statement, and a loosened rule, are each a new value: the approval no longer matches
    approved = digest()
    declare(root, [{**ROW, 'statement': 'ナビゲーションは置かない（例外あり）'}], NO_NAV)
    assert digest() != approved
    assert any('unapproved design prohibitions 1' in e for e in check(project, ship=True)['errors'])
    declare(root, [ROW], {**NO_NAV, 'maximum': 3})
    assert digest() != approved


def test_an_approval_for_a_prohibition_that_is_gone_is_a_warning_not_a_crash(project):
    root, _ = project
    declare(root, [ROW], NO_NAV)
    approve_rows(project, [{'key': key(ROW), 'value_hash': manifest(project)['prohibition_hashes']['no-nav']}])
    declare(root, [], NO_NAV)
    result = check(project)
    assert result['exit_code'] != 2, result['errors']
    assert any('no longer' in w and 'prohibition' in w for w in result['warnings'])


# ---- real browser: the build loop's measurement fails on a broken prohibition and stays green without one ----

def with_prohibition(project):
    root = prepare(project)
    declare(root, [ROW], NO_NAV)
    assert check(project)['exit_code'] != 2
    return root


@browser
def test_a_page_that_uses_a_prohibited_part_is_red_and_a_page_without_it_is_green(project):
    root = with_prohibition(project)
    ok, report = measure(root, NORMAL.replace('CONTENT', BODY), 'lp')
    assert ok.returncode == 0, ok.stdout + ok.stderr
    assert report['failures'] == []
    bad, broken = measure(root, NORMAL.replace('CONTENT', '<nav><a href="#a">Menu</a></nav>' + BODY), 'lp')
    assert bad.returncode == 1, bad.stdout + bad.stderr
    rows = [r for r in broken['failures'] if r['id'] == 'design-prohibition']
    assert rows and rows[0]['prohibition'] == 'no-nav' and rows[0]['count'] == 1, broken['failures']


@browser
def test_a_prohibition_that_could_not_be_measured_is_not_a_pass(project):
    root = with_prohibition(project)
    manifest_path = lp.path(root, 'resolved', 'web.json')
    data = json.loads(manifest_path.read_text())
    data['item_checks']['collect'] = [c for c in data['item_checks']['collect'] if c != 'rule:no-nav']
    manifest_path.write_text(json.dumps(data))
    stopped, report = measure(root, NORMAL.replace('CONTENT', BODY), 'lp')
    assert stopped.returncode == 1, stopped.stdout + stopped.stderr
    assert any(r['id'] == 'design-prohibition' and 'not measured' in r['note'] for r in report['failures']), report['failures']


@browser
def test_an_uncovered_prohibition_is_reported_as_uncovered_on_a_green_page(project):
    root = prepare(project)
    declare(root, [{'id': 'no-quote', 'statement': '引用の枠を置かない'}])
    assert check(project)['exit_code'] != 2
    ok, report = measure(root, NORMAL.replace('CONTENT', BODY), 'lp')
    assert 'prohibition:no-quote' in report['uncovered']
