"""A design decision passes only through a check the repository declares for it, run on every rendered page.

Writing the decision in DESIGN.md and approving it is not verification: an approved decision with no declared
check is uncovered and stops ship.
"""
import json

import pytest

from test_core import project, write, check, approve_rows  # noqa: F401  (fixture reuse)
from item_checks import load_map  # noqa: E402  (test_core puts scripts/ on the path)
from test_item_coverage import catalog_item, use_catalog_item, set_items, doc_section, rendered, stages  # noqa: F401

DECISION = ('components/accordion', 'ユースケース')


def declare(root, *rules):
    write(root / 'design/checks/web.yaml', {'rules': list(rules)})


def cite(root, *refs):
    set_items(root, {'dads-000001': {'state': 'defined', 'evidence': list(refs)}})


def approve_all(project):
    manifest = check(project)['manifest']['web']
    approve_rows(project, [{'key': 'ds:web:item:dads-000001', 'value_hash': manifest['items']['dads-000001']['value_hash']},
                           {'key': 'ds:web:token:color.primary', 'value_hash': manifest['tokens']['color.primary']['value_hash']}])


def routes(project):
    return check(project)['manifest']['web']['coverage']


COUNT = {'id': 'one-open', 'kind': 'count', 'scope': ['details'], 'targets': ['summary'], 'maximum': 1}


def test_approved_decision_without_a_declared_check_is_uncovered_and_stops_ship(project):
    use_catalog_item(project, catalog_item(*DECISION))
    approve_all(project)
    coverage = routes(project)
    assert coverage['item_routes']['dads-000001']['route'] == 'uncovered'
    assert 'declared check' in coverage['item_routes']['dads-000001']['reason']
    assert coverage['decisions'] == ['dads-000001']
    shipped = check(project, ship=True)
    assert shipped['exit_code'] == 2 and any('without a verification route 1' in e for e in shipped['errors'])


def test_decision_with_a_declared_page_rule_is_checked_by_code_and_still_needs_approval(project):
    root, _ = project
    use_catalog_item(project, catalog_item(*DECISION))
    declare(root, COUNT)
    cite(root, 'rule:one-open')
    manifest = check(project)['manifest']['web']
    assert manifest['coverage']['item_routes']['dads-000001'] == {
        'route': 'machine', 'checkers': ['rule:one-open'], 'document': 'components/accordion', 'decision': True}
    assert 'rule:one-open' in manifest['item_checks']['collect']
    assert not any(r['id'] == 'one-open' for r in check(project)['report']['uncovered'])  # measured on pages, not L0
    assert any('unapproved design decisions 1' in e for e in check(project, ship=True)['errors'])
    approve_all(project)
    shipped = check(project, ship=True)
    assert shipped['exit_code'] == 0, shipped['errors']


def test_changing_the_declared_rule_invalidates_the_approval(project):
    root, _ = project
    use_catalog_item(project, catalog_item(*DECISION))
    declare(root, COUNT)
    cite(root, 'rule:one-open')
    approve_all(project)
    declare(root, {**COUNT, 'maximum': 5})
    result = check(project, ship=True)
    assert result['exit_code'] == 2


def test_decision_declared_as_a_judgment_goes_to_afm_with_the_repositorys_statement(project):
    root, _ = project
    use_catalog_item(project, catalog_item(*DECISION))
    declare(root, {'id': 'faq-only', 'kind': 'judgment', 'targets': ['details'],
                   'statement': 'Accordions hold only FAQ answers, never primary content.'})
    cite(root, 'rule:faq-only')
    result = check(project)
    row = result['manifest']['web']['coverage']['item_routes']['dads-000001']
    assert row['route'] == 'afm' and row['presence'] == ['rule:faq-only'] and row['decision'] is True
    assert 'never primary content' in result['afm_candidates'][0]['rule']
    assert 'rule:faq-only' in result['manifest']['web']['item_checks']['collect']


@pytest.mark.parametrize('path,expected', [(DECISION, 'uncovered'), (('components/card', '外観'), 'afm')])
def test_a_cited_token_does_not_verify_a_decision_or_a_judgment(project, path, expected):
    root, _ = project
    use_catalog_item(project, catalog_item(*path))
    cite(root, 'token:color.primary')
    assert routes(project)['item_routes']['dads-000001']['route'] == expected


@pytest.mark.parametrize('rule', [
    {'id': 'r', 'kind': 'count', 'targets': ['summary']},                                     # no bound
    {'id': 'r', 'kind': 'count', 'targets': ['summary'], 'maximum': -1},
    {'id': 'r', 'kind': 'text', 'targets': ['main'], 'pattern': '('},                         # bad regex
    {'id': 'r', 'kind': 'text', 'targets': ['main']},
    {'id': 'r', 'kind': 'style', 'targets': ['a'], 'property': 'color'},                      # no allowlist
    {'id': 'r', 'kind': 'style', 'targets': ['a'], 'allowlist': ['red']},                     # no property
    {'id': 'r', 'kind': 'style', 'targets': ['a'], 'property': 'color', 'allowlist': ['red'], 'state': 'hover'},
    {'id': 'r', 'kind': 'judgment', 'targets': ['details']},                                  # no statement
    {'id': 'r', 'kind': 'count', 'targets': ['summary'], 'maximum': 1, 'statement': 'x'},     # field of another kind
])
def test_malformed_page_rules_are_rejected(project, rule):
    root, _ = project
    declare(root, rule)
    result = check(project)
    assert result['exit_code'] == 2 and any('rule' in e for e in result['errors']), result['errors']


def test_style_rule_token_reference_must_exist(project):
    root, _ = project
    use_catalog_item(project, catalog_item(*DECISION))
    declare(root, {'id': 'focus', 'kind': 'style', 'targets': ['summary'], 'property': 'outline-color',
                   'state': 'focus', 'allowlist': ['token:color.absent']})
    cite(root, 'rule:focus')
    assert any('style rule token' in e for e in check(project)['errors'])


def test_example_rules_shipped_with_the_skill_are_valid(project):
    from test_core import SKILL
    root, _ = project
    examples = json.loads((SKILL / 'references/decision-rule-examples.json').read_text())
    assert {r['kind'] for r in examples['rules']} == {'count', 'text', 'style', 'judgment'}
    mapping = load_map()
    assert set(examples['for']) == {r['id'] for r in examples['rules']}
    for target in examples['for'].values():  # every example is written for a real catalog decision
        document, _, path = target.partition(' > ')
        assert mapping['_index'][(document, path)]['route'] == 'decision', target
    declare(root, *examples['rules'])
    errors = check(project)['errors']
    assert not any('rule' in e for e in errors), errors


# --- on rendered pages ---------------------------------------------------------------------------

def run_rule(rendered, project, rule, body):
    run, root, _ = rendered
    use_catalog_item(project, catalog_item(*DECISION, grade='nice-to-have'))
    declare(root, rule)
    cite(root, 'rule:' + rule['id'])
    code, result = run(body)
    items = json.loads((root / 'design/.resolved/items.json').read_text()) if (root / 'design/.resolved/items.json').exists() else {}
    return code, result, items.get('machine', {}).get('dads-000001', {}).get('DESIGN_FIXTURE_URL')


DETAILS = '<details><summary style="display:block;padding:16px">Q{}</summary>A</details>'


def test_count_rule_passes_and_fails_on_the_page(rendered, project):
    rule = {'id': 'few', 'kind': 'count', 'targets': ['details'], 'maximum': 2}
    code, result, status = run_rule(rendered, project, rule, DETAILS.format(1) + DETAILS.format(2))
    assert status == 'pass', result
    code, result, status = run_rule(rendered, project, rule, ''.join(DETAILS.format(i) for i in range(3)))
    assert code == 2 and status == 'fail'
    assert any(e.startswith('DESIGN_FIXTURE_URL: rule:few') and e.endswith('breaks adopted items dads-000001')
               for e in result['errors']), result['errors']


def test_text_rule_finds_a_forbidden_phrase(rendered, project):
    rule = {'id': 'polite', 'kind': 'text', 'targets': ['main'], 'pattern': '[Cc]lick here'}
    code, result, status = run_rule(rendered, project, rule, '<p>Apply now.</p>')
    assert status == 'pass', (result['errors'], result['stages'])
    code, result, status = run_rule(rendered, project, rule, '<p>Click here.</p>')
    assert code == 2 and status == 'fail', result


def test_style_rule_reads_the_focused_state_against_a_token(rendered, project):
    rule = {'id': 'focus-ring', 'kind': 'style', 'targets': ['a'], 'property': 'outline-color', 'state': 'focus',
            'allowlist': ['token:color.primary']}
    code, result, status = run_rule(rendered, project, rule, '<a href="#x" style="display:inline-block;padding:16px;color:black">link</a>')
    assert status == 'pass', (result['errors'], result['stages'])
    body = '<style>a:focus-visible,a:focus{outline-color:red}</style><a href="#x" style="display:inline-block;padding:16px;color:black">link</a>'
    code, result, status = run_rule(rendered, project, rule, body)
    assert status == 'fail', result


def test_pattern_the_browser_rejects_fails_the_item_instead_of_passing(rendered, project):
    rule = {'id': 'py-only', 'kind': 'text', 'targets': ['main'], 'pattern': '(?i)click'}  # Python-only inline flag
    code, result, status = run_rule(rendered, project, rule, '<p>Apply now.</p>')
    assert code == 2 and status == 'fail', result
    assert any('rule:py-only was not measured' in e for e in result['errors']), result['errors']


def test_rule_that_matches_nothing_is_not_applicable_and_warned(rendered, project):
    rule = {'id': 'typo', 'kind': 'count', 'scope': ['detials'], 'targets': ['summary'], 'maximum': 1}
    code, result, status = run_rule(rendered, project, rule, DETAILS.format(1))
    assert status == 'n/a', result
    assert any('rule typo never applied' in w for w in result['warnings']), result['warnings']


def test_judgment_rule_skips_afm_where_its_targets_are_absent(rendered, project):
    run, root, tmp_path = rendered
    rule = {'id': 'faq-only', 'kind': 'judgment', 'targets': ['details'], 'statement': 'Accordions hold FAQ answers only.'}
    code, result, _ = run_rule(rendered, project, rule, '<p>no accordion</p>')
    assert stages(result) == ['L0', 'L1', 'L2', 'ITEMS'], result['errors']
    code, result, _ = run_rule(rendered, project, rule, DETAILS.format(1))
    assert stages(result) == ['L0', 'L1', 'L2', 'ITEMS', 'AFM'], result['errors']
