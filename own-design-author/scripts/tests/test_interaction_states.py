"""Interaction states: a URL may declare named states reached by declarative steps; each is rendered like a page.

Components that appear only after an operation (dialogs, menus, error text after submit) exist in no static
render, so a rule on them never applies. A declared state renders the page after its steps and reports it as
`<env>#<state>`. A step that cannot run (no unique match) fails the render; it never passes silently.
"""
import copy
import hashlib
import json

import pytest

from test_core import project, write  # noqa: F401  (fixture reuse)
from test_item_coverage import catalog_item, use_catalog_item, set_items, rendered, stages  # noqa: F401
from web_check import render_targets
from ship_contract import validate_web_result

DECISION = ('components/accordion', 'ユースケース')
OPEN = [{'click': '#open'}]


def config(**states):
    row = {'env': 'DESIGN_WEB_URL'}
    if states:
        row['states'] = states
    return {'schema_version': 1, 'html': ['index.html'], 'urls': [row], 'viewports': [375], 'schemes': ['light']}


# --- declaration ---------------------------------------------------------------------------------

def test_a_url_without_states_renders_only_the_page():
    assert render_targets(config()) == [{'label': 'DESIGN_WEB_URL', 'env': 'DESIGN_WEB_URL', 'profile': 'lp', 'state': None}]


def test_each_declared_state_renders_after_its_page_as_env_hash_state():
    steps = [{'fill': '#email', 'value': 'not-an-email'}, {'submit': '#email'}, {'focus': '#email'}]
    targets = render_targets(config(**{'dialog-open': OPEN, 'error': steps}))
    assert [t['label'] for t in targets] == ['DESIGN_WEB_URL', 'DESIGN_WEB_URL#dialog-open', 'DESIGN_WEB_URL#error']
    assert targets[1]['state'] == {'name': 'dialog-open', 'steps': OPEN}
    assert targets[2]['state'] == {'name': 'error', 'steps': steps} and targets[2]['env'] == 'DESIGN_WEB_URL'


@pytest.mark.parametrize('states', [
    {},                                                    # declared but empty
    {'open': []},                                          # a state with no step is the page itself
    {'Open Dialog': OPEN},                                 # name outside [a-z0-9-]
    {'open': [{'hover': '#open'}]},                        # unknown action
    {'open': [{'click': ''}]},                             # empty selector
    {'open': [{'click': '#a', 'focus': '#b'}]},            # two actions in one step
    {'open': [{'fill': '#email'}]},                        # fill without a value
    {'open': [{'click': '#open', 'value': 'x'}]},          # value on a non-fill step
    {'open': [{'eval': 'document.body.innerHTML=""'}]},    # no arbitrary code
    {'open': ['#open']},                                   # step must be a mapping
    {'open': OPEN * 21},                                   # bounded
])
def test_malformed_states_are_rejected(states):
    with pytest.raises(ValueError):
        render_targets({**config(), 'urls': [{'env': 'DESIGN_WEB_URL', 'states': states}]})


# --- ship evidence -------------------------------------------------------------------------------

def ship_fixture(tmp_path, states, recorded):
    write(tmp_path / 'design/validation.json', {**config(), 'html': ['dist/index.html'],
          'urls': [{'env': 'DESIGN_WEB_URL', 'states': states}]})
    env = {'runtime': {'browser': 'fixture'}, 'viewports': [{'width': 375, 'height': 812}], 'schemes': ['light']}
    write(tmp_path / 'design/render-environment.json', env)
    digest = hashlib.sha256((tmp_path / 'design/render-environment.json').read_bytes()).hexdigest()
    row = {'scheme': 'light', 'width': 375, 'render_evidence': {'screenshot_sha256': 'a' * 64}}
    labels = ['DESIGN_WEB_URL'] + ['DESIGN_WEB_URL#' + name for name in states]
    for i, state in enumerate([None] + recorded):
        write(tmp_path / f'design/.resolved/l2-{i}.json', {'url': 'http://127.0.0.1:12345/', 'state': state,
              'environment': {'runtime': env['runtime'], 'config_sha256': digest, 'repeat_equal': True},
              'runs': [{**row, 'repeat': 0}, {**row, 'repeat': 1}], 'failures': [], 'uncovered': []})
    write(tmp_path / 'design/.resolved/items.json', {'schema_version': 1, 'pages': labels,
          'machine': {}, 'afm_pairs': [], 'failures': []})
    return {'exit_code': 0, 'errors': [], 'stages': [{'stage': 'L0', 'exit_code': 0, 'counts': {'items': 1}},
            {'stage': 'L1', 'exit_code': 0, 'targets': 1},
            *({'stage': 'L2', 'exit_code': 0, 'target_env': label} for label in labels),
            {'stage': 'ITEMS', 'exit_code': 0, 'afm_pairs': 0, 'failures': 0}]}


def test_ship_requires_a_render_of_every_declared_state_bound_to_its_steps(tmp_path):
    urls = {'DESIGN_WEB_URL': 'http://127.0.0.1:12345/'}
    result = ship_fixture(tmp_path, {'open': OPEN}, [{'name': 'open', 'steps': OPEN}])
    assert validate_web_result(tmp_path, result, urls)
    # The state's report must have run the declared steps, not some other interaction.
    result = ship_fixture(tmp_path, {'open': OPEN}, [{'name': 'open', 'steps': [{'click': '#other'}]}])
    with pytest.raises(ValueError):
        validate_web_result(tmp_path, result, urls)
    # A state report missing from the stages is a partial run.
    result = ship_fixture(tmp_path, {'open': OPEN}, [{'name': 'open', 'steps': OPEN}])
    broken = copy.deepcopy(result)
    del broken['stages'][3]
    with pytest.raises(ValueError):
        validate_web_result(tmp_path, broken, urls)


# --- on rendered pages ---------------------------------------------------------------------------

BUTTON = ('<style>h1,h2,p{margin:0 0 16px}</style><button id="open" type="button" style="display:block;min-height:48px;padding:12px 16px;font:inherit;'
          'color:black;background:white;border:1px solid black">Open</button>')
DIALOG = ('<script>document.getElementById("open").addEventListener("click",()=>{{const d=document.createElement("div");'
          'd.setAttribute("role","dialog");d.setAttribute("aria-label","Notice");'
          'd.style.cssText="padding:16px;border:1px solid black;background:white;color:black";'
          'd.innerHTML={inner};document.querySelector("main").append(d)}})</script>')
RULE = {'id': 'dialog-title', 'kind': 'count', 'scope': ["[role='dialog']"], 'targets': ['h2'], 'minimum': 1}


def run_states(rendered, project, inner, states):
    run, root, _ = rendered
    use_catalog_item(project, catalog_item(*DECISION, grade='nice-to-have'))
    write(root / 'design/checks/web.yaml', {'rules': [RULE]})
    set_items(root, {'dads-000001': {'state': 'defined', 'evidence': ['rule:dialog-title']}})
    validation = json.loads((root / 'design/validation.json').read_text())
    validation['urls'] = [{'env': 'DESIGN_FIXTURE_URL', **({'states': states} if states else {})}]
    write(root / 'design/validation.json', validation)
    code, result = run(BUTTON + DIALOG.format(inner=json.dumps(inner)))
    path = root / 'design/.resolved/items.json'
    items = json.loads(path.read_text()) if path.exists() else {}
    return code, result, items.get('machine', {}).get('dads-000001', {})


def test_a_rule_on_a_dialog_applies_only_in_the_declared_state(rendered, project):
    # Without a state the dialog never renders: the rule checks nothing (n/a and warned), never a pass.
    code, result, status = run_states(rendered, project, '<h2>Notice</h2><p>Saved.</p>', None)
    assert status == {'DESIGN_FIXTURE_URL': 'n/a'}, result
    assert any('rule dialog-title never applied' in w for w in result['warnings']), result['warnings']
    # With the state, the page is still checked as before and the opened dialog is checked too.
    code, result, status = run_states(rendered, project, '<h2>Notice</h2><p>Saved.</p>', {'dialog-open': OPEN})
    assert status == {'DESIGN_FIXTURE_URL': 'n/a', 'DESIGN_FIXTURE_URL#dialog-open': 'pass'}, (result['errors'], result['stages'])
    assert [s.get('target_env') for s in result['stages'] if s['stage'] == 'L2'] == ['DESIGN_FIXTURE_URL', 'DESIGN_FIXTURE_URL#dialog-open']
    assert not any('never applied' in w for w in result.get('warnings', [])), result['warnings']
    # A dialog that breaks the rule fails on the state, named as <env>#<state>.
    code, result, status = run_states(rendered, project, '<p>Saved.</p>', {'dialog-open': OPEN})
    assert code == 2 and status.get('DESIGN_FIXTURE_URL#dialog-open') == 'fail', result
    assert any(e.startswith('DESIGN_FIXTURE_URL#dialog-open: rule:dialog-title') for e in result['errors']), result['errors']


def test_a_step_whose_selector_matches_nothing_cannot_reach_the_state(rendered, project):
    code, result, status = run_states(rendered, project, '<h2>Notice</h2>', {'dialog-open': [{'click': '#missing'}]})
    assert code == 2, result
    assert stages(result) == ['L0', 'L1', 'L2', 'L2'] and result['stages'][-1]['exit_code'] == 2, result['stages']
    assert any('could not reach state DESIGN_FIXTURE_URL#dialog-open' in e and '#missing' in e for e in result['errors']), result['errors']
    assert 'DESIGN_FIXTURE_URL#dialog-open' not in status  # nothing recorded as passing
