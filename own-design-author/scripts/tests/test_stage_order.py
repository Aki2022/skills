"""Cheap checks run before expensive ones: L0 before the gate's preparation, and code checks per page before
rendering the next page, so a failure never waits for work it makes useless."""
import json

from test_core import project, write, approve  # noqa: F401  (fixture reuse)
from test_ship_contract import channel
from test_ship_integration import browser_env, commit, run_git
from test_item_coverage import catalog_item, use_catalog_item, rendered, stages  # noqa: F401


def local_gate_with_prepare(root):
    config = json.loads((root / 'design/design.yaml').read_text())
    row = {k: v for k, v in channel().items() if k in ('media', 'merge_base', 'render_servers')}
    # The marker lands in the ignored .resolved/ directory, so the preparation leaves the committed source clean.
    config['integration_gate'] = {**row, 'prepare_command': ['python3', '-c', 'open("design/.resolved/prepared", "w").write("ran")']}
    write(root / 'design/design.yaml', config)
    return config['integration_gate']


def test_l0_failure_stops_before_the_preparation_runs(project, tmp_path, monkeypatch):
    import sys
    from ship import gate
    monkeypatch.setenv('DESIGN_AUTHOR_PYTHON', sys.executable)  # the L0 subprocess uses this test's interpreter
    root, catalogs = project
    env = browser_env()
    run_git(root, env, 'config', 'user.name', 'Fixture Owner')
    run_git(root, env, 'config', 'user.email', 'fixture-owner')  # git accepts any string; no address needed
    (root / '.gitignore').write_text('design/.resolved/\n')
    marker = root / 'design/.resolved/prepared'
    row = local_gate_with_prepare(root)
    commit(root, env, 'fixture unapproved must item')  # proposed must token: ship L0 fails
    result = gate(root, 'local-integration', row, catalogs, local=True)
    assert result['exit_code'] == 2 and any('proposed must' in e for e in result['errors']), result['errors']
    assert not marker.exists()  # the preparation never ran for a design system that cannot ship
    approve(project)
    commit(root, env, 'fixture approved')
    try:
        passed = gate(root, 'local-integration', row, catalogs, local=True)
    except (ValueError, OSError, KeyError) as error:
        passed = error  # rendering is not set up here; only the order matters
    assert marker.exists() and marker.read_text() == 'ran', passed


def test_item_failure_on_the_first_page_stops_before_rendering_the_next(rendered, project):
    run, root, tmp_path = rendered
    use_catalog_item(project, catalog_item('foundations/spacing', '余白ルールの例', grade='nice-to-have'))
    validation = json.loads((root / 'design/validation.json').read_text())
    validation['urls'] = [{'env': 'DESIGN_FIXTURE_URL'}, {'env': 'DESIGN_FIXTURE_URL_2'}]
    write(root / 'design/validation.json', validation)
    code, result = run('<p style="padding:13px">off scale</p>', DESIGN_FIXTURE_URL_2='http://127.0.0.1:9/')
    assert code == 2, result
    assert stages(result) == ['L0', 'L1', 'L2', 'ITEMS'], stages(result)  # the second page was never rendered
    assert any('palette:spacing' in e and 'dads-000001' in e for e in result['errors']), result['errors']
    assert not (root / 'design/.resolved/l2-1.json').exists()


def test_pages_render_in_parallel_but_are_judged_in_declared_order(rendered, project):
    run, root, tmp_path = rendered
    use_catalog_item(project, catalog_item('foundations/spacing', '余白ルールの例', grade='nice-to-have'))
    validation = json.loads((root / 'design/validation.json').read_text())
    validation.update(urls=[{'env': 'DESIGN_FIXTURE_URL'}, {'env': 'DESIGN_FIXTURE_URL_2'}], parallel=2)
    write(root / 'design/validation.json', validation)
    code, result = run('<p style="padding:16px">ok</p>')
    assert code in (0, 1), result
    assert [(s['stage'], s.get('target_env')) for s in result['stages']][2:] == [
        ('L2', 'DESIGN_FIXTURE_URL'), ('L2', 'DESIGN_FIXTURE_URL_2'), ('ITEMS', None)]
    items = json.loads((root / 'design/.resolved/items.json').read_text())
    assert items['machine'] == {'dads-000001': {'DESIGN_FIXTURE_URL': 'pass', 'DESIGN_FIXTURE_URL_2': 'pass'}}
    code, result = run('<p style="padding:13px">off scale</p>')
    assert code == 2 and stages(result) == ['L0', 'L1', 'L2', 'ITEMS'], stages(result)
