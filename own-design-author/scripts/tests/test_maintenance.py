"""Maintenance runs real L0, observes git changes, and never writes approvals."""
from datetime import date
import json
import subprocess

from test_core import project, approve, write, SKILL, check
from maintenance import maintain_design


def commit(root, message='fixture baseline'):
    subprocess.run(['git', 'add', '.'], cwd=root, check=True, capture_output=True)
    subprocess.run(['git', '-c', 'user.name=Fixture Owner', '-c',
                    'user.email=fixture@example.org', 'commit', '-qm', message],
                   cwd=root, check=True, capture_output=True)


def ready(project):
    root, catalogs = project
    (root / '.gitignore').write_text('design/.resolved/\ndesign/assets.json\n')
    approve(project)
    commit(root)
    return root, catalogs


def run(project, **kwargs):
    root, catalogs = project
    return maintain_design(root, catalog_dir=catalogs, today=date(2026, 10, 23), **kwargs)


def test_no_design_skips_as_one_cli_line(tmp_path):
    repo = tmp_path / 'empty'
    subprocess.run(['git', 'init', '-q', str(repo)], check=True)
    result = maintain_design(repo)
    assert result['status'] == 'skipped' and result['reason'] == 'design/ absent'
    from test_core import sys
    cli = subprocess.run([sys.executable, str(SKILL / 'scripts/design_author.py'),
                          'maintain', '--repo', str(repo)], capture_output=True, text=True)
    assert cli.returncode == 0 and len(cli.stdout.splitlines()) == 1
    assert json.loads(cli.stdout)['status'] == 'skipped'


def test_first_run_generates_real_manifest_then_clean_run_skips(project):
    root, _ = ready(project)
    result = run(project)
    assert result['status'] == 'checked' and result['exit_code'] == 0
    assert 'first-run' in result['triggers']
    assert (root / 'design/.resolved/web.json').exists()
    assert run(project)['status'] == 'skipped'


def test_dirty_and_committed_changes_trigger_without_writing_approvals(project):
    root, _ = ready(project)
    before = (root / 'design/approvals.yaml').read_bytes()
    run(project)
    (root / 'example.txt').write_text('changed source')
    assert 'working-tree' in run(project)['triggers']
    commit(root, 'fixture source addition')
    result = run(project)
    assert 'commit-diff' in result['triggers'] and result['changed_paths'] == ['example.txt']
    assert (root / 'design/approvals.yaml').read_bytes() == before


def test_rename_and_delete_are_seen(project):
    root, _ = ready(project)
    (root / 'old.txt').write_text('some source')
    commit(root)
    run(project)
    (root / 'old.txt').rename(root / 'new.txt')
    result = run(project)
    assert set(result['changed_paths']) == {'old.txt', 'new.txt'}
    commit(root)
    assert 'commit-diff' in run(project)['triggers']


def test_due_review_runs_with_zero_diff(project):
    root, _ = ready(project)
    path = root / 'design/design.yaml'
    data = json.loads(path.read_text())
    data['maintenance'] = {'within_days': 7, 'review_due_dates': {'dads-000001': '2026-10-29'}}
    write(path, data)
    commit(root)
    assert run(project)['exit_code'] == 0
    result = run(project)
    assert result['changed_paths'] == [] and result['triggers'] == ['review-due']
    assert result['due_items'] == [{'id': 'dads-000001', 'review_by': '2026-10-29'}]


def test_policy_or_catalog_bytes_change_forces_regeneration(project):
    root, catalogs = ready(project)
    run(project)
    catalog = catalogs / 'test-v1.json'
    data = json.loads(catalog.read_text())
    data['items'][0]['title'] = 'New upstream heading'
    write(catalog, data)
    result = run(project)
    assert 'versions-changed' in result['triggers']


def test_stale_approval_is_listed_and_failed_run_does_not_advance_stamp(project):
    root, _ = ready(project)
    run(project)
    stamp = root / 'design/.resolved/maintenance-stamp.json'
    before = json.loads(stamp.read_text())
    path = root / 'design/tokens/brand.tokens.json'
    data = json.loads(path.read_text())
    data['color']['primary']['$value'].update(components=[1, 1, 1], hex='#ffffff')
    write(path, data)
    ledger = (root / 'design/approvals.yaml').read_bytes()
    result = run(project)
    assert result['exit_code'] == 2 and result['invalidated_approvals'] == ['ds:web:token:color.primary']
    assert not stamp.exists() or json.loads(stamp.read_text()) == before
    assert (root / 'design/approvals.yaml').read_bytes() == ledger
    assert not (root / 'design/.resolved/web.json').exists()


def test_corrupt_stamp_cannot_make_a_clean_repo_skip(project):
    root, _ = ready(project)
    run(project)
    (root / 'design/.resolved/maintenance-stamp.json').write_text('{"commit":"bad"}')
    result = run(project)
    assert result['status'] == 'checked' and 'invalid-stamp' in result['triggers']


def test_missing_generated_manifest_cannot_be_skipped(project):
    root, _ = ready(project)
    run(project)
    (root / 'design/.resolved/web.json').unlink()
    result = run(project)
    assert result['status'] == 'checked' and 'resolved-missing' in result['triggers']
    assert (root / 'design/.resolved/web.json').exists()


def test_regular_l0_does_not_erase_the_maintenance_commit_stamp(project):
    root, _ = ready(project)
    run(project)
    original = (root / 'design/.resolved/maintenance-stamp.json').read_bytes()
    assert check(project)['exit_code'] == 0
    assert (root / 'design/.resolved/maintenance-stamp.json').read_bytes() == original
    assert run(project)['status'] == 'skipped'


def test_invalid_due_id_or_date_is_red(project):
    root, _ = ready(project)
    for due in [{'unknown': '2026-10-23'}, {'dads-000001': 'tomorrow'}]:
        config = json.loads((root / 'design/design.yaml').read_text())
        config['maintenance'] = {'within_days': 7, 'review_due_dates': due}
        write(root / 'design/design.yaml', config)
        result = run(project)
        assert result['exit_code'] == 2


def test_unborn_head_is_visible_failure(project):
    result = run(project)
    assert result['exit_code'] == 2 and 'commit' in ' '.join(result['errors']).lower()


def test_invalid_configuration_is_a_reported_failure(project):
    root, _ = ready(project)
    write(root / 'design/design.yaml', [])
    result = run(project)
    assert result['exit_code'] == 2 and result['status'] == 'failed'
    assert 'object' in ' '.join(result['errors'])
