"""Every other entry of own-design-author (migrate, approve, the judgment CLI and the importable roads) stops on a
repository whose layers moved to a layout this version does not read, instead of skipping or reading around it."""
import json
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from layer_testkit import FOREIGN_CODE, SKILLS, assert_foreign_failure, moved_repo, run  # noqa: E402

import design_author  # noqa: E402
import design_modules  # noqa: E402
import layer_paths  # noqa: E402
import maintenance  # noqa: E402
import owner_approval  # noqa: E402
from reporting import empty_report  # noqa: E402

KINDS = ['foreign_only', 'both']
CLI = SKILLS / 'own-design-author/scripts/design-author'
LEDGER_CLI = SKILLS / 'own-design-author/scripts/judgment_ledger.py'
SYNTHETIC_CATALOG = SKILLS / 'own-design-author/catalog'


@pytest.mark.parametrize('kind', KINDS)
def test_migrate_stops_on_a_moved_repository(tmp_path, kind):
    repo = moved_repo(tmp_path, kind)
    result = run([CLI, 'migrate', '--repo', repo, '--catalog-version', 'test-v1'])
    assert_foreign_failure(result)


@pytest.mark.parametrize('kind', KINDS)
def test_approve_prepare_stops_on_a_moved_repository(tmp_path, kind):
    repo = moved_repo(tmp_path, kind)
    result = run([CLI, 'approve', '--prepare', '--must', '--repo', repo])
    assert_foreign_failure(result)


@pytest.mark.parametrize('kind', KINDS)
def test_approve_record_stops_on_a_moved_repository_and_writes_nothing(tmp_path, kind):
    repo = moved_repo(tmp_path, kind)
    request = tmp_path / 'request.json'
    request.write_text(json.dumps({'rows': []}))
    result = run([CLI, 'approve', '--request', request, '--owner-words', 'fixture', '--approved-on', '2026-11-21',
                  '--repo', repo])
    assert_foreign_failure(result)
    assert not (repo / layer_paths.rel('resolved')).exists(), 'the approval lock directory was created in a moved repo'


@pytest.mark.parametrize('kind', KINDS)
@pytest.mark.parametrize('command', ['rekey'])
def test_judgment_ledger_cli_stops_on_a_moved_repository(tmp_path, kind, command):
    repo = moved_repo(tmp_path, kind)
    result = run([sys.executable, LEDGER_CLI, command, '--repo', repo])
    assert_foreign_failure(result)
    assert not (repo / layer_paths.rel('resolved')).exists()


@pytest.mark.parametrize('kind', KINDS)
def test_check_design_reports_the_layout_error_and_leaves_the_repo_alone(tmp_path, kind):
    repo = moved_repo(tmp_path, kind)
    before = sorted(p.relative_to(repo).as_posix() for p in repo.rglob('*') if '.git' not in p.parts)
    result = design_author.check_design(repo)
    assert result['exit_code'] == 2
    assert FOREIGN_CODE in ' '.join(result['errors'])
    assert sorted(p.relative_to(repo).as_posix() for p in repo.rglob('*') if '.git' not in p.parts) == before


@pytest.mark.parametrize('kind', KINDS)
def test_maintain_design_does_not_skip_a_moved_repository(tmp_path, kind):
    repo = moved_repo(tmp_path, kind)
    result = maintenance.maintain_design(repo)
    assert result['exit_code'] == 2 and result['status'] != 'skipped', result
    assert FOREIGN_CODE in json.dumps(result, ensure_ascii=False)


@pytest.mark.parametrize('kind', KINDS)
def test_maintain_all_does_not_skip_either_stage_on_a_moved_repository(tmp_path, kind):
    repo = moved_repo(tmp_path, kind)
    result = maintenance.maintain_all(repo)
    assert result['exit_code'] == 2, result
    assert result['status'] != 'skipped'
    assert FOREIGN_CODE in json.dumps(result, ensure_ascii=False)


@pytest.mark.parametrize('kind', KINDS)
def test_resolve_ia_does_not_skip_a_moved_repository(tmp_path, kind):
    repo = moved_repo(tmp_path, kind)
    with pytest.raises(layer_paths.LayerLayoutError):
        design_modules.resolve_ia(repo, {}, {}, {}, empty_report())


@pytest.mark.parametrize('kind', KINDS)
def test_owner_approval_roads_stop_on_a_moved_repository(tmp_path, kind):
    repo = moved_repo(tmp_path, kind)
    with pytest.raises(layer_paths.LayerLayoutError):
        owner_approval.prepare(repo, SYNTHETIC_CATALOG, [], True)
    with pytest.raises(layer_paths.LayerLayoutError):
        owner_approval.approve_request(repo, SYNTHETIC_CATALOG, {}, 'owner', owner_words='x', approved_on='2026-11-21')
    with pytest.raises(layer_paths.LayerLayoutError):
        owner_approval._ia_only(repo)


@pytest.mark.parametrize('kind', KINDS)
def test_init_and_sync_roads_stop_on_a_moved_repository(tmp_path, kind):
    repo = moved_repo(tmp_path, kind)
    with pytest.raises(layer_paths.LayerLayoutError):
        design_author.migrate_design(repo, 'test-v1', catalog_dir=SYNTHETIC_CATALOG)
    with pytest.raises(layer_paths.LayerLayoutError):
        design_author.sync_artifacts(repo, {'artifacts': {}, 'asset_exports': [{'source': 'assets/x.svg', 'targets': ['x.svg']}]})


def test_ia_manifest_input_keys_follow_the_rule():
    """The infoarch input keys a resolved manifest may carry come from the entry, not from a spelled root."""
    key = layer_paths.rel('infoarch', 'info_arch.yaml')
    assert design_modules._ia_input_key_ok(key)
    assert not design_modules._ia_input_key_ok(key.replace(layer_paths.rel('infoarch'), 'elsewhere', 1))
