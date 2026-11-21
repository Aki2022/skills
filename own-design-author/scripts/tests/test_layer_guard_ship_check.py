"""design-check, design-ship and the resolved-manifest / fixed-render readers follow the layer-paths rule and refuse a
repository they cannot read (ISSUE-02 of the ux move). A moved repository must stop each road with
LAYER_LAYOUT_FOREIGN_ROOT; none may skip, report "0 found" or swallow the reason into a generic failure."""
import json
import os
import subprocess
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from layer_testkit import (ALL_SKILLS, FOREIGN, FOREIGN_CODE, LAYOUT, ROOT, SKILLS, assert_foreign_failure,
                           copy_skills, git_init, moved_repo, run)

import layer_paths as lp  # noqa: E402

SCRIPTS = SKILLS / 'own-design-author/scripts'
NODE = os.environ.get('DESIGN_TEST_NODE', 'node')
RESOLVED = LAYOUT['entries']['resolved']

KINDS = ['foreign_only', 'both']
CHECK_ARGS = {'default': [], 'report_all': ['--report-all']}
SHIP_ARGS = {'default': [], 'channel': ['--channel', 'web'], 'pre_push': ['--pre-push'], 'doctor': ['--doctor'],
             'integration': ['--integration']}


@pytest.mark.parametrize('kind', KINDS)
@pytest.mark.parametrize('mode', sorted(CHECK_ARGS))
def test_design_check_stops_on_a_moved_repository(tmp_path, kind, mode):
    repo = moved_repo(tmp_path, kind)
    assert_foreign_failure(run([SCRIPTS / 'design-check', '--repo', repo, *CHECK_ARGS[mode]]))


@pytest.mark.parametrize('kind', KINDS)
@pytest.mark.parametrize('mode', sorted(SHIP_ARGS))
def test_design_ship_stops_on_a_moved_repository(tmp_path, kind, mode):
    repo = moved_repo(tmp_path, kind)
    assert_foreign_failure(run([SCRIPTS / 'design-ship', '--repo', repo, *SHIP_ARGS[mode]]))


@pytest.mark.parametrize('kind', KINDS)
def test_design_ship_probe_stops_on_a_moved_repository(tmp_path, kind):
    repo = moved_repo(tmp_path, kind)
    result = run([SCRIPTS / 'design-ship', '--probe', '--repo', repo])
    assert result.returncode == 2, result.stdout + result.stderr
    assert FOREIGN_CODE in result.stderr


def test_design_ship_probe_says_present_or_absent(tmp_path):
    empty = git_init(tmp_path / 'empty')
    assert run([SCRIPTS / 'design-ship', '--probe', '--repo', empty]).returncode == 3
    (empty / LAYOUT['entries']['design_config']).parent.mkdir(parents=True)
    (empty / LAYOUT['entries']['design_config']).write_text('{}')
    assert run([SCRIPTS / 'design-ship', '--probe', '--repo', empty]).returncode == 0


def test_design_ship_follows_a_renamed_layer_root(tmp_path):
    """The config is read from the place the rule names, not from design/: rename every entry in a copy of the skills
    and put design.yaml only at the new place."""
    renamed = 'zzmoved'

    def mutate(data):
        data['entries'] = {k: v.replace(ROOT, renamed, 1) for k, v in data['entries'].items()}
    skills = copy_skills(tmp_path / 'skills', ALL_SKILLS, mutate)
    repo = git_init(tmp_path / 'repo')
    (repo / renamed).mkdir()
    (repo / renamed / 'design.yaml').write_text('{}')
    result = run([skills / 'own-design-author/scripts/design-ship', '--repo', repo, '--doctor'])
    out = json.loads(result.stdout)
    assert any('no ship:true channels' in e for e in out['errors']), 'the renamed config was not read: ' + result.stdout
    assert not (repo / ROOT).exists(), 'a command wrote the old root'
    plain = git_init(tmp_path / 'plain')
    (plain / ROOT).mkdir()
    (plain / ROOT / 'design.yaml').write_text('{}')
    baseline = json.loads(run([SCRIPTS / 'design-ship', '--repo', plain, '--doctor']).stdout)
    assert baseline['errors'] == out['errors']  # same behaviour before and after the rename


@pytest.mark.parametrize('kind', KINDS)
def test_load_manifest_stops_in_a_moved_repository(tmp_path, kind):
    from resolved_contract import load_manifest
    repo = moved_repo(tmp_path, kind)
    manifest = repo / RESOLVED / 'web.json'
    with pytest.raises(lp.LayerLayoutError) as caught:  # the file is missing: the guard speaks first
        load_manifest(manifest)
    assert FOREIGN_CODE in str(caught.value)
    manifest.parent.mkdir(parents=True, exist_ok=True)
    manifest.write_text('{}')
    with pytest.raises(lp.LayerLayoutError) as caught:  # present: still not read
        load_manifest(manifest)
    assert FOREIGN_CODE in str(caught.value)


def test_load_manifest_under_a_foreign_root_names_the_cause(tmp_path):
    from resolved_contract import load_manifest
    manifest = tmp_path / FOREIGN[0] / RESOLVED / 'web.json'
    manifest.parent.mkdir(parents=True)
    manifest.write_text('{}')
    with pytest.raises(lp.LayerLayoutError) as caught:
        load_manifest(manifest)
    assert FOREIGN_CODE in str(caught.value)


def test_load_manifest_keeps_rejecting_a_manifest_in_the_wrong_place(tmp_path):
    from resolved_contract import load_manifest
    (tmp_path / 'elsewhere').mkdir()
    (tmp_path / 'elsewhere/web.json').write_text('{}')
    with pytest.raises(ValueError, match='manifest must be'):
        load_manifest(tmp_path / 'elsewhere/web.json')


@pytest.mark.parametrize('kind', KINDS)
def test_check_asset_exports_stops_in_a_moved_repository(tmp_path, kind):
    from resolved_contract import check_asset_exports
    repo = moved_repo(tmp_path, kind)
    with pytest.raises(lp.LayerLayoutError):
        check_asset_exports(repo, {'asset_exports': [{'source': 'assets/x.svg', 'sha256': '0' * 64, 'targets': []}]})


@pytest.mark.parametrize('kind', KINDS)
def test_ship_evidence_validation_stops_in_a_moved_repository(tmp_path, kind):
    from ship_contract import validate_web_result
    repo = moved_repo(tmp_path, kind)
    with pytest.raises(lp.LayerLayoutError):
        validate_web_result(repo, {'exit_code': 0}, {})


def _node(code, *args):
    return subprocess.run([NODE, '--input-type=module', '-e', code, *args], capture_output=True, text=True)


@pytest.mark.parametrize('kind', KINDS)
def test_javascript_load_manifest_stops_in_a_moved_repository(tmp_path, kind):
    repo = moved_repo(tmp_path, kind)
    manifest = repo / RESOLVED / 'web.json'
    code = ('import {loadManifest} from "' + (SCRIPTS / 'resolved_contract.mjs').as_uri() + '";'
            'try{loadManifest(process.argv[1]);console.log("no error")}catch(e){console.log(e.message)}')
    missing = _node(code, str(manifest))
    assert missing.returncode == 0, missing.stderr
    assert FOREIGN_CODE in missing.stdout, 'a missing manifest must still get the layout message: ' + missing.stdout
    manifest.parent.mkdir(parents=True, exist_ok=True)
    manifest.write_text('{}')
    present = _node(code, str(manifest))
    assert present.returncode == 0, present.stderr
    assert FOREIGN_CODE in present.stdout, present.stdout


def test_javascript_load_manifest_under_a_foreign_root_names_the_cause(tmp_path):
    manifest = tmp_path / FOREIGN[0] / RESOLVED / 'web.json'
    manifest.parent.mkdir(parents=True)
    manifest.write_text('{}')
    code = ('import {loadManifest} from "' + (SCRIPTS / 'resolved_contract.mjs').as_uri() + '";'
            'try{loadManifest(process.argv[1]);console.log("no error")}catch(e){console.log(e.message)}')
    result = _node(code, str(manifest))
    assert FOREIGN_CODE in result.stdout, result.stdout + result.stderr


@pytest.mark.parametrize('kind', KINDS)
def test_javascript_load_environment_stops_in_a_moved_repository(tmp_path, kind):
    repo = moved_repo(tmp_path, kind)
    environment = repo / LAYOUT['entries']['render_environment']
    environment.parent.mkdir(parents=True, exist_ok=True)
    environment.write_text('{"schema_version":1}')
    code = ('import {loadEnvironment} from "' + (SCRIPTS / 'fixed_render.mjs').as_uri() + '";'
            'try{loadEnvironment(process.argv[1],process.argv[2],[375],["light"]);console.log("no error")}'
            'catch(e){console.log(e.message)}')
    result = _node(code, str(environment), str(repo))
    assert result.returncode == 0, result.stderr
    assert FOREIGN_CODE in result.stdout, 'loadEnvironment must stop with the layout message: ' + result.stdout
