"""The one place that knows where the layers live (references/layer-paths.json), and the guard that stops this
version of the skills from silently reading around a repository that has already moved."""
import json
import shutil
import subprocess
from pathlib import Path

import pytest
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import layer_paths as lp  # noqa: E402

SKILL = Path(__file__).resolve().parents[2]
SCRIPTS = SKILL / 'scripts'
LAYOUT = json.loads((SKILL / 'references/layer-paths.json').read_text())
FOREIGN = LAYOUT['foreign_roots']
ROOT = LAYOUT['entries']['layers']


def _node(code, *args):
    run = subprocess.run(['node', '--input-type=module', '-e', code, *args], capture_output=True, text=True)
    return run


def test_every_entry_is_a_relative_path_under_the_layers_root():
    for key, rel in LAYOUT['entries'].items():
        assert not rel.startswith('/') and '..' not in rel.split('/'), key
        assert rel == ROOT or rel.startswith(ROOT + '/'), key


def test_rel_and_path_join_entry_with_the_rest():
    repo = Path('/somewhere/repo')
    assert lp.rel('ledger') == LAYOUT['entries']['ledger']
    assert lp.rel('infoarch', 'blueprint', 'x.yaml') == LAYOUT['entries']['infoarch'] + '/blueprint/x.yaml'
    assert lp.path(repo, 'resolved', 'web.json') == repo / LAYOUT['entries']['resolved'] / 'web.json'


def test_unknown_key_is_an_error_not_a_guess():
    with pytest.raises(KeyError):
        lp.rel('nope')


def test_rest_may_not_escape_the_entry():
    with pytest.raises(ValueError):
        lp.rel('ledger', '..', 'x')
    with pytest.raises(ValueError):
        lp.rel('ledger', '/abs')


@pytest.mark.parametrize('layers', ['foreign_only', 'both'])
def test_guard_fails_when_the_repository_has_already_moved(tmp_path, layers):
    for name in FOREIGN:
        (tmp_path / name).mkdir()
    if layers == 'both':
        (tmp_path / ROOT).mkdir()
    with pytest.raises(lp.LayerLayoutError) as caught:
        lp.guard(tmp_path)
    assert 'LAYER_LAYOUT_FOREIGN_ROOT' in str(caught.value)
    assert FOREIGN[0] + '/' in str(caught.value)
    with pytest.raises(lp.LayerLayoutError):
        lp.path(tmp_path, 'ledger')  # every road goes through the guard


def test_guard_passes_for_the_current_layout_and_for_an_empty_repository(tmp_path):
    lp.guard(tmp_path)
    (tmp_path / ROOT).mkdir()
    lp.guard(tmp_path)
    assert lp.path(tmp_path, 'ledger') == tmp_path / LAYOUT['entries']['ledger']


def test_a_plain_file_named_like_the_foreign_root_is_not_a_layer_root(tmp_path):
    (tmp_path / FOREIGN[0]).write_text('not a directory')
    lp.guard(tmp_path)


def test_manifest_path_gives_the_repository_even_when_the_file_is_missing(tmp_path):
    manifest = tmp_path / LAYOUT['entries']['resolved'] / 'web.json'
    assert lp.repo_from_resolved(manifest) == tmp_path.resolve()


def test_manifest_path_of_a_moved_repository_fails_with_the_guard_message(tmp_path):
    (tmp_path / FOREIGN[0]).mkdir()
    manifest = tmp_path / LAYOUT['entries']['resolved'] / 'web.json'
    with pytest.raises(lp.LayerLayoutError) as caught:
        lp.repo_from_resolved(manifest)
    assert 'LAYER_LAYOUT_FOREIGN_ROOT' in str(caught.value)


def test_manifest_path_pointing_into_a_foreign_root_names_the_cause(tmp_path):
    manifest = tmp_path / FOREIGN[0] / 'design' / '.resolved' / 'web.json'
    with pytest.raises(lp.LayerLayoutError) as caught:
        lp.repo_from_resolved(manifest)
    assert 'LAYER_LAYOUT_FOREIGN_ROOT' in str(caught.value)


def test_manifest_path_in_the_wrong_place_is_rejected(tmp_path):
    with pytest.raises(ValueError):
        lp.repo_from_resolved(tmp_path / 'elsewhere' / 'web.json')


def test_node_module_agrees_with_the_python_module(tmp_path):
    repo = tmp_path / 'repo'
    repo.mkdir()
    cases = [['ledger'], ['infoarch', 'blueprint', 'x.yaml'], ['resolved', 'web.json'], ['render_environment']]
    code = ('import {layerRel, layerPath} from "' + (SCRIPTS / 'layer_paths.mjs').as_uri() + '";'
            'const out=JSON.parse(process.argv[1]).map(c=>[layerRel(...c),layerPath(process.argv[2],...c)]);'
            'console.log(JSON.stringify(out));')
    run = _node(code, json.dumps(cases), str(repo))
    assert run.returncode == 0, run.stderr
    got = json.loads(run.stdout)
    for case, (rel, full) in zip(cases, got):
        assert rel == lp.rel(*case)
        assert full == str(lp.path(repo, *case))


def test_node_guard_fails_like_the_python_guard(tmp_path):
    (tmp_path / FOREIGN[0]).mkdir()
    code = ('import {layerPath} from "' + (SCRIPTS / 'layer_paths.mjs').as_uri() + '";'
            'try{layerPath(process.argv[1],"ledger");console.log("no error")}catch(e){console.log(e.message)}')
    run = _node(code, str(tmp_path))
    assert run.returncode == 0, run.stderr
    assert 'LAYER_LAYOUT_FOREIGN_ROOT' in run.stdout


def test_node_manifest_root_matches_python(tmp_path):
    manifest = tmp_path / LAYOUT['entries']['resolved'] / 'web.json'
    code = ('import {repoFromResolved} from "' + (SCRIPTS / 'layer_paths.mjs').as_uri() + '";'
            'console.log(repoFromResolved(process.argv[1]))')
    run = _node(code, str(manifest))
    assert run.returncode == 0, run.stderr
    assert run.stdout.strip() == str(lp.repo_from_resolved(manifest))
