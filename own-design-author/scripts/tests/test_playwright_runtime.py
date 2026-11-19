"""Playwright is found per machine by version, never through NODE_PATH or another repo's checkout."""
import json
import os
from pathlib import Path
import subprocess

NODE = os.environ.get('DESIGN_TEST_NODE', 'node')
HELPER = Path(__file__).resolve().parents[1] / 'playwright_runtime.mjs'


def fake(dir, version, marker):
    pkg = dir / 'node_modules/playwright'
    pkg.mkdir(parents=True)
    (pkg / 'package.json').write_text(json.dumps({'name': 'playwright', 'version': version, 'main': 'index.js'}))
    (pkg / 'index.js').write_text('module.exports={chromium:{},marker:%s};' % json.dumps(marker))
    return dir


def run(cwd, env, version=None):
    js = ('import {loadPlaywright} from ' + json.dumps(HELPER.as_uri()) + ';'
          'const r=loadPlaywright({cwd:process.cwd(),version:process.argv.at(-1)||null});'
          'console.log(JSON.stringify({marker:r.playwright.marker,source:r.source,version:r.version}));')
    base = {k: v for k, v in os.environ.items() if k not in ('NODE_PATH', 'OWN_PLAYWRIGHT_DIR', 'OWN_PLAYWRIGHT_HOME', 'INFOARCH_PLAYWRIGHT_DIR')}
    proc = subprocess.run([NODE, '--input-type=module', '-e', js, version or ''], cwd=cwd,
                          env={**base, **env}, capture_output=True, text=True)
    return proc.returncode, (json.loads(proc.stdout) if proc.returncode == 0 else proc.stderr)


def test_shared_runtime_is_found_by_pinned_version(tmp_path):
    home = tmp_path / 'home'
    fake(home / '1.63.0', '1.63.0', 'shared')
    repo = tmp_path / 'repo'; repo.mkdir()
    rc, out = run(repo, {'OWN_PLAYWRIGHT_HOME': str(home)}, '1.63.0')
    assert rc == 0, out
    assert out == {'marker': 'shared', 'source': 'shared', 'version': '1.63.0'}


def test_repo_local_playwright_wins_when_its_version_matches(tmp_path):
    home = tmp_path / 'home'; fake(home / '1.63.0', '1.63.0', 'shared')
    repo = fake(tmp_path / 'repo', '1.63.0', 'local')
    rc, out = run(repo, {'OWN_PLAYWRIGHT_HOME': str(home)}, '1.63.0')
    assert rc == 0, out
    assert out['marker'] == 'local' and out['source'] == 'project'


def test_repo_local_playwright_of_another_version_is_skipped_for_the_pin(tmp_path):
    home = tmp_path / 'home'; fake(home / '1.63.0', '1.63.0', 'shared')
    repo = fake(tmp_path / 'repo', '1.50.0', 'local')
    rc, out = run(repo, {'OWN_PLAYWRIGHT_HOME': str(home)}, '1.63.0')
    assert rc == 0, out
    assert out['marker'] == 'shared'


def test_node_path_pointing_at_another_repo_is_ignored(tmp_path):
    other = fake(tmp_path / 'other_repo', '1.63.0', 'other')
    repo = tmp_path / 'repo'; repo.mkdir()
    rc, out = run(repo, {'NODE_PATH': str(other / 'node_modules'), 'OWN_PLAYWRIGHT_HOME': str(tmp_path / 'empty')}, '1.63.0')
    assert rc == 2
    assert 'npm install --prefix' in out and '1.63.0' in out


def test_missing_runtime_stops_with_the_install_command(tmp_path):
    repo = tmp_path / 'repo'; repo.mkdir()
    rc, out = run(repo, {'OWN_PLAYWRIGHT_HOME': str(tmp_path / 'home')}, '1.63.0')
    assert rc == 2
    assert str(tmp_path / 'home' / '1.63.0') in out


def test_explicit_dir_is_authoritative_and_must_match_the_pin(tmp_path):
    home = tmp_path / 'home'; fake(home / '1.63.0', '1.63.0', 'shared')
    explicit = fake(tmp_path / 'explicit', '1.50.0', 'explicit')
    repo = tmp_path / 'repo'; repo.mkdir()
    rc, out = run(repo, {'OWN_PLAYWRIGHT_HOME': str(home), 'OWN_PLAYWRIGHT_DIR': str(explicit)}, '1.63.0')
    assert rc == 2 and '1.50.0' in out
    rc, out = run(repo, {'OWN_PLAYWRIGHT_HOME': str(home), 'OWN_PLAYWRIGHT_DIR': str(explicit)})
    assert rc == 0 and out['marker'] == 'explicit' and out['source'] == 'explicit'


def test_without_a_pin_only_the_project_is_used(tmp_path):
    home = tmp_path / 'home'; fake(home / '1.63.0', '1.63.0', 'shared')
    repo = tmp_path / 'repo'; repo.mkdir()
    rc, out = run(repo, {'OWN_PLAYWRIGHT_HOME': str(home)})
    assert rc == 2
    repo2 = fake(tmp_path / 'repo2', '1.70.0', 'local')
    rc, out = run(repo2, {'OWN_PLAYWRIGHT_HOME': str(home)})
    assert rc == 0 and out['marker'] == 'local'


def test_project_root_override_is_used_for_snapshots_without_node_modules(tmp_path):
    home = tmp_path / 'home'; fake(home / '1.63.0', '1.63.0', 'shared')
    checkout = fake(tmp_path / 'checkout', '1.63.0', 'checkout')
    snapshot = tmp_path / 'snapshot'; snapshot.mkdir()
    rc, out = run(snapshot, {'OWN_PLAYWRIGHT_HOME': str(home), 'INFOARCH_PLAYWRIGHT_DIR': str(checkout)}, '1.63.0')
    assert rc == 0 and out['marker'] == 'checkout'


def companion(cwd, runtime_dir, env):
    js = ('import {locateCompanion} from ' + json.dumps(HELPER.as_uri()) + ';'
          'console.log(JSON.stringify(locateCompanion("axe-core/axe.min.js",{cwd:process.cwd(),runtimeDir:process.argv.at(-1)||null})));')
    base = {k: v for k, v in os.environ.items() if k not in ('NODE_PATH', 'INFOARCH_PLAYWRIGHT_DIR')}
    proc = subprocess.run([NODE, '--input-type=module', '-e', js, str(runtime_dir or '')], cwd=cwd,
                          env={**base, **env}, capture_output=True, text=True)
    assert proc.returncode == 0, proc.stderr
    return json.loads(proc.stdout)


def axe(dir):
    path = dir / 'node_modules/axe-core/axe.min.js'
    path.parent.mkdir(parents=True); path.write_text('//axe')
    return str(path)


def test_axe_comes_from_the_project_then_the_runtime_never_node_path(tmp_path):
    runtime = tmp_path / 'runtime'; in_runtime = axe(runtime)
    other = tmp_path / 'other_repo'; axe(other)
    repo = tmp_path / 'repo'; repo.mkdir()
    assert companion(repo, runtime, {'NODE_PATH': str(other / 'node_modules')}) == in_runtime
    assert companion(repo, None, {'NODE_PATH': str(other / 'node_modules')}) is None
    in_repo = axe(repo)
    assert companion(repo, runtime, {}) == in_repo
