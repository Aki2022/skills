"""Real git push through an additive hook harness; the real scanner is retained."""
import hashlib
import os
from pathlib import Path
import shlex
import shutil
import subprocess
import sys

import pytest
from test_core import SKILL, project, approve, check


@pytest.mark.skipif(not os.environ.get('DESIGN_TEST_VIBE_ROOT'), reason='real vibe-guard fixture run required')
def test_push_and_mise_share_ship_gate_without_bypassing_scanner(project, tmp_path):
    root, catalogs = project
    guard = Path(os.environ['DESIGN_TEST_VIBE_ROOT']).resolve()
    global_config = Path(os.environ.get('GIT_CONFIG_GLOBAL', str(Path.home() / '.gitconfig')))
    before = hashlib.sha256(global_config.read_bytes()).hexdigest()
    harness = tmp_path / 'guard'
    (harness / 'bin').mkdir(parents=True)
    shutil.copytree(guard / 'git-hooks', harness / 'git-hooks')
    hook = harness / 'git-hooks/pre-push'
    lines = hook.read_text().splitlines()
    assert any(line.startswith('ROOT=') for line in lines)
    hook.write_text('\n'.join('ROOT=' + shlex.quote(str(harness)) if line.startswith('ROOT=') else line for line in lines) + '\n')
    scanner_marker = tmp_path / 'scanner-called'
    wrapper = harness / 'bin/vibe-guard'
    wrapper.write_text('#!/bin/sh\nset -eu\nprintf called >> ' + shlex.quote(str(scanner_marker)) +
                       '\nexec ' + shlex.quote(str(guard / 'bin/vibe-guard')) + ' "$@"\n')
    wrapper.chmod(0o755)
    extension = harness / 'git-hooks/pre-push.d/60-design-ship'
    shutil.copy2(SKILL / 'hooks/pre-push.d/60-design-ship', extension)
    # This temporary config retains the user's config and real scanner, then adds the gate.
    # The user's config, global hooks, and repository hook settings are never edited.
    config = tmp_path / 'gitconfig'
    config.write_text('[include]\n path = ' + str(global_config) + '\n[core]\n hooksPath = ' + str(harness / 'git-hooks') + '\n')
    fixture_l0 = tmp_path / 'fixture-l0'
    fixture_l0.write_text('#!' + sys.executable + '\nimport subprocess,sys\nraise SystemExit(subprocess.call(' +
                         repr([str(SKILL / 'scripts/design-author')]) + ' + sys.argv[1:] + ' +
                         repr(['--catalog-dir', str(catalogs)]) + '))\n')
    fixture_l0.chmod(0o755)
    env = dict(os.environ, GIT_CONFIG_GLOBAL=str(config), VIBE_GUARD_ROOT=str(guard),
               DESIGN_AUTHOR_COMMAND=str(fixture_l0), DESIGN_AUTHOR_PYTHON=sys.executable,
               PATH=str(SKILL / 'scripts') + os.pathsep + os.environ['PATH'])
    def git(*args, allowed=(0,)):
        run = subprocess.run(['git', *args], cwd=root, env=env, capture_output=True, text=True)
        assert run.returncode in allowed, run.stderr
        return run
    (root / '.vibe-guard').mkdir()
    shutil.copy2(SKILL / 'templates/design-ship.enabled', root / '.vibe-guard/design-ship.enabled')
    shutil.copy2(SKILL / 'templates/mise.design.toml', root / '.mise.toml')
    remote = tmp_path / 'remote.git'
    subprocess.run(['git','init','--bare','-q',str(remote)], check=True, env=env)
    git('config','user.name','Fixture Owner')
    git('config','user.email','fixture@example.org')
    git('checkout','-b','fixture-design')
    git('add','.')
    git('commit','-qm','fixture with proposed must')
    git('remote','add','fixture',str(remote))
    blocked = git('push','fixture','HEAD:refs/heads/main',allowed=(1,))
    assert 'ship gate' in blocked.stdout + blocked.stderr
    assert scanner_marker.exists()
    assert not (remote / 'refs/heads/main').exists()
    # Real mise invocation uses the same PATH command, not an inline imitation.
    mise = os.environ.get('DESIGN_TEST_MISE')
    if mise:
        mise_env = dict(env, MISE_TRUSTED_CONFIG_PATHS=str(root))
        run = subprocess.run([mise,'run','design:ship'], cwd=root, env=mise_env, capture_output=True)
        assert run.returncode != 0
    approve(project)
    git('add','design/approvals.yaml')
    git('commit','-qm','fixture owner approval')
    git('push','fixture','HEAD:refs/heads/main')
    assert (remote / 'refs/heads/main').read_text().strip() == git('rev-parse','HEAD').stdout.strip()
    if mise:
        run = subprocess.run([mise,'run','design:ship'], cwd=root, env=mise_env, capture_output=True)
        assert run.returncode == 0, run.stderr
    assert scanner_marker.read_text().count('called') >= 2
    assert hashlib.sha256(global_config.read_bytes()).hexdigest() == before
    assert not git('config','--local','--get','core.hooksPath',allowed=(1,)).stdout
