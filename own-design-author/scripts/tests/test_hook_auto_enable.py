"""The pre-push extension runs design-ship in every repository that has a design system, marker or not."""
import os
from pathlib import Path
import subprocess
import sys

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from layer_testkit import FOREIGN_CODE, ROOT, moved_repo
import layer_paths as lp  # noqa: E402

HOOK = Path(__file__).resolve().parents[2] / 'hooks/pre-push.d/60-design-ship'
REAL_SHIP = Path(__file__).resolve().parents[1] / 'design-ship'


@pytest.fixture
def repo(tmp_path):
    root = tmp_path / 'repo'
    root.mkdir()
    subprocess.run(['git', 'init', '-q', str(root)], check=True)
    fake = tmp_path / 'design-ship'
    # The presence probe is answered by the real design-ship (it owns the layer rule); a push check is only logged.
    fake.write_text('#!/bin/sh\ncase " $* " in *" --probe "*) exec "$REAL_SHIP" "$@";; esac\n'
                    'echo "called $*" > "$FAKE_LOG"\nexit "${FAKE_EXIT:-0}"\n')
    fake.chmod(0o755)
    refs = tmp_path / 'refs'
    refs.write_text('refs/heads/x 0 refs/heads/x 0\n')
    return root, fake, refs, tmp_path / 'called.log'


def run(root, fake, refs, log, **env):
    return subprocess.run(['sh', str(HOOK), str(refs)], cwd=root, capture_output=True, text=True,
                          env={**os.environ, 'DESIGN_SHIP_COMMAND': str(fake), 'FAKE_LOG': str(log),
                               'REAL_SHIP': str(REAL_SHIP), 'DESIGN_AUTHOR_PYTHON': sys.executable, **env})


def test_design_system_without_marker_runs_ship(repo):
    root, fake, refs, log = repo
    config = lp.path(root, 'design_config')
    config.parent.mkdir(parents=True)
    config.write_text('{}')
    result = run(root, fake, refs, log)
    assert result.returncode == 0, result.stderr
    assert log.read_text().startswith('called --repo')


def test_ship_failure_blocks_the_push(repo):
    root, fake, refs, log = repo
    config = lp.path(root, 'design_config')
    config.parent.mkdir(parents=True)
    config.write_text('{}')
    assert run(root, fake, refs, log, FAKE_EXIT='2').returncode == 2


def test_repository_without_design_system_is_not_checked(repo):
    root, fake, refs, log = repo
    assert run(root, fake, refs, log).returncode == 0
    assert not log.exists()


def test_broken_marker_still_stops(repo):
    root, fake, refs, log = repo
    (root / '.vibe-guard').mkdir()
    (root / '.vibe-guard/design-ship.enabled').write_text('0')
    assert run(root, fake, refs, log).returncode == 2
    assert not log.exists()


@pytest.mark.parametrize('kind', ['foreign_only', 'both'])
def test_moved_repository_stops_the_push_instead_of_skipping(tmp_path, kind):
    root = moved_repo(tmp_path, kind)
    fake = tmp_path / 'design-ship'
    fake.write_text('#!/bin/sh\ncase " $* " in *" --probe "*) exec "$REAL_SHIP" "$@";; esac\n'
                    'echo "called $*" > "$FAKE_LOG"\nexit 0\n')
    fake.chmod(0o755)
    refs = tmp_path / 'refs'
    refs.write_text('refs/heads/x 0 refs/heads/x 0\n')
    log = tmp_path / 'called.log'
    result = run(root, fake, refs, log)
    assert result.returncode != 0, 'a repository this version does not read passed the hook: ' + result.stderr
    assert FOREIGN_CODE in result.stderr
    assert not log.exists()


def test_the_hook_does_not_spell_the_design_config_path():
    text = HOOK.read_text()
    assert ROOT + '/' not in text, 'the hook must ask design-ship --probe, not test a literal layer path'
    assert '--probe' in text


def test_a_failing_probe_is_a_stop_not_a_skip(repo):
    root, _, refs, log = repo
    broken = log.parent / 'broken-ship'
    broken.write_text('#!/bin/sh\nexit 7\n')
    broken.chmod(0o755)
    result = run(root, broken, refs, log)
    assert result.returncode == 2
    assert 'design layout probe failed; shipment stopped' in result.stderr
