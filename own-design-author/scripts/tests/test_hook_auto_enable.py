"""The pre-push extension runs design-ship in every repository that has a design system, marker or not."""
import os
from pathlib import Path
import subprocess

import pytest

HOOK = Path(__file__).resolve().parents[2] / 'hooks/pre-push.d/60-design-ship'


@pytest.fixture
def repo(tmp_path):
    root = tmp_path / 'repo'
    root.mkdir()
    subprocess.run(['git', 'init', '-q', str(root)], check=True)
    fake = tmp_path / 'design-ship'
    fake.write_text('#!/bin/sh\necho "called $*" > "$FAKE_LOG"\nexit "${FAKE_EXIT:-0}"\n')
    fake.chmod(0o755)
    refs = tmp_path / 'refs'
    refs.write_text('refs/heads/x 0 refs/heads/x 0\n')
    return root, fake, refs, tmp_path / 'called.log'


def run(root, fake, refs, log, **env):
    return subprocess.run(['sh', str(HOOK), str(refs)], cwd=root, capture_output=True, text=True,
                          env={**os.environ, 'DESIGN_SHIP_COMMAND': str(fake), 'FAKE_LOG': str(log), **env})


def test_design_system_without_marker_runs_ship(repo):
    root, fake, refs, log = repo
    (root / 'design').mkdir()
    (root / 'design/design.yaml').write_text('{}')
    result = run(root, fake, refs, log)
    assert result.returncode == 0, result.stderr
    assert log.read_text().startswith('called --repo')


def test_ship_failure_blocks_the_push(repo):
    root, fake, refs, log = repo
    (root / 'design').mkdir()
    (root / 'design/design.yaml').write_text('{}')
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
