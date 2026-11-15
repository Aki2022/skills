"""Shared shipping base for design-ship and the IA ship gate (ADR-20261115-design-ia-two-skills-shared-base, biz_ops).

The IA gate already checked the predicted merge into what the same push sets the base branch to, and verified by
running it that the pre-push dispatcher propagates a failing extension. Moving those into one base must keep them
for IA and give them to design-ship too.
"""
import subprocess
from pathlib import Path

import pytest

import test_core  # noqa: F401 -- scripts/ on sys.path
from test_maintenance import commit as fixture_commit  # the fixture identity lives there
import ship
import ship_base


def git(root, *args):
    if args[:1] == ('commit',):
        return fixture_commit(root, args[-1])
    return subprocess.run(['git', '-C', str(root), *args], check=True, capture_output=True, text=True).stdout.strip()


@pytest.fixture
def repo(tmp_path):
    root = tmp_path / 'repo'
    root.mkdir()
    git(root, 'init', '-q', '-b', 'main')
    (root / 'a.txt').write_text('a\n')
    git(root, 'add', '-A'); git(root, 'commit', '-qm', 'base')
    git(root, 'update-ref', 'refs/remotes/origin/main', 'HEAD')
    git(root, 'checkout', '-qb', 'feature')
    (root / 'b.txt').write_text('b\n'); git(root, 'add', '-A'); git(root, 'commit', '-qm', 'feature')
    feature = git(root, 'rev-parse', 'HEAD')
    git(root, 'checkout', '-q', 'main')
    (root / 'c.txt').write_text('c\n'); git(root, 'add', '-A'); git(root, 'commit', '-qm', 'new main')
    new_main = git(root, 'rev-parse', 'HEAD')
    return root, feature, new_main


def refs(tmp_path, rows):
    path = tmp_path / 'refs'
    path.write_text(''.join(' '.join(r) + '\n' for r in rows))
    return path


def test_read_push_refs_skips_deletions_and_rejects_non_commits(repo, tmp_path):
    root, feature, new_main = repo
    zero = '0' * 40
    updates, pushed = ship_base.read_push_refs(root, refs(tmp_path, [
        ('refs/heads/feature', feature, 'refs/heads/feature', zero), ('(delete)', zero, 'refs/heads/old', feature)]))
    assert updates == [(feature, 'refs/heads/feature', zero)] and pushed == {'refs/heads/feature': feature}
    tree = git(root, 'rev-parse', 'HEAD^{tree}')
    with pytest.raises(ValueError):
        ship_base.read_push_refs(root, refs(tmp_path, [('x', tree, 'refs/heads/x', zero)]))


def test_merge_bases_include_what_the_same_push_sets_the_base_to(repo):
    root, feature, new_main = repo
    current = git(root, 'rev-parse', 'origin/main')
    assert ship_base.merge_bases(root, 'origin/main', {}) == [current]
    assert ship_base.merge_bases(root, 'origin/main', {'refs/heads/main': new_main}) == [current, new_main]


def test_design_pre_push_gates_the_same_push_base_too(repo, tmp_path, monkeypatch):
    root, feature, new_main = repo
    channel = {'merge_base': 'origin/main'}
    seen = []
    monkeypatch.setattr(ship, 'read_data', lambda path: {})
    monkeypatch.setattr(ship, 'validate_channels', lambda config: {'web': channel})
    monkeypatch.setattr(ship, 'validate_integration_gate', lambda config: None)

    def fake_gate(snapshot, name, row, catalogs, **kw):
        parents = git(snapshot, 'log', '-1', '--format=%P').split()
        seen.append(parents[-1])  # the base the source was merged into
        return {'exit_code': 0}
    monkeypatch.setattr(ship, 'gate', fake_gate)
    zero = '0' * 40
    result = ship.pre_push(root, {'web': channel}, tmp_path, refs(tmp_path, [
        ('refs/heads/main', new_main, 'refs/heads/main', git(root, 'rev-parse', 'origin/main')),
        ('refs/heads/feature', feature, 'refs/heads/feature', zero)]))
    assert result['exit_code'] == 0
    current = git(root, 'rev-parse', 'origin/main')
    assert seen.count(new_main) >= 1 and seen.count(current) >= 1  # the feature merged into both the old and the new main


SWALLOWING = '''#!/bin/sh
ROOT=/nowhere
for ext in "$ROOT/git-hooks/pre-push.d/"*; do "$ext" "$@" || true; done  # "$ext" pre-push.d
vibe-guard scan-push
exit 0
'''


def test_dispatcher_that_swallows_a_failing_extension_is_rejected():
    with pytest.raises(ValueError, match='propagate'):
        ship_base.dispatcher_propagates(SWALLOWING)


PROPAGATING = '''#!/bin/sh
ROOT=/nowhere
for ext in "$ROOT/git-hooks/pre-push.d/"*; do "$ext" "$@" || exit 1; done  # "$ext" pre-push.d
vibe-guard scan-push
exit 0
'''


@pytest.mark.parametrize('dispatcher,ok', [(SWALLOWING, False), (PROPAGATING, True)])
def test_design_doctor_runs_the_dispatcher(repo, dispatcher, ok, monkeypatch):
    """The string check alone passed a dispatcher that swallows a failing extension; design doctor now runs it."""
    import shutil
    root, _, _ = repo
    (root / '.vibe-guard').mkdir()
    (root / '.vibe-guard/design-ship.enabled').write_text('1\n')
    hooks = root / '.git/hooks'
    (hooks / 'pre-push.d').mkdir(parents=True, exist_ok=True)
    (hooks / 'pre-push').write_text(dispatcher)
    (hooks / 'pre-push').chmod(0o755)
    extension = hooks / 'pre-push.d/60-design-ship'
    shutil.copy2(ship.SCRIPT.parent / 'hooks/pre-push.d/60-design-ship', extension)
    real_git = ship.git
    # The runner's global hooksPath must not decide the fixture: point the hook lookup at the fixture's hook.
    monkeypatch.setattr(ship, 'git', lambda r, *a, **k: str(hooks / 'pre-push') if a == ('rev-parse', '--git-path', 'hooks/pre-push') else real_git(r, *a, **k))
    errors = ship.doctor(root, {'web': {}})['errors']
    assert any('propagate' in e for e in errors) != ok, errors
