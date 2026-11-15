"""Shared shipping base for design-ship and the IA ship gate (ADR-20261115-design-ia-two-skills-shared-base, biz_ops).

Git in a hook-free environment, the pre-push refs, the bases a push must be merged into (the base as it is now and
what the same push sets it to), the synthetic predicted merge, and a run-time check that the pre-push dispatcher
propagates a failing extension. Both gates use these, so a safeguard added here protects both.
"""
from __future__ import annotations

import functools
import os
from pathlib import Path
import shlex
import subprocess
import tempfile


@functools.cache
def _local_git_env():
    return frozenset(subprocess.run(['git', 'rev-parse', '--local-env-vars'], capture_output=True, text=True, check=True).stdout.split())


def hook_free_env(base=None):
    """Drop the repository variables git exports to hooks, so every command targets its own checkout."""
    return {key: value for key, value in (os.environ if base is None else base).items() if key not in _local_git_env()}


def git(root, *args, input=None):
    result = subprocess.run(['git', '-C', str(root), *args], input=input, env=hook_free_env(), capture_output=True, text=True)
    if result.returncode:
        raise ValueError('git ' + args[0] + ' failed')
    return result.stdout.strip()


def read_push_refs(root, refs_file):
    """-> ([(sha, remote ref, remote sha)] for every non-deleting update, {remote ref: sha})."""
    if refs_file is None:
        raise ValueError('pre-push requires the preserved refs file')
    records = Path(refs_file).read_text().splitlines()
    if not records:
        raise ValueError('zero push refs')
    updates, pushed = [], {}
    for row in records:
        parts = row.split()
        if len(parts) != 4:
            raise ValueError('invalid push ref record')
        _, sha, remote_ref, remote_sha = parts
        if set(sha) == {'0'}:
            continue  # a deleted ref ships nothing
        if git(root, 'rev-parse', '--verify', sha + '^{commit}') != sha:
            raise ValueError('push source must be a commit')
        updates.append((sha, remote_ref, remote_sha))
        pushed[remote_ref] = sha
    return updates, pushed


def merge_bases(root, base_ref, pushed):
    """The merge base as it is now, plus what this same push sets the branch behind it to."""
    bases = [git(root, 'rev-parse', '--verify', base_ref + '^{commit}')]
    try:
        full = git(root, 'rev-parse', '--symbolic-full-name', base_ref)
    except ValueError:
        full = ''
    if full.startswith('refs/remotes/') and full.count('/') >= 3:
        full = 'refs/heads/' + full.split('/', 3)[3]
    if full.startswith('refs/heads/') and pushed.get(full) and pushed[full] not in bases:
        bases.append(pushed[full])
    return bases


def predicted_commit(root, base, source, *, name, message, email=None):
    """A synthetic merge of `source` into `base`; never referenced or pushed. A conflict stops the push."""
    tree = git(root, 'merge-tree', '--write-tree', base, source).splitlines()[0]
    identity = ['-c', 'user.name=' + name] + (['-c', 'user.email=' + email] if email else [])
    return git(root, *identity, 'commit-tree', tree, '-p', source, '-p', base, input=message + '\n')


def dispatcher_propagates(text):
    """Run a copy of the dispatcher against probe extensions: a failing one must make the whole hook fail.

    The copy is re-rooted (its ROOT= line) into a temporary vibe-guard-shaped directory with a stub scanner, so
    neither the real scanner nor the installed extensions run. A dispatcher that also fails with only passing
    probes is rejected too, so a hook that always exits nonzero does not read as propagating.
    """
    lines = text.splitlines()
    if not any(line.startswith('ROOT=') for line in lines):
        raise ValueError('cannot verify that the pre-push dispatcher propagates an extension exit (no ROOT= line)')
    with tempfile.TemporaryDirectory(prefix='ship-doctor-') as temporary:
        sandbox = Path(temporary).resolve()
        hooks = sandbox / 'git-hooks'
        (hooks / 'pre-push.d').mkdir(parents=True)
        (sandbox / 'bin').mkdir()
        calls = sandbox / 'calls'

        def script(path, body):
            path.write_text('#!/bin/sh\n' + body)
            path.chmod(0o755)

        script(sandbox / 'bin/vibe-guard', 'cat >/dev/null\n')
        copy = hooks / 'pre-push'
        copy.write_text('\n'.join('ROOT=' + shlex.quote(str(sandbox)) if line.startswith('ROOT=') else line
                                  for line in lines) + '\n')
        copy.chmod(0o755)
        env = hook_free_env()
        env['PATH'] = str(sandbox / 'bin') + os.pathsep + env.get('PATH', '')

        def run(first):
            script(hooks / 'pre-push.d/00-probe', f'echo 00 >> {shlex.quote(str(calls))}\nexit {first}\n')
            script(hooks / 'pre-push.d/99-probe', f'echo 99 >> {shlex.quote(str(calls))}\nexit 0\n')
            calls.unlink(missing_ok=True)
            code = subprocess.run([str(copy), 'probe', 'probe'], input='', capture_output=True, text=True,
                                  env=env, cwd=sandbox, timeout=60).returncode
            return code, calls.read_text().split() if calls.exists() else []

        code, called = run(0)
        if code != 0 or '00' not in called:
            raise ValueError('pre-push dispatcher does not run passing extensions with exit 0; its exit cannot be trusted')
        code, called = run(3)
        if code == 0 or '00' not in called:
            raise ValueError('pre-push dispatcher does not propagate a failing extension exit (a gate could not stop a push)')
