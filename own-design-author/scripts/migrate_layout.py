"""One-time move of a repository's layers from the old `design/` tree to the layout in references/layer-paths.json
(`ux/positioning`, `ux/infoarch`, `ux/design`; the approval ledger and the judgment record directly under `ux/`).

The move changes places only. Every file keeps its bytes; tracked files are renamed in git, ignored or untracked ones
(generated manifests) are carried along; a layer the repository does not have is not created; text that names the old
place (prose in DESIGN.md, rule targets) is reported and left alone because approvals bind to it. Dry run by default;
a failure part-way puts everything back.

This file is the only place that spells the old layout: it is the record of what is being migrated away from, so it is
exempt from the single-source scan (and goes away with the old layout).
"""
from __future__ import annotations

import os
import re
import subprocess
from pathlib import Path, PurePosixPath

import layer_paths

OLD_ROOT = 'design'
OLD_INFOARCH = OLD_ROOT + '/info_arch'
OLD_POSITIONING = OLD_INFOARCH + '/info_arch.yaml'
OLD_LEDGER = OLD_ROOT + '/approvals.yaml'
OLD_JUDGMENTS = OLD_ROOT + '/judgments.json'
PROSE_SUFFIXES = {'.md', '.yaml', '.yml'}
OLD_ASSETS = OLD_ROOT + '/assets/'
OLD_ASSETS_RE = re.compile(r'(?<![A-Za-z0-9_.-])' + re.escape(OLD_ASSETS))


class MigrationError(ValueError):
    """The move cannot start, or failed and was rolled back."""


def map_path(old):
    """The new repository-relative path of a file that lives at `old` in the old tree."""
    parts = PurePosixPath(old).parts
    if not parts or parts[0] != OLD_ROOT or len(parts) < 2:
        raise ValueError(f'not under the old layer root: {old}')
    text = '/'.join(parts)
    if text == OLD_POSITIONING:
        return layer_paths.rel('positioning_file')
    if text.startswith(OLD_INFOARCH + '/'):
        return layer_paths.rel('infoarch', *parts[2:])
    if text == OLD_LEDGER:
        return layer_paths.rel('ledger')
    if text == OLD_JUDGMENTS:
        return layer_paths.rel('judgments')
    return layer_paths.rel('design_dir', *parts[1:])


def _git(repo, *args, check=True):
    run = subprocess.run(['git', '-C', str(repo), *args], capture_output=True, text=True)
    if check and run.returncode:
        raise MigrationError(f'git {args[0]} failed: {run.stderr.strip()}')
    return run


def _tracked(repo):
    out = _git(repo, 'ls-files', '-z', '--', OLD_ROOT).stdout
    return {p for p in out.split('\0') if p}


def _files(repo):
    base = Path(repo) / OLD_ROOT
    found = []
    for path in sorted(base.rglob('*')):
        if path.is_symlink():
            raise MigrationError(f'a symlink inside the old tree cannot be moved safely: {path.relative_to(repo).as_posix()}')
        if path.is_file():
            found.append(path.relative_to(repo).as_posix())
    return found


def _check_start(repo):
    repo = Path(repo)
    if _git(repo, 'rev-parse', '--show-toplevel', check=False).returncode:
        raise MigrationError('migrate-layout needs a git repository')
    old = repo / OLD_ROOT
    if not old.is_dir() or old.is_symlink():
        raise MigrationError(f'nothing to migrate: {OLD_ROOT}/ is not a directory here (already migrated, or this repository has no layers)')
    if (repo / layer_paths.LAYERS_ROOT).exists():
        raise MigrationError(f'{layer_paths.LAYERS_ROOT}/ already exists next to {OLD_ROOT}/; resolve by hand, nothing was moved')


def _prose_mentions(repo, files):
    found = []
    pattern = re.compile(r'(?<![A-Za-z0-9_.-])' + re.escape(OLD_ROOT) + r'/')
    for old in files:
        if PurePosixPath(old).suffix not in PROSE_SUFFIXES and not old.endswith('NOTICE'):
            continue
        try:
            text = (Path(repo) / old).read_text(encoding='utf-8')
        except (OSError, UnicodeDecodeError):
            continue
        lines = sum(1 for line in text.splitlines() if pattern.search(line))
        if lines:
            found.append({'file': map_path(old), 'lines': lines})
    return found


def _gitignore_edits(repo):
    """New text of the root .gitignore (or None) and the patterns left alone because they use wildcards."""
    path = Path(repo) / '.gitignore'
    if not path.is_file():
        return None, []
    changed, skipped, out = False, [], []
    for line in path.read_text(encoding='utf-8').split('\n'):
        stripped = line.strip()
        body = stripped.lstrip('/').rstrip('/')
        if stripped and not stripped.startswith('#') and (body == OLD_ROOT or body.startswith(OLD_ROOT + '/')):
            if any(ch in body for ch in '*?[!'):
                skipped.append(stripped)
                out.append(line)
                continue
            new = map_path(body) if body != OLD_ROOT else layer_paths.LAYERS_ROOT
            lead = '/' if stripped.startswith('/') else ''
            tail = '/' if stripped.endswith('/') else ''
            out.append(lead + new + tail)
            changed = True
        else:
            out.append(line)
    return ('\n'.join(out) if changed else None), skipped


def _rule_target_rewrites(repo, moves):
    """Rule files whose asset targets name the old assets folder: they must follow the files (the rule is approved in a
    layout-independent form, see layer_paths.identity, so the approval stays)."""
    found = []
    checks_old = OLD_ROOT + '/checks/'
    for move in moves:
        if not move['from'].startswith(checks_old):
            continue
        try:
            text = (Path(repo) / move['from']).read_text(encoding='utf-8')
        except (OSError, UnicodeDecodeError):
            continue
        new_text, count = OLD_ASSETS_RE.subn(layer_paths.rel('assets') + '/', text)
        if count:
            found.append({'file': move['to'], 'count': count, 'old_text': text, 'new_text': new_text})
    return found


def plan(repo, apply_=False):
    """What would move (nothing is changed)."""
    _check_start(repo)
    files = _files(repo)
    tracked = _tracked(repo)
    moves = []
    for old in files:
        new = map_path(old)
        if (Path(repo) / new).exists():
            raise MigrationError(f'the destination already exists: {new}')
        moves.append({'from': old, 'to': new, 'tracked': old in tracked})
    if not moves:
        raise MigrationError(f'nothing to migrate: {OLD_ROOT}/ holds no files')
    new_ignore, skipped = _gitignore_edits(repo)
    rewrites = _rule_target_rewrites(repo, moves)
    destinations = sorted({m['to'].rsplit('/', 1)[0] for m in moves})
    return {'applied': False, 'moves': moves, 'count': len(moves),
            'tracked': sum(m['tracked'] for m in moves), 'carried_untracked': sum(not m['tracked'] for m in moves),
            'gitignore': 'update' if new_ignore is not None else 'unchanged',
            'gitignore_wildcards_left_alone': skipped,
            'prose_mentions': _prose_mentions(repo, files), 'destinations': destinations,
            'rule_target_rewrites': [{'file': r['file'], 'count': r['count']} for r in rewrites]}


def _move(repo, src, dst, tracked):
    repo = Path(repo)
    (repo / dst).parent.mkdir(parents=True, exist_ok=True)
    if tracked:
        _git(repo, 'mv', '--', src, dst)
    else:
        os.rename(repo / src, repo / dst)


def _remove_empty_dirs(base):
    for path in sorted(Path(base).rglob('*'), key=lambda p: len(p.parts), reverse=True):
        if path.is_dir() and not any(path.iterdir()):
            path.rmdir()
    if Path(base).is_dir() and not any(Path(base).iterdir()):
        Path(base).rmdir()


def apply(repo):
    """Move everything; on any failure move it all back and raise MigrationError."""
    repo = Path(repo)
    result = plan(repo)
    rewrites = _rule_target_rewrites(repo, result['moves'])
    ignore_path = repo / '.gitignore'
    old_ignore = ignore_path.read_text(encoding='utf-8') if ignore_path.is_file() else None
    new_ignore, _ = _gitignore_edits(repo)
    done = []

    try:
        for move in result['moves']:
            _move(repo, move['from'], move['to'], move['tracked'])
            done.append(move)
        if new_ignore is not None:
            ignore_path.write_text(new_ignore, encoding='utf-8')
        for rewrite in rewrites:
            (repo / rewrite['file']).write_text(rewrite['new_text'], encoding='utf-8')
        _remove_empty_dirs(repo / OLD_ROOT)
        left = repo / OLD_ROOT
        if left.exists():
            raise MigrationError(f'{OLD_ROOT}/ still holds entries after the move: {sorted(p.name for p in left.iterdir())}')
    except (OSError, MigrationError) as exc:
        for rewrite in rewrites:  # put the rule text back before the files go back
            target = repo / rewrite['file']
            if target.exists():
                target.write_text(rewrite['old_text'], encoding='utf-8')
        for move in reversed(done):
            try:
                _move(repo, move['to'], move['from'], move['tracked'])
            except (OSError, MigrationError):  # keep going: restore as much as possible
                pass
        if old_ignore is not None:
            ignore_path.write_text(old_ignore, encoding='utf-8')
        new_root = repo / layer_paths.LAYERS_ROOT
        if new_root.exists():
            _remove_empty_dirs(new_root)
        raise MigrationError(f'migration failed and was rolled back: {exc}') from exc
    result['applied'] = True
    return result
