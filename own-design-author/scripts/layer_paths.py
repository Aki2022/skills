"""Where a repository keeps its layers: the single reading of references/layer-paths.json.

Every command that touches positioning, infoarch, design, the approval ledger or the judgment record asks this
module for the path instead of spelling it. Every road goes through `guard`: a repository that already moved to a
layout this version does not read (a foreign root exists) stops the command with LAYER_LAYOUT_FOREIGN_ROOT, so an
old skills checkout can never read around a migrated repository and report "0 found".

The Node twin is layer_paths.mjs; a test pins the two together.
"""
from __future__ import annotations

import json
from pathlib import Path, PurePosixPath

LAYOUT_FILE = Path(__file__).resolve().parents[1] / 'references' / 'layer-paths.json'
FOREIGN_CODE = 'LAYER_LAYOUT_FOREIGN_ROOT'


class LayerLayoutError(ValueError):
    """The repository's layout is not the one this version of the skills reads."""


def _load():
    data = json.loads(LAYOUT_FILE.read_text(encoding='utf-8'))
    if data.get('schema_version') != 1 or not isinstance(data.get('entries'), dict) \
            or not isinstance(data.get('foreign_roots'), list) or not data['foreign_roots']:
        raise LayerLayoutError('layer-paths.json is malformed')
    return data


LAYOUT = _load()
ENTRIES = LAYOUT['entries']
FOREIGN_ROOTS = tuple(LAYOUT['foreign_roots'])
LAYERS_ROOT = ENTRIES['layers']


def _foreign_message(root):
    text = (f'{FOREIGN_CODE}: {root}/ がある。この repo はこの版の skills が読む置き場所（{LAYERS_ROOT}/）と違う並びに移っている。'
            f'ここで止める（{root}/ を読み飛ばして 0 件にしないため）。skills をこの repo の並びに合う版へ揃える。'
            f' / {root}/ exists: this repository follows a layout this version of the skills does not read; stopping '
            f'instead of reading around it. Update the skills to the version that matches it.')
    hint = LAYOUT.get('hint') or ''
    return f'{text} {hint}'.strip()


def guard(repo):
    """Raise LayerLayoutError if `repo` has a foreign root (a real directory, not a file or a link to nowhere)."""
    base = Path(repo)
    for name in FOREIGN_ROOTS:
        if (base / name).is_dir():
            raise LayerLayoutError(_foreign_message(name))


def rel(key, *rest):
    """The repository-relative POSIX path of entry `key` joined with `rest` (no filesystem access, no guard)."""
    base = PurePosixPath(ENTRIES[key])
    for part in rest:
        parts = PurePosixPath(part)
        if parts.is_absolute() or '..' in parts.parts:
            raise ValueError(f'path escapes the layer entry {key}: {part}')
        base = base / parts
    return str(base)


def path(repo, key, *rest):
    """`repo`/entry/rest, after the guard."""
    guard(repo)
    return Path(repo) / rel(key, *rest)


def repo_from_resolved(manifest):
    """The repository root a resolved manifest path belongs to, from the path text alone (the file may be missing).

    A manifest under a foreign root, or in a repository that has one, stops with the guard message."""
    full = Path(manifest).resolve()
    want = PurePosixPath(ENTRIES['resolved']).parts
    have = full.parent.parts
    if len(have) > len(want) and have[len(have) - len(want):] == want:
        root = Path(*have[:len(have) - len(want)])
        if root.name in FOREIGN_ROOTS and not (root / '.git').exists():  # <repo>/<foreign root>/…/.resolved/x.json
            raise LayerLayoutError(_foreign_message(root.name))
        guard(root)
        return root
    for name in FOREIGN_ROOTS:
        if name in full.parts:
            raise LayerLayoutError(_foreign_message(name))
    raise ValueError(f'manifest must be {ENTRIES["resolved"]}/<media>.json')
