#!/usr/bin/env python3
"""Safely synchronize one per-skill alias directory with the canonical root."""

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent))
from skill_catalog import load_skill_catalog


IGNORED_CANONICAL_DIRS = {".git", "docs", "node_modules"}


def lexical(path: Path) -> Path:
    return Path(os.path.abspath(os.path.expanduser(str(path))))


def resolved(path: Path) -> Path:
    return Path(os.path.realpath(path))


def canonical_names(root: Path) -> set[str]:
    catalog = load_skill_catalog(root)
    if catalog.errors:
        raise ValueError("; ".join(catalog.errors))
    return catalog.active


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Synchronize direct per-skill symlinks from one canonical root"
    )
    parser.add_argument("--canonical", required=True, type=Path)
    parser.add_argument("--alias-root", required=True, type=Path)
    parser.add_argument("--apply", action="store_true", help="apply the reported changes")
    parser.add_argument(
        "--prune-stale",
        action="store_true",
        help="remove broken unregistered symlinks; valid foreign links are protected",
    )
    parser.add_argument(
        "--ignore-entry",
        action="append",
        default=[],
        help="exact direct-child name to preserve outside canonical sync (repeatable)",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    canonical = lexical(args.canonical)
    alias_root = lexical(args.alias_root)

    if canonical.is_symlink() or not canonical.is_dir():
        print(f"ERROR canonical root must be a regular directory: {canonical}")
        return 2
    if alias_root.is_symlink() or not alias_root.is_dir():
        print(f"ERROR alias root must be a regular directory: {alias_root}")
        return 2

    catalog = load_skill_catalog(canonical)
    for error in catalog.errors:
        print(f"ERROR canonical catalog: {error}")
    if catalog.errors:
        return 2
    names = catalog.active
    if not names:
        print(f"ERROR canonical root has no skill directories: {canonical}")
        return 2
    if catalog.retired_physical:
        retired = ", ".join(sorted(catalog.retired_physical))
        print(f"NOTE retired physical skills excluded from sync: {retired}")

    ignored = set(args.ignore_entry)
    if any(not name or Path(name).name != name or name in {".", ".."} for name in ignored):
        print("ERROR --ignore-entry must be an exact direct-child name")
        return 2
    overlap = names & ignored
    if overlap:
        print(f"ERROR canonical skills cannot be ignored: {', '.join(sorted(overlap))}")
        return 2

    creates: list[tuple[Path, Path]] = []
    prunes: list[Path] = []
    stale: list[Path] = []
    conflicts: list[str] = []

    for name in sorted(names):
        alias = alias_root / name
        target = canonical / name
        if not alias.exists() and not alias.is_symlink():
            creates.append((alias, target))
        elif not alias.is_symlink():
            conflicts.append(f"{alias}: existing entry is not a symlink")
        elif not alias.exists():
            conflicts.append(f"{alias}: expected skill link is broken")
        elif resolved(alias) != resolved(target):
            conflicts.append(f"{alias}: target is not {target}")

    for child in sorted(alias_root.iterdir(), key=lambda path: path.name):
        if child.name in names:
            continue
        if child.name in ignored:
            print(f"IGNORE {child}")
            continue
        if child.is_symlink() and not child.exists():
            if args.prune_stale:
                prunes.append(child)
            else:
                stale.append(child)
        elif child.is_symlink():
            conflicts.append(f"{child}: valid unregistered symlink is protected")
        else:
            conflicts.append(f"{child}: unregistered non-symlink entry is protected")

    for alias, target in creates:
        print(f"CREATE {alias} -> {target}")
    for path in prunes:
        print(f"PRUNE {path} (broken unregistered symlink)")
    for path in stale:
        print(f"STALE {path} (use --prune-stale to remove)")
    for message in conflicts:
        print(f"CONFLICT {message}")

    if conflicts:
        print("RESULT: CONFLICT (no changes applied)")
        return 2

    changes = bool(creates or prunes or stale)
    if not args.apply:
        print("RESULT: DRIFT" if changes else "RESULT: OK (no changes)")
        return 1 if changes else 0

    for alias, target in creates:
        os.symlink(target, alias, target_is_directory=True)
    for path in prunes:
        path.unlink()

    if stale:
        print("RESULT: DRIFT (stale links retained)")
        return 1
    print("RESULT: OK" if creates or prunes else "RESULT: OK (no changes)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
