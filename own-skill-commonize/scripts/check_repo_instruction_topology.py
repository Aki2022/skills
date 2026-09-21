#!/usr/bin/env python3
"""Validate repository-level AGENTS.md/CLAUDE.md topology without writes."""

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path


def lexical(path: Path) -> Path:
    """Normalize a path without resolving symlink targets."""

    return Path(os.path.abspath(os.path.expanduser(str(path))))


def resolved(path: Path) -> Path:
    """Resolve a path for symlink target comparison."""

    return Path(os.path.realpath(path))


def present(path: Path) -> bool:
    """Return true for existing paths and broken symlinks."""

    return path.exists() or path.is_symlink()


def entry_for(repo: Path, name: str) -> Path | None:
    """Find an immediate entry, including a case-only spelling mismatch."""

    try:
        entries = sorted(repo.iterdir(), key=lambda path: path.name)
    except OSError:
        return None
    for entry in entries:
        if entry.name == name:
            return entry
    for entry in entries:
        if entry.name.casefold() == name.casefold():
            return entry
    return None


def fail(failures: list[str], message: str) -> None:
    failures.append(message)
    print(f"FAIL {message}")


def check_repo(repo: Path, mode: str, failures: list[str]) -> None:
    """Check one repository in native or compatibility mode."""

    before = len(failures)
    if not repo.is_dir():
        fail(failures, f"repo {repo}: directory missing")
        return

    agents = entry_for(repo, "AGENTS.md") or (repo / "AGENTS.md")
    claude = entry_for(repo, "CLAUDE.md")

    if agents.name != "AGENTS.md" or agents.is_symlink() or not agents.is_file():
        fail(failures, f"repo {repo}: AGENTS.md must be a regular file")

    if mode == "native":
        if claude is not None and present(claude):
            fail(failures, f"repo {repo}: native mode requires CLAUDE.md absent")
    elif claude is None or claude.name != "CLAUDE.md" or not claude.is_symlink():
        fail(failures, f"repo {repo}: compat mode requires CLAUDE.md -> AGENTS.md")
    elif not claude.exists():
        fail(failures, f"repo {repo}: CLAUDE.md is a broken symlink")
    elif resolved(claude) != resolved(agents):
        fail(failures, f"repo {repo}: CLAUDE.md does not target AGENTS.md")

    for name in ("AGENTS.override.md", "CLAUDE.local.md"):
        shadow = entry_for(repo, name)
        if shadow is not None and present(shadow):
            fail(failures, f"repo {repo}: shadowing file is present: {name}")

    if len(failures) == before:
        print(f"OK repo {repo} mode={mode}")


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Check repository AGENTS.md and CLAUDE.md topology"
    )
    parser.add_argument("--repo", action="append", required=True, type=Path)
    parser.add_argument("--mode", choices=("native", "compat"), required=True)
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    failures: list[str] = []
    for raw in args.repo:
        check_repo(lexical(raw), args.mode, failures)
    print("RESULT:", "FAIL" if failures else "OK")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
