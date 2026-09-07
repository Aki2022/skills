#!/usr/bin/env python3
"""Validate the global skill alias topology without changing any files.

The canonical skills directory is the only shared source.  Claude and
Antigravity expose that source through one directory symlink, while Codex and
Gemini keep a regular skills directory containing one symlink per skill.  A
Codex-only adapted skill and explicitly named seat-local paths can be supplied
as narrow exceptions; everything else must be a canonical skill symlink.

The checker deliberately takes every root as an argument.  It therefore does
not guess account names, inspect mutable account state, or follow a shell's
current working directory.  It is suitable for CI and for a read-only audit
of several global accounts in one invocation.
"""

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path


IGNORED_CANONICAL_DIRS = {".git", "docs", "node_modules"}


def lexical(path: Path) -> Path:
    """Normalize a user-supplied path without resolving symlink targets."""

    return Path(os.path.abspath(os.path.expanduser(str(path))))


def resolved(path: Path) -> Path:
    """Resolve a path for target comparison, including a relative symlink."""

    return Path(os.path.realpath(path))


def canonical_names(root: Path) -> tuple[set[str], list[str]]:
    """Return skill names and non-skill directories that need lint attention."""

    names: set[str] = set()
    warnings: list[str] = []
    try:
        children = sorted(root.iterdir(), key=lambda p: p.name)
    except OSError as exc:
        return names, [f"cannot read canonical root: {exc}"]
    for child in children:
        if child.name in IGNORED_CANONICAL_DIRS or child.name.startswith("."):
            continue
        if child.is_dir() and (child / "SKILL.md").is_file():
            names.add(child.name)
        elif child.is_dir():
            # The skill linter owns the detailed S1 failure.  Keeping this as
            # a warning avoids duplicating its policy while making the scope
            # visible in an audit.
            warnings.append(f"canonical non-skill directory ignored: {child.name}")
    return names, warnings


def report_failure(failures: list[str], message: str) -> None:
    failures.append(message)
    print(f"FAIL {message}")


def check_directory_alias(
    kind: str, alias: Path, canonical: Path, failures: list[str]
) -> None:
    """Check a Claude/Antigravity whole-directory alias."""

    if not alias.is_symlink():
        if alias.exists():
            report_failure(failures, f"{kind} {alias}: expected directory symlink")
        else:
            report_failure(failures, f"{kind} {alias}: missing directory symlink")
        return
    if not alias.exists():
        report_failure(failures, f"{kind} {alias}: broken directory symlink")
        return
    if not alias.is_dir() or resolved(alias) != resolved(canonical):
        report_failure(
            failures,
            f"{kind} {alias}: target is not canonical root ({canonical})",
        )
        return
    print(f"OK {kind} {alias}")


def check_per_skill_root(
    kind: str,
    alias_root: Path,
    canonical: Path,
    canonical_skill_names: set[str],
    adapted_targets: dict[str, Path],
    local_only_paths: set[Path],
    failures: list[str],
) -> None:
    """Check one regular Codex/Gemini skills directory."""

    failure_count_before = len(failures)
    expected: dict[str, Path] = {
        name: canonical / name for name in canonical_skill_names
    }
    expected.update(adapted_targets)

    if alias_root.is_symlink():
        report_failure(
            failures,
            f"{kind} {alias_root}: root must be a regular directory (per-skill symlinks)",
        )
        return
    if not alias_root.exists():
        report_failure(failures, f"{kind} {alias_root}: missing skills directory")
        return
    if not alias_root.is_dir():
        report_failure(failures, f"{kind} {alias_root}: not a directory")
        return

    try:
        children = sorted(alias_root.iterdir(), key=lambda p: p.name)
    except OSError as exc:
        report_failure(failures, f"{kind} {alias_root}: cannot read directory: {exc}")
        return

    for child in children:
        name = child.name
        child_lexical = lexical(child)

        if kind == "codex" and name == ".system":
            # Codex owns this product/system tree; it is intentionally not a
            # shared skill and must not be rewritten by this checker.
            print(f"OK {kind} {child}: .system exception")
            continue
        if child_lexical in local_only_paths:
            print(f"OK {kind} {child}: seat-local exception")
            continue

        if name not in expected:
            if child.is_symlink() and not child.exists():
                report_failure(failures, f"{kind} {child}: broken unregistered symlink")
            else:
                report_failure(failures, f"{kind} {child}: unregistered entry")
            continue

        target = expected[name]
        if not child.is_symlink():
            report_failure(failures, f"{kind} {child}: expected per-skill symlink")
            continue
        if not child.exists():
            report_failure(failures, f"{kind} {child}: broken symlink")
            continue
        if resolved(child) != resolved(target):
            report_failure(
                failures,
                f"{kind} {child}: target mismatch (expected {target})",
            )
            continue
        print(f"OK {kind} {child}")

    for name, target in sorted(expected.items()):
        alias = alias_root / name
        if not alias.exists() and not alias.is_symlink():
            report_failure(failures, f"{kind} {alias}: missing per-skill symlink")
            continue
        # A broken link is already reported in the child loop; this branch
        # keeps an expected entry missing from being silently accepted.
        if not alias.is_symlink():
            continue

    if len(failures) == failure_count_before:
        print(f"OK {kind} {alias_root}: per-skill topology checked")


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Check global Claude/Codex/Gemini/Antigravity skill aliases"
    )
    parser.add_argument("--canonical", required=True, type=Path)
    parser.add_argument("--claude", action="append", default=[], type=Path)
    parser.add_argument("--codex", action="append", default=[], type=Path)
    parser.add_argument("--gemini", action="append", default=[], type=Path)
    parser.add_argument("--antigravity", action="append", default=[], type=Path)
    parser.add_argument(
        "--codex-adapted",
        action="append",
        default=[],
        type=Path,
        help="path to a Codex-only adapted skill directory (repeatable)",
    )
    parser.add_argument(
        "--local-only",
        action="append",
        default=[],
        type=Path,
        help="exact seat-local skill path exempt from shared topology (repeatable)",
    )
    args = parser.parse_args(argv)
    if not (args.claude or args.codex or args.gemini or args.antigravity):
        parser.error("at least one alias root is required")
    return args


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    canonical = lexical(args.canonical)
    if canonical.is_symlink() or not canonical.is_dir():
        print(f"ERROR canonical root must be a regular directory: {canonical}")
        return 2

    names, canonical_notes = canonical_names(canonical)
    if not names:
        print(f"ERROR canonical root has no skill directories: {canonical}")
        return 2
    print(f"scope: canonical={canonical} skills={len(names)}")
    for note in canonical_notes:
        print(f"WARN {note}")

    adapted_targets: dict[str, Path] = {}
    failures: list[str] = []
    for raw in args.codex_adapted:
        path = lexical(raw)
        name = path.name
        if name in names or name in adapted_targets:
            report_failure(failures, f"codex-adapted {path}: duplicate skill name")
            continue
        if not path.is_dir() or not (path / "SKILL.md").is_file():
            report_failure(failures, f"codex-adapted {path}: skill directory/SKILL.md missing")
            continue
        adapted_targets[name] = path

    per_skill_roots = [*(lexical(p) for p in args.codex), *(lexical(p) for p in args.gemini)]
    local_only_paths = {lexical(p) for p in args.local_only}
    for path in sorted(local_only_paths):
        if not any(path.parent == root for root in per_skill_roots):
            report_failure(
                failures,
                f"local-only {path}: must be a direct child of a --codex/--gemini root",
            )
        elif not path.exists() and not path.is_symlink():
            report_failure(failures, f"local-only {path}: path is missing")

    for raw in args.claude:
        check_directory_alias("claude", lexical(raw), canonical, failures)
    for raw in args.antigravity:
        check_directory_alias("antigravity", lexical(raw), canonical, failures)
    for raw in args.codex:
        check_per_skill_root(
            "codex",
            lexical(raw),
            canonical,
            names,
            adapted_targets,
            local_only_paths,
            failures,
        )
    for raw in args.gemini:
        check_per_skill_root(
            "gemini",
            lexical(raw),
            canonical,
            names,
            {},
            local_only_paths,
            failures,
        )

    print("RESULT:", "FAIL" if failures else "OK")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
