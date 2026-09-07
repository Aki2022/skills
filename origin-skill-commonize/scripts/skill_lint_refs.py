#!/usr/bin/env python3
"""Emit the resource references that ``skill_lint.sh`` must validate.

The skills repository uses both ordinary relative references (for example
``references/guide.md``) and cross-skill references (``/other-skill`` followed
by ``references/guide.md``).  The latter are resolved from the canonical root,
not from the skill currently being inspected.

Output is tab-separated and intentionally contains only relative names:

``same <TAB> resource``
    A resource relative to the current skill.
``cross <TAB> skill <TAB> resource``
    A resource relative to another canonical skill.
``unsafe <TAB> scope <TAB> resource``
    A resource containing a parent traversal.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path


GROUPS = "scripts|references|assets"
RESOURCE = re.compile(
    rf"(?<![A-Za-z0-9_/-])(?:{GROUPS})/[A-Za-z0-9._/-]+"
)
FULL_CROSS = re.compile(
    rf"/(?P<skill>[A-Za-z0-9][A-Za-z0-9._-]*)/(?P<resource>(?:{GROUPS})/[A-Za-z0-9._/-]+)"
)
SKILL_TOKEN = re.compile(r"/(?P<skill>[A-Za-z0-9][A-Za-z0-9._-]*)")


def canonical_skill_names(root: Path) -> set[str]:
    return {
        path.name
        for path in root.iterdir()
        if path.name not in {".git", "docs", "node_modules"}
        and (path.is_dir() or path.is_symlink())
    }


def unsafe(resource: str) -> bool:
    return ".." in Path(resource).parts


def emit(kind: str, *values: str) -> None:
    print("\t".join((kind, *values)))


def scan(root: Path, markdown: Path) -> None:
    names = canonical_skill_names(root)
    seen: set[tuple[str, ...]] = set()

    for line in markdown.read_text(encoding="utf-8").splitlines():
        # A full /skill/group/path form is unambiguous and takes precedence.
        full_spans: list[tuple[int, int]] = []
        for match in FULL_CROSS.finditer(line):
            full_spans.append(match.span())
            skill = match.group("skill")
            if skill not in names:
                # A path-like phrase that is not a canonical skill is not a
                # skill resource reference; the normal relative scan below
                # still handles any standalone resource on the line.
                continue
            resource = match.group("resource").rstrip(".,)")
            key = ("cross", skill, resource)
            if key in seen:
                continue
            seen.add(key)
            if unsafe(resource):
                emit("unsafe", "cross", f"{skill}/{resource}")
            else:
                emit("cross", skill, resource)

        # For the suite's `/skill` → `references/file` spelling, associate a
        # resource with the nearest canonical skill token to its left.  A full
        # path span is skipped so it is not emitted twice.
        skill_tokens = [
            (match.start(), match.group("skill"))
            for match in SKILL_TOKEN.finditer(line)
            if match.group("skill") in names
        ]
        for match in RESOURCE.finditer(line):
            if any(match.start() >= start and match.end() <= end for start, end in full_spans):
                continue
            resource = match.group(0).rstrip(".,)")
            preceding = [item for item in skill_tokens if item[0] < match.start()]
            skill = preceding[-1][1] if preceding else None
            if skill is None:
                key = ("same", resource)
                if key in seen:
                    continue
                seen.add(key)
                if unsafe(resource):
                    emit("unsafe", "same", resource)
                else:
                    emit("same", resource)
            else:
                key = ("cross", skill, resource)
                if key in seen:
                    continue
                seen.add(key)
                if unsafe(resource):
                    emit("unsafe", "cross", f"{skill}/{resource}")
                else:
                    emit("cross", skill, resource)


def main(argv: list[str]) -> int:
    if len(argv) != 3:
        print(f"usage: {argv[0]} CANONICAL_ROOT SKILL_MD", file=sys.stderr)
        return 2
    root = Path(argv[1])
    markdown = Path(argv[2])
    if not root.is_dir() or not markdown.is_file():
        print("resource reference inputs must exist", file=sys.stderr)
        return 2
    scan(root, markdown)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
