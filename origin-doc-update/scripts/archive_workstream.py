#!/usr/bin/env python3
"""Archive a completed workstream and remove its active index entry."""
from __future__ import annotations

import argparse
import re
import sys
from datetime import date
from pathlib import Path

from archive_transaction import apply_archive, cleanup_staged, stage_text
from index_entries import find_index_entry_lines, remove_index_entry
from validate_repo_docs import parse_workstream_issue_blocks, validate_repo


# `- [ ] x` was the only spelling this matched, so `- [] x`, an indented
# sub-item, `* [ ] x` and a bare `- [ ]` all counted as done. The gate's own
# message claims every item is complete, which it could not measure.
UNCHECKED_BOX = re.compile(r"^[ \t]*[-*+][ \t]+\[[ \t]*\].*$", re.MULTILINE)


def update_scalar(content: str, field: str, value: str) -> str:
    """Set a front-matter scalar, raising when the field is not there to set.

    Substituting nothing used to be silent, so a file with no `status:` line
    moved into archive/ having never been marked archived — and the validator
    accepts an empty status under archive/, so nothing downstream disagreed.
    """
    pattern = rf"(^---\n.*?)(^{re.escape(field)}:[ \t]*).*?([ \t]*\n)(.*?^---)"
    replacement = rf"\g<1>\g<2>{value}\g<3>\g<4>"
    updated, count = re.subn(
        pattern, replacement, content, count=1, flags=re.DOTALL | re.MULTILINE
    )
    if not count:
        raise KeyError(field)
    return updated


def main() -> None:
    parser = argparse.ArgumentParser(description="Archive a completed workstream.")
    parser.add_argument("workstream", help="Workstream filename, ID, or path")
    parser.add_argument("--repo", default=".")
    args = parser.parse_args()

    repo = Path(args.repo).resolve()
    active_dir = repo / "docs/workstreams"
    archive_dir = active_dir / "archive"
    candidate = Path(args.workstream)
    if candidate.is_file():
        source = candidate.resolve()
    else:
        name = candidate.name
        source = active_dir / (name if name.endswith(".md") else f"{name}.md")
    if not source.is_file() or source.parent != active_dir:
        print(f"Error: active workstream not found: {source}", file=sys.stderr)
        raise SystemExit(1)

    original_content = source.read_text()
    content = original_content

    # The template's "Workstream archived when complete" box is the very action
    # this script performs, so requiring it pre-checked made every straight run
    # fail (self-reference). archive_issue.py already ticks its equivalent box on
    # the script's behalf; do the same here before gating.
    content = re.sub(
        r"^([ \t]*[-*+][ \t]+\[)[ \t]*(\][ \t]*Workstream archived)",
        r"\g<1>x\g<2>",
        content,
        flags=re.MULTILINE,
    )

    completion = content.split("## Completion", 1)
    # Two different problems, two different messages. Conflating them sent a
    # reader hunting for unchecked boxes in a file that had no checklist at all
    # — and validate_repo_docs.py passes without the section, so a hand-authored
    # workstream only discovers the requirement here, at archive time.
    if len(completion) != 2:
        print(
            "Error: this workstream has no '## Completion' section, so there is "
            "nothing to verify before archiving. Add the section from "
            "references/workstream.template.md and tick its items.",
            file=sys.stderr,
        )
        raise SystemExit(1)
    unchecked = UNCHECKED_BOX.findall(completion[1])
    if unchecked:
        print(
            "Error: complete every workstream checklist item before archiving. "
            f"{len(unchecked)} still unchecked:",
            file=sys.stderr,
        )
        for item in unchecked:
            print(f"  {item}", file=sys.stderr)
        raise SystemExit(1)

    incomplete = [
        issue_id
        for issue_id, metadata, _body in parse_workstream_issue_blocks(content)
        if metadata.get("status") != "complete"
    ]
    if incomplete:
        print(
            f"Error: mark every embedded issue complete before archiving: {', '.join(incomplete)}",
            file=sys.stderr,
        )
        raise SystemExit(1)

    errors, _warnings = validate_repo(repo)
    target = str(source.relative_to(repo))
    target_errors = [error for error in errors if error.startswith(target)]
    if target_errors:
        print("Error: resolve workstream documentation errors before archiving", file=sys.stderr)
        for error in target_errors:
            print(f"  - {error}", file=sys.stderr)
        raise SystemExit(1)

    try:
        archive_dir.mkdir(parents=True, exist_ok=True)
    except OSError as error:
        print(f"Error: cannot prepare archive directory: {error}", file=sys.stderr)
        raise SystemExit(1)
    destination = archive_dir / source.name
    if destination.exists():
        print(f"Error: already archived: {destination}", file=sys.stderr)
        raise SystemExit(1)

    today = date.today().isoformat()
    try:
        content = update_scalar(content, "status", "archived")
        content = update_scalar(content, "updated_at", today)
    except KeyError as missing:
        print(
            f"Error: front matter has no '{missing.args[0]}:' field to update, so archiving "
            "would move the file without recording that it was archived. Add the field first.",
            file=sys.stderr,
        )
        raise SystemExit(1)
    index = repo / "docs/00_index.md"
    workstream_id = source.stem
    index_plan = None
    staged_paths = []
    try:
        if index.exists() and not index.is_file():
            raise OSError(f"index path is not a regular file: {index}")
        if index.is_file():
            index_content = index.read_text()
            target_lines = find_index_entry_lines(
                index_content, "docs/workstreams", workstream_id
            )
            new_index, removed = remove_index_entry(
                index_content, "docs/workstreams", workstream_id
            )
            if removed or target_lines:
                if removed != len(target_lines):
                    raise ValueError(
                        "index target count mismatch: "
                        f"removed={removed}, reported={len(target_lines)}"
                    )
                new_index = re.sub(
                    r"(updated_at:[ \t]*)[\d-]+", rf"\g<1>{today}", new_index, count=1
                )
                index_plan = (index_content, new_index, target_lines)

        staged_destination = stage_text(destination, content, source)
        staged_paths.append(staged_destination)
        staged_source_restore = stage_text(source, original_content, source)
        staged_paths.append(staged_source_restore)

        staged_index = None
        staged_index_restore = None
        if index_plan is not None:
            original_index, new_index, _target_lines = index_plan
            staged_index = stage_text(index, new_index, index)
            staged_paths.append(staged_index)
            staged_index_restore = stage_text(index, original_index, index)
            staged_paths.append(staged_index_restore)

        apply_archive(
            source,
            destination,
            staged_destination,
            staged_source_restore,
            index if index_plan is not None else None,
            staged_index,
            staged_index_restore,
        )
    except (OSError, RuntimeError, ValueError) as error:
        print(f"Error: {error}", file=sys.stderr)
        raise SystemExit(1)
    finally:
        cleanup_staged(staged_paths)

    if index_plan is not None:
        target_lines = index_plan[2]
        print(
            f"Updated: docs/00_index.md (removed {len(target_lines)} active reference(s))"
        )
        for line_number, line in target_lines:
            print(f"  line {line_number}: {line}")
    else:
        print(
            f"Note: {workstream_id} not found in docs/00_index.md "
            "active references (check manually if needed)"
        )

    print(f"Archived: {source.name} -> docs/workstreams/archive/{source.name}")


if __name__ == "__main__":
    main()
