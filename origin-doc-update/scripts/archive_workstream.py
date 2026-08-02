#!/usr/bin/env python3
"""Archive a completed workstream and remove its active index entry."""
from __future__ import annotations

import argparse
import re
import shutil
import sys
from datetime import date
from pathlib import Path

from index_entries import remove_index_entry
from validate_repo_docs import parse_workstream_issue_blocks, validate_repo


def update_scalar(content: str, field: str, value: str) -> str:
    pattern = rf"(^---\n.*?)(^{re.escape(field)}:[ \t]*).*?([ \t]*\n)(.*?^---)"
    replacement = rf"\g<1>\g<2>{value}\g<3>\g<4>"
    return re.sub(pattern, replacement, content, count=1, flags=re.DOTALL | re.MULTILINE)


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

    content = source.read_text()
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
    unchecked = re.findall(r"^- \[ \] .*$", completion[1], re.MULTILINE)
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
        for issue_id, metadata in parse_workstream_issue_blocks(content)
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

    archive_dir.mkdir(parents=True, exist_ok=True)
    destination = archive_dir / source.name
    if destination.exists():
        print(f"Error: already archived: {destination}", file=sys.stderr)
        raise SystemExit(1)

    today = date.today().isoformat()
    content = update_scalar(content, "status", "archived")
    content = update_scalar(content, "updated_at", today)
    source.write_text(content)
    shutil.move(str(source), str(destination))

    index = repo / "docs/00_index.md"
    if index.is_file():
        index_content = index.read_text()
        workstream_id = source.stem
        index_content, removed = remove_index_entry(
            index_content, "docs/workstreams", workstream_id
        )
        index_content = re.sub(
            r"(updated_at:[ \t]*)[\d-]+", rf"\g<1>{today}", index_content, count=1
        )
        index.write_text(index_content)
        if not removed:
            print(
                f"Note: {workstream_id} not found in docs/00_index.md "
                "active references (check manually if needed)"
            )

    print(f"Archived: {source.name} -> docs/workstreams/archive/{source.name}")


if __name__ == "__main__":
    main()
