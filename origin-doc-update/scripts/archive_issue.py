#!/usr/bin/env python3
"""Archive a completed issue: move to archive/, update status, update 00_index.md."""
import argparse
import os
import re
import sys
from datetime import date
from pathlib import Path
from typing import Optional

from archive_workstream import UNCHECKED_BOX
from archive_transaction import apply_archive, cleanup_staged, stage_text
from index_entries import find_index_entry_lines, normalize_entry_id, remove_index_entry
from validate_repo_docs import validate_repo


def update_front_matter_field(content: str, field: str, value: str) -> str:
    """Set a front-matter scalar, raising when the field is not there to set."""
    pattern = rf"(^---\n.*?)(^{re.escape(field)}:[ \t]*).*?([ \t]*\n)(.*?^---)"
    replacement = rf"\g<1>\g<2>{value}\g<3>\g<4>"
    updated, count = re.subn(
        pattern, replacement, content, count=1, flags=re.DOTALL | re.MULTILINE
    )
    if not count:
        raise KeyError(field)
    return updated


def read_front_matter_field(content: str, field: str) -> str:
    """Read a scalar field's raw value from the first YAML front matter block, or ''."""
    m = re.search(rf"^{re.escape(field)}:[ \t]*(.*?)[ \t]*$", content, flags=re.MULTILINE)
    if not m:
        return ""
    return m.group(1).strip().strip('"').strip("'")


def prepare_index_update(
    index_path: str, issue_id: str, today: Optional[str] = None
) -> Optional[tuple[str, str, list[tuple[int, str]]]]:
    """Prepare an index update without writing it or changing the issue."""
    if not os.path.exists(index_path):
        return None
    if not os.path.isfile(index_path):
        raise OSError(f"index path is not a regular file: {index_path}")

    with open(index_path) as f:
        content = f.read()

    target_lines = find_index_entry_lines(content, "docs/issues", issue_id)
    new_content, removed = remove_index_entry(content, "docs/issues", issue_id)

    if not removed and not target_lines:
        return None
    if removed != len(target_lines):
        raise ValueError(
            f"index target count mismatch: removed={removed}, reported={len(target_lines)}"
        )

    today = today or date.today().isoformat()
    new_content = re.sub(r"(updated_at:[ \t]*)[\d-]+", f"\\g<1>{today}", new_content)
    return content, new_content, target_lines


def remove_issue_from_index(index_path: str, issue_id: str) -> tuple[bool, list[tuple[int, str]]]:
    """Remove target rows atomically and return the exact rows that changed."""
    plan = prepare_index_update(index_path, issue_id)
    if plan is None:
        return False, []

    _original, new_content, target_lines = plan
    path = Path(index_path)
    staged = stage_text(path, new_content, path)
    try:
        os.replace(staged, path)
    finally:
        cleanup_staged([staged])
    return True, target_lines


def main():
    parser = argparse.ArgumentParser(description="Archive a completed issue.")
    parser.add_argument("issue", help="Issue filename or path (e.g. ISSUE-20260612-foo.md or the full path)")
    parser.add_argument("--repo", default=".", help="Repository root (default: cwd)")
    args = parser.parse_args()

    repo = os.path.abspath(args.repo)
    issues_dir = os.path.join(repo, "docs", "issues")
    archive_dir = os.path.join(issues_dir, "archive")

    issue_arg = args.issue
    candidate = os.path.abspath(os.path.expanduser(issue_arg)) if os.path.isabs(issue_arg) else issue_arg
    if os.path.isfile(candidate):
        src = os.path.abspath(candidate)
    else:
        issue_id = normalize_entry_id(issue_arg)
        src = os.path.join(issues_dir, f"{issue_id}.md")

    if not os.path.isfile(src):
        print(f"Error: issue not found: {src}", file=sys.stderr)
        sys.exit(1)

    src = os.path.abspath(src)
    if os.path.dirname(src) != os.path.abspath(issues_dir):
        print(f"Error: active issue must be directly under {issues_dir}: {src}", file=sys.stderr)
        sys.exit(1)

    fname = os.path.basename(src)
    issue_id = normalize_entry_id(fname)
    dest = os.path.join(archive_dir, fname)

    if os.path.exists(dest):
        print(f"Error: already archived: {dest}", file=sys.stderr)
        sys.exit(1)

    try:
        os.makedirs(archive_dir, exist_ok=True)
    except OSError as error:
        print(f"Error: cannot prepare archive directory: {error}", file=sys.stderr)
        sys.exit(1)

    with open(src) as f:
        original_content = f.read()
    content = original_content

    # The template's "Moved to docs/issues/archive/" box is the very action this
    # script performs, so requiring it pre-checked made every straight run fail
    # (self-reference). Check it here on the script's behalf before gating.
    content = re.sub(
        r"^([ \t]*[-*+][ \t]+\[)[ \t]*(\][ \t]*Moved to docs/issues/archive)",
        r"\g<1>x\g<2>",
        content,
        flags=re.MULTILINE,
    )

    # An issue had no completion gate at all, while a workstream had one — so an
    # improvement issue filed precisely so an observation would not be lost could
    # be archived unstarted, dropped from the index, and reported as archived.
    completion = content.split("## Completion", 1)
    if len(completion) == 2:
        unchecked = UNCHECKED_BOX.findall(completion[1])
        if unchecked:
            print(
                "Error: complete every issue checklist item before archiving. "
                f"{len(unchecked)} still unchecked:",
                file=sys.stderr,
            )
            for item in unchecked:
                print(f"  {item}", file=sys.stderr)
            sys.exit(1)

    errors, _warnings = validate_repo(repo)
    target = os.path.relpath(src, repo)
    target_errors = [error for error in errors if error.startswith(target)]
    if target_errors:
        print("Error: resolve issue documentation errors before archiving", file=sys.stderr)
        for error in target_errors:
            print(f"  - {error}", file=sys.stderr)
        sys.exit(1)

    today = date.today().isoformat()
    branch = read_front_matter_field(content, "branch")
    try:
        content = update_front_matter_field(content, "status", "archived")
        content = update_front_matter_field(content, "updated_at", today)
    except KeyError as missing:
        print(
            f"Error: front matter has no '{missing.args[0]}:' field to update, so archiving "
            "would move the file without recording that it was archived. Add the field first.",
            file=sys.stderr,
        )
        sys.exit(1)

    index_path = os.path.join(repo, "docs", "00_index.md")
    index_plan = None
    staged_paths = []
    try:
        index_plan = prepare_index_update(index_path, issue_id, today)
        staged_destination = stage_text(Path(dest), content, Path(src))
        staged_paths.append(staged_destination)
        staged_source_restore = stage_text(Path(src), original_content, Path(src))
        staged_paths.append(staged_source_restore)

        staged_index = None
        staged_index_restore = None
        if index_plan is not None:
            original_index, new_index, _target_lines = index_plan
            index_file = Path(index_path)
            staged_index = stage_text(index_file, new_index, index_file)
            staged_paths.append(staged_index)
            staged_index_restore = stage_text(index_file, original_index, index_file)
            staged_paths.append(staged_index_restore)

        apply_archive(
            Path(src),
            Path(dest),
            staged_destination,
            staged_source_restore,
            Path(index_path) if index_plan is not None else None,
            staged_index,
            staged_index_restore,
        )
    except (OSError, RuntimeError, ValueError) as error:
        print(f"Error: {error}", file=sys.stderr)
        sys.exit(1)
    finally:
        cleanup_staged(staged_paths)

    print(f"Archived: {fname} -> docs/issues/archive/{fname}")
    if branch:
        print(f"Note: this issue's branch was '{branch}'.")
        print(f"  If it still exists, run origin-git-cleanup to check/remove it.")
    if index_plan is not None:
        target_lines = index_plan[2]
        print(f"Updated: docs/00_index.md (removed {len(target_lines)} active reference(s))")
        for line_number, line in target_lines:
            print(f"  line {line_number}: {line}")
    else:
        print(f"Note: {issue_id} not found in docs/00_index.md Active Issues (check manually if needed)")


if __name__ == "__main__":
    main()
