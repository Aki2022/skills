#!/usr/bin/env python3
"""Archive a completed issue: move to archive/, update status, update 00_index.md."""
import argparse
import os
import re
import shutil
import sys
from datetime import date

from archive_workstream import UNCHECKED_BOX
from index_entries import remove_index_entry
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


def remove_issue_from_index(index_path: str, issue_id: str) -> bool:
    """Remove the issue reference line from docs/00_index.md and bump updated_at."""
    if not os.path.isfile(index_path):
        return False

    with open(index_path) as f:
        content = f.read()

    new_content, removed = remove_index_entry(content, "docs/issues", issue_id)

    if not removed:
        return False

    today = date.today().isoformat()
    new_content = re.sub(r"(updated_at:[ \t]*)[\d-]+", f"\\g<1>{today}", new_content)

    with open(index_path, "w") as f:
        f.write(new_content)
    return True


def main():
    parser = argparse.ArgumentParser(description="Archive a completed issue.")
    parser.add_argument("issue", help="Issue filename or path (e.g. ISSUE-20260612-foo.md or the full path)")
    parser.add_argument("--repo", default=".", help="Repository root (default: cwd)")
    args = parser.parse_args()

    repo = os.path.abspath(args.repo)
    issues_dir = os.path.join(repo, "docs", "issues")
    archive_dir = os.path.join(issues_dir, "archive")

    issue_arg = args.issue
    if os.path.isabs(issue_arg) and os.path.isfile(issue_arg):
        src = issue_arg
    elif os.path.isfile(issue_arg):
        src = os.path.abspath(issue_arg)
    else:
        src = os.path.join(issues_dir, os.path.basename(issue_arg))

    if not os.path.isfile(src):
        print(f"Error: issue not found: {src}", file=sys.stderr)
        sys.exit(1)

    fname = os.path.basename(src)
    issue_id = fname.replace(".md", "")
    dest = os.path.join(archive_dir, fname)

    if os.path.exists(dest):
        print(f"Error: already archived: {dest}", file=sys.stderr)
        sys.exit(1)

    os.makedirs(archive_dir, exist_ok=True)

    with open(src) as f:
        content = f.read()

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

    with open(src, "w") as f:
        f.write(content)

    shutil.move(src, dest)
    print(f"Archived: {fname} -> docs/issues/archive/{fname}")

    if branch:
        print(f"Note: this issue's branch was '{branch}'.")
        print(f"  If it still exists, run origin-git-cleanup to check/remove it.")

    index_path = os.path.join(repo, "docs", "00_index.md")
    if remove_issue_from_index(index_path, issue_id):
        print(f"Updated: docs/00_index.md (removed active reference)")
    else:
        print(f"Note: {issue_id} not found in docs/00_index.md Active Issues (check manually if needed)")


if __name__ == "__main__":
    main()
