#!/usr/bin/env python3
"""Cross-check own-doc-update active issues against real git branch/worktree state.

Read-only. Never mutates git or docs. Bridges own-git-clean (owns git state) with
own-doc-update (owns issue front matter) so branch classification in own-git-clean's
Stage 2 can be driven by actual active-issue references instead of name
guessing, and so resuming an issue can detect a branch that went missing.
"""
import argparse
import os
import re
import subprocess
import sys
from typing import Optional


def strip_inline_comment(raw: str) -> str:
    """Remove a trailing YAML `# comment` from a scalar value.

    YAML only starts a comment at a `#` preceded by whitespace (or at the start of
    the value), and never inside a quoted scalar. So `feat/x#frag` and
    `"has # inside"` keep their `#`, while `"" # memo` and `feat/x # memo` lose the memo.
    Mirrored verbatim in own-git-clean/scripts/check_active_issue_branches.py (the two
    skills are independent, so the function is copied rather than imported; a test
    pins them to the same answer).
    """
    quote = ""
    for index, char in enumerate(raw):
        if quote:
            if char == quote:
                quote = ""
        elif char in "\"'" and not raw[:index].strip(" \t"):
            quote = char
        elif char == "#" and (index == 0 or raw[index - 1] in " \t"):
            return raw[:index].rstrip()
    return raw.rstrip()


def parse_front_matter(path: str) -> Optional[dict]:
    """Parse top-level scalar YAML front matter fields. Mirrors own-doc-update's
    validate_repo_docs.py parser: only reads unindented `key: value` lines,
    so inline lists like `related_specs: []` are captured as their raw
    string value and multi-line list items are ignored (not needed here)."""
    try:
        with open(path) as f:
            content = f.read()
    except OSError:
        return None

    if not content.startswith("---"):
        return {}

    end = content.find("\n---", 3)
    if end == -1:
        return None

    fm = content[3:end].strip()
    result: dict = {}
    for line in fm.splitlines():
        if ":" in line and not line.startswith(" ") and not line.startswith("-"):
            key, _, val = line.partition(":")
            result[key.strip()] = strip_inline_comment(val.strip()).strip('"').strip("'")
    return result


def run_git(repo: str, *args: str) -> str:
    try:
        out = subprocess.run(
            ["git", "-C", repo, *args],
            capture_output=True, text=True, check=False,
        )
        return out.stdout
    except OSError:
        return ""


def section(title: str) -> None:
    print(f"\n=== {title} ===")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Cross-check active issue `branch` fields against real git branches/worktrees."
    )
    parser.add_argument("repo", nargs="?", default=".", help="Repository root (default: cwd)")
    args = parser.parse_args()

    repo = os.path.abspath(args.repo)
    issues_dir = os.path.join(repo, "docs", "issues")

    if not os.path.isdir(issues_dir):
        print("(skipped: no docs/issues/ found — own-doc-update not in use here)")
        return

    is_git_repo = subprocess.run(
        ["git", "-C", repo, "rev-parse", "--is-inside-work-tree"],
        capture_output=True, text=True, check=False,
    ).returncode == 0
    if not is_git_repo:
        print("(skipped: not inside a git work tree)")
        return

    # Collect active issues (exclude archive/) and their recorded branch.
    issue_branches: dict[str, str] = {}  # issue_id -> branch
    issue_status: dict[str, str] = {}  # issue_id -> status
    branch_owners: dict[str, list[str]] = {}
    for fname in sorted(os.listdir(issues_dir)):
        fpath = os.path.join(issues_dir, fname)
        if not fname.endswith(".md") or fname.startswith(".") or os.path.isdir(fpath):
            continue
        fm = parse_front_matter(fpath)
        if not fm:
            continue
        issue_id = fm.get("id", fname.replace(".md", ""))
        branch = fm.get("branch", "").strip()
        if branch:
            issue_branches[issue_id] = branch
            issue_status[issue_id] = fm.get("status", "").strip()
            branch_owners.setdefault(branch, []).append(issue_id)

    local_branches = {
        b.strip() for b in run_git(repo, "branch", "--format=%(refname:short)").splitlines() if b.strip()
    }
    remote_branches_raw = run_git(repo, "branch", "-r", "--format=%(refname:short)").splitlines()
    # Strip the remote prefix (e.g. "origin/foo" -> "foo") for name comparison.
    remote_branches = {
        re.sub(r"^[^/]+/", "", b.strip()) for b in remote_branches_raw if b.strip() and "HEAD" not in b
    }

    worktree_branches: set[str] = set()
    for line in run_git(repo, "worktree", "list", "--porcelain").splitlines():
        if line.startswith("branch refs/heads/"):
            worktree_branches.add(line[len("branch refs/heads/"):])

    section("DUPLICATE BRANCH REFERENCES")
    dupes = {b: owners for b, owners in branch_owners.items() if len(owners) > 1}
    if dupes:
        for b, owners in dupes.items():
            print(f"branch '{b}' recorded on multiple active issues: {', '.join(owners)}")
    else:
        print("(none)")

    section("ACTIVE ISSUE -> BRANCH STATUS")
    if not issue_branches:
        print("(no active issues record a branch field)")
    unchecked = 0
    for issue_id, branch in sorted(issue_branches.items()):
        if branch in local_branches or branch in remote_branches:
            print(f"OK: {issue_id} -> {branch}")
        elif issue_status.get(issue_id) == "in_progress":
            print(f"MISSING: {issue_id} -> {branch} (not found locally or on remote — the branch of work in progress is gone: resume needs a fresh branch, or it was cleaned up without archiving the issue)")
        else:
            # 枝は merge 後に削除する運用なので、完了・保留・未着手の issue の枝が
            # 無いのは正常。ここを MISSING にすると、枝を記録した issue は全件が
            # MISSING になり、母集団の 100% を指す検査は何も指さなくなる。
            unchecked += 1
    if unchecked:
        print(f"(not checked: {unchecked} issue(s) not in_progress record a branch that no longer exists — expected after merge cleanup)")

    section("ORPHAN CANDIDATE BRANCHES")
    known_branches = set(issue_branches.values())
    orphans = []
    for b in sorted(local_branches):
        if not b.startswith("ISSUE-"):
            continue
        if b in known_branches:
            continue
        if b in worktree_branches:
            continue  # checked out elsewhere; own-git-clean's worktree-protection rule already covers this
        orphans.append(b)
    if orphans:
        for b in orphans:
            print(f"{b} (looks issue-shaped, no active issue references it — investigate before deleting)")
    else:
        print("(none)")

    print("\n=== END CHECK ===")


if __name__ == "__main__":
    main()
