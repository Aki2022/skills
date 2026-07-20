#!/usr/bin/env python3
"""Create a new issue file following doc-governance naming convention."""
import argparse
import os
import sys
from datetime import date

SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE_DIR = os.path.join(SKILL_DIR, "references")


def main():
    parser = argparse.ArgumentParser(description="Create a new issue file.")
    parser.add_argument("slug", help="Short slug (lowercase, hyphens): e.g. auth-token-refresh")
    parser.add_argument("--date", default=None, help="Date override YYYYMMDD (default: today)")
    parser.add_argument("--repo", default=".", help="Repository root (default: cwd)")
    parser.add_argument("--title", default="", help="Issue title (default: derived from slug)")
    args = parser.parse_args()

    slug = args.slug.lower().replace(" ", "-")
    if not re_slug_ok(slug):
        print(f"Error: slug must be alphanumeric with hyphens: {slug!r}", file=sys.stderr)
        sys.exit(1)

    date_str = args.date or date.today().strftime("%Y%m%d")
    if len(date_str) != 8 or not date_str.isdigit():
        print(f"Error: date must be YYYYMMDD: {date_str!r}", file=sys.stderr)
        sys.exit(1)

    issue_id = f"ISSUE-{date_str}-{slug}"
    repo = os.path.abspath(args.repo)
    issues_dir = os.path.join(repo, "docs", "issues")

    if not os.path.isdir(issues_dir):
        print(f"Error: {issues_dir} does not exist. Run init_repo_docs.py first.", file=sys.stderr)
        sys.exit(1)

    dest = os.path.join(issues_dir, f"{issue_id}.md")
    if os.path.exists(dest):
        print(f"Error: already exists: {dest}", file=sys.stderr)
        sys.exit(1)

    today_iso = date.today().isoformat()
    title = args.title or slug.replace("-", " ").title()

    try:
        tmpl_path = os.path.join(TEMPLATE_DIR, "issue.template.md")
        with open(tmpl_path) as f:
            content = f.read()
        content = content.replace("ISSUE-YYYYMMDD-short-slug", issue_id)
        content = content.replace("YYYY-MM-DD", today_iso)
        content = content.replace("# Title", f"# {title}", 1)
    except FileNotFoundError:
        content = (
            f"---\nid: {issue_id}\nstatus: active\n"
            f"created_at: {today_iso}\nupdated_at: {today_iso}\n"
            f"branch: {issue_id}\npr: \"\"\n"
            "related_specs: []\nrelated_guides: []\n---\n\n"
            f"# {title}\n\n## Goal\n\n## Current Status\n\n## Next Actions\n\n## Notes\n\n"
            "## Completion\n\n"
            "- [ ] Implementation completed or intentionally not needed\n"
            "- [ ] Specs updated if direction or requirements changed\n"
            "- [ ] Guides updated if implemented behavior changed\n"
            "- [ ] Branch merged and cleaned up (or intentionally kept — note why)\n"
            "- [ ] 00_index.md updated\n"
            "- [ ] Moved to docs/issues/archive/ when complete\n"
        )

    with open(dest, "w") as f:
        f.write(content)

    print(f"Created: {dest}")
    print(f"\nAdd to docs/00_index.md Active Issues:")
    print(f"  - docs/issues/{issue_id}.md — {title}")
    print(f"\nSuggested branch (convention: branch name = issue id):")
    print(f"  git checkout -b {issue_id}")
    print(f"  # or, for a separate worktree: git worktree add ../{issue_id} -b {issue_id}")


def re_slug_ok(slug: str) -> bool:
    import re
    return bool(re.match(r"^[a-z0-9][a-z0-9-]*[a-z0-9]$", slug) or re.match(r"^[a-z0-9]$", slug))


if __name__ == "__main__":
    main()
