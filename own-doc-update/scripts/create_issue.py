#!/usr/bin/env python3
"""Create a new issue file following doc-governance naming convention."""
import argparse
import os
import sys
from datetime import date
from pathlib import Path

import ownership
from validate_repo_docs import parse_front_matter

SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE_DIR = os.path.join(SKILL_DIR, "references")


def main():
    parser = argparse.ArgumentParser(description="Create a new issue file.")
    parser.add_argument("slug", help="Short slug (lowercase, hyphens): e.g. auth-token-refresh")
    parser.add_argument("--date", default=None, help="Date override YYYYMMDD (default: today)")
    parser.add_argument("--repo", default=".", help="Repository root (default: cwd)")
    parser.add_argument("--title", default="", help="Issue title (default: derived from slug)")
    # Mirrors create_workstream.py: classify guide impact at creation. Without
    # this the generated file always carries guide_impact: required against an
    # empty related_guides, which validate_repo_docs.py rejects, so a fresh
    # issue could never validate.
    impact = parser.add_mutually_exclusive_group(required=True)
    impact.add_argument("--guide", help="Guide ID this issue must update, e.g. GUIDE-data-map")
    impact.add_argument(
        "--no-guide-reason", help="Why this issue changes no implemented behavior"
    )
    # Acceptance is classified at creation for the same reason as guide impact:
    # an unstated acceptance cannot be self-verified, so an agent either stalls
    # at a gate nobody set or overclaims completion.
    verify = parser.add_mutually_exclusive_group(required=True)
    verify.add_argument(
        "--verify-machine",
        help="Machine-verifiable acceptance: command and expected result",
    )
    verify.add_argument(
        "--verify-human",
        help="Human-review acceptance: who reviews what",
    )
    # 生まれた時点で次の一手を持たせる。これがあれば `Next Actions` 空の検査に
    # 免除が要らなくなり、時刻にも git にも依存しなくなる。免除の条件だった
    # `updated_at` は人手で書き換える値で、更新を忘れた作業単位は着手済みでも
    # 永久に免除され続けた（実測 2026-09-20: 477 件中 56 件・13 リポジトリ）。
    parser.add_argument(
        "--next-action",
        required=True,
        help="The very next command or step (or what unblocks a blocked issue)",
    )
    # Ownership decides the route (SPEC-doc-governance): an owned issue is listed by
    # its workstream, a standalone one by the index. Declaring it here and writing the
    # route row in the same run is what keeps a new issue from being born unrouted.
    owner = parser.add_mutually_exclusive_group(required=True)
    owner.add_argument("--workstream", help="Owning workstream id, e.g. WS-20260923-example")
    owner.add_argument("--standalone", action="store_true", help="Not under any workstream's envelope")
    parser.add_argument("--priority", choices=ownership.PRIORITIES, help="Required with --standalone")
    parser.add_argument("--due", help="YYYY-MM-DD or none; required with --standalone")
    args = parser.parse_args()

    if args.standalone and (not args.priority or not args.due):
        parser.error("--standalone requires --priority and --due (use --due none when there is no deadline)")
    if args.due and not ownership.valid_due(args.due):
        parser.error("--due must be YYYY-MM-DD or none")

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

    ws_path = None
    if args.workstream:
        ws_path = Path(repo) / "docs/workstreams" / f"{args.workstream}.md"
        if not ws_path.is_file():
            print(f"Error: {args.workstream} is not an active workstream ({ws_path} not found)", file=sys.stderr)
            sys.exit(1)

    today_iso = date.today().isoformat()
    title = args.title or slug.replace("-", " ").title()
    guide_impact = "required" if args.guide else "none"
    related_guides = f"[{args.guide}]" if args.guide else "[]"
    guide_reason = "" if args.guide else args.no_guide_reason.replace('"', "'")
    if args.verify_machine:
        verify_line = f"machine — {args.verify_machine}"
    else:
        verify_line = f"human-review — {args.verify_human}"

    try:
        tmpl_path = os.path.join(TEMPLATE_DIR, "issue.template.md")
        with open(tmpl_path) as f:
            content = f.read()
        content = content.replace("ISSUE-YYYYMMDD-short-slug", issue_id)
        content = content.replace("YYYY-MM-DD", today_iso)
        content = content.replace("# Title", f"# {title}", 1)
        content = content.replace("related_guides: []", f"related_guides: {related_guides}", 1)
        content = content.replace("guide_impact: required", f"guide_impact: {guide_impact}", 1)
        content = content.replace(
            'guide_impact_reason: ""', f'guide_impact_reason: "{guide_reason}"', 1
        )
        content = content.replace(
            "- Decision: required | none", f"- Decision: {guide_impact}", 1
        )
        content = content.replace("- verify:", f"- verify: {verify_line}", 1)
        content = content.replace(
            "## Next Actions\n", f"## Next Actions\n\n1. {args.next_action}\n", 1
        )
    except FileNotFoundError:
        content = (
            f"---\nid: {issue_id}\nstatus: active\n"
            f"created_at: {today_iso}\nupdated_at: {today_iso}\n"
            f"branch: {issue_id}\npr: \"\"\n"
            f"related_specs: []\nrelated_guides: {related_guides}\n"
            f"guide_impact: {guide_impact}\n"
            f'guide_impact_reason: "{guide_reason}"\n---\n\n'
            f"# {title}\n\n## Goal\n\n"
            f"## Acceptance\n\n- verify: {verify_line}\n\n"
            f"## Current Status\n\n## Next Actions\n\n## Notes\n\n"
            "## Completion\n\n"
            "- [ ] Implementation completed or intentionally not needed\n"
            "- [ ] Specs updated if direction or requirements changed\n"
            "- [ ] Guides updated if implemented behavior changed\n"
            "- [ ] Branch merged and cleaned up (or intentionally kept — note why)\n"
            "- [ ] 00_index.md updated\n"
            "- [ ] Moved to docs/issues/archive/ when complete\n"
        )

    ownership_lines = f"workstream: {args.workstream or 'none'}\n"
    if args.standalone:
        ownership_lines += f"priority: {args.priority}\ndue: {args.due}\n"
    content = content.replace('workstream: ""\n', "", 1)
    content = content.replace("status: active\n", f"status: active\n{ownership_lines}", 1)

    with open(dest, "w") as f:
        f.write(content)

    try:
        route = write_route(Path(repo), ws_path)
    except Exception as exc:  # the issue must not survive without its route row
        os.remove(dest)
        print(f"Error: could not write the route row, issue not created: {exc}", file=sys.stderr)
        sys.exit(1)

    print(f"Created: {dest}")
    print(f"Routed from: {route}")
    print(f"\nSuggested branch (convention: branch name = issue id):")
    print(f"  git checkout -b {issue_id}")
    print(f"  # or, for a separate worktree: git worktree add ../{issue_id} -b {issue_id}")


def write_route(repo: Path, ws_path) -> str:
    """Regenerate the one list this issue is routed from; returns its path."""
    model = ownership.load_model(repo, parse_front_matter)
    if ws_path is not None:
        rows = ownership.split_rows(model, ws_path.stem)
        ws_path.write_text(
            ownership.write_block(ws_path.read_text(), ownership.SPLIT, rows, ownership.HEADINGS[ownership.SPLIT])
        )
        return str(ws_path.relative_to(repo))
    index = repo / "docs/00_index.md"
    rows = ownership.active_issue_rows(model)
    index.write_text(
        ownership.write_block(index.read_text(), ownership.ACTIVE_ISSUES, rows, ownership.HEADINGS[ownership.ACTIVE_ISSUES])
    )
    return "docs/00_index.md"


def re_slug_ok(slug: str) -> bool:
    import re
    return bool(re.match(r"^[a-z0-9][a-z0-9-]*[a-z0-9]$", slug) or re.match(r"^[a-z0-9]$", slug))


if __name__ == "__main__":
    main()
