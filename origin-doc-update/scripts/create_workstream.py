#!/usr/bin/env python3
"""Create a human-bounded workstream after its authorization envelope is confirmed."""
from __future__ import annotations

import argparse
import re
import sys
from datetime import date
from pathlib import Path


SKILL_DIR = Path(__file__).resolve().parent.parent
TEMPLATE = SKILL_DIR / "references/workstream.template.md"


def valid_slug(value: str) -> str:
    slug = value.lower().replace(" ", "-")
    if not re.fullmatch(r"[a-z0-9]+(?:-[a-z0-9]+)*", slug):
        raise argparse.ArgumentTypeError("use lowercase letters, numbers, and hyphens")
    return slug


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Create a workstream only after human boundaries are confirmed."
    )
    parser.add_argument("slug", type=valid_slug)
    parser.add_argument("--issue", required=True, type=valid_slug, help="Initial vertical-slice issue slug")
    parser.add_argument("--title", default="")
    parser.add_argument("--scope", required=True, help="Human-approved scope summary")
    parser.add_argument("--confirmed-at", required=True, help="Confirmation date, YYYY-MM-DD")
    parser.add_argument("--next-human-gate", required=True, help="Named next human checkpoint")
    parser.add_argument("--date", default=None, help="ID date override, YYYYMMDD")
    parser.add_argument("--repo", default=".")
    impact = parser.add_mutually_exclusive_group(required=True)
    impact.add_argument("--guide", help="Guide ID updated by the initial issue")
    impact.add_argument("--no-guide-reason", help="Why the initial issue changes no implemented behavior")
    args = parser.parse_args()

    if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", args.confirmed_at):
        parser.error("--confirmed-at must be YYYY-MM-DD")
    date_str = args.date or date.today().strftime("%Y%m%d")
    if not re.fullmatch(r"\d{8}", date_str):
        parser.error("--date must be YYYYMMDD")

    repo = Path(args.repo).resolve()
    workstreams_dir = repo / "docs/workstreams"
    if not workstreams_dir.is_dir():
        print("Error: docs/workstreams does not exist; run init_repo_docs.py first", file=sys.stderr)
        raise SystemExit(1)

    workstream_id = f"WS-{date_str}-{args.slug}"
    issue_id = f"ISSUE-01-{args.issue}"
    destination = workstreams_dir / f"{workstream_id}.md"
    if destination.exists():
        print(f"Error: already exists: {destination}", file=sys.stderr)
        raise SystemExit(1)

    title = args.title or args.slug.replace("-", " ").title()
    guide_impact = "required" if args.guide else "none"
    related_guides = f"[{args.guide}]" if args.guide else "[]"
    guide_reason = "" if args.guide else args.no_guide_reason.replace('"', "'")

    content = TEMPLATE.read_text()
    replacements = {
        "WS-YYYYMMDD-short-slug": workstream_id,
        "YYYY-MM-DD": args.confirmed_at,
        "short-gate-name": args.next_human_gate,
        "# Title": f"# {title}",
        "ISSUE-01-short-slug": issue_id,
        "One vertical slice": args.issue.replace("-", " "),
        "- Approved scope:": f"- Approved scope: {args.scope}",
        "- guide_impact: required": f"- guide_impact: {guide_impact}",
        "- related_guides: [GUIDE-short-slug]": f"- related_guides: {related_guides}",
        '- guide_impact_reason: ""': f'- guide_impact_reason: "{guide_reason}"',
    }
    for old, new in replacements.items():
        content = content.replace(old, new)
    destination.write_text(content)

    print(f"Created: {destination}")
    print("Add to docs/00_index.md Active Workstreams:")
    print(f"  - docs/workstreams/{workstream_id}.md — {title}")
    print("Suggested branch:")
    print(f"  git checkout -b {workstream_id}")
    if args.guide:
        print(f"Add {workstream_id} to {args.guide}'s source_workstreams before completing the issue.")


if __name__ == "__main__":
    main()
