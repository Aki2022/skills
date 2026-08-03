#!/usr/bin/env python3
"""Create a repository ADR from the origin-doc-update template."""
from __future__ import annotations

import argparse
import re
import sys
from datetime import date
from pathlib import Path


SKILL_DIR = Path(__file__).resolve().parent.parent
TEMPLATE = SKILL_DIR / "references/adr.template.md"


def valid_slug(value: str) -> str:
    slug = value.lower().replace(" ", "-")
    if not re.fullmatch(r"[a-z0-9]+(?:-[a-z0-9]+)*", slug):
        raise argparse.ArgumentTypeError("use lowercase letters, numbers, and hyphens")
    return slug


def main() -> None:
    parser = argparse.ArgumentParser(description="Create a repository ADR.")
    parser.add_argument("slug", type=valid_slug, help="Short ADR slug")
    parser.add_argument("--scope", required=True, choices=("spec", "development"))
    parser.add_argument(
        "--status", default="proposed", choices=("proposed", "accepted", "rejected")
    )
    parser.add_argument("--title", default="", help="Decision title")
    parser.add_argument("--date", default=None, help="ID date override, YYYYMMDD")
    parser.add_argument("--repo", default=".", help="Repository root (default: cwd)")
    args = parser.parse_args()

    date_str = args.date or date.today().strftime("%Y%m%d")
    if not re.fullmatch(r"\d{8}", date_str):
        parser.error("--date must be YYYYMMDD")

    repo = Path(args.repo).resolve()
    adrs_dir = repo / "docs/adrs"
    adrs_dir.mkdir(parents=True, exist_ok=True)

    adr_id = f"ADR-{date_str}-{args.slug}"
    destination = adrs_dir / f"{adr_id}.md"
    if destination.exists():
        print(f"Error: already exists: {destination}", file=sys.stderr)
        raise SystemExit(1)

    today_iso = date.today().isoformat()
    title = args.title or args.slug.replace("-", " ").title()
    content = TEMPLATE.read_text()
    replacements = {
        "ADR-YYYYMMDD-short-slug": adr_id,
        "status: proposed": f"status: {args.status}",
        "scope: spec": f"scope: {args.scope}",
        "YYYY-MM-DD": today_iso,
        "# Decision: Title": f"# Decision: {title}",
    }
    for old, new in replacements.items():
        content = content.replace(old, new)
    destination.write_text(content)

    print(f"Created: {destination}")
    print("Add source_workstreams/source_issues and related_specs/related_guides links.")
    print("Add the ADR to the relevant spec, workstream/issue, or guide when applicable.")


if __name__ == "__main__":
    main()
