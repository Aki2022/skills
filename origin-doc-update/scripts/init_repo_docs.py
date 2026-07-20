#!/usr/bin/env python3
"""Initialize doc-governance scaffold in a repository."""
import argparse
import os
import re
import sys
from datetime import date

SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE_DIR = os.path.join(SKILL_DIR, "references")


def read_template(name: str) -> str:
    path = os.path.join(TEMPLATE_DIR, name)
    with open(path) as f:
        return f.read()


def main():
    parser = argparse.ArgumentParser(description="Initialize docs scaffold in a repository.")
    parser.add_argument("repo", nargs="?", default=".", help="Repository root (default: cwd)")
    args = parser.parse_args()

    repo = os.path.abspath(args.repo)
    if not os.path.isdir(repo):
        print(f"Error: {repo} is not a directory", file=sys.stderr)
        sys.exit(1)

    print(f"Initializing docs scaffold in: {repo}")

    for subdir in [
        "docs/specs",
        "docs/issues/archive",
        "docs/workstreams/archive",
        "docs/guides",
    ]:
        d = os.path.join(repo, subdir)
        os.makedirs(d, exist_ok=True)
        entries = [e for e in os.listdir(d) if not e.startswith(".")]
        if not entries:
            gitkeep = os.path.join(d, ".gitkeep")
            if not os.path.exists(gitkeep):
                open(gitkeep, "w").close()
                print(f"  created: {os.path.relpath(gitkeep, repo)}")
        else:
            print(f"  exists:  {subdir}/")

    index_path = os.path.join(repo, "docs", "00_index.md")
    if os.path.exists(index_path):
        print(f"  skip (exists): docs/00_index.md")
    else:
        today = date.today().isoformat()
        try:
            content = read_template("00_index.template.md")
            content = content.replace("YYYY-MM-DD", today, 1)
            # Remove routing examples so the generated index starts clean.
            content = re.sub(
                r"^.*(?:WS-YYYYMMDD-short-slug|ISSUE-YYYYMMDD-short-slug).*$\n?",
                "",
                content,
                flags=re.MULTILINE,
            )
        except FileNotFoundError:
            content = (
                f"---\nupdated_at: {today}\ncurrent_focus: []\n---\n\n"
                "# 00 Index\n\n"
                "## Read Policy\n\nRead this file first. Do not scan all docs unless needed.\n\n"
                "## Active Workstreams\n\n## Specs\n\n## Active Issues\n\n## Guides\n\n"
                "## Archive Policy\n\n"
                "Completed workstreams and issues are stored under their archive directories. "
                "Archive files are historical context, not current truth.\n"
            )
        with open(index_path, "w") as f:
            f.write(content)
        print(f"  created: docs/00_index.md")

    print("Done.")


if __name__ == "__main__":
    main()
