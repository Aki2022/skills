#!/usr/bin/env python3
"""Turn a hygiene result into this session's bounded curation plan.

Input: the JSON printed by `docs_hygiene.py <repo> --fix --report --json`, plus the
review file it wrote. Output (JSON on stdout):

  judge   1 when the sweep or the review has archive candidates, else 0
  docs    up to --budget-docs guides/specs to split or condense this session,
          hot before warm before cold, larger first; `mode` is "split" when the
          review lists two or more H2 sections over 4 KB, otherwise "condense"
  carry   candidates left for the next close-session

The budget exists because a split costs 150-280k subagent tokens (measured
2026-09-20/21); unbounded, one close-session on a 12-oversized-doc repository
would spend millions. Two per session drains that repository in six sessions
while the review keeps re-listing what is left, so nothing is dropped.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

TIER_RANK = {"hot": 0, "warm": 1, "cold": 2}
BIG_SECTION_BYTES = 4 * 1024


def parse_review(path: Path) -> tuple[dict[str, dict], list[dict], list[str]]:
    """Return (tier by file, split candidates with sections, demote candidate files)."""
    text = path.read_text() if path.is_file() else ""
    tiers: dict[str, dict] = {}
    for m in re.finditer(r"^\| (hot|warm|cold) \| `([^`]+)` \| (\d+) \|", text, re.MULTILINE):
        tiers[m.group(2)] = {"tier": m.group(1), "kb": int(m.group(3))}
    splits: list[dict] = []
    block = re.search(r"^## Split candidates\n(.*?)(?=^## )", text, re.MULTILINE | re.DOTALL)
    if block:
        current = None
        for line in block.group(1).splitlines():
            m = re.match(r"^- \[[ xX]\] split `([^`]+)` — (\d+)KB, dated headings (\d+)", line)
            if m:
                current = {"file": m.group(1), "kb": int(m.group(2)), "dated_headings": int(m.group(3)), "sections": []}
                splits.append(current)
                continue
            m = re.match(r"^  - (\d+)KB  (.+)$", line)
            if m and current is not None:
                current["sections"].append({"kb": int(m.group(1)), "heading": m.group(2)})
    demote = re.findall(r"^- \[ \] archive `([^`]+)`", text, re.MULTILINE)
    return tiers, splits, demote


def plan(hygiene: dict, review_path: Path, budget_docs: int) -> dict:
    tiers, splits, demote = parse_review(review_path)
    report = hygiene.get("report", {})
    sweep_candidates = sum(
        report.get(k, {}).get("count", 0) for k in ("R1_stale_branch_gone", "R2_untouched_60d", "R7_says_done_but_active")
    )
    judge = 1 if (sweep_candidates or demote) else 0

    ranked = sorted(
        splits,
        key=lambda s: (TIER_RANK.get(tiers.get(s["file"], {}).get("tier", "cold"), 2), -s["kb"]),
    )
    docs = []
    for s in ranked[:budget_docs]:
        big = [x for x in s["sections"] if x["kb"] * 1024 >= BIG_SECTION_BYTES]
        kind = "spec" if s["file"].startswith("docs/specs/") else "guide"
        # A spec is one policy document: its bulk is decision records and revision
        # history, which condense into current policy + ADRs + history rather than
        # splitting into several specs. A guide splits by task when several
        # sections are large, and condenses when one section carries the bulk.
        mode = "condense" if kind == "spec" or len(big) < 2 else "split"
        docs.append({
            "file": s["file"],
            "kind": kind,
            "kb": s["kb"],
            "tier": tiers.get(s["file"], {}).get("tier", "cold"),
            "mode": mode,
            "sections": s["sections"],
        })
    r9 = report.get("R9_context_budget", {}).get("items", [{}])[0]
    return {
        "repo": hygiene.get("repo"),
        "judge": judge,
        "sweep_candidates": sweep_candidates,
        "demote_candidates": len(demote),
        "docs": docs,
        "carry": max(0, len(ranked) - budget_docs),
        "context_budget_before": {
            "digest_dropped_entries": r9.get("digest_dropped_entries"),
            "oversized_docs": r9.get("oversized_docs"),
            "hot_docs_kb": r9.get("hot_docs_kb"),
        },
    }


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("repo")
    ap.add_argument("hygiene_json", help="output of docs_hygiene.py --fix --report --json")
    ap.add_argument("--budget-docs", type=int, default=2)
    args = ap.parse_args()
    hygiene = json.loads(Path(args.hygiene_json).read_text())
    review = hygiene.get("review_path")
    review_path = Path(args.repo) / review if review else Path("/nonexistent")
    print(json.dumps(plan(hygiene, review_path, args.budget_docs), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
