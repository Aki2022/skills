import importlib.util
import json
import tempfile
import unittest
from pathlib import Path

SCRIPT = Path(__file__).resolve().parents[1] / "plan.py"
SPEC = importlib.util.spec_from_file_location("plan", SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
SPEC.loader.exec_module(MODULE)

REVIEW = """---
updated_at: 2026-09-21
kind: review
---

# docs review 2026-09-21

## Demote candidates

- [ ] archive `docs/specs/old.md`
  - no living document links here

## Split candidates

- [ ] split `docs/guides/big-hot.md` — 90KB, dated headings 12
  - 30KB  Run
  - 40KB  Report
  - 20KB  Lexicon
- [ ] split `docs/guides/big-cold.md` — 120KB, dated headings 0
  - 118KB  Everything
  - 1KB  Notes
- [ ] split `docs/specs/mid-warm.md` — 40KB, dated headings 7
  - 20KB  Policy
  - 20KB  History

## All current-truth documents

| tier | file | KB | last commit | linked from |
|---|---|---|---|---|
| hot | `docs/guides/big-hot.md` | 90 | 2026-09-01 | issues×2 |
| warm | `docs/specs/mid-warm.md` | 40 | 2026-08-01 | index×1 |
| cold | `docs/guides/big-cold.md` | 120 | 2026-03-01 | — |
| cold | `docs/specs/old.md` | 3 | 2026-01-01 | — |
"""


class PlanTest(unittest.TestCase):
    def setUp(self):
        self.root = Path(tempfile.mkdtemp())
        (self.root / "docs/log").mkdir(parents=True)
        (self.root / "docs/log/review-20260921.md").write_text(REVIEW)
        self.hygiene = {
            "repo": str(self.root), "review_path": "docs/log/review-20260921.md",
            "report": {
                "R1_stale_branch_gone": {"count": 3}, "R2_untouched_60d": {"count": 0},
                "R7_says_done_but_active": {"count": 2},
                "R9_context_budget": {"items": [{"digest_dropped_entries": 5, "oversized_docs": 3, "hot_docs_kb": 400}]},
            },
        }

    def test_budget_orders_hot_first_and_carries_the_rest(self):
        p = MODULE.plan(self.hygiene, self.root / "docs/log/review-20260921.md", 2)
        self.assertEqual(p["judge"], 1)
        self.assertEqual(p["sweep_candidates"], 5)
        self.assertEqual(p["demote_candidates"], 1)
        self.assertEqual([d["file"] for d in p["docs"]], ["docs/guides/big-hot.md", "docs/specs/mid-warm.md"])
        self.assertEqual(p["docs"][0]["mode"], "split")      # three sections over 4KB
        self.assertEqual(p["docs"][1]["mode"], "condense")  # a spec condenses, never splits
        self.assertEqual(p["carry"], 1)
        self.assertEqual(p["context_budget_before"]["digest_dropped_entries"], 5)

    def test_single_giant_section_is_condensed_not_split(self):
        p = MODULE.plan(self.hygiene, self.root / "docs/log/review-20260921.md", 3)
        cold = [d for d in p["docs"] if d["file"] == "docs/guides/big-cold.md"][0]
        self.assertEqual(cold["mode"], "condense")
        self.assertEqual(p["carry"], 0)

    def test_nothing_to_do_is_zero_not_missing(self):
        hygiene = {"repo": "x", "review_path": None, "report": {}}
        p = MODULE.plan(hygiene, Path("/nonexistent"), 2)
        self.assertEqual((p["judge"], p["docs"], p["carry"]), (0, [], 0))
        self.assertEqual(p["context_budget_before"]["oversized_docs"], None)


if __name__ == "__main__":
    unittest.main()
