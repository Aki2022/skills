"""A target the router cannot classify must stop, not be checked as an LP.

SPEC-deliverable-production-layers (biz_ops, Requirements 6): the old fallback
"迷ったら LP として扱う" ran the LP checks on whatever was handed over and reported
them as if they applied, which is a silent wrong result. The router now stops and
asks, or reports the target as unchecked.

The router is a prose skill run by an agent, so the contract lives in SKILL.md.
The red case is a SKILL.md that still carries the LP fallback or lacks the stop row.
"""
import re
import unittest
from pathlib import Path

SKILL = Path(__file__).resolve().parents[2] / "SKILL.md"

LP_FALLBACK = re.compile(r"(迷|判定でき|分から)[^\n。]*LP\s*として(扱う|検査する|進める)")


def step1(text: str) -> str:
    start = text.index("### 1. 対象を判定する")
    end = text.index("### 2.", start)
    return text[start:end]


def undetermined_rows(section: str) -> list[str]:
    return [line for line in section.splitlines()
            if line.startswith("|") and "判定できない" in line]


def contract_holds(text: str) -> list[str]:
    """Return the broken parts of the contract (empty list = holds)."""
    broken = []
    section = step1(text)
    if LP_FALLBACK.search(text):
        broken.append("LP への既定の振り分けが残っている")
    rows = undetermined_rows(section)
    if len(rows) != 1:
        broken.append(f"判定できない対象の行が 1 つでない（{len(rows)} 行）")
    else:
        row = rows[0]
        if "止" not in row or "聞" not in row:
            broken.append("判定できない対象の行が、止まって聞くと書いていない")
        if "確かめていない" not in row:
            broken.append("判定できない対象の行が、確かめていないと報告すると書いていない")
    return broken


class UndeterminedTargetStops(unittest.TestCase):
    def test_skill_md_keeps_the_stop_contract(self):
        self.assertEqual(contract_holds(SKILL.read_text(encoding="utf-8")), [])

    def test_the_old_lp_fallback_is_red(self):
        old = SKILL.read_text(encoding="utf-8").replace(
            "### 1. 対象を判定する",
            "### 1. 対象を判定する\n\n迷ったら聞かずに「LP として扱う」と宣言して進める（検査が最も厚いため）。\n",
            1,
        )
        self.assertIn("LP への既定の振り分けが残っている", contract_holds(old))

    def test_a_missing_stop_row_is_red(self):
        text = SKILL.read_text(encoding="utf-8")
        stripped = "\n".join(line for line in text.splitlines() if "判定できない" not in line)
        self.assertTrue(contract_holds(stripped))


if __name__ == "__main__":
    unittest.main()
