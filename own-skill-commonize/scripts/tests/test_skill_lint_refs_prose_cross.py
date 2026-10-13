"""skill_lint_refs の cross 判定が、散文の「他 skill の参照」を拾うことの回帰テスト。

欠陥（2026-10-13 に発見）: cross 判定は `/skill` 表記だけを他 skill として扱った。
cloudflare の SKILL.md は `` `agents-sdk` skill, its `references/mcp.md` `` と書いており、
これが cloudflare 自身の `references/mcp.md` と解決されて S4 が偽陽性で FAIL した。
参照先 `agents-sdk/references/mcp.md` は実在する。

修正の方針: 散文でも `` `<canonical-skill>` skill `` の形を skill トークンとして扱い、
直後の resource を、その skill に属するものとして cross で出す。
既存の `/skill/group/path` と、skill を名指ししない同 skill 参照の出力は変えない。
"""
from __future__ import annotations

import subprocess
import sys
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

PARSER = Path(__file__).resolve().parents[1] / "skill_lint_refs.py"


def _scan(skill_md_text: str, skills: list[str]) -> list[str]:
    with TemporaryDirectory() as tmp:
        root = Path(tmp)
        for name in skills:
            (root / name).mkdir()
        md = root / "SKILL.md"
        md.write_text(skill_md_text, encoding="utf-8")
        out = subprocess.run(
            [sys.executable, str(PARSER), str(root), str(md)],
            capture_output=True, text=True, timeout=30,
        )
        if out.returncode != 0:
            raise AssertionError(f"parser failed: {out.stderr}")
        return [line for line in out.stdout.splitlines() if line]


class ProseCrossReferenceTest(unittest.TestCase):
    def test_backticked_skill_name_with_skill_word_is_cross(self) -> None:
        # cloudflare の実際の書き方（表のセル内）。
        line = "| Expose tools | `agents-sdk` skill, its `references/mcp.md` | x |"
        got = _scan(line + "\n", ["cloudflare", "agents-sdk"])
        self.assertIn("cross\tagents-sdk\treferences/mcp.md", got)
        self.assertNotIn("same\treferences/mcp.md", got)

    def test_slash_form_still_cross(self) -> None:
        got = _scan("see /agents-sdk/references/mcp.md\n", ["cloudflare", "agents-sdk"])
        self.assertIn("cross\tagents-sdk\treferences/mcp.md", got)

    def test_same_skill_reference_without_skill_name_stays_same(self) -> None:
        got = _scan("read `references/guide.md` first\n", ["cloudflare", "agents-sdk"])
        self.assertIn("same\treferences/guide.md", got)
        self.assertFalse(any(x.startswith("cross") for x in got))

    def test_unknown_backticked_word_is_not_a_skill(self) -> None:
        # canonical に無い語は skill 扱いしない（従来どおり same）。
        got = _scan("the `widget` skill, `references/guide.md`\n", ["cloudflare"])
        self.assertIn("same\treferences/guide.md", got)


if __name__ == "__main__":
    unittest.main()
