"""skill_lint.sh の S8（各 skill の pytest 実行）を陽性対照で確かめるテスト。

実装前にこのテストを走らせて赤を見ること。偽の skills root を一時ディレクトリに作り、
skill_lint.sh <root> を subprocess で呼んで exit code と出力を検証する。

偽 skill には SKILL.md（name: と description: がディレクトリ名と一致）を必ず置く —
そうしないと S1〜S3 で FAIL し、S8 の検証にならない。
"""
from __future__ import annotations

import os
import subprocess
import textwrap
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

SCRIPTS_DIR = Path(__file__).resolve().parents[1]
SKILL_LINT = SCRIPTS_DIR / "skill_lint.sh"


def _write(p: Path, text: str) -> None:
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(text)


def _write_skill_md(root: Path, name: str) -> None:
    _write(
        root / name / "SKILL.md",
        textwrap.dedent(f"""\
            ---
            name: {name}
            description: fake skill for S8 test
            ---
            body
            """),
    )


FAILING_TEST = textwrap.dedent(
    """\
    def test_always_fails():
        assert False, "intentional failure for S8 positive control"
    """
)

PASSING_TEST = textwrap.dedent(
    """\
    def test_always_passes():
        assert True
    """
)


class SkillLintPytestTest(unittest.TestCase):
    def _run(self, root: Path, env_extra: dict | None = None):
        env = dict(os.environ)
        env.pop("SKILL_LINT_SKIP_PYTEST", None)
        if env_extra:
            env.update(env_extra)
        proc = subprocess.run(
            ["bash", str(SKILL_LINT), str(root)],
            capture_output=True,
            text=True,
            env=env,
        )
        return proc.returncode, proc.stdout + proc.stderr

    def test_positive_control_failing_pytest_is_red(self):
        """陽性対照: 失敗するテストを持つ偽 skill → exit 1、出力に S8 と skill 名。"""
        with TemporaryDirectory() as tmp:
            root = Path(tmp) / "skills"
            _write_skill_md(root, "fake-skill-fail")
            _write(root / "fake-skill-fail" / "scripts" / "tests" / "test_x.py", FAILING_TEST)

            rc, out = self._run(root)

            self.assertEqual(rc, 1, out)
            self.assertIn("S8", out)
            self.assertIn("fake-skill-fail", out)

    def test_passing_pytest_is_green(self):
        """通るテストを持つ偽 skill → exit 0、出力に S8 の成功行。"""
        with TemporaryDirectory() as tmp:
            root = Path(tmp) / "skills"
            _write_skill_md(root, "fake-skill-pass")
            _write(root / "fake-skill-pass" / "scripts" / "tests" / "test_x.py", PASSING_TEST)

            rc, out = self._run(root)

            self.assertEqual(rc, 0, out)
            self.assertIn("S8", out)
            self.assertIn("fake-skill-pass", out)

    def test_no_skill_has_tests_is_explicit(self):
        """scripts/tests/ を持つ skill が無い root → exit 0、無いことを明示する行。"""
        with TemporaryDirectory() as tmp:
            root = Path(tmp) / "skills"
            _write_skill_md(root, "fake-skill-no-tests")

            rc, out = self._run(root)

            self.assertEqual(rc, 0, out)
            self.assertIn("S8: pytest を持つ skill が無い", out)

    def test_skip_env_var_bypasses_even_failing_tests(self):
        """SKILL_LINT_SKIP_PYTEST=1 → skip の旨を出力し、失敗するテストがあっても exit 0。"""
        with TemporaryDirectory() as tmp:
            root = Path(tmp) / "skills"
            _write_skill_md(root, "fake-skill-fail")
            _write(root / "fake-skill-fail" / "scripts" / "tests" / "test_x.py", FAILING_TEST)

            rc, out = self._run(root, env_extra={"SKILL_LINT_SKIP_PYTEST": "1"})

            self.assertEqual(rc, 0, out)
            self.assertIn("S8", out)
            self.assertIn("skip", out.lower())


if __name__ == "__main__":
    unittest.main()
