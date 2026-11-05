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


# uv の代わりに PATH 先頭へ置く stub。受け取った引数をログに残し、
# `--` の後ろのコマンドをそのまま実行する（失敗させたい時は UV_STUB_EXIT に非0を入れる）。
UV_STUB = textwrap.dedent("""\
    #!/bin/sh
    printf '%s\\n' "$*" >> "$UV_STUB_LOG"
    if [ "${UV_STUB_EXIT:-0}" != "0" ]; then exit "$UV_STUB_EXIT"; fi
    while [ $# -gt 0 ]; do
        if [ "$1" = "--" ]; then shift; exec "$@"; fi
        shift
    done
    echo "uv stub: no -- separator" >&2
    exit 97
    """)


def _install_uv_stub(bin_dir: Path) -> Path:
    bin_dir.mkdir(parents=True, exist_ok=True)
    stub = bin_dir / "uv"
    stub.write_text(UV_STUB)
    stub.chmod(0o755)
    return bin_dir


def _uv_env(bin_dir: Path, log: Path, exit_code: int = 0) -> dict:
    return {
        "PATH": f"{bin_dir}{os.pathsep}{os.environ['PATH']}",
        "UV_STUB_LOG": str(log),
        "UV_STUB_EXIT": str(exit_code),
    }


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

    def test_requirements_file_runs_pytest_through_uv(self):
        """scripts/tests/requirements.txt を持つ skill は uv run（--no-project、依存宣言つき）で pytest を走らせる。"""
        with TemporaryDirectory() as tmp:
            root = Path(tmp) / "skills"
            _write_skill_md(root, "fake-skill-deps")
            tests = root / "fake-skill-deps" / "scripts" / "tests"
            _write(tests / "test_x.py", PASSING_TEST)
            _write(tests / "requirements.txt", "pytest==9.2.0\n")
            log = Path(tmp) / "uv.log"
            bin_dir = _install_uv_stub(Path(tmp) / "bin")

            rc, out = self._run(root, env_extra=_uv_env(bin_dir, log))

            self.assertEqual(rc, 0, out)
            self.assertIn("fake-skill-deps", out)
            logged = log.read_text()
            self.assertIn("--no-project", logged)
            self.assertIn("--with-requirements", logged)
            self.assertIn("requirements.txt", logged)

    def test_uv_failure_is_red_not_skipped(self):
        """依存付きの skill で uv が失敗したら、黙って合格にせず S8 を FAIL にする。"""
        with TemporaryDirectory() as tmp:
            root = Path(tmp) / "skills"
            _write_skill_md(root, "fake-skill-deps")
            tests = root / "fake-skill-deps" / "scripts" / "tests"
            _write(tests / "test_x.py", PASSING_TEST)
            _write(tests / "requirements.txt", "pytest==9.2.0\n")
            log = Path(tmp) / "uv.log"
            bin_dir = _install_uv_stub(Path(tmp) / "bin")

            rc, out = self._run(root, env_extra=_uv_env(bin_dir, log, exit_code=3))

            self.assertEqual(rc, 1, out)
            self.assertIn("S8", out)
            self.assertIn("fake-skill-deps", out)


if __name__ == "__main__":
    unittest.main()
