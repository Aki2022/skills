"""Positive and negative controls for repository instruction topology."""

from __future__ import annotations

import importlib.util
import io
import os
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from tempfile import TemporaryDirectory


SCRIPT = Path(__file__).resolve().parents[1] / "check_repo_instruction_topology.py"
SPEC = importlib.util.spec_from_file_location("check_repo_instruction_topology", SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
SPEC.loader.exec_module(MODULE)


class RepoInstructionTopologyTest(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = TemporaryDirectory()
        self.repo = Path(self.tmp.name) / "repo"
        self.repo.mkdir()
        (self.repo / "AGENTS.md").write_text("project rules\n", encoding="utf-8")

    def tearDown(self) -> None:
        self.tmp.cleanup()

    def _run(self, mode: str) -> tuple[int, str]:
        output = io.StringIO()
        with redirect_stdout(output):
            rc = MODULE.main(["--repo", str(self.repo), "--mode", mode])
        return rc, output.getvalue()

    def test_native_without_claude_is_green(self):
        rc, output = self._run("native")
        self.assertEqual(rc, 0, output)
        self.assertIn("RESULT: OK", output)

    def test_compat_symlink_is_green(self):
        os.symlink("AGENTS.md", self.repo / "CLAUDE.md")
        rc, output = self._run("compat")
        self.assertEqual(rc, 0, output)
        self.assertIn("RESULT: OK", output)

    def test_native_rejects_real_claude_file(self):
        (self.repo / "CLAUDE.md").write_text("duplicate\n", encoding="utf-8")
        rc, output = self._run("native")
        self.assertEqual(rc, 1, output)
        self.assertIn("requires CLAUDE.md absent", output)

    def test_native_rejects_compatibility_symlink(self):
        os.symlink("AGENTS.md", self.repo / "CLAUDE.md")
        rc, output = self._run("native")
        self.assertEqual(rc, 1, output)
        self.assertIn("requires CLAUDE.md absent", output)

    def test_native_detects_case_only_claude_spelling(self):
        (self.repo / "claude.md").write_text("duplicate\n", encoding="utf-8")
        rc, output = self._run("native")
        self.assertEqual(rc, 1, output)
        self.assertIn("requires CLAUDE.md absent", output)

    def test_compat_rejects_broken_symlink(self):
        os.symlink("missing.md", self.repo / "CLAUDE.md")
        rc, output = self._run("compat")
        self.assertEqual(rc, 1, output)
        self.assertIn("broken symlink", output)

    def test_rejects_noncanonical_agents_file(self):
        agents = self.repo / "AGENTS.md"
        agents.unlink()
        os.symlink("other.md", agents)
        rc, output = self._run("native")
        self.assertEqual(rc, 1, output)
        self.assertIn("AGENTS.md must be a regular file", output)

    def test_rejects_shadowing_files(self):
        (self.repo / "AGENTS.override.md").write_text("override\n", encoding="utf-8")
        (self.repo / "CLAUDE.local.md").write_text("local\n", encoding="utf-8")
        rc, output = self._run("native")
        self.assertEqual(rc, 1, output)
        self.assertIn("AGENTS.override.md", output)
        self.assertIn("CLAUDE.local.md", output)


if __name__ == "__main__":
    unittest.main()
