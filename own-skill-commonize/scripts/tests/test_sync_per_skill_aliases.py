"""Tests for the explicit per-skill alias synchronizer."""

from __future__ import annotations

import importlib.util
import io
import os
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from tempfile import TemporaryDirectory


SCRIPT = Path(__file__).resolve().parents[1] / "sync_per_skill_aliases.py"
SPEC = importlib.util.spec_from_file_location("sync_per_skill_aliases", SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
SPEC.loader.exec_module(MODULE)


def _skill(root: Path, name: str) -> Path:
    path = root / name
    path.mkdir(parents=True, exist_ok=True)
    (path / "SKILL.md").write_text(
        f"---\nname: {name}\ndescription: fixture\n---\nbody\n",
        encoding="utf-8",
    )
    return path


class SyncPerSkillAliasesTest(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = TemporaryDirectory()
        root = Path(self.tmp.name)
        self.canonical = root / "canonical"
        self.aliases = root / "aliases"
        _skill(self.canonical, "alpha")
        _skill(self.canonical, "beta")
        self.aliases.mkdir()
        os.symlink(self.canonical / "alpha", self.aliases / "alpha")

    def tearDown(self) -> None:
        self.tmp.cleanup()

    def _run(self, *args: str) -> tuple[int, str]:
        output = io.StringIO()
        argv = [
            "--canonical",
            str(self.canonical),
            "--alias-root",
            str(self.aliases),
            *args,
        ]
        with redirect_stdout(output):
            rc = MODULE.main(argv)
        return rc, output.getvalue()

    def test_dry_run_reports_drift_without_mutating(self):
        os.symlink(Path(self.tmp.name) / "missing", self.aliases / "retired")

        rc, output = self._run("--prune-stale")

        self.assertEqual(rc, 1, output)
        self.assertIn("CREATE", output)
        self.assertIn("PRUNE", output)
        self.assertFalse((self.aliases / "beta").exists())
        self.assertTrue((self.aliases / "retired").is_symlink())

    def test_apply_creates_missing_and_prunes_broken_stale_link(self):
        os.symlink(Path(self.tmp.name) / "missing", self.aliases / "retired")

        rc, output = self._run("--apply", "--prune-stale")

        self.assertEqual(rc, 0, output)
        self.assertEqual((self.aliases / "beta").resolve(), (self.canonical / "beta").resolve())
        self.assertFalse((self.aliases / "retired").is_symlink())
        self.assertIn("RESULT: OK", output)

    def test_refuses_regular_entry_and_valid_foreign_link(self):
        (self.aliases / "beta").mkdir()
        foreign = _skill(Path(self.tmp.name) / "foreign", "gamma")
        os.symlink(foreign, self.aliases / "gamma")

        rc, output = self._run("--apply", "--prune-stale")

        self.assertEqual(rc, 2, output)
        self.assertIn("CONFLICT", output)
        self.assertTrue((self.aliases / "beta").is_dir())
        self.assertEqual((self.aliases / "gamma").resolve(), foreign.resolve())

    def test_apply_is_idempotent(self):
        first_rc, first_output = self._run("--apply")
        second_rc, second_output = self._run("--apply")

        self.assertEqual(first_rc, 0, first_output)
        self.assertEqual(second_rc, 0, second_output)
        self.assertIn("no changes", second_output)

    def test_explicit_ignored_entries_are_preserved(self):
        (self.aliases / ".system").mkdir()
        adapted = _skill(Path(self.tmp.name) / "adapted", "gamma")
        os.symlink(adapted, self.aliases / "gamma")

        rc, output = self._run(
            "--apply",
            "--ignore-entry",
            ".system",
            "--ignore-entry",
            "gamma",
        )

        self.assertEqual(rc, 0, output)
        self.assertTrue((self.aliases / ".system").is_dir())
        self.assertEqual((self.aliases / "gamma").resolve(), adapted.resolve())
        self.assertEqual((self.aliases / "beta").resolve(), (self.canonical / "beta").resolve())


if __name__ == "__main__":
    unittest.main()
