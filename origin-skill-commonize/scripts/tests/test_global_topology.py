"""Positive and negative controls for the global skill topology checker."""

from __future__ import annotations

import importlib.util
import io
import os
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from tempfile import TemporaryDirectory


SCRIPT = Path(__file__).resolve().parents[1] / "check_global_topology.py"
SPEC = importlib.util.spec_from_file_location("check_global_topology", SCRIPT)
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


def _link(source: Path, target: Path) -> None:
    target.parent.mkdir(parents=True, exist_ok=True)
    os.symlink(source, target, target_is_directory=True)


class GlobalTopologyTest(unittest.TestCase):
    def _fixture(self) -> tuple[Path, dict[str, Path]]:
        tmp = self.tmp.name
        canonical = Path(tmp) / "canonical"
        _skill(canonical, "alpha")
        _skill(canonical, "beta")

        adapted = Path(tmp) / "adapted" / "gamma"
        _skill(Path(tmp) / "adapted", "gamma")

        roots = {
            "canonical": canonical,
            "adapted": adapted,
            "claude": Path(tmp) / "claude" / "skills",
            "antigravity": Path(tmp) / "antigravity" / "skills",
            "codex": Path(tmp) / "codex" / "skills",
            "gemini": Path(tmp) / "gemini" / "skills",
        }
        _link(canonical, roots["claude"])
        _link(canonical, roots["antigravity"])
        roots["codex"].mkdir(parents=True)
        roots["gemini"].mkdir(parents=True)
        for root in (roots["codex"], roots["gemini"]):
            for name in ("alpha", "beta"):
                _link(canonical / name, root / name)
        _link(adapted, roots["codex"] / "gamma")
        (roots["codex"] / ".system").mkdir()
        _skill(roots["codex"], "grilling")
        return canonical, roots

    def setUp(self) -> None:
        self.tmp = TemporaryDirectory()

    def tearDown(self) -> None:
        self.tmp.cleanup()

    def _args(self, roots: dict[str, Path], *extra: str) -> list[str]:
        return [
            "--canonical",
            str(roots["canonical"]),
            "--claude",
            str(roots["claude"]),
            "--antigravity",
            str(roots["antigravity"]),
            "--codex",
            str(roots["codex"]),
            "--gemini",
            str(roots["gemini"]),
            "--codex-adapted",
            str(roots["adapted"]),
            "--local-only",
            str(roots["codex"] / "grilling"),
            *extra,
        ]

    def _run(self, args: list[str]) -> tuple[int, str]:
        output = io.StringIO()
        with redirect_stdout(output):
            rc = MODULE.main(args)
        return rc, output.getvalue()

    def test_valid_topology_with_allowed_exceptions_is_green(self):
        canonical, roots = self._fixture()
        rc, output = self._run(self._args(roots))
        self.assertEqual(rc, 0, output)
        self.assertIn("RESULT: OK", output)

    def test_regular_copy_is_red(self):
        _, roots = self._fixture()
        (roots["codex"] / "alpha").unlink()
        _skill(roots["codex"], "alpha")
        rc, output = self._run(self._args(roots))
        self.assertEqual(rc, 1, output)
        self.assertIn("expected per-skill symlink", output)

    def test_broken_link_is_red(self):
        _, roots = self._fixture()
        (roots["gemini"] / "beta").unlink()
        os.symlink(Path(self.tmp.name) / "missing", roots["gemini"] / "beta")
        rc, output = self._run(self._args(roots))
        self.assertEqual(rc, 1, output)
        self.assertIn("broken symlink", output)

    def test_missing_per_skill_link_is_red(self):
        _, roots = self._fixture()
        (roots["gemini"] / "beta").unlink()
        rc, output = self._run(self._args(roots))
        self.assertEqual(rc, 1, output)
        self.assertIn("missing per-skill symlink", output)

    def test_unregistered_entry_is_red(self):
        _, roots = self._fixture()
        _skill(roots["gemini"], "not-shared")
        rc, output = self._run(self._args(roots))
        self.assertEqual(rc, 1, output)
        self.assertIn("unregistered entry", output)

    def test_directory_alias_copy_is_red(self):
        canonical, roots = self._fixture()
        roots["claude"].unlink()
        roots["claude"].mkdir(parents=True)
        for name in ("alpha", "beta"):
            _link(canonical / name, roots["claude"] / name)
        rc, output = self._run(self._args(roots))
        self.assertEqual(rc, 1, output)
        self.assertIn("expected directory symlink", output)


if __name__ == "__main__":
    unittest.main()
