from __future__ import annotations

import os
import subprocess
import tempfile
import unittest
from pathlib import Path

SCRIPT = Path(__file__).resolve().parents[1] / "mirror_skill.sh"


def _skill(root: Path, name: str, body: str = "fixture") -> Path:
    path = root / name
    path.mkdir(parents=True, exist_ok=True)
    (path / "SKILL.md").write_text(
        f"---\nname: {name}\ndescription: fixture\n---\n{body}\n", encoding="utf-8"
    )
    return path


class MirrorSkillLocalTest(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.base = Path(self.tmp.name)
        self.home = self.base / "home"
        self.canonical = self.home / ".agents" / "skills"
        self.canonical.mkdir(parents=True)
        self.ledger = self.canonical / "mirrors.yaml"
        self.ledger.write_text(
            "mirrors: []\n\nretired:\n  - dir: retired-snapshot\n"
            "    upstream: example/skills\n    upstream_path: retired-snapshot\n"
            "    upstream_version: abc\n    fetched_at: 2026-01-01\n"
            "    license: Apache-2.0\n    reinstall: example\n",
            encoding="utf-8",
        )
        _skill(self.canonical, "own-fixture-run")
        _skill(self.canonical, "retired-snapshot", "retired body")
        self.roots = [
            self.home / ".codex" / "skills",
            self.home / ".codex-private" / "skills",
            self.home / ".codex-seat2" / "skills",
            self.home / ".gemini" / "config" / "skills",
        ]
        for root in self.roots:
            root.mkdir(parents=True)
            (root / "own-fixture-run").symlink_to(self.canonical / "own-fixture-run")
        self.source_root = self.base / "sources"
        self.source_root.mkdir()
        self.source = _skill(self.source_root, "sample-skill")
        self.env = os.environ.copy()
        self.env.update({"AGENTS_SKILLS_ROOT": str(self.canonical)})

    def tearDown(self) -> None:
        self.tmp.cleanup()

    def _run(self, name: str, source: Path | None = None, *, reactivate: bool = False):
        argv = [
            "bash", str(SCRIPT), "--upstream", "example/skills", "--path", name,
            "--name", name, "--license", "Apache-2.0", "--version", "abc",
        ]
        if source is not None:
            argv += ["--source-dir", str(source)]
        if reactivate:
            argv.append("--reactivate-retired")
        for root in self.roots:
            argv += ["--alias-root", str(root)]
        return subprocess.run(argv, env=self.env, text=True, capture_output=True, check=False)

    def test_import_local_and_reactivate_retired_excludes_retired_from_aliases(self):
        result = self._run("sample-skill", self.source)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertEqual((self.canonical / "sample-skill" / "SKILL.md").read_text(),
                         (self.source / "SKILL.md").read_text())
        for root in self.roots:
            self.assertEqual((root / "sample-skill").resolve(),
                             (self.canonical / "sample-skill").resolve())
            self.assertFalse((root / "retired-snapshot").exists())

        # A frontmatter name mismatch fails before writing a new destination.
        bad = _skill(self.source_root, "wrong-source")
        mismatch = self._run("different-name", bad)
        self.assertNotEqual(mismatch.returncode, 0)
        self.assertFalse((self.canonical / "different-name").exists())

        reactivated_source = _skill(self.source_root, "retired-snapshot", "fresh source")
        revived = self._run("retired-snapshot", reactivated_source, reactivate=True)
        self.assertEqual(revived.returncode, 0, revived.stdout + revived.stderr)
        self.assertIn("fresh source", (self.canonical / "retired-snapshot" / "SKILL.md").read_text())
        self.assertIn("retired-snapshot", self.ledger.read_text(encoding="utf-8").split("retired:", 1)[0])
        self.assertNotIn("  - dir: retired-snapshot", self.ledger.read_text(encoding="utf-8").split("retired:", 1)[1])
        for root in self.roots:
            self.assertEqual((root / "retired-snapshot").resolve(),
                             (self.canonical / "retired-snapshot").resolve())


if __name__ == "__main__":
    unittest.main()
