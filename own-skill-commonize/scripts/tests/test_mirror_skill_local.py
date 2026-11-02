from __future__ import annotations

import os
import shutil
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
            "mirrors:\n\nretired:\n  - dir: retired-snapshot\n"
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
        fake_bin = self.base / "fake-bin"
        fake_bin.mkdir()
        fake_npx = fake_bin / "npx"
        fake_npx.write_text(
            "#!/bin/sh\ncp -R \"$FAKE_SKILL_SOURCE\" \"$4\"\n",
            encoding="utf-8",
        )
        fake_npx.chmod(0o755)
        self.env["PATH"] = str(fake_bin) + os.pathsep + self.env.get("PATH", "")

    def tearDown(self) -> None:
        self.tmp.cleanup()

    def _run(
        self,
        name: str,
        source: Path | None = None,
        *,
        reactivate: bool = False,
        update_existing: bool = False,
        version: str = "abc",
    ):
        argv = [
            "bash", str(SCRIPT), "--upstream", "example/skills", "--path", name,
            "--name", name, "--license", "Apache-2.0", "--version", version,
        ]
        if source is not None:
            argv += ["--source-dir", str(source)]
        if reactivate:
            argv.append("--reactivate-retired")
        if update_existing:
            argv.append("--update-existing")
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

        shutil.rmtree(self.canonical / "retired-snapshot")
        reactivated_source = _skill(self.source_root, "retired-snapshot", "fresh source")
        self.env["FAKE_SKILL_SOURCE"] = str(reactivated_source)
        revived = self._run("retired-snapshot", reactivate=True)
        self.assertEqual(revived.returncode, 0, revived.stdout + revived.stderr)
        self.assertIn("fresh source", (self.canonical / "retired-snapshot" / "SKILL.md").read_text())
        self.assertIn("retired-snapshot", self.ledger.read_text(encoding="utf-8").split("retired:", 1)[0])
        self.assertNotIn("  - dir: retired-snapshot", self.ledger.read_text(encoding="utf-8").split("retired:", 1)[1])
        for root in self.roots:
            self.assertEqual((root / "retired-snapshot").resolve(),
                             (self.canonical / "retired-snapshot").resolve())

    def test_update_existing_mirror_replaces_contents_and_metadata(self):
        imported = self._run("sample-skill", self.source)
        self.assertEqual(imported.returncode, 0, imported.stdout + imported.stderr)

        old_content = (self.canonical / "sample-skill" / "SKILL.md").read_text()
        self.assertIn("fixture", old_content)
        rejected = self._run("sample-skill", self.source)
        self.assertNotEqual(rejected.returncode, 0)
        self.assertEqual((self.canonical / "sample-skill" / "SKILL.md").read_text(), old_content)

        bad = _skill(self.source_root, "different-name", "wrong name")
        bad_update = self._run("sample-skill", bad, update_existing=True)
        self.assertNotEqual(bad_update.returncode, 0)
        self.assertEqual((self.canonical / "sample-skill" / "SKILL.md").read_text(), old_content)

        (self.source / "SKILL.md").write_text(
            "---\nname: sample-skill\ndescription: fixture\n---\nupdated source\n",
            encoding="utf-8",
        )
        updated = self._run(
            "sample-skill", self.source, update_existing=True, version="def456"
        )
        self.assertEqual(updated.returncode, 0, updated.stdout + updated.stderr)
        self.assertIn("updated source", (self.canonical / "sample-skill" / "SKILL.md").read_text())
        ledger = self.ledger.read_text(encoding="utf-8")
        self.assertEqual(ledger.count("  - dir: sample-skill"), 1)
        self.assertIn('upstream_version: "def456"', ledger)
        for root in self.roots:
            self.assertEqual((root / "sample-skill").resolve(),
                             (self.canonical / "sample-skill").resolve())


if __name__ == "__main__":
    unittest.main()
