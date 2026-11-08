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


FULL_SHA = "0123456789abcdef0123456789abcdef01234567"


class MirrorSkillUpdateTest(unittest.TestCase):
    """--update-existing: plugin の版から正典の現役 mirror を置き換える同期。"""

    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.base = Path(self.tmp.name)
        self.canonical = self.base / "home" / ".agents" / "skills"
        self.canonical.mkdir(parents=True)
        _skill(self.canonical, "sample-skill", "old body")
        self.ledger = self.canonical / "mirrors.yaml"
        self.ledger_text = (
            "mirrors:\n"
            "  - dir: sample-skill\n"
            '    upstream: "example/skills"\n'
            '    upstream_path: "skills/sample-skill"\n'
            '    upstream_version: "1111111111111111111111111111111111111111"\n'
            "    fetched_at: 2026-01-01\n"
            '    license: "Apache-2.0"\n'
            '    reinstall: "example"\n'
            '    plugin: "sample@example-market"\n'
            '    local_copy: "~/plugin-cache/sample-skill"\n'
            '    note: "pinned note that must survive"\n'
            "\nretired:\n"
        )
        self.ledger.write_text(self.ledger_text, encoding="utf-8")
        self.root = self.base / "home" / ".codex" / "skills"
        self.root.mkdir(parents=True)
        (self.root / "sample-skill").symlink_to(self.canonical / "sample-skill")
        self.source = _skill(self.base / "plugin", "sample-skill", "new body")
        self.env = os.environ.copy()
        self.env.update({"AGENTS_SKILLS_ROOT": str(self.canonical)})

    def tearDown(self) -> None:
        self.tmp.cleanup()

    def _update(self, *extra: str, name: str = "sample-skill", source: Path | None = None,
                version: str | None = FULL_SHA):
        argv = [
            "bash", str(SCRIPT), "--upstream", "example/skills",
            "--path", f"skills/{name}", "--name", name, "--license", "Apache-2.0",
            "--source-dir", str(source or self.source), "--update-existing",
            "--alias-root", str(self.root),
        ]
        if version is not None:
            argv += ["--version", version]
        argv += list(extra)
        return subprocess.run(argv, env=self.env, text=True, capture_output=True, check=False)

    def _body(self) -> str:
        return (self.canonical / "sample-skill" / "SKILL.md").read_text(encoding="utf-8")

    def test_without_apply_shows_diff_and_changes_nothing(self):
        result = self._update()
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("CHANGED SKILL.md", result.stdout)
        self.assertIn("--apply", result.stdout)
        self.assertIn("old body", self._body())
        self.assertEqual(self.ledger.read_text(encoding="utf-8"), self.ledger_text)

    def test_apply_replaces_contents_and_keeps_unspecified_ledger_fields(self):
        result = self._update("--apply")
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("new body", self._body())
        ledger = self.ledger.read_text(encoding="utf-8")
        self.assertEqual(ledger.count("  - dir: sample-skill"), 1)
        self.assertIn(f'upstream_version: "{FULL_SHA}"', ledger)
        self.assertIn('note: "pinned note that must survive"', ledger)
        self.assertIn('local_copy: "~/plugin-cache/sample-skill"', ledger)
        self.assertIn('plugin: "sample@example-market"', ledger)
        self.assertNotIn("fetched_at: 2026-01-01", ledger)
        self.assertEqual((self.root / "sample-skill").resolve(),
                         (self.canonical / "sample-skill").resolve())

    def test_apply_can_record_plugin_link_for_the_drift_check(self):
        result = self._update("--apply", "--plugin", "sample@other-market",
                              "--plugin-path", "skills/sample-skill")
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        ledger = self.ledger.read_text(encoding="utf-8")
        self.assertIn('plugin: "sample@other-market"', ledger)
        self.assertIn('plugin_path: "skills/sample-skill"', ledger)
        self.assertEqual(ledger.count("plugin:"), 1)

    def test_update_requires_a_full_pin(self):
        for version in (None, "unknown", "2b1ffd2"):
            with self.subTest(version=version):
                result = self._update("--apply", version=version)
                self.assertNotEqual(result.returncode, 0, result.stdout)
                self.assertIn("old body", self._body())
                self.assertEqual(self.ledger.read_text(encoding="utf-8"), self.ledger_text)

    def test_name_mismatch_changes_nothing(self):
        wrong = _skill(self.base / "wrong", "other-skill", "wrong body")
        result = self._update("--apply", source=wrong)
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("old body", self._body())
        self.assertEqual(self.ledger.read_text(encoding="utf-8"), self.ledger_text)

    def test_update_of_inactive_name_is_refused(self):
        _skill(self.base / "plugin", "not-mirrored")
        result = self._update("--apply", name="not-mirrored",
                              source=self.base / "plugin" / "not-mirrored")
        self.assertNotEqual(result.returncode, 0)
        self.assertFalse((self.canonical / "not-mirrored").exists())

    def test_update_and_reactivate_cannot_combine(self):
        result = self._update("--apply", "--reactivate-retired")
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("old body", self._body())


if __name__ == "__main__":
    unittest.main()
