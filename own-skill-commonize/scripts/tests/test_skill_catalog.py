from __future__ import annotations

import sys
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory


SCRIPTS = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SCRIPTS))
from skill_catalog import load_skill_catalog  # noqa: E402


def _skill(root: Path, name: str) -> None:
    path = root / name
    path.mkdir(parents=True, exist_ok=True)
    (path / "SKILL.md").write_text(
        f"---\nname: {name}\ndescription: fixture\n---\nbody\n",
        encoding="utf-8",
    )


class SkillCatalogTest(unittest.TestCase):
    def test_retired_physical_skill_is_excluded_from_active_set(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            _skill(root, "own-example-run")
            _skill(root, "active-third-party")
            _skill(root, "retired-third-party")
            (root / "mirrors.yaml").write_text(
                "mirrors:\n  - dir: active-third-party\n"
                "retired:\n  - dir: retired-third-party\n",
                encoding="utf-8",
            )

            catalog = load_skill_catalog(root)

            self.assertEqual(catalog.active, {"own-example-run", "active-third-party"})
            self.assertEqual(catalog.retired_physical, {"retired-third-party"})
            self.assertEqual(catalog.errors, [])

    def test_unregistered_physical_skill_is_an_error(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            _skill(root, "own-example-run")
            _skill(root, "unregistered-third-party")
            (root / "mirrors.yaml").write_text("mirrors: []\nretired: []\n")

            catalog = load_skill_catalog(root)

            self.assertIn("unregistered physical skill: unregistered-third-party", catalog.errors)

    def test_active_ledger_entry_without_directory_is_an_error(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            _skill(root, "own-example-run")
            (root / "mirrors.yaml").write_text(
                "mirrors:\n  - dir: missing-third-party\nretired: []\n"
            )

            catalog = load_skill_catalog(root)

            self.assertIn("active mirror directory missing: missing-third-party", catalog.errors)

    def test_legacy_repository_without_manifest_uses_skill_directories(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            _skill(root, "own-example-run")
            _skill(root, "repo-specific-skill")

            catalog = load_skill_catalog(root)

            self.assertEqual(catalog.active, {"own-example-run", "repo-specific-skill"})
            self.assertTrue(catalog.legacy)
            self.assertEqual(catalog.errors, [])


if __name__ == "__main__":
    unittest.main()
