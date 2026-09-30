import importlib.util
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch


SCRIPTS = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SCRIPTS))
SPEC = importlib.util.spec_from_file_location("archive_issue", SCRIPTS / "archive_issue.py")
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
SPEC.loader.exec_module(MODULE)
import archive_transaction as TRANSACTION


class ArchiveIssueInputTest(unittest.TestCase):
    def test_issue_id_filename_and_path_are_normalized_to_the_same_id(self):
        self.assertEqual(MODULE.normalize_entry_id("ISSUE-20260828-example"), "ISSUE-20260828-example")
        self.assertEqual(MODULE.normalize_entry_id("ISSUE-20260828-example.md"), "ISSUE-20260828-example")
        self.assertEqual(
            MODULE.normalize_entry_id("docs/issues/ISSUE-20260828-example.md"),
            "ISSUE-20260828-example",
        )

    def test_index_removal_returns_count_and_exact_target_lines(self):
        content = (
            "## Current Focus\n\n"
            "- [ISSUE-a](issues/ISSUE-a.md) — focus\n\n"
            "## Active Issues\n\n"
            "- [ISSUE-a](issues/ISSUE-a.md) — active\n"
        )
        with tempfile.TemporaryDirectory() as directory:
            index = Path(directory) / "00_index.md"
            index.write_text(content)

            changed, target_lines = MODULE.remove_issue_from_index(str(index), "ISSUE-a")

            self.assertTrue(changed)
            self.assertEqual(
                target_lines,
                [
                    (3, "- [ISSUE-a](issues/ISSUE-a.md) — focus"),
                    (7, "- [ISSUE-a](issues/ISSUE-a.md) — active"),
                ],
            )
            self.assertNotIn("ISSUE-a", index.read_text())


class ArchiveTransactionTest(unittest.TestCase):
    def test_late_index_failure_restores_document_and_index(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            active = root / "docs/issues"
            archive = active / "archive"
            active.mkdir(parents=True)
            archive.mkdir()
            source = active / "ISSUE-a.md"
            destination = archive / source.name
            index = root / "docs/00_index.md"
            original_document = "status: active\n"
            original_index = "- [ISSUE-a](issues/ISSUE-a.md)\n"
            source.write_text(original_document)
            index.write_text(original_index)

            staged_destination = TRANSACTION.stage_text(
                destination, "status: archived\n", source
            )
            staged_source_restore = TRANSACTION.stage_text(
                source, original_document, source
            )
            staged_index = TRANSACTION.stage_text(
                index, "", index
            )
            staged_index_restore = TRANSACTION.stage_text(
                index, original_index, index
            )
            staged = [
                staged_destination,
                staged_source_restore,
                staged_index,
                staged_index_restore,
            ]
            real_replace = TRANSACTION.os.replace
            failed = False

            def fail_once(source_path, destination_path):
                nonlocal failed
                if Path(destination_path) == index and not failed:
                    failed = True
                    raise OSError("simulated index replacement failure")
                return real_replace(source_path, destination_path)

            try:
                with patch.object(TRANSACTION.os, "replace", side_effect=fail_once):
                    with self.assertRaises(RuntimeError) as raised:
                        TRANSACTION.apply_archive(
                            source,
                            destination,
                            staged_destination,
                            staged_source_restore,
                            index,
                            staged_index,
                            staged_index_restore,
                        )
                self.assertIn("archive transaction failed", str(raised.exception))
                self.assertEqual(source.read_text(), original_document)
                self.assertFalse(destination.exists())
                self.assertEqual(index.read_text(), original_index)
            finally:
                TRANSACTION.cleanup_staged(staged)


if __name__ == "__main__":
    unittest.main()


ISSUE_DOC = """---
schema_version: 2
id: {id}
status: complete
created_at: 2026-07-01
updated_at: 2026-09-01
branch: ""
pr: ""
related_specs: []
related_guides: []
guide_impact: none
guide_impact_reason: docs only
---

# {id}

## Acceptance

- verify: machine — true

## Completion

- [x] Implementation completed or intentionally not needed
- [x] Specs updated if direction or requirements changed
- [x] Guide impact classified before implementation
- [x] Guides updated in the same slice if implemented behavior changed
- [ ] Moved to docs/issues/archive/
"""


class IndexLinkSurvivesArchivingTest(unittest.TestCase):
    """The move must never leave the index pointing at the old path.

    Measured 2026-09-19 on a real repository: the issue was listed as a prose
    paragraph that linked it (not as a `- [id](...)` bullet), so the row matcher
    found nothing, the index was excluded from link rewriting as usual, and the
    script printed "not found in Active Issues (check manually if needed)" while
    its own move turned that link into a broken one. Nothing downstream failed;
    the validator's link check caught it three commits later.
    """

    SCRIPT = SCRIPTS / "archive_issue.py"
    ISSUE_ID = "ISSUE-20260901-example"
    # Where this script's documents live. The subclass below re-runs every test
    # in this class against `archive_workstream.py`, so anything that touches a
    # document path has to ask rather than hardcode `docs/issues`.
    DOC_DIR = "docs/issues"

    def make_repo(self, index_body: str) -> Path:
        root = Path(tempfile.mkdtemp())
        for path in ("docs/adrs", "docs/specs", "docs/issues/archive",
                     "docs/workstreams/archive", "docs/guides"):
            (root / path).mkdir(parents=True, exist_ok=True)
        (root / "docs/00_index.md").write_text(
            "---\nupdated_at: 2026-09-01\ncurrent_focus: x\n---\n\n# 00 Index\n\n"
            "## Active Issues\n\n" + index_body
        )
        (root / "docs/issues" / f"{self.ISSUE_ID}.md").write_text(
            ISSUE_DOC.format(id=self.ISSUE_ID)
        )
        return root

    def run_archive(self, root: Path, *args: str):
        import subprocess
        return subprocess.run(
            [sys.executable, str(self.SCRIPT), self.ISSUE_ID, "--repo", str(root), *args],
            capture_output=True, text=True,
        )

    def broken_links(self, root: Path) -> list[str]:
        import validate_repo_docs
        errors, _warnings = validate_repo_docs.validate_repo(root)
        return [e for e in errors if "broken relative link" in e]

    def test_an_issue_without_a_completion_section_is_refused(self):
        """The checklist gate must not be skipped by the absence of the checklist.

        Measured 2026-09-26: guarding only the "section present" branch meant an
        issue with NO `## Completion` section skipped the gate entirely and
        archived in silence. In one repository 8 of 11 open issues, and 38 of 74
        already archived, had no section -- so for all of them the gate had never
        run. `archive_workstream.py` already refused this case; only the issue
        script did not. Turns red if that branch goes back to `if len(...) == 2`.
        """
        root = self.make_repo(f"- [{self.ISSUE_ID}]({self.ISSUE_ID}.md) -- x\n")
        doc = root / self.DOC_DIR / f"{self.ISSUE_ID}.md"
        doc.write_text(doc.read_text().split("## Completion", 1)[0])

        result = self.run_archive(root)

        self.assertEqual(result.returncode, 1, result.stdout)
        self.assertIn("no '## Completion' section", result.stderr)
        self.assertTrue(doc.is_file(), "a refused archive must not move the file")
        self.assertFalse(
            (root / self.DOC_DIR / "archive" / f"{self.ISSUE_ID}.md").exists(),
            "a refused archive must not leave a copy in archive/",
        )

    def test_a_prose_row_is_repointed_instead_of_left_broken(self):
        prose = (
            "consumer feedback (2026-08-09) — "
            f"[docs/issues/{self.ISSUE_ID}](issues/{self.ISSUE_ID}.md): what it was about\n"
        )
        root = self.make_repo(prose)

        result = self.run_archive(root)

        self.assertEqual(result.returncode, 0, result.stderr)
        index = (root / "docs/00_index.md").read_text()
        self.assertIn(f"](issues/archive/{self.ISSUE_ID}.md)", index)
        self.assertNotIn(f"](issues/{self.ISSUE_ID}.md)", index)
        self.assertEqual(self.broken_links(root), [])
        self.assertIn("docs/00_index.md", result.stdout)

    def test_a_bullet_row_is_still_removed_by_default(self):
        bullet = f"- [{self.ISSUE_ID}](issues/{self.ISSUE_ID}.md) — one line\n"
        root = self.make_repo(bullet)

        result = self.run_archive(root)

        self.assertEqual(result.returncode, 0, result.stderr)
        index = (root / "docs/00_index.md").read_text()
        self.assertNotIn(self.ISSUE_ID, index)
        self.assertEqual(self.broken_links(root), [])

    def test_keep_row_repoints_the_bullet_rather_than_deleting_it(self):
        # An index whose own Read Policy says completed rows stay (repointed at
        # the archive) had its row deleted and restored by hand.
        bullet = f"- [{self.ISSUE_ID}](issues/{self.ISSUE_ID}.md) — one line\n"
        root = self.make_repo(bullet)

        result = self.run_archive(root, "--keep-row")

        self.assertEqual(result.returncode, 0, result.stderr)
        index = (root / "docs/00_index.md").read_text()
        self.assertIn(f"- [{self.ISSUE_ID}](issues/archive/{self.ISSUE_ID}.md) — one line", index)
        self.assertEqual(self.broken_links(root), [])

    def test_completed_policy_moves_row_out_of_generated_active_section(self):
        rel_dir = "workstreams" if self.ISSUE_ID.startswith("WS-") else "issues"
        bullet = f"- [{self.ISSUE_ID}]({rel_dir}/{self.ISSUE_ID}.md) — one line\n"
        root = self.make_repo(
            "<!-- own-doc-update:generated active-issues begin -->\n"
            + bullet
            + "<!-- own-doc-update:generated active-issues end -->\n\n"
            + "## Completed (archive)\n\n"
        )
        index_path = root / "docs/00_index.md"
        index_path.write_text(index_path.read_text().replace(
            "current_focus: x\n", "current_focus: x\narchive_index_rows: completed\n"
        ))

        result = self.run_archive(root)

        self.assertEqual(result.returncode, 0, result.stderr)
        index = index_path.read_text()
        active, completed = index.split("## Completed (archive)")
        self.assertNotIn(self.ISSUE_ID, active)
        self.assertIn(
            f"- [{self.ISSUE_ID}]({rel_dir}/archive/{self.ISSUE_ID}.md) — one line",
            completed,
        )
        self.assertEqual(self.broken_links(root), [])

    def test_completed_policy_takes_precedence_over_keep_row_flag(self):
        rel_dir = "workstreams" if self.ISSUE_ID.startswith("WS-") else "issues"
        root = self.make_repo(
            f"- [{self.ISSUE_ID}]({rel_dir}/{self.ISSUE_ID}.md) — one line\n"
            "\n## Completed (archive)\n\n"
        )
        index_path = root / "docs/00_index.md"
        index_path.write_text(index_path.read_text().replace(
            "current_focus: x\n", "current_focus: x\narchive_index_rows: completed\n"
        ))

        result = self.run_archive(root, "--keep-row")

        self.assertEqual(result.returncode, 0, result.stderr)
        active, completed = index_path.read_text().split("## Completed (archive)")
        self.assertNotIn(self.ISSUE_ID, active)
        self.assertIn(f"{rel_dir}/archive/{self.ISSUE_ID}.md", completed)

    def test_invalid_policy_stops_before_moving_issue(self):
        rel_dir = "workstreams" if self.ISSUE_ID.startswith("WS-") else "issues"
        root = self.make_repo(f"- [{self.ISSUE_ID}]({rel_dir}/{self.ISSUE_ID}.md)\n")
        index_path = root / "docs/00_index.md"
        index_path.write_text(index_path.read_text().replace(
            "current_focus: x\n", "current_focus: x\narchive_index_rows: unexpected\n"
        ))

        result = self.run_archive(root)

        self.assertNotEqual(result.returncode, 0)
        self.assertTrue((root / "docs" / rel_dir / f"{self.ISSUE_ID}.md").exists())
        self.assertFalse((root / "docs" / rel_dir / "archive" / f"{self.ISSUE_ID}.md").exists())

    def test_completed_policy_requires_destination_heading_before_move(self):
        rel_dir = "workstreams" if self.ISSUE_ID.startswith("WS-") else "issues"
        root = self.make_repo(f"- [{self.ISSUE_ID}]({rel_dir}/{self.ISSUE_ID}.md)\n")
        index_path = root / "docs/00_index.md"
        index_path.write_text(index_path.read_text().replace(
            "current_focus: x\n", "current_focus: x\narchive_index_rows: completed\n"
        ))

        result = self.run_archive(root)

        self.assertNotEqual(result.returncode, 0)
        self.assertIn("Completed (archive)", result.stderr)
        self.assertTrue((root / "docs" / rel_dir / f"{self.ISSUE_ID}.md").exists())

    def test_completed_policy_moves_plain_path_row(self):
        rel_dir = "workstreams" if self.ISSUE_ID.startswith("WS-") else "issues"
        root = self.make_repo(
            f"- docs/{rel_dir}/{self.ISSUE_ID}.md — one line\n"
            "\n## Completed (archive)\n\n"
        )
        index_path = root / "docs/00_index.md"
        index_path.write_text(index_path.read_text().replace(
            "current_focus: x\n", "current_focus: x\narchive_index_rows: completed\n"
        ))

        result = self.run_archive(root)

        self.assertEqual(result.returncode, 0, result.stderr)
        active, completed = index_path.read_text().split("## Completed (archive)")
        self.assertNotIn(self.ISSUE_ID, active)
        self.assertIn(f"docs/{rel_dir}/archive/{self.ISSUE_ID}.md", completed)


class WorkstreamIndexLinkSurvivesArchivingTest(IndexLinkSurvivesArchivingTest):
    """The twin script carried the identical defect; fixing only one leaves it live."""

    SCRIPT = SCRIPTS / "archive_workstream.py"
    ISSUE_ID = "WS-20260901-example"
    DOC_DIR = "docs/workstreams"

    WS_DOC = """---
schema_version: 2
id: {id}
status: complete
created_at: 2026-07-01
updated_at: 2026-09-01
branch: ""
pr: ""
human_boundary_confirmed_at: 2026-07-01
next_human_gate: none
related_specs: []
related_guides: []
---

# {id}

## Authorization Envelope

- Autonomous actions allowed: edit docs
- Confirm first: external sends
- Merge policy: CD on green

## Human Gates

none

## Issue Queue

| Issue | Status |
| --- | --- |
| ISSUE-01-a | complete |

### ISSUE-01-a

- status: complete
- depends_on: []
- runnability: ready
- guide_impact: none
- related_guides: []
- guide_impact_reason: docs only

#### Acceptance

- verify: machine — true

## Completion

- [x] All issues complete or explicitly dropped
- [x] Specs updated if direction changed
- [x] Guides updated for implemented behavior
- [x] Human gate reached or stop reason recorded
- [x] Moved to docs/workstreams/archive/
"""

    def make_repo(self, index_body: str) -> Path:
        root = Path(tempfile.mkdtemp())
        for path in ("docs/adrs", "docs/specs", "docs/issues/archive",
                     "docs/workstreams/archive", "docs/guides"):
            (root / path).mkdir(parents=True, exist_ok=True)
        (root / "docs/00_index.md").write_text(
            "---\nupdated_at: 2026-09-01\ncurrent_focus: x\n---\n\n# 00 Index\n\n"
            "## Active Workstreams\n\n"
            + index_body.replace("issues/", "workstreams/")
        )
        (root / "docs/workstreams" / f"{self.ISSUE_ID}.md").write_text(
            self.WS_DOC.format(id=self.ISSUE_ID)
        )
        return root

    def test_a_prose_row_is_repointed_instead_of_left_broken(self):
        prose = (
            "the delivery round — "
            f"[docs/workstreams/{self.ISSUE_ID}](workstreams/{self.ISSUE_ID}.md): what it did\n"
        )
        root = self.make_repo(prose)

        result = self.run_archive(root)

        self.assertEqual(result.returncode, 0, result.stderr)
        index = (root / "docs/00_index.md").read_text()
        self.assertIn(f"](workstreams/archive/{self.ISSUE_ID}.md)", index)
        self.assertNotIn(f"](workstreams/{self.ISSUE_ID}.md)", index)
        self.assertEqual(self.broken_links(root), [])

    def test_a_bullet_row_is_still_removed_by_default(self):
        bullet = f"- [{self.ISSUE_ID}](workstreams/{self.ISSUE_ID}.md) — one line\n"
        root = self.make_repo(bullet)

        result = self.run_archive(root)

        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertNotIn(self.ISSUE_ID, (root / "docs/00_index.md").read_text())
        self.assertEqual(self.broken_links(root), [])

    def test_keep_row_repoints_the_bullet_rather_than_deleting_it(self):
        bullet = f"- [{self.ISSUE_ID}](workstreams/{self.ISSUE_ID}.md) — one line\n"
        root = self.make_repo(bullet)

        result = self.run_archive(root, "--keep-row")

        self.assertEqual(result.returncode, 0, result.stderr)
        index = (root / "docs/00_index.md").read_text()
        self.assertIn(
            f"- [{self.ISSUE_ID}](workstreams/archive/{self.ISSUE_ID}.md) — one line", index
        )

    def test_completed_policy_moves_workstream_row(self):
        root = self.make_repo(
            f"- [{self.ISSUE_ID}](workstreams/{self.ISSUE_ID}.md) — one line\n"
            "\n## Completed (archive)\n\n"
        )
        index_path = root / "docs/00_index.md"
        index_path.write_text(index_path.read_text().replace(
            "current_focus: x\n", "current_focus: x\narchive_index_rows: completed\n"
        ))

        result = self.run_archive(root)

        self.assertEqual(result.returncode, 0, result.stderr)
        active, completed = index_path.read_text().split("## Completed (archive)")
        self.assertNotIn(self.ISSUE_ID, active)
        self.assertIn(
            f"- [{self.ISSUE_ID}](workstreams/archive/{self.ISSUE_ID}.md) — one line",
            completed,
        )
        self.assertEqual(self.broken_links(root), [])
