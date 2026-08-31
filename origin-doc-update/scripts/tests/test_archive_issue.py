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
