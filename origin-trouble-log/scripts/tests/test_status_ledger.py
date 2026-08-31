#!/usr/bin/env python3

import importlib.util
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


SCRIPT = Path(__file__).resolve().parents[1] / "status_ledger.py"
SPEC = importlib.util.spec_from_file_location("status_ledger", SCRIPT)
assert SPEC and SPEC.loader
status_ledger = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = status_ledger
SPEC.loader.exec_module(status_ledger)


class StatusLedgerTests(unittest.TestCase):
    def setUp(self):
        self.tempdir = tempfile.TemporaryDirectory()
        self.root = Path(self.tempdir.name)
        (self.root / "entries/2026-08").mkdir(parents=True)
        (self.root / "triage").mkdir()
        self.entry_a = "2026-08-27-alpha.md"
        self.entry_b = "2026-08-28-beta.md"
        self.entry_c = "2026-08-28-gamma.md"
        for name in (self.entry_a, self.entry_b, self.entry_c):
            (self.root / "entries/2026-08" / name).write_text("---\n---\n", encoding="utf-8")

    def tearDown(self):
        self.tempdir.cleanup()

    def run_cli(self, *args):
        return subprocess.run(
            [sys.executable, "-B", str(SCRIPT), "--root", str(self.root), *args],
            check=False,
            capture_output=True,
            text=True,
        )

    def test_sync_uses_only_target_section_and_marks_pending(self):
        (self.root / "triage/2026-08-16.md").write_text(
            """# report

参照 2026-08-27-not-a-target.md

## 対象 entry 一覧（1件）
- 2026-08-27-alpha.md
```
- 2026-08-28-not-a-target.md
```
## 対策
- 2026-08-28-not-a-target.md
""",
            encoding="utf-8",
        )
        (self.root / "triage/2026-08-26.md").write_text(
            """# report
## 対象 entry 一覧（1件）
- 2026-08-28-beta.md
## 対策
- 2026-08-27-not-a-target.md
""",
            encoding="utf-8",
        )
        result = self.run_cli("sync")
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        rows = status_ledger.read_ledger(self.root / "triage/status.tsv")
        self.assertEqual(rows[self.entry_a]["triage_status"], "triaged")
        self.assertEqual(rows[self.entry_a]["last_triaged"], "2026-08-16")
        self.assertEqual(rows[self.entry_b]["triage_status"], "triaged")
        self.assertEqual(rows[self.entry_b]["last_triaged"], "2026-08-26")
        self.assertEqual(rows[self.entry_c]["triage_status"], "untriaged")
        self.assertEqual(set(rows), {self.entry_a, self.entry_b, self.entry_c})

    def test_sync_preserves_manually_updated_status(self):
        self.assertEqual(self.run_cli("sync").returncode, 0)
        (self.root / "triage/responses.tsv").write_text(
            "response_id\towner\tstate\timplemented_at\tevidence\tnext_evaluation\n"
            "hook-h1-pipeline-evidence\thooks\timplemented_unverified\t2026-08-28\t13 tests\tnext triage\n",
            encoding="utf-8",
        )
        update = self.run_cli(
            "update",
            "--entry",
            self.entry_a,
            "--triage-status",
            "triaged",
            "--response-status",
            "effective_provisional",
            "--response-ids",
            "hook-h1-pipeline-evidence",
            "--last-triaged",
            "2026-08-28",
            "--status-basis",
            "回帰測定で再発なし",
            "--next-action",
            "次回も的中率を測る",
        )
        self.assertEqual(update.returncode, 0, update.stdout + update.stderr)
        self.assertEqual(self.run_cli("sync").returncode, 0)
        rows = status_ledger.read_ledger(self.root / "triage/status.tsv")
        self.assertEqual(rows[self.entry_a]["response_status"], "effective_provisional")
        self.assertEqual(rows[self.entry_a]["response_ids"], "hook-h1-pipeline-evidence")

    def test_validate_rejects_invalid_status_and_absolute_path(self):
        self.assertEqual(self.run_cli("sync").returncode, 0)
        ledger = self.root / "triage/status.tsv"
        rows = status_ledger.read_ledger(ledger)
        rows[self.entry_a]["response_status"] = "not-a-status"
        rows[self.entry_a]["status_basis"] = Path("/", "Users", "example", "private").as_posix()
        status_ledger.write_ledger(ledger, rows)
        result = self.run_cli("validate")
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("invalid response_status", result.stdout)
        self.assertIn("user-named absolute path", result.stdout)

    def test_validate_detects_missing_row(self):
        self.assertEqual(self.run_cli("sync").returncode, 0)
        ledger = self.root / "triage/status.tsv"
        rows = status_ledger.read_ledger(ledger)
        del rows[self.entry_b]
        status_ledger.write_ledger(ledger, rows)
        result = self.run_cli("validate")
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("missing ledger rows: 1", result.stdout)

    def test_sync_migrates_legacy_ledger_column(self):
        ledger = self.root / "triage/status.tsv"
        ledger.write_text(
            "entry\ttriage_status\tresponse_status\tresponse_ids\tlast_triaged\tstatus_basis\tnext_action\n"
            f"{self.entry_a}\ttriaged\tlegacy_unrecorded\t\t2026-08-16\tlegacy\tevaluate\n",
            encoding="utf-8",
        )
        result = self.run_cli("sync")
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("status_updated_at", ledger.read_text(encoding="utf-8").splitlines()[0])
        rows = status_ledger.read_ledger(ledger)
        self.assertEqual(rows[self.entry_a]["status_updated_at"], "2026-08-16")


if __name__ == "__main__":
    unittest.main()
