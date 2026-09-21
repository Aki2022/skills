#!/usr/bin/env python3

import csv
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


SCRIPT = Path(__file__).resolve().parents[1] / "historical_index.py"
STATUS_SCRIPT = Path(__file__).resolve().parents[1] / "status_ledger.py"


class HistoricalIndexTests(unittest.TestCase):
    def setUp(self):
        self.tempdir = tempfile.TemporaryDirectory()
        self.root = Path(self.tempdir.name)
        (self.root / "entries/2026-08").mkdir(parents=True)
        (self.root / "triage").mkdir()
        self.entry_a = "2026-08-16-alpha.md"
        self.entry_b = "2026-08-16-beta.md"
        (self.root / "entries/2026-08" / self.entry_a).write_text(
            """---
date: 2026-08-16
summary: 課金: なしと誤って表示した
skills: [own-trouble-log]
repo: sample
canon: deadbeef
paths: []
---

## 意図 — 何をしようとしていたか
集計しようとした。

## 実際にやったこと — 実行したコマンド・書いたコードを実文で
```
echo ok
```
""",
            encoding="utf-8",
        )
        (self.root / "entries/2026-08" / self.entry_b).write_text(
            "本文だけの古い記録\n", encoding="utf-8"
        )
        (self.root / "triage/2026-08-16.md").write_text(
            """# report
## 対象 entry 一覧
- 2026-08-16-alpha.md
- 2026-08-16-beta.md
## clusters
""",
            encoding="utf-8",
        )
        (self.root / "triage/responses.tsv").write_text(
            "response_id\towner\tstate\timplemented_at\tevidence\tnext_evaluation\n"
            "known-response\towner\tproposed\t\treport\tnext\n",
            encoding="utf-8",
        )
        sync = subprocess.run(
            [sys.executable, str(STATUS_SCRIPT), "--root", str(self.root), "sync"],
            check=False,
            capture_output=True,
            text=True,
        )
        self.assertEqual(sync.returncode, 0, sync.stdout + sync.stderr)

    def tearDown(self):
        self.tempdir.cleanup()

    def run_cli(self, *args):
        return subprocess.run(
            [sys.executable, "-B", str(SCRIPT), "--root", str(self.root), *args],
            check=False,
            capture_output=True,
            text=True,
        )

    def valid_payload(self):
        return {
            "snapshot_entries": [self.entry_a, self.entry_b],
            "patterns": [
                {
                    "pattern_id": "report-value-loss",
                    "title": "表示用派生値の消失",
                    "failure_shape": "集計値は残るが表示値だけ失われる",
                    "detection_mode": "runtime_state",
                    "owner": "own-trouble-log",
                    "formalization_state": "candidate",
                    "basis": "出力と元データの差を比較できる",
                },
                {
                    "pattern_id": "unclassified",
                    "title": "材料不足",
                    "failure_shape": "本文だけでは観測可能な失敗形を決められない",
                    "detection_mode": "manual_review",
                    "owner": "own-trouble-log",
                    "formalization_state": "manual_only",
                    "basis": "不足理由をentryごとに残す",
                },
            ],
            "links": [
                {
                    "entry": self.entry_a,
                    "pattern_id": "report-value-loss",
                    "response_id": "known-response",
                    "confidence": "mechanism",
                    "basis": "表示値の消失という観測可能な機構が一致",
                    "analyzed_at": "2026-09-21",
                },
                {
                    "entry": self.entry_b,
                    "pattern_id": "unclassified",
                    "response_id": "",
                    "confidence": "unclassified",
                    "basis": "固定7節と実文が無く機構を判定できない",
                    "analyzed_at": "2026-09-21",
                },
            ],
            "report_name": "2026-09-21-legacy-mining.md",
            "report_content": "# Historical mining\n\n## Snapshot\n2 entries\n",
        }

    def write_payload(self, payload):
        path = self.root / "payload.json"
        path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
        return path

    def test_snapshot_tolerates_invalid_yaml_and_missing_frontmatter(self):
        result = self.run_cli("snapshot")
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        rows = [json.loads(line) for line in result.stdout.splitlines()]
        self.assertEqual([row["entry"] for row in rows], [self.entry_a, self.entry_b])
        self.assertEqual(rows[0]["summary"], "課金: なしと誤って表示した")
        self.assertTrue(rows[0]["frontmatter_present"])
        self.assertFalse(rows[1]["frontmatter_present"])
        self.assertIn("本文だけの古い記録", rows[1]["text"])

    def test_snapshot_can_be_written_atomically_without_editing_entries(self):
        entry_before = {
            path: path.read_bytes() for path in (self.root / "entries").glob("*/*.md")
        }
        output = self.root / "snapshot.jsonl"
        result = self.run_cli("snapshot", "--output", str(output))
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertEqual(len(output.read_text(encoding="utf-8").splitlines()), 2)
        self.assertEqual(
            entry_before,
            {path: path.read_bytes() for path in (self.root / "entries").glob("*/*.md")},
        )

    def test_apply_validate_and_summary_cover_snapshot(self):
        payload = self.write_payload(self.valid_payload())
        applied = self.run_cli("apply", "--input", str(payload))
        self.assertEqual(applied.returncode, 0, applied.stdout + applied.stderr)
        validated = self.run_cli("validate")
        self.assertEqual(validated.returncode, 0, validated.stdout + validated.stderr)
        summary = self.run_cli("summary")
        self.assertEqual(summary.returncode, 0, summary.stdout + summary.stderr)
        self.assertIn("indexed_entries=2", summary.stdout)
        self.assertIn("missing_legacy_entries=0", summary.stdout)
        self.assertTrue((self.root / "triage/history/2026-09-21-legacy-mining.md").is_file())

    def test_check_validates_a_batch_without_writing(self):
        payload = self.valid_payload()
        payload.pop("report_name")
        payload.pop("report_content")
        checked = self.run_cli("check", "--input", str(self.write_payload(payload)))
        self.assertEqual(checked.returncode, 0, checked.stdout + checked.stderr)
        self.assertIn("checked entries=2", checked.stdout)
        self.assertFalse((self.root / "triage/patterns.tsv").exists())

    def test_apply_rejects_invalid_rows_without_writing(self):
        cases = {
            "unknown entry": lambda p: p["links"][0].update(entry="2026-08-16-missing.md"),
            "unknown pattern": lambda p: p["links"][0].update(pattern_id="missing-pattern"),
            "unknown response": lambda p: p["links"][0].update(response_id="missing-response"),
            "invalid confidence": lambda p: p["links"][0].update(confidence="certain"),
            "empty basis": lambda p: p["links"][0].update(basis=""),
            "absolute path": lambda p: p["links"][0].update(
                basis="path=" + "/" + "Users/example/private"
            ),
        }
        for label, mutate in cases.items():
            with self.subTest(label=label):
                payload = self.valid_payload()
                mutate(payload)
                result = self.run_cli("apply", "--input", str(self.write_payload(payload)))
                self.assertNotEqual(result.returncode, 0)
                self.assertFalse((self.root / "triage/patterns.tsv").exists())

    def test_apply_rejects_duplicate_and_missing_snapshot_coverage(self):
        duplicate = self.valid_payload()
        duplicate["links"].append(dict(duplicate["links"][0]))
        result = self.run_cli("apply", "--input", str(self.write_payload(duplicate)))
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("duplicate link", result.stdout)

        missing = self.valid_payload()
        missing["links"] = missing["links"][:1]
        result = self.run_cli("apply", "--input", str(self.write_payload(missing)))
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("snapshot entries without links: 1", result.stdout)

    def test_candidate_links_do_not_change_status_ledger(self):
        before = (self.root / "triage/status.tsv").read_text(encoding="utf-8")
        result = self.run_cli("apply", "--input", str(self.write_payload(self.valid_payload())))
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        after = (self.root / "triage/status.tsv").read_text(encoding="utf-8")
        self.assertEqual(after, before)
        with (self.root / "triage/entry-pattern-links.tsv").open(encoding="utf-8", newline="") as f:
            rows = list(csv.DictReader(f, delimiter="\t"))
        self.assertEqual(len(rows), 2)

    def test_status_update_requires_a_matching_direct_link(self):
        payload = self.valid_payload()
        payload["status_updates"] = [
            {
                "entry": self.entry_a,
                "response_id": "known-response",
                "response_status": "recurred",
                "status_basis": "同じ失敗形の再発を本文で確認",
                "next_action": "coverage gapを調べる",
            }
        ]
        rejected = self.run_cli("apply", "--input", str(self.write_payload(payload)))
        self.assertNotEqual(rejected.returncode, 0)
        self.assertIn("lacks matching direct link", rejected.stdout)

        payload["links"][0]["confidence"] = "direct"
        accepted = self.run_cli("apply", "--input", str(self.write_payload(payload)))
        self.assertEqual(accepted.returncode, 0, accepted.stdout + accepted.stderr)
        with (self.root / "triage/status.tsv").open(encoding="utf-8", newline="") as handle:
            rows = {row["entry"]: row for row in csv.DictReader(handle, delimiter="\t")}
        self.assertEqual(rows[self.entry_a]["response_status"], "recurred")
        self.assertEqual(rows[self.entry_a]["response_ids"], "known-response")

    def test_apply_can_replace_links_and_retire_a_redundant_pattern(self):
        first = self.run_cli("apply", "--input", str(self.write_payload(self.valid_payload())))
        self.assertEqual(first.returncode, 0, first.stdout + first.stderr)
        payload = self.valid_payload()
        payload["report_name"] = "2026-09-22-legacy-mining.md"
        payload["report_content"] = "# Pattern merge\n\nMerged duplicate runtime context pattern.\n"
        payload["replace_entries"] = [self.entry_a, self.entry_b]
        payload["replace_patterns"] = ["unclassified"]
        payload["retire_patterns"] = ["report-value-loss"]
        payload["patterns"] = [payload["patterns"][1]]
        payload["patterns"][0]["title"] = "統合後の材料不足"
        payload["patterns"][0]["basis"] = "統合後の根拠を更新する"
        payload["links"][0]["pattern_id"] = "unclassified"
        payload["links"][0]["response_id"] = ""
        second = self.run_cli("apply", "--input", str(self.write_payload(payload)))
        self.assertEqual(second.returncode, 0, second.stdout + second.stderr)
        with (self.root / "triage/entry-pattern-links.tsv").open(encoding="utf-8", newline="") as handle:
            rows = list(csv.DictReader(handle, delimiter="\t"))
        self.assertEqual(len(rows), 2)
        self.assertEqual({row["pattern_id"] for row in rows}, {"unclassified"})
        with (self.root / "triage/patterns.tsv").open(encoding="utf-8", newline="") as handle:
            patterns = list(csv.DictReader(handle, delimiter="\t"))
        self.assertEqual([row["pattern_id"] for row in patterns], ["unclassified"])

    def test_validate_rejects_all_guardrail_violations(self):
        result = self.run_cli("apply", "--input", str(self.write_payload(self.valid_payload())))
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        links_path = self.root / "triage/entry-pattern-links.tsv"
        original = links_path.read_text(encoding="utf-8")
        cases = {
            "unknown entry": lambda rows: rows[0].update(entry="2026-08-16-missing.md"),
            "unknown pattern": lambda rows: rows[0].update(pattern_id="missing-pattern"),
            "unknown response": lambda rows: rows[0].update(response_id="missing-response"),
            "invalid confidence": lambda rows: rows[0].update(confidence="certain"),
            "empty basis": lambda rows: rows[0].update(basis=""),
            "absolute path": lambda rows: rows[0].update(
                basis="path=" + "/" + "Users/example/private"
            ),
            "duplicate link": lambda rows: rows.append(dict(rows[0])),
        }
        for label, mutate in cases.items():
            with self.subTest(label=label):
                with links_path.open(encoding="utf-8", newline="") as handle:
                    rows = list(csv.DictReader(handle, delimiter="\t"))
                mutate(rows)
                with links_path.open("w", encoding="utf-8", newline="") as handle:
                    writer = csv.DictWriter(
                        handle,
                        fieldnames=("entry", "pattern_id", "response_id", "confidence", "basis", "analyzed_at"),
                        delimiter="\t",
                        lineterminator="\n",
                    )
                    writer.writeheader()
                    writer.writerows(rows)
                checked = self.run_cli("validate")
                self.assertNotEqual(checked.returncode, 0)
                links_path.write_text(original, encoding="utf-8")


if __name__ == "__main__":
    unittest.main()
