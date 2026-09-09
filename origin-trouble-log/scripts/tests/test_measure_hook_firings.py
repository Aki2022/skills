#!/usr/bin/env python3
"""Tests for the warn-guard firing measurement.

The measurement exists to fill three numbers the status ledger asks for on every
hook response — 是正率 / 同型再発数 / 誤警告数 — and the first two are derived from
what agent transcripts actually recorded. These tests pin the shapes that were
measured by hand against real transcripts on 2026-09-09:

- one firing writes three attachment records (hook_success / hook_additional_context /
  hook_system_message) carrying the same text, so only hook_success may be counted;
- the agent grepping the guard script also puts the literal string into user and
  assistant records, and those are not firings;
- `attachment.toolUseID` joins a firing to the tool_use that triggered it.
"""

import importlib.util
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


SCRIPT = Path(__file__).resolve().parents[1] / "measure_hook_firings.py"
SPEC = importlib.util.spec_from_file_location("measure_hook_firings", SCRIPT)
assert SPEC and SPEC.loader
measure = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = measure
SPEC.loader.exec_module(measure)


def hook_success(tool_use_id, message, *, session="s1", ts="2026-09-01T00:00:00.000Z",
                 tool="Bash", duration=96, sidechain=False):
    stdout = json.dumps(
        {
            "systemMessage": message,
            "hookSpecificOutput": {"hookEventName": "PreToolUse", "additionalContext": message},
        },
        ensure_ascii=False,
    )
    return {
        "type": "attachment",
        "timestamp": ts,
        "sessionId": session,
        "isSidechain": sidechain,
        "attachment": {
            "type": "hook_success",
            "hookName": f"PreToolUse:{tool}",
            "hookEvent": "PreToolUse",
            "toolUseID": tool_use_id,
            "durationMs": duration,
            "exitCode": 0,
            "stdout": stdout,
            "stderr": "",
        },
    }


def hook_echo(tool_use_id, message, kind, *, session="s1", ts="2026-09-01T00:00:00.000Z", tool="Bash"):
    """The other two attachment records the harness writes for the same firing."""
    return {
        "type": "attachment",
        "timestamp": ts,
        "sessionId": session,
        "attachment": {
            "type": kind,
            "hookName": f"PreToolUse:{tool}",
            "hookEvent": "PreToolUse",
            "toolUseID": tool_use_id,
            "content": message,
        },
    }


def tool_use(tool_use_id, command, *, session="s1", ts="2026-09-01T00:00:00.000Z",
             tool="Bash", sidechain=False):
    key = "file_path" if tool in ("Write", "Edit", "NotebookEdit") else "command"
    return {
        "type": "assistant",
        "timestamp": ts,
        "sessionId": session,
        "isSidechain": sidechain,
        "message": {"content": [{"type": "tool_use", "id": tool_use_id, "name": tool,
                                 "input": {key: command}}]},
    }


def tool_result(tool_use_id, *, is_error=False, session="s1", ts="2026-09-01T00:00:00.000Z"):
    return {
        "type": "user",
        "timestamp": ts,
        "sessionId": session,
        "message": {"content": [{"type": "tool_result", "tool_use_id": tool_use_id,
                                 "is_error": is_error}]},
    }


GUARD = "origin-warn-guards: "


class ParseTranscriptTests(unittest.TestCase):
    def setUp(self):
        self.tempdir = tempfile.TemporaryDirectory()
        self.root = Path(self.tempdir.name)

    def tearDown(self):
        self.tempdir.cleanup()

    def write(self, name, records):
        path = self.root / name
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("w", encoding="utf-8") as handle:
            for record in records:
                handle.write(json.dumps(record, ensure_ascii=False) + "\n")
        return path

    def test_one_firing_per_hook_success_not_three(self):
        message = GUARD + "[H2] パイプ後の $? は末尾コマンドの exit code。"
        path = self.write("a.jsonl", [
            tool_use("toolu_1", 'pytest | tail -1; echo "$?"'),
            hook_success("toolu_1", message),
            hook_echo("toolu_1", message, "hook_additional_context"),
            hook_echo("toolu_1", message, "hook_system_message"),
            tool_result("toolu_1"),
        ])
        scan = measure.parse_transcript(path)
        self.assertEqual(len(scan.firings), 1)
        self.assertEqual(scan.firings[0]["tags"], ["H2"])
        # Prove the fixture actually contains the two echo records the dedup must drop,
        # so this test cannot pass because they were never there.
        self.assertEqual(scan.echo_records, 2)

    def test_session_start_hook_output_is_not_a_tool_firing(self):
        """Measured 2026-09-09: 11 of 2,773 marker-carrying hook_success records are
        SessionStart, carry no toolUseID, and must not enter a per-tool rate."""
        record = hook_success("", GUARD + "[H1] x")
        record["attachment"]["hookEvent"] = "SessionStart"
        record["attachment"]["hookName"] = "SessionStart"
        del record["attachment"]["toolUseID"]
        scan = measure.parse_transcript(self.write("a.jsonl", [record]))
        self.assertEqual(scan.firings, [])

    def test_echo_records_are_counted_by_type_and_tool_use_id(self):
        """`hook_additional_context` does not keep the text in a top-level string field
        (measured: 2,762 such records, 0 with the marker in `content`), so the dedup
        invariant must be counted by record type plus toolUseID, not by text."""
        message = GUARD + "[H1] x"
        additional = hook_echo("toolu_1", message, "hook_additional_context")
        additional["attachment"].pop("content")
        additional["attachment"]["hookSpecificOutput"] = {"additionalContext": message}
        scan = measure.parse_transcript(self.write("a.jsonl", [
            tool_use("toolu_1", "cmd | tail"),
            hook_success("toolu_1", message),
            additional,
            hook_echo("toolu_1", message, "hook_system_message"),
        ]))
        self.assertEqual(len(scan.firings), 1)
        self.assertEqual(scan.echo_records, 2)

    def test_agent_grepping_the_guard_script_is_not_a_firing(self):
        path = self.write("a.jsonl", [
            tool_use("toolu_1", "grep -n origin-warn-guards ~/.agents/hooks/origin_warn_guards.sh"),
            tool_result("toolu_1"),
            {"type": "user", "timestamp": "2026-09-01T00:00:00.000Z", "sessionId": "s1",
             "message": {"content": [{"type": "tool_result", "tool_use_id": "toolu_1",
                                      "content": "76: add_warn \"origin-warn-guards: [H1] ...\""}]}},
        ])
        scan = measure.parse_transcript(path)
        self.assertEqual(scan.firings, [])

    def test_firing_joins_to_its_triggering_command_and_result(self):
        path = self.write("a.jsonl", [
            tool_use("toolu_1", "git add -A"),
            hook_success("toolu_1", GUARD + "[H5f] untracked 2 件がある状態で"),
            tool_result("toolu_1", is_error=True),
        ])
        firing = measure.parse_transcript(path).firings[0]
        self.assertEqual(firing["command"], "git add -A")
        self.assertTrue(firing["is_error"])
        self.assertEqual(firing["tool_name"], "Bash")

    def test_next_call_is_the_following_use_of_the_same_tool(self):
        path = self.write("a.jsonl", [
            tool_use("toolu_1", "cmd | tail -1; echo $?"),
            hook_success("toolu_1", GUARD + "[H2] x"),
            tool_result("toolu_1"),
            tool_use("toolu_w", "/tmp/x.md", tool="Write"),
            tool_use("toolu_2", "cmd > log 2>&1; rc=${PIPESTATUS[0]}"),
            tool_result("toolu_2"),
        ])
        firing = measure.parse_transcript(path).firings[0]
        self.assertEqual(firing["next_command"], "cmd > log 2>&1; rc=${PIPESTATUS[0]}")
        self.assertEqual(firing["next_tags"], [])

    def test_next_call_records_the_tags_that_fired_on_it(self):
        path = self.write("a.jsonl", [
            tool_use("toolu_1", "cmd | tail -1; echo $?"),
            hook_success("toolu_1", GUARD + "[H2] x"),
            tool_use("toolu_2", "other | tail -1; echo $?"),
            hook_success("toolu_2", GUARD + "[H2] x"),
        ])
        firings = measure.parse_transcript(path).firings
        self.assertEqual(firings[0]["next_tags"], ["H2"])
        self.assertIsNone(firings[1]["next_command"])

    def test_multi_tag_firing_reports_every_tag(self):
        message = GUARD + "[H5e] main 上で / [H5f] untracked がある状態で"
        path = self.write("a.jsonl", [
            tool_use("toolu_1", "git add -A && git commit -m x"),
            hook_success("toolu_1", message),
        ])
        self.assertEqual(measure.parse_transcript(path).firings[0]["tags"], ["H5e", "H5f"])

    def test_denominator_counts_tool_calls_per_tool_name(self):
        path = self.write("a.jsonl", [
            tool_use("toolu_1", "ls"),
            tool_use("toolu_2", "pwd"),
            tool_use("toolu_w", "/tmp/x.md", tool="Write"),
            tool_use("toolu_s", "ls", sidechain=True),
        ])
        scan = measure.parse_transcript(path)
        self.assertEqual(scan.tool_calls["Bash"], 2)
        self.assertEqual(scan.tool_calls["Write"], 1)
        self.assertEqual(scan.sidechain_tool_calls["Bash"], 1)


class AggregateTests(unittest.TestCase):
    def setUp(self):
        self.tempdir = tempfile.TemporaryDirectory()
        self.root = Path(self.tempdir.name)

    def tearDown(self):
        self.tempdir.cleanup()

    def write(self, name, records):
        path = self.root / name
        with path.open("w", encoding="utf-8") as handle:
            for record in records:
                handle.write(json.dumps(record, ensure_ascii=False) + "\n")
        return path

    def test_window_excludes_firings_outside_the_interval(self):
        path = self.write("a.jsonl", [
            tool_use("toolu_1", "a", ts="2026-09-05T00:00:00.000Z"),
            hook_success("toolu_1", GUARD + "[H1] x", ts="2026-09-05T00:00:00.000Z"),
            tool_use("toolu_2", "b", ts="2026-09-08T00:00:00.000Z"),
            hook_success("toolu_2", GUARD + "[H1] x", ts="2026-09-08T00:00:00.000Z"),
        ])
        scan = measure.parse_transcript(path)
        stats = measure.aggregate([scan], since="2026-09-07", until="2026-09-30")
        self.assertEqual(stats.tags["H1"]["firings"], 1)

    def test_session_is_uncorrected_when_the_tag_fires_again(self):
        path = self.write("a.jsonl", [
            tool_use("toolu_1", "a"),
            hook_success("toolu_1", GUARD + "[H1] x"),
            tool_use("toolu_2", "b"),
            hook_success("toolu_2", GUARD + "[H1] x"),
            tool_use("toolu_3", "c", session="s2"),
            hook_success("toolu_3", GUARD + "[H1] x", session="s2"),
            tool_use("toolu_4", "d", session="s2"),
        ])
        stats = measure.aggregate([measure.parse_transcript(path)])
        h1 = stats.tags["H1"]
        self.assertEqual(h1["sessions"], 2)
        self.assertEqual(h1["sessions_repeated"], 1)
        self.assertAlmostEqual(h1["correction_rate"], 0.5)

    def test_firing_rate_uses_the_denominator_of_the_firing_tool(self):
        records = [tool_use(f"toolu_b{i}", "ls") for i in range(9)]
        records += [
            tool_use("toolu_1", "git add -A"),
            hook_success("toolu_1", GUARD + "[H5f] x"),
            tool_use("toolu_w", "/tmp/x.md", tool="Write"),
        ]
        stats = measure.aggregate([measure.parse_transcript(self.write("a.jsonl", records))])
        self.assertEqual(stats.tags["H5f"]["denominator"], 10)
        self.assertAlmostEqual(stats.tags["H5f"]["firing_rate"], 0.1)

    def test_self_referential_firings_are_counted_separately(self):
        """Sessions that instrument the guard fire it on purpose. Measured 2026-09-09:
        the 09-07..09-08 window was dominated by hook-development sessions, so a rate
        that does not separate them reads as a behaviour change that never happened."""
        path = self.write("a.jsonl", [
            tool_use("toolu_1", "jq -n ... | bash ~/.agents/hooks/origin_warn_guards.sh"),
            hook_success("toolu_1", GUARD + "[H1] x"),
            tool_use("toolu_2", "pytest -q | tail -1"),
            hook_success("toolu_2", GUARD + "[H1] x"),
        ])
        stats = measure.aggregate([measure.parse_transcript(path)])
        self.assertEqual(stats.tags["H1"]["firings"], 2)
        self.assertEqual(stats.tags["H1"]["self_referential"], 1)
        self.assertEqual(stats.self_referential, 1)

    def test_duration_summary_is_reported(self):
        path = self.write("a.jsonl", [
            tool_use("toolu_1", "a"),
            hook_success("toolu_1", GUARD + "[H1] x", duration=60),
            tool_use("toolu_2", "b"),
            hook_success("toolu_2", GUARD + "[H1] x", duration=148),
        ])
        stats = measure.aggregate([measure.parse_transcript(path)])
        self.assertEqual(stats.duration_max, 148)
        self.assertEqual(stats.duration_median, 104)


class RenderAndSampleTests(unittest.TestCase):
    def setUp(self):
        self.tempdir = tempfile.TemporaryDirectory()
        self.root = Path(self.tempdir.name)
        records = []
        for i in range(5):
            records.append(tool_use(f"toolu_{i}", f"secret-company-name-{i} | tail -1; echo $?"))
            records.append(hook_success(f"toolu_{i}", GUARD + "[H2] x"))
        path = self.root / "a.jsonl"
        with path.open("w", encoding="utf-8") as handle:
            for record in records:
                handle.write(json.dumps(record, ensure_ascii=False) + "\n")
        self.scan = measure.parse_transcript(path)

    def tearDown(self):
        self.tempdir.cleanup()

    def test_report_never_contains_command_text(self):
        """The report is pasted into the cloud-synced triage report; commands are not."""
        rendered = measure.render_report(measure.aggregate([self.scan]))
        self.assertNotIn("secret-company-name", rendered)
        self.assertIn("H2", rendered)

    def test_report_maps_tags_to_response_ids(self):
        rendered = measure.render_report(measure.aggregate([self.scan]))
        self.assertIn("hook-h2-pipeline-status", rendered)

    def test_unmapped_tag_is_reported_as_unregistered(self):
        self.assertEqual(measure.response_id_for("H4"), None)
        self.assertEqual(measure.response_id_for("H5e"), "hook-h5-git-state")

    def test_sample_is_deterministic_for_a_seed_and_carries_commands(self):
        first = measure.sample_firings([self.scan], tag="H2", n=3, seed=7)
        second = measure.sample_firings([self.scan], tag="H2", n=3, seed=7)
        self.assertEqual(len(first), 3)
        self.assertEqual([f["command"] for f in first], [f["command"] for f in second])
        self.assertTrue(all("secret-company-name" in f["command"] for f in first))

    def test_sample_returns_everything_when_n_exceeds_the_population(self):
        self.assertEqual(len(measure.sample_firings([self.scan], tag="H2", n=99, seed=1)), 5)


class CliTests(unittest.TestCase):
    def setUp(self):
        self.tempdir = tempfile.TemporaryDirectory()
        self.root = Path(self.tempdir.name)
        self.transcripts = self.root / "projects/-Users-x-repo"
        self.transcripts.mkdir(parents=True)
        path = self.transcripts / "session.jsonl"
        with path.open("w", encoding="utf-8") as handle:
            for record in [
                tool_use("toolu_1", "cmd | tail -1; echo $?"),
                hook_success("toolu_1", GUARD + "[H2] x"),
            ]:
                handle.write(json.dumps(record, ensure_ascii=False) + "\n")

    def tearDown(self):
        self.tempdir.cleanup()

    def run_cli(self, *args):
        return subprocess.run(
            [sys.executable, "-B", str(SCRIPT), "--transcripts", str(self.root / "projects"), *args],
            check=False,
            capture_output=True,
            text=True,
        )

    def test_report_subcommand_prints_a_table(self):
        result = self.run_cli("report")
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn("H2", result.stdout)
        self.assertIn("hook-h2-pipeline-status", result.stdout)

    def test_report_is_empty_but_succeeds_when_no_firing_is_in_the_window(self):
        result = self.run_cli("report", "--since", "2027-01-01")
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn("0", result.stdout)

    def test_sample_writes_to_a_local_file_and_not_to_stdout(self):
        out = self.root / "sample.json"
        result = self.run_cli("sample", "--tag", "H2", "--n", "2", "--out", str(out))
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertTrue(out.exists())
        self.assertNotIn("cmd | tail", result.stdout)
        payload = json.loads(out.read_text(encoding="utf-8"))
        self.assertEqual(payload["tag"], "H2")
        self.assertEqual(len(payload["firings"]), 1)


if __name__ == "__main__":
    unittest.main()
