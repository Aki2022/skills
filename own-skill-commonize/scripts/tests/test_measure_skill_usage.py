"""Tests for the secret-safe Claude/Codex skill usage collector."""

from __future__ import annotations

import importlib.util
import json
import sys
import tempfile
import unittest
from contextlib import redirect_stderr
from io import StringIO
from datetime import datetime, timezone
from pathlib import Path


SCRIPT = Path(__file__).resolve().parents[1] / "measure_skill_usage.py"
SPEC = importlib.util.spec_from_file_location("measure_skill_usage", SCRIPT)
assert SPEC and SPEC.loader
measure = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = measure
SPEC.loader.exec_module(measure)


CANONICAL = {"alpha-skill", "beta-skill"}
SINCE = datetime(2026, 9, 1, tzinfo=timezone.utc)
UNTIL = datetime(2026, 10, 1, tzinfo=timezone.utc)


def claude_user(
    text, *, timestamp="2026-09-10T00:00:00Z", prompt_source="sdk", sidechain=False
):
    return {
        "type": "user",
        "timestamp": timestamp,
        "sessionId": "session-secret",
        "isSidechain": sidechain,
        "promptSource": prompt_source,
        "message": {"role": "user", "content": text},
    }


def claude_tool_result(text):
    return {
        "type": "user",
        "timestamp": "2026-09-10T00:00:01Z",
        "message": {
            "role": "user",
            "content": [
                {"type": "tool_result", "content": text, "tool_use_id": "tool-secret"}
            ],
        },
    }


def codex_message(text, *, timestamp="2026-09-10T00:00:00Z", actual=True):
    kinds = ["user.text"] if actual else ["agents_md.instructions"]
    return {
        "timestamp": timestamp,
        "payload": {
            "type": "message",
            "role": "user",
            "id": "message-secret",
            "internal_chat_message_metadata_passthrough": {
                "content_item_kinds": kinds,
            },
            "content": [{"type": "input_text", "text": text}],
        },
    }


def write_jsonl(path: Path, records):
    with path.open("w", encoding="utf-8") as handle:
        for record in records:
            handle.write(json.dumps(record, ensure_ascii=False) + "\n")
    return path


class ExtractInvocationTests(unittest.TestCase):
    def test_explicit_dollar_and_slash_tokens_are_counted_once(self):
        text = "$alpha-skill then /beta-skill and `$alpha-skill`."
        self.assertEqual(
            measure.extract_invocations(text, CANONICAL), {"alpha-skill", "beta-skill"}
        )

    def test_path_like_slash_is_not_a_skill_invocation(self):
        self.assertEqual(
            measure.extract_invocations("/alpha-skill/SKILL.md", CANONICAL), set()
        )

    def test_command_name_tag_is_supported(self):
        text = "<command-name>/alpha-skill</command-name>"
        self.assertEqual(measure.extract_invocations(text, CANONICAL), {"alpha-skill"})


class ClaudeScanTests(unittest.TestCase):
    def setUp(self):
        self.tempdir = tempfile.TemporaryDirectory()
        self.root = Path(self.tempdir.name)

    def tearDown(self):
        self.tempdir.cleanup()

    def test_skips_injected_and_tool_result_text(self):
        path = write_jsonl(
            self.root / "session.jsonl",
            [
                claude_user("$alpha-skill", prompt_source="sdk"),
                claude_user(
                    "Caveat: injected $beta-skill", prompt_source=None
                ),
                claude_tool_result("tool output mentions $beta-skill"),
                claude_user(
                    "<command-name>/beta-skill</command-name>", prompt_source=None
                ),
            ],
        )
        result = measure.scan_claude(path, CANONICAL, SINCE, UNTIL)
        self.assertEqual(
            [event.skill for event in result.events], ["alpha-skill", "beta-skill"]
        )
        self.assertEqual(result.stats.prompt_messages, 2)
        self.assertEqual(result.stats.matched_messages, 2)

    def test_window_and_sidechain_are_recorded_without_leaking_ids(self):
        path = write_jsonl(
            self.root / "session.jsonl",
            [
                claude_user("$alpha-skill", timestamp="2026-08-31T23:59:59Z"),
                claude_user("$alpha-skill", sidechain=True),
            ],
        )
        result = measure.scan_claude(path, CANONICAL, SINCE, UNTIL)
        self.assertEqual(len(result.events), 1)
        self.assertTrue(result.events[0].sidechain)
        self.assertNotIn("session-secret", repr(result.events[0]))


class CodexScanTests(unittest.TestCase):
    def setUp(self):
        self.tempdir = tempfile.TemporaryDirectory()
        self.root = Path(self.tempdir.name)

    def tearDown(self):
        self.tempdir.cleanup()

    def test_only_user_text_metadata_is_counted(self):
        path = write_jsonl(
            self.root / "rollout.jsonl",
            [
                codex_message("$alpha-skill", actual=True),
                codex_message("Injected $beta-skill", actual=False),
            ],
        )
        result = measure.scan_codex(path, CANONICAL, SINCE, UNTIL)
        self.assertEqual([event.skill for event in result.events], ["alpha-skill"])
        self.assertEqual(result.stats.ambiguous_messages, 0)

    def test_missing_metadata_is_ambiguous_not_use(self):
        record = codex_message("$alpha-skill", actual=True)
        del record["payload"]["internal_chat_message_metadata_passthrough"]
        result = measure.scan_codex(
            write_jsonl(self.root / "rollout.jsonl", [record]), CANONICAL, SINCE, UNTIL
        )
        self.assertEqual(result.events, [])
        self.assertEqual(result.stats.ambiguous_messages, 1)


class ReportTests(unittest.TestCase):
    def test_dependency_only_and_unknown_are_distinct(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp) / "canonical"
            for name in CANONICAL:
                skill = root / name
                skill.mkdir(parents=True)
                body = "---\nname: %s\ndescription: fixture\n---\n" % name
                if name == "alpha-skill":
                    body += "See /beta-skill/references/guide.md.\n"
                (skill / "SKILL.md").write_text(body, encoding="utf-8")
            source_root = Path(tmp) / "claude-projects"
            source_root.mkdir()
            write_jsonl(source_root / "session.jsonl", [claude_user("$alpha-skill")])
            report = measure.build_report(
                root,
                [measure.Source("claude", "seat1", source_root, "claude")],
                SINCE,
                UNTIL,
                generated_at="2026-09-30T00:00:00Z",
            )
        rows = {row["skill"]: row for row in report["skills"]}
        self.assertEqual(rows["alpha-skill"]["status"], "used")
        self.assertEqual(rows["beta-skill"]["status"], "dependency-only")

    def test_rendered_report_contains_no_raw_prompt_or_path(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp) / "canonical"
            skill = root / "alpha-skill"
            skill.mkdir(parents=True)
            (skill / "SKILL.md").write_text(
                "---\nname: alpha-skill\ndescription: fixture\n---\n", encoding="utf-8"
            )
            source_root = Path(tmp) / "claude-projects"
            source_root.mkdir()
            write_jsonl(
                source_root / "session.jsonl", [claude_user("secret prompt $alpha-skill")]
            )
            report = measure.build_report(
                root,
                [measure.Source("claude", "seat1", source_root, "claude")],
                SINCE,
                UNTIL,
                generated_at="2026-09-30T00:00:00Z",
            )
        rendered = json.dumps(report, ensure_ascii=False)
        self.assertNotIn("secret prompt", rendered)
        self.assertNotIn(str(source_root), rendered)
        self.assertEqual(report["privacy"]["raw_prompts_persisted"], False)


class CLITests(unittest.TestCase):
    def test_append_writes_one_json_snapshot_per_line(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp) / "canonical"
            skill = root / "alpha-skill"
            skill.mkdir(parents=True)
            (skill / "SKILL.md").write_text(
                "---\nname: alpha-skill\ndescription: fixture\n---\n", encoding="utf-8"
            )
            source_root = Path(tmp) / "claude-projects"
            source_root.mkdir()
            write_jsonl(source_root / "session.jsonl", [claude_user("secret $alpha-skill")])
            output = Path(tmp) / "snapshots.jsonl"
            errors = StringIO()
            with redirect_stderr(errors):
                rc = measure.main(
                    [
                        "--canonical",
                        str(root),
                        "--claude",
                        f"seat1={source_root}",
                        "--since",
                        "2026-09-01T00:00:00Z",
                        "--until",
                        "2026-10-01T00:00:00Z",
                        "--append",
                        str(output),
                    ]
                )
            self.assertEqual(rc, 0, errors.getvalue())
            lines = output.read_text(encoding="utf-8").splitlines()
            self.assertEqual(len(lines), 1)
            report = json.loads(lines[0])
            self.assertEqual(report["summary"]["used"], 1)
            self.assertNotIn("secret", lines[0])
            self.assertNotIn(str(source_root), lines[0])


if __name__ == "__main__":
    unittest.main()
