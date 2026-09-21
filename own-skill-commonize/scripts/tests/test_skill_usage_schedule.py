"""Tests for the locked aggregate runner and macOS LaunchAgent installer."""

from __future__ import annotations

import fcntl
import importlib.util
import json
import os
import plistlib
import sys
import tempfile
import unittest
from datetime import datetime, timezone
from pathlib import Path
from unittest.mock import patch


SCRIPTS = Path(__file__).resolve().parents[1]


def _load(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, path)
    assert spec and spec.loader
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


runner = _load("run_skill_usage_metrics_test_module", SCRIPTS / "run_skill_usage_metrics.py")
installer = _load(
    "install_skill_usage_launchd_test_module",
    SCRIPTS / "install_skill_usage_launchd.py",
)


def _snapshot(generated_at: str) -> dict[str, object]:
    return {
        "schema_version": 1,
        "generated_at": generated_at,
        "window": {"since": generated_at, "until": generated_at},
        "summary": {
            "canonical_skills": 0,
            "used": 0,
            "dependency_only": 0,
            "unknown": 0,
            "candidate": 0,
        },
        "sources": [],
        "dependency_scan_errors": 0,
        "skills": [],
        "privacy": {
            "raw_prompts_persisted": False,
            "paths_persisted": False,
            "session_ids_persisted": False,
            "credentials_persisted": False,
        },
    }


def _write_snapshots(path: Path, *generated_at: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        "".join(
            json.dumps(_snapshot(value), ensure_ascii=False, separators=(",", ":"), sort_keys=True)
            + "\n"
            for value in generated_at
        ),
        encoding="utf-8",
    )


class SnapshotTests(unittest.TestCase):
    def test_prune_removes_old_entries_atomically(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "skill-usage.jsonl"
            _write_snapshots(path, "2026-01-01T00:00:00Z", "2026-09-01T00:00:00Z")

            removed = runner.prune_snapshot(
                path,
                retention_days=30,
                now=datetime(2026, 9, 9, tzinfo=timezone.utc),
            )

            self.assertEqual(removed, 1)
            lines = path.read_text(encoding="utf-8").splitlines()
            self.assertEqual(len(lines), 1)
            self.assertEqual(json.loads(lines[0])["generated_at"], "2026-09-01T00:00:00Z")

    def test_prune_rejects_malformed_snapshot_without_rewriting(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "skill-usage.jsonl"
            _write_snapshots(path, "2026-09-01T00:00:00Z")
            with path.open("a", encoding="utf-8") as handle:
                handle.write("not-json\n")
            original = path.read_bytes()

            with self.assertRaises(ValueError):
                runner.prune_snapshot(
                    path,
                    retention_days=30,
                    now=datetime(2026, 9, 9, tzinfo=timezone.utc),
                )

            self.assertEqual(path.read_bytes(), original)

    def test_prune_rejects_non_aggregate_json_without_rewriting(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "skill-usage.jsonl"
            path.write_text(
                json.dumps(
                    {
                        "generated_at": "2026-09-01T00:00:00Z",
                        "privacy": {
                            "raw_prompts_persisted": False,
                            "paths_persisted": False,
                            "session_ids_persisted": False,
                            "credentials_persisted": False,
                        },
                    }
                )
                + "\n",
                encoding="utf-8",
            )
            original = path.read_bytes()

            with self.assertRaises(ValueError):
                runner.prune_snapshot(
                    path,
                    retention_days=30,
                    now=datetime(2026, 9, 9, tzinfo=timezone.utc),
                )

            self.assertEqual(path.read_bytes(), original)

    def test_collector_arguments_are_explicit_four_account_vector(self):
        args = runner.collector_arguments("/tmp/example-home", "/tmp/snapshot.jsonl", 30)

        self.assertEqual(args.count("--claude"), 2)
        self.assertEqual(args.count("--codex"), 2)
        self.assertEqual(
            args[args.index("--canonical") + 1], "/tmp/example-home/.agents/skills"
        )
        self.assertIn("seat1=/tmp/example-home/.claude/projects", args)
        self.assertIn("seat2=/tmp/example-home/.claude-seat2/projects", args)
        self.assertIn("seat1=/tmp/example-home/.codex/sessions", args)
        self.assertIn("seat2=/tmp/example-home/.codex-seat2/sessions", args)
        self.assertEqual(args[-2:], ["--append", "/tmp/snapshot.jsonl"])


class RunnerTests(unittest.TestCase):
    def test_run_once_appends_through_collector_and_uses_private_state(self):
        class FakeCollector:
            @staticmethod
            def main(args):
                output = Path(args[args.index("--append") + 1])
                output.parent.mkdir(parents=True, exist_ok=True)
                output.write_text(
                    json.dumps(_snapshot("2026-09-09T00:00:00Z"), separators=(",", ":")) + "\n",
                    encoding="utf-8",
                )
                return 0

            @staticmethod
            def _timestamp(value):
                return runner.UTC if value == "not-used" else datetime.fromisoformat(
                    value.replace("Z", "+00:00")
                )

        with tempfile.TemporaryDirectory() as tmp, patch.object(
            runner, "_collector_module", return_value=FakeCollector
        ) as load_collector, patch.object(runner, "prune_snapshot", return_value=0):
            home = Path(tmp)
            self.assertEqual(
                runner.run_once(home=home, retention_days=180),
                0,
            )

            snapshot = home / ".local/state/origin-skill-usage/skill-usage.jsonl"
            lock = home / ".local/state/origin-skill-usage/collector.lock"
            self.assertEqual(len(snapshot.read_text(encoding="utf-8").splitlines()), 1)
            self.assertEqual(snapshot.stat().st_mode & 0o777, 0o600)
            self.assertEqual(lock.stat().st_mode & 0o777, 0o600)
            self.assertEqual(load_collector.call_count, 1)

    def test_run_once_skips_when_another_process_holds_lock(self):
        with tempfile.TemporaryDirectory() as tmp, patch.object(
            runner, "_collector_module"
        ) as load:
            home = Path(tmp)
            state = home / ".local/state/origin-skill-usage"
            state.mkdir(parents=True)
            lock = state / "collector.lock"
            with lock.open("a+", encoding="utf-8") as handle:
                fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                self.assertEqual(runner.run_once(home=home), 0)
            load.assert_not_called()

    def test_run_once_fails_closed_when_collector_does_not_append(self):
        class EmptyCollector:
            @staticmethod
            def main(_args):
                return 0

        with tempfile.TemporaryDirectory() as tmp, patch.object(
            runner, "_collector_module", return_value=EmptyCollector
        ):
            with self.assertRaises(RuntimeError):
                runner.run_once(home=Path(tmp))


class InstallerTests(unittest.TestCase):
    def test_template_is_home_independent_and_has_authorized_schedule(self):
        parsed = plistlib.loads(installer.template_bytes())

        self.assertEqual(parsed["Label"], installer.LABEL)
        self.assertEqual(parsed["StartCalendarInterval"], {"Hour": 3, "Minute": 30})
        self.assertIs(parsed["RunAtLoad"], False)
        arguments = parsed["ProgramArguments"]
        self.assertEqual(arguments[:2], ["/bin/sh", "-c"])
        command = arguments[2]
        self.assertIn("$HOME", command)
        self.assertNotIn("/Users/", command)

    def test_install_is_idempotent_and_check_is_read_only(self):
        with tempfile.TemporaryDirectory() as tmp, patch.object(
            installer, "is_loaded", return_value=False
        ):
            home = Path(tmp)
            destination = installer.install_plist(home)
            self.assertEqual(destination.read_bytes(), installer.template_bytes())
            self.assertEqual(destination.stat().st_mode & 0o777, 0o600)
            self.assertEqual(
                (home / ".local/state/origin-skill-usage").stat().st_mode & 0o777,
                0o700,
            )

            self.assertEqual(installer.install_plist(home), destination)
            result = installer.check(home)
            self.assertEqual(
                result,
                {
                    "template_valid": True,
                    "plist_installed": True,
                    "plist_exact": True,
                    "loaded": False,
                },
            )

    def test_install_refuses_different_existing_plist_without_replace(self):
        with tempfile.TemporaryDirectory() as tmp:
            home = Path(tmp)
            destination = installer.installed_path(home)
            destination.parent.mkdir(parents=True)
            destination.write_bytes(b"different")

            with self.assertRaises(RuntimeError):
                installer.install_plist(home)


if __name__ == "__main__":
    unittest.main()
