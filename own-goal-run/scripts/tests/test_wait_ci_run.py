"""A CI waiter must never accept a run from the wrong commit."""

import importlib.util
import json
import subprocess
import sys
import unittest
from pathlib import Path


SCRIPT = Path(__file__).resolve().parents[1] / "wait_ci_run.py"
SPEC = importlib.util.spec_from_file_location("wait_ci_run", SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = MODULE
SPEC.loader.exec_module(MODULE)

SHA = "a" * 40
OLD_SHA = "b" * 40


def run(sha=SHA, status="completed", conclusion="success", run_id=17):
    return {"headSha": sha, "status": status, "conclusion": conclusion,
            "databaseId": run_id}


class WaitCiRunTests(unittest.TestCase):
    def test_query_pins_commit_and_workflow(self):
        command = MODULE.build_command(SHA, "CI/CD Pipeline", "owner/repo", 100)
        self.assertEqual(command[:3], ["gh", "run", "list"])
        self.assertEqual(command[command.index("--commit") + 1], SHA)
        self.assertEqual(command[command.index("--workflow") + 1], "CI/CD Pipeline")
        self.assertEqual(command[command.index("--repo") + 1], "owner/repo")

    def test_empty_and_old_commit_are_not_success(self):
        self.assertEqual(MODULE.classify([], SHA), ("waiting", []))
        self.assertEqual(MODULE.classify([run(OLD_SHA)], SHA), ("waiting", []))

    def test_pending_target_keeps_waiting_even_with_completed_run(self):
        rows = [run(run_id=1), run(status="in_progress", conclusion="", run_id=2)]
        self.assertEqual(MODULE.classify(rows, SHA), ("waiting", [1, 2]))

    def test_failed_target_is_not_success(self):
        self.assertEqual(MODULE.classify([run(conclusion="failure")], SHA),
                         ("failed", [17]))
        self.assertEqual(MODULE.classify([run(conclusion="skipped")], SHA),
                         ("failed", [17]))

    def test_failed_target_stops_the_wait(self):
        def fake_run(command, **kwargs):
            return subprocess.CompletedProcess(command, 0,
                                               json.dumps([run(conclusion="failure")]), "")

        with self.assertRaises(MODULE.CiFailed):
            MODULE.wait_for_ci(SHA, "CI/CD Pipeline", None, timeout=5,
                               interval=0, limit=100, run_command=fake_run,
                               sleep=lambda _: None)

    def test_old_run_then_matching_run_requires_second_poll(self):
        replies = [[run(OLD_SHA)], [run(SHA)]]
        calls = []

        def fake_run(command, **kwargs):
            calls.append(command)
            return subprocess.CompletedProcess(command, 0, json.dumps(replies.pop(0)), "")

        outcome = MODULE.wait_for_ci(SHA, "CI/CD Pipeline", None, timeout=5,
                                     interval=0, limit=100, run_command=fake_run,
                                     sleep=lambda _: None)
        self.assertEqual(outcome, [17])
        self.assertEqual(len(calls), 2)

    def test_gh_error_stops_without_reporting_success(self):
        def fake_run(command, **kwargs):
            return subprocess.CompletedProcess(command, 4, "", "auth failed")

        with self.assertRaises(MODULE.QueryError):
            MODULE.wait_for_ci(SHA, "CI/CD Pipeline", None, timeout=5,
                               interval=0, limit=100, run_command=fake_run,
                               sleep=lambda _: None)

    def test_malformed_result_stops_without_reporting_success(self):
        with self.assertRaises(MODULE.QueryError):
            MODULE.classify([{"status": "completed", "conclusion": "success"}], SHA)
        with self.assertRaises(MODULE.QueryError):
            MODULE.classify([run(run_id=None)], SHA)

    def test_no_matching_run_times_out(self):
        ticks = iter([0, 1])

        def fake_run(command, **kwargs):
            return subprocess.CompletedProcess(command, 0, "[]", "")

        with self.assertRaises(TimeoutError):
            MODULE.wait_for_ci(SHA, "CI/CD Pipeline", None, timeout=1,
                               interval=0, limit=100, run_command=fake_run,
                               sleep=lambda _: None, monotonic=lambda: next(ticks))


if __name__ == "__main__":
    unittest.main()
