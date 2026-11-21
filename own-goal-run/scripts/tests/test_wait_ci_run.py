"""A CI waiter must never accept a run from the wrong commit."""

import importlib.util
import json
import os
import subprocess
import sys
import tempfile
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


# --- transient query retry -------------------------------------------------


def ok(rows):
    return subprocess.CompletedProcess(["gh"], 0, json.dumps(rows), "")


def fail(code=1, stderr=""):
    return subprocess.CompletedProcess(["gh"], code, "", stderr)


class Clock:
    """Fake time: sleeping is the only thing that advances it."""

    def __init__(self):
        self.now = 0.0
        self.sleeps = []

    def monotonic(self):
        return self.now

    def sleep(self, seconds):
        self.sleeps.append(seconds)
        self.now += seconds


class FakeGh:
    """Replays outcomes in order; the last one repeats forever."""

    def __init__(self, *outcomes):
        self.outcomes = list(outcomes)
        self.calls = 0

    def __call__(self, command, **kwargs):
        self.calls += 1
        item = self.outcomes.pop(0) if len(self.outcomes) > 1 else self.outcomes[0]
        if isinstance(item, BaseException):
            raise item
        return item


def wait(gh, clock=None, **overrides):
    clock = clock or Clock()
    params = dict(timeout=600, interval=15, limit=100,
                  query_retries=3, query_retry_interval=10)
    params.update(overrides)
    outcome = MODULE.wait_for_ci(SHA, "CI/CD Pipeline", None, run_command=gh,
                                 sleep=clock.sleep, monotonic=clock.monotonic,
                                 **params)
    return outcome, clock


TRANSIENT = {
    "dns": lambda: fail(1, "dial tcp: lookup api.github.com: no such host"),
    "enotfound": lambda: fail(1, "getaddrinfo ENOTFOUND api.github.com"),
    "tls": lambda: fail(1, "Post https://api.github.com/graphql: net/http: TLS handshake timeout"),
    "bad-gateway": lambda: fail(1, "HTTP 502: Bad Gateway (https://api.github.com/x)"),
    "rate-limited": lambda: fail(1, "HTTP 429: Too Many Requests"),
    "silent": lambda: fail(1, ""),
    "gh-timeout": lambda: subprocess.TimeoutExpired(["gh"], 30),
}

PERMANENT = {
    "auth": lambda: fail(4, "To get started with GitHub CLI, please run: gh auth login"),
    "not-found": lambda: fail(1, "HTTP 404: Not Found (https://api.github.com/repos/o/r)"),
    "forbidden": lambda: fail(1, "HTTP 403: Resource not accessible"),
    "no-workflow": lambda: fail(1, "could not find any workflows named nope"),
    "gh-missing": lambda: FileNotFoundError("gh"),
    "gh-not-executable": lambda: PermissionError("gh"),
    "invalid-json": lambda: subprocess.CompletedProcess(["gh"], 0, "<html>", ""),
    "bad-schema": lambda: ok([{"status": "completed"}]),
}


class TransientRetryTests(unittest.TestCase):
    def test_transient_error_then_success_succeeds(self):
        for name, make in TRANSIENT.items():
            with self.subTest(name):
                gh = FakeGh(make(), ok([run()]))
                outcome, clock = wait(gh)
                self.assertEqual(outcome, [17])
                self.assertEqual(gh.calls, 2)
                self.assertEqual(clock.sleeps, [10])

    def test_retry_interval_is_configurable(self):
        gh = FakeGh(fail(1, ""), ok([run()]))
        _, clock = wait(gh, query_retry_interval=3)
        self.assertEqual(clock.sleeps, [3])

    def test_transient_errors_beyond_the_limit_stop_as_query_error(self):
        gh = FakeGh(fail(1, "net/http: TLS handshake timeout"))
        with self.assertRaises(MODULE.QueryError):
            wait(gh, query_retries=3)
        self.assertEqual(gh.calls, 4)  # first try + 3 retries, no more

    def test_zero_retries_keeps_the_old_stop_on_first_error(self):
        gh = FakeGh(fail(1, ""), ok([run()]))
        with self.assertRaises(MODULE.QueryError):
            wait(gh, query_retries=0)
        self.assertEqual(gh.calls, 1)

    def test_exhausted_retries_never_report_success(self):
        gh = FakeGh(fail(1, ""), fail(1, ""), ok([run()]))
        with self.assertRaises(MODULE.QueryError):
            wait(gh, query_retries=1)

    def test_permanent_errors_are_not_retried(self):
        for name, make in PERMANENT.items():
            with self.subTest(name):
                gh = FakeGh(make(), ok([run()]))
                clock = Clock()
                with self.assertRaises(MODULE.QueryError):
                    wait(gh, clock)
                self.assertEqual(gh.calls, 1)
                self.assertEqual(clock.sleeps, [])

    def test_ci_failure_after_a_transient_error_is_not_retried(self):
        for conclusion in ("failure", "cancelled", "skipped"):
            with self.subTest(conclusion):
                gh = FakeGh(fail(1, ""), ok([run(conclusion=conclusion)]), ok([run()]))
                with self.assertRaises(MODULE.CiFailed):
                    wait(gh)
                self.assertEqual(gh.calls, 2)

    def test_ci_failure_on_first_poll_is_not_retried(self):
        gh = FakeGh(ok([run(conclusion="failure")]), ok([run()]))
        clock = Clock()
        with self.assertRaises(MODULE.CiFailed):
            wait(gh, clock)
        self.assertEqual(gh.calls, 1)
        self.assertEqual(clock.sleeps, [])

    def test_timeout_is_not_extended_by_retries(self):
        for retry_interval in (10, 45, 100):
            with self.subTest(retry_interval):
                gh = FakeGh(fail(1, ""))
                clock = Clock()
                with self.assertRaises(TimeoutError):
                    wait(gh, clock, timeout=60, query_retries=1000,
                         query_retry_interval=retry_interval)
                self.assertLessEqual(clock.now, 60)
                self.assertLess(gh.calls, 10)

    def test_timeout_with_the_default_retry_cap_still_reports_timeout(self):
        gh = FakeGh(fail(1, ""))
        with self.assertRaises(TimeoutError):
            wait(gh, timeout=5, query_retries=100, query_retry_interval=10)

    def test_failure_budget_is_consecutive_not_cumulative(self):
        gh = FakeGh(fail(1, ""), ok([run(status="in_progress", conclusion="")]),
                    fail(1, ""), ok([run()]))
        outcome, _ = wait(gh, query_retries=1)
        self.assertEqual(outcome, [17])
        self.assertEqual(gh.calls, 4)

    def test_stale_commit_never_satisfies_after_a_transient_error(self):
        gh = FakeGh(fail(1, ""), ok([run(OLD_SHA)]))
        with self.assertRaises(TimeoutError):
            wait(gh, timeout=60)

    def test_progress_is_reported_for_each_retry(self):
        seen = []
        gh = FakeGh(fail(1, ""), fail(1, ""), ok([run()]))
        clock = Clock()
        MODULE.wait_for_ci(SHA, "CI/CD Pipeline", None, timeout=600, interval=15,
                           limit=100, query_retries=3, query_retry_interval=10,
                           run_command=gh, sleep=clock.sleep,
                           monotonic=clock.monotonic,
                           retry_notice=lambda attempt, limit, delay: seen.append(
                               (attempt, limit, delay)))
        self.assertEqual(seen, [(1, 3, 10), (2, 3, 10)])


class CommandLineTests(unittest.TestCase):
    """Run the real script against a fake `gh` on PATH, as agents do."""

    def run_cli(self, fails, *extra, rows=None):
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            bin_dir = tmp / "bin"
            bin_dir.mkdir()
            gh = bin_dir / "gh"
            gh.write_text(
                "#!/bin/sh\n"
                'n=$(cat "$GH_FAKE_DIR/count" 2>/dev/null || echo 0)\n'
                'n=$((n+1)); echo "$n" > "$GH_FAKE_DIR/count"\n'
                'if [ "$n" -le "$GH_FAKE_FAILS" ]; then\n'
                '  echo "net/http: TLS handshake timeout" >&2; exit 1\n'
                "fi\n"
                'cat "$GH_FAKE_DIR/rows.json"\n')
            gh.chmod(0o755)
            (tmp / "rows.json").write_text(json.dumps(rows if rows is not None else [run()]))
            env = dict(os.environ, GH_FAKE_DIR=str(tmp), GH_FAKE_FAILS=str(fails),
                       PATH=f"{bin_dir}{os.pathsep}{os.environ['PATH']}")
            proc = subprocess.run(
                [sys.executable, "-I", str(SCRIPT), "--commit", SHA,
                 "--workflow", "CI/CD Pipeline", "--interval", "1", *extra],
                capture_output=True, text=True, env=env, timeout=60)
            count = tmp / "count"
            calls = int(count.read_text()) if count.exists() else 0
        return proc, calls

    def test_cli_recovers_from_a_transient_error(self):
        proc, calls = self.run_cli(1, "--query-retry-interval", "1")
        self.assertEqual(proc.returncode, 0, proc.stderr)
        self.assertEqual(calls, 2)
        self.assertIn("CI succeeded", proc.stdout)

    def test_cli_exits_3_when_the_error_persists(self):
        proc, calls = self.run_cli(99, "--query-retries", "1",
                                   "--query-retry-interval", "1")
        self.assertEqual(proc.returncode, 3, proc.stdout)
        self.assertEqual(calls, 2)
        self.assertNotIn("CI succeeded", proc.stdout)

    def test_cli_default_retries_are_modest(self):
        proc, calls = self.run_cli(99, "--query-retry-interval", "1")
        self.assertEqual(proc.returncode, 3)
        self.assertEqual(calls, 1 + MODULE.DEFAULT_QUERY_RETRIES)
        self.assertLessEqual(MODULE.DEFAULT_QUERY_RETRIES, 5)
        self.assertLessEqual(MODULE.DEFAULT_QUERY_RETRY_INTERVAL, 30)

    def test_cli_failed_ci_exits_1_without_retry(self):
        proc, calls = self.run_cli(0, rows=[run(conclusion="failure")])
        self.assertEqual(proc.returncode, 1)
        self.assertEqual(calls, 1)

    def test_cli_timeout_exits_2(self):
        proc, _ = self.run_cli(99, "--timeout", "2", "--query-retries", "50",
                               "--query-retry-interval", "1")
        self.assertEqual(proc.returncode, 2, proc.stderr)

    def test_cli_rejects_negative_retries_and_accepts_zero(self):
        proc, _ = self.run_cli(0, "--query-retries", "-1")
        self.assertEqual(proc.returncode, 2)
        proc, calls = self.run_cli(1, "--query-retries", "0")
        self.assertEqual(proc.returncode, 3)
        self.assertEqual(calls, 1)


if __name__ == "__main__":
    unittest.main()
