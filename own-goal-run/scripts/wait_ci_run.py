#!/usr/bin/env python3
"""Wait for a named GitHub Actions workflow on one exact commit SHA."""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
import time
from typing import Callable


class QueryError(RuntimeError):
    """The run list could not be read or did not have the expected schema."""


class CiFailed(RuntimeError):
    """At least one matching run completed without success."""


def commit_sha(value: str) -> str:
    if not re.fullmatch(r"[0-9a-fA-F]{40,64}", value):
        raise argparse.ArgumentTypeError("--commit requires a full 40- or 64-character hex SHA")
    return value.lower()


def positive_int(value: str) -> int:
    try:
        number = int(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError("value must be a positive integer") from exc
    if number <= 0:
        raise argparse.ArgumentTypeError("value must be a positive integer")
    return number


def build_command(sha: str, workflow: str, repo: str | None, limit: int) -> list[str]:
    command = ["gh", "run", "list", "--commit", sha, "--workflow", workflow,
               "--limit", str(limit), "--json", "headSha,databaseId,status,conclusion"]
    if repo:
        command += ["--repo", repo]
    return command


def classify(rows: object, sha: str) -> tuple[str, list[int]]:
    """Return waiting, failed, or success for matching runs only."""
    if not isinstance(rows, list):
        raise QueryError("gh returned a non-list JSON result")
    matching = []
    for row in rows:
        if not isinstance(row, dict) or not all(
            key in row for key in ("headSha", "databaseId", "status", "conclusion")
        ):
            raise QueryError("gh returned a run with missing fields")
        head_sha = row["headSha"]
        if not isinstance(head_sha, str):
            raise QueryError("gh returned a run without a commit SHA")
        if not isinstance(row["databaseId"], int) or row["databaseId"] <= 0:
            raise QueryError("gh returned a run without a valid ID")
        if not isinstance(row["status"], str):
            raise QueryError("gh returned a run without a status")
        if head_sha.lower() == sha.lower():
            matching.append(row)

    if not matching:
        return "waiting", []
    ids = [row["databaseId"] for row in matching]
    if any(row["status"] == "completed" and row["conclusion"] != "success"
           for row in matching):
        return "failed", ids
    if any(row["status"] != "completed" for row in matching):
        return "waiting", ids
    return "success", ids


def wait_for_ci(
    sha: str,
    workflow: str,
    repo: str | None,
    *,
    timeout: int,
    interval: int,
    limit: int,
    run_command: Callable = subprocess.run,
    sleep: Callable = time.sleep,
    monotonic: Callable = time.monotonic,
    progress: Callable = lambda _state, _ids: None,
) -> list[int]:
    deadline = monotonic() + timeout
    command = build_command(sha, workflow, repo, limit)
    while True:
        try:
            result = run_command(command, capture_output=True, text=True,
                                 check=False, timeout=30)
        except (OSError, subprocess.TimeoutExpired) as exc:
            raise QueryError("gh run list could not be executed") from exc
        if result.returncode != 0:
            raise QueryError(f"gh run list failed with exit code {result.returncode}")
        try:
            rows = json.loads(result.stdout)
        except json.JSONDecodeError as exc:
            raise QueryError("gh run list returned invalid JSON") from exc
        state, ids = classify(rows, sha)
        progress(state, ids)
        if state == "success":
            return ids
        if state == "failed":
            raise CiFailed(f"matching run failed: {','.join(map(str, ids))}")
        remaining = deadline - monotonic()
        if remaining <= 0:
            raise TimeoutError("matching run did not succeed before the timeout")
        sleep(min(interval, remaining))


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--commit", required=True, type=commit_sha,
                        help="full commit SHA to verify; a prior SHA never satisfies this wait")
    parser.add_argument("--workflow", required=True,
                        help="workflow name or file recognized by gh run list")
    parser.add_argument("--repo", help="[HOST/]OWNER/REPO; defaults to the current repository")
    parser.add_argument("--timeout", type=positive_int, default=1800,
                        help="maximum seconds to wait (default: 1800)")
    parser.add_argument("--interval", type=positive_int, default=15,
                        help="poll interval in seconds (default: 15)")
    parser.add_argument("--limit", type=positive_int, default=100,
                        help="maximum runs read per poll (default: 100)")
    args = parser.parse_args(argv)

    def report(state: str, ids: list[int]) -> None:
        print(f"{state}: matching runs {','.join(map(str, ids)) or 'none'}", flush=True)

    try:
        ids = wait_for_ci(args.commit, args.workflow, args.repo,
                          timeout=args.timeout, interval=args.interval,
                          limit=args.limit, progress=report)
    except CiFailed as exc:
        print(f"CI failed: {exc}", file=sys.stderr)
        return 1
    except TimeoutError as exc:
        print(f"CI wait timed out: {exc}", file=sys.stderr)
        return 2
    except QueryError as exc:
        print(f"CI query error: {exc}", file=sys.stderr)
        return 3
    print(f"CI succeeded for {args.commit[:12]}: {','.join(map(str, ids))}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
