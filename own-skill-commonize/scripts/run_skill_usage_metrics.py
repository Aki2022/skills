#!/usr/bin/env python3
"""Run the aggregate skill-usage collector once with a local lock.

This is the small operational wrapper used by the macOS LaunchAgent.  It keeps
the collector itself useful as a read-only command while adding two properties
needed by a recurring job: only one run may write at a time, and the aggregate
JSONL history has a finite retention window.
"""

from __future__ import annotations

import argparse
import contextlib
import datetime as _datetime
import fcntl
import importlib.util
import io
import json
import os
import sys
import tempfile
from pathlib import Path
from typing import Sequence


UTC = _datetime.timezone.utc
DEFAULT_WINDOW_DAYS = 30
DEFAULT_RETENTION_DAYS = 180
LABEL = "com.origin.skill-usage-metrics"
SCRIPT_DIR = Path(__file__).resolve().parent
COLLECTOR_PATH = SCRIPT_DIR / "measure_skill_usage.py"


def _collector_module():
    spec = importlib.util.spec_from_file_location(
        "_scheduled_measure_skill_usage", COLLECTOR_PATH
    )
    if spec is None or spec.loader is None:
        raise RuntimeError("skill usage collector cannot be loaded")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def _ensure_private_directory(path: Path) -> None:
    if path.is_symlink():
        raise ValueError("state directory must not be a symlink")
    path.mkdir(parents=True, exist_ok=True, mode=0o700)
    # A pre-existing directory may have been created with a broader umask.  The
    # snapshots are aggregate but still account metadata, so tighten it without
    # changing ownership or any sibling paths.
    mode = path.stat().st_mode & 0o777
    if mode & 0o077:
        path.chmod(0o700)


def _ensure_private_file(path: Path) -> None:
    """Reject symlinks and tighten an existing regular file to owner-only."""

    if path.is_symlink():
        raise ValueError("state file must not be a symlink")
    if path.exists():
        if not path.is_file():
            raise ValueError("state file is not a regular file")
        mode = path.stat().st_mode & 0o777
        if mode & 0o077:
            path.chmod(0o600)


def _valid_report_timestamp(value: object) -> bool:
    if not isinstance(value, str) or not value.strip():
        return False
    text = value.strip()
    if text.endswith(("Z", "z")):
        text = text[:-1] + "+00:00"
    try:
        parsed = _datetime.datetime.fromisoformat(text)
    except ValueError:
        return False
    return parsed.tzinfo is not None


def _snapshot_documents(path: Path) -> list[tuple[str, dict]]:
    """Read and validate one aggregate JSONL file without rewriting it."""

    if path.is_symlink():
        raise ValueError("snapshot path must not be a symlink")
    if not path.exists():
        return []
    if not path.is_file():
        raise ValueError("snapshot path is not a regular file")
    try:
        lines = path.read_text(encoding="utf-8").splitlines()
    except OSError as error:
        raise ValueError("snapshot cannot be read") from error
    documents: list[tuple[str, dict]] = []
    for line in lines:
        if not line.strip():
            raise ValueError("snapshot contains an empty line")
        try:
            document = json.loads(line)
        except (TypeError, ValueError):
            raise ValueError("snapshot contains malformed JSON") from None
        if not isinstance(document, dict):
            raise ValueError("snapshot line is not an object")
        required = {
            "schema_version",
            "generated_at",
            "window",
            "summary",
            "sources",
            "dependency_scan_errors",
            "skills",
            "privacy",
        }
        if not required.issubset(document):
            raise ValueError("snapshot is not an aggregate report")
        if document.get("schema_version") != 1:
            raise ValueError("snapshot schema version is unsupported")
        window = document.get("window")
        if (
            not isinstance(window, dict)
            or not isinstance(window.get("since"), str)
            or not isinstance(window.get("until"), str)
            or not _valid_report_timestamp(window.get("since"))
            or not _valid_report_timestamp(window.get("until"))
        ):
            raise ValueError("snapshot window is missing")
        summary = document.get("summary")
        if not isinstance(summary, dict):
            raise ValueError("snapshot summary is missing")
        summary_fields = (
            "canonical_skills",
            "used",
            "dependency_only",
            "unknown",
            "candidate",
        )
        if any(
            not isinstance(summary.get(key), int)
            or isinstance(summary.get(key), bool)
            or summary[key] < 0
            for key in summary_fields
        ):
            raise ValueError("snapshot summary is invalid")
        if not isinstance(document.get("sources"), list) or not isinstance(
            document.get("skills"), list
        ):
            raise ValueError("snapshot aggregate rows are missing")
        dependency_errors = document.get("dependency_scan_errors")
        if (
            not isinstance(dependency_errors, int)
            or isinstance(dependency_errors, bool)
            or dependency_errors < 0
        ):
            raise ValueError("snapshot dependency error count is invalid")
        if summary["canonical_skills"] != len(document["skills"]):
            raise ValueError("snapshot skill count is inconsistent")
        if sum(summary[key] for key in summary_fields[1:]) != summary["canonical_skills"]:
            raise ValueError("snapshot status counts are inconsistent")
        for row in document["skills"]:
            if not isinstance(row, dict) or not isinstance(row.get("skill"), str):
                raise ValueError("snapshot skill rows are invalid")
            if row.get("status") not in {"used", "dependency-only", "unknown"}:
                raise ValueError("snapshot skill status is invalid")
        privacy = document.get("privacy")
        if not isinstance(privacy, dict) or any(
            privacy.get(key) is not False
            for key in (
                "raw_prompts_persisted",
                "paths_persisted",
                "session_ids_persisted",
                "credentials_persisted",
            )
        ):
            raise ValueError("snapshot privacy contract is missing or unsafe")
        generated = document.get("generated_at")
        if not _valid_report_timestamp(generated):
            raise ValueError("snapshot generated_at is missing")
        documents.append((generated, document))
    return documents


def prune_snapshot(
    path: str | Path,
    retention_days: int = DEFAULT_RETENTION_DAYS,
    *,
    now: _datetime.datetime | None = None,
) -> int:
    """Drop valid aggregate snapshots older than ``retention_days`` atomically."""

    if retention_days <= 0:
        raise ValueError("retention_days must be greater than zero")
    snapshot = Path(path).expanduser()
    documents = _snapshot_documents(snapshot)
    if not documents:
        return 0
    collector = _collector_module()
    current = now or _datetime.datetime.now(tz=UTC)
    if current.tzinfo is None:
        current = current.replace(tzinfo=UTC)
    current = current.astimezone(UTC)
    cutoff = current - _datetime.timedelta(days=retention_days)
    keep: list[str] = []
    removed = 0
    for generated, document in documents:
        parsed = collector._timestamp(generated)
        if parsed is None:
            raise ValueError("snapshot generated_at is invalid")
        if parsed < cutoff:
            removed += 1
        else:
            keep.append(
                json.dumps(
                    document,
                    ensure_ascii=False,
                    separators=(",", ":"),
                    sort_keys=True,
                )
            )
    if not removed:
        return 0
    if snapshot.parent.is_symlink() or not snapshot.parent.is_dir():
        raise ValueError("snapshot parent directory is unavailable")
    temporary_name: str | None = None
    try:
        with tempfile.NamedTemporaryFile(
            "w",
            encoding="utf-8",
            dir=snapshot.parent,
            prefix=f".{snapshot.name}.",
            delete=False,
        ) as handle:
            temporary_name = handle.name
            handle.write("\n".join(keep))
            if keep:
                handle.write("\n")
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary_name, snapshot)
        temporary_name = None
    finally:
        if temporary_name:
            try:
                os.unlink(temporary_name)
            except OSError:
                pass
    return removed


def collector_arguments(
    home: str | Path,
    snapshot: str | Path,
    window_days: int = DEFAULT_WINDOW_DAYS,
) -> list[str]:
    """Build the deterministic four-account collector argument vector."""

    if window_days <= 0:
        raise ValueError("window_days must be greater than zero")
    home_path = Path(home).expanduser()
    snapshot_path = Path(snapshot).expanduser()
    return [
        "--canonical",
        str(home_path / ".agents/skills"),
        "--claude",
        f"seat1={home_path / '.claude/projects'}",
        "--claude",
        f"seat2={home_path / '.claude-seat2/projects'}",
        "--codex",
        f"seat1={home_path / '.codex/sessions'}",
        "--codex",
        f"seat2={home_path / '.codex-seat2/sessions'}",
        "--days",
        str(window_days),
        "--append",
        str(snapshot_path),
    ]


def run_once(
    *,
    home: str | Path | None = None,
    snapshot: str | Path | None = None,
    lock_path: str | Path | None = None,
    window_days: int = DEFAULT_WINDOW_DAYS,
    retention_days: int = DEFAULT_RETENTION_DAYS,
) -> int:
    """Run one guarded collection; return zero when another run owns the lock."""

    home_path = Path(home).expanduser() if home is not None else Path.home()
    state_dir = home_path / ".local/state/origin-skill-usage"
    snapshot_path = (
        Path(snapshot).expanduser()
        if snapshot is not None
        else state_dir / "skill-usage.jsonl"
    )
    lock = (
        Path(lock_path).expanduser()
        if lock_path is not None
        else state_dir / "collector.lock"
    )
    if not snapshot_path.is_absolute():
        snapshot_path = state_dir / snapshot_path
    if not lock.is_absolute():
        lock = state_dir / lock
    if retention_days <= 0 or window_days <= 0:
        raise ValueError("window_days and retention_days must be greater than zero")
    _ensure_private_directory(state_dir)
    if snapshot_path.parent != state_dir:
        _ensure_private_directory(snapshot_path.parent)
    if lock.parent != state_dir:
        _ensure_private_directory(lock.parent)
    _ensure_private_file(snapshot_path)
    _ensure_private_file(lock)
    lock.parent.mkdir(parents=True, exist_ok=True)
    with lock.open("a+", encoding="utf-8") as handle:
        lock.chmod(0o600)
        try:
            fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError:
            return 0
        try:
            # Validate the existing file before appending.  A malformed line is
            # never silently carried forward or discarded by a pruning run.
            before_count = len(_snapshot_documents(snapshot_path))
            collector = _collector_module()
            # The collector normally writes no stdout/stderr in append mode,
            # but capture both streams so a schema error can never copy a
            # source path, prompt, or other diagnostic into launchd logs.
            with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(
                io.StringIO()
            ):
                result = collector.main(
                    collector_arguments(home_path, snapshot_path, window_days)
                )
            if result != 0:
                raise RuntimeError("skill usage collection failed")
            _ensure_private_file(snapshot_path)
            after_count = len(_snapshot_documents(snapshot_path))
            if after_count != before_count + 1:
                raise RuntimeError(
                    "collector did not append exactly one aggregate snapshot"
                )
            prune_snapshot(snapshot_path, retention_days)
            return 0
        finally:
            fcntl.flock(handle.fileno(), fcntl.LOCK_UN)


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Run one locked skill usage snapshot collection.")
    parser.add_argument("--window-days", type=int, default=DEFAULT_WINDOW_DAYS)
    parser.add_argument("--retention-days", type=int, default=DEFAULT_RETENTION_DAYS)
    parser.add_argument("--snapshot", help="override the aggregate JSONL path (tests/manual use)")
    parser.add_argument("--lock", help="override the lock path (tests/manual use)")
    args = parser.parse_args(argv)
    try:
        return run_once(
            snapshot=args.snapshot,
            lock_path=args.lock,
            window_days=args.window_days,
            retention_days=args.retention_days,
        )
    except (OSError, RuntimeError, ValueError, TypeError):
        # Do not echo exception details: scheduler logs must not become a side
        # channel for account paths or transcript-derived text.
        print("error: scheduled skill usage collection failed", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
