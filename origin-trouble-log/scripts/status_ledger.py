#!/usr/bin/env python3
"""Maintain the mutable status ledger for origin-trouble-log entries.

Entry bodies are evidence and remain append-only.  This file is the small,
machine-readable mutable layer used to record triage and response state.
"""

from __future__ import annotations

import argparse
import csv
import os
import re
import sys
import tempfile
from collections import Counter
from datetime import date
from pathlib import Path


ENTRY_NAME_RE = re.compile(r"^[0-9]{4}-[0-9]{2}-[0-9]{2}-[a-z0-9-]+\.md$")
DATE_RE = re.compile(r"^[0-9]{4}-[0-9]{2}-[0-9]{2}$")
TARGET_RE = re.compile(r"^- ([0-9]{4}-[0-9]{2}-[0-9]{2}-[a-z0-9-]+\.md)$")

LEDGER_FIELDS = (
    "entry",
    "triage_status",
    "response_status",
    "response_ids",
    "last_triaged",
    "status_updated_at",
    "status_basis",
    "next_action",
)
REGISTRY_FIELDS = (
    "response_id",
    "owner",
    "state",
    "implemented_at",
    "evidence",
    "next_evaluation",
)

TRIAGE_STATUSES = {"untriaged", "triaged"}
RESPONSE_STATUSES = {
    "unknown",
    "legacy_unrecorded",
    "none",
    "proposed",
    "implemented_unverified",
    "effective_provisional",
    "recurred",
    "ineffective",
    "manual_only",
    "repo_specific_pending",
}
RESPONSE_STATES = RESPONSE_STATUSES - {"unknown", "legacy_unrecorded"}


def _today() -> str:
    return date.today().isoformat()


def resolve_root(value: str | None) -> Path:
    if value:
        root = Path(value).expanduser()
    else:
        env_root = os.environ.get("ORIGIN_TROUBLE_LOG_ROOT", "").strip()
        if env_root:
            root = Path(env_root).expanduser()
        else:
            pointer = Path.home() / ".config/origin-trouble-log/root"
            try:
                root = Path(pointer.read_text(encoding="utf-8").strip()).expanduser()
            except OSError as exc:
                raise SystemExit(f"保管ルート未設定: {exc}") from exc
    if not root.is_dir():
        raise SystemExit(f"保管ルートがディレクトリではありません: {root}")
    return root


def entry_names(root: Path) -> set[str]:
    entries = root / "entries"
    result: set[str] = set()
    if not entries.is_dir():
        return result
    for path in entries.glob("*/*.md"):
        if path.is_file() and ENTRY_NAME_RE.fullmatch(path.name):
            result.add(path.name)
    return result


def target_names(report: Path) -> set[str]:
    """Extract only literal list items in the target-entry section."""

    lines = report.read_text(encoding="utf-8").splitlines()
    start = None
    for index, line in enumerate(lines):
        if line.startswith("## 対象 entry 一覧"):
            start = index + 1
            break
    if start is None:
        return set()

    result: set[str] = set()
    in_fence = False
    for line in lines[start:]:
        if line.startswith("## "):
            break
        if line.startswith("```"):
            in_fence = not in_fence
            continue
        if in_fence:
            continue
        match = TARGET_RE.fullmatch(line.strip())
        if match:
            result.add(match.group(1))
    return result


def report_targets(root: Path) -> tuple[set[str], dict[str, str]]:
    """Return the union and the latest report date for each target entry."""

    union: set[str] = set()
    latest: dict[str, str] = {}
    reports = sorted(
        (root / "triage").glob("[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9].md")
    )
    for report in reports:
        report_date = report.stem
        if not DATE_RE.fullmatch(report_date):
            continue
        for name in target_names(report):
            union.add(name)
            if report_date > latest.get(name, ""):
                latest[name] = report_date
    return union, latest


def _blank_row(
    entry: str,
    *,
    triaged: bool,
    last_triaged: str = "",
    status_updated_at: str | None = None,
) -> dict[str, str]:
    return {
        "entry": entry,
        "triage_status": "triaged" if triaged else "untriaged",
        "response_status": "legacy_unrecorded" if triaged else "unknown",
        "response_ids": "",
        "last_triaged": last_triaged if triaged else "",
        "status_updated_at": status_updated_at or _today(),
        "status_basis": (
            "過去レポートの対象欄に列挙済みだが、当時はstatus欄が無かった"
            if triaged
            else "対象欄に未列挙。次回トリアージ対象"
        ),
        "next_action": (
            "過去文を参照して対策の実装・有効性を評価する。再実装しない"
            if triaged
            else "次回トリアージで読み、既存responseとの再発か新規形かを判定する"
        ),
    }


def read_ledger(path: Path) -> dict[str, dict[str, str]]:
    if not path.exists():
        return {}
    with path.open(encoding="utf-8", newline="") as handle:
        reader = csv.DictReader(handle, delimiter="\t")
        fieldnames = set(reader.fieldnames or ())
        rows = [row for row in reader if row.get("entry")]
    if not rows:
        return {}
    missing = set(LEDGER_FIELDS) - fieldnames
    # status_updated_at was added after the first backfill.  Derive it once
    # from the existing triage date so the ledger can be migrated atomically.
    if missing == {"status_updated_at"}:
        for row in rows:
            row["status_updated_at"] = row.get("last_triaged", "") or _today()
        missing = set()
    if missing:
        raise SystemExit(f"status ledger の列が不足しています: {', '.join(sorted(missing))}")
    result: dict[str, dict[str, str]] = {}
    for row in rows:
        normalized = {field: row.get(field, "") for field in LEDGER_FIELDS}
        if normalized["entry"] in result:
            raise SystemExit(f"status ledger に重複entryがあります: {normalized['entry']}")
        result[normalized["entry"]] = normalized
    return result


def write_ledger(path: Path, rows: dict[str, dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temp_name = tempfile.mkstemp(prefix=f".{path.name}.", suffix=".tmp", dir=path.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="") as handle:
            writer = csv.DictWriter(
                handle, fieldnames=LEDGER_FIELDS, delimiter="\t", lineterminator="\n"
            )
            writer.writeheader()
            for entry in sorted(rows):
                row = rows[entry]
                if any("\t" in row.get(field, "") or "\n" in row.get(field, "") for field in LEDGER_FIELDS):
                    raise SystemExit(f"status ledger の値にtab/newlineがあります: {entry}")
                writer.writerow({field: row.get(field, "") for field in LEDGER_FIELDS})
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temp_name, path)
    finally:
        try:
            os.unlink(temp_name)
        except FileNotFoundError:
            pass


def read_registry(path: Path) -> dict[str, dict[str, str]]:
    if not path.exists():
        return {}
    with path.open(encoding="utf-8", newline="") as handle:
        rows = [row for row in csv.DictReader(handle, delimiter="\t") if row.get("response_id")]
    if not rows:
        return {}
    missing = set(REGISTRY_FIELDS) - set(rows[0])
    if missing:
        raise SystemExit(f"response registry の列が不足しています: {', '.join(sorted(missing))}")
    result: dict[str, dict[str, str]] = {}
    for row in rows:
        normalized = {field: row.get(field, "") for field in REGISTRY_FIELDS}
        response_id = normalized["response_id"]
        if response_id in result:
            raise SystemExit(f"response registry に重複response_idがあります: {response_id}")
        result[response_id] = normalized
    return result


def sync(root: Path) -> int:
    ledger_path = root / "triage/status.tsv"
    rows = read_ledger(ledger_path)
    reported, latest = report_targets(root)
    current = entry_names(root)
    added = 0
    for entry in sorted(current):
        if entry in rows:
            continue
        is_triaged = entry in reported
        rows[entry] = _blank_row(entry, triaged=is_triaged, last_triaged=latest.get(entry, ""))
        added += 1
    write_ledger(ledger_path, rows)
    print(f"synced={added} ledger_rows={len(rows)} entries={len(current)}")
    return 0


def validate(root: Path) -> int:
    ledger_path = root / "triage/status.tsv"
    rows = read_ledger(ledger_path)
    current = entry_names(root)
    registry = read_registry(root / "triage/responses.tsv")
    missing = sorted(current - set(rows))
    stale = sorted(set(rows) - current)
    errors: list[str] = []
    if missing:
        errors.append(f"missing ledger rows: {len(missing)}")
    if stale:
        errors.append(f"ledger rows without entry files: {len(stale)}")
    for entry, row in rows.items():
        if not ENTRY_NAME_RE.fullmatch(entry):
            errors.append(f"invalid entry name: {entry}")
        if row["triage_status"] not in TRIAGE_STATUSES:
            errors.append(f"invalid triage_status for {entry}: {row['triage_status']}")
        if row["response_status"] not in RESPONSE_STATUSES:
            errors.append(f"invalid response_status for {entry}: {row['response_status']}")
        if row["triage_status"] == "triaged" and not DATE_RE.fullmatch(row["last_triaged"]):
            errors.append(f"triaged row has invalid last_triaged: {entry}")
        if row["triage_status"] == "untriaged" and row["last_triaged"]:
            errors.append(f"untriaged row has last_triaged: {entry}")
        if not DATE_RE.fullmatch(row["status_updated_at"]):
            errors.append(f"row has invalid status_updated_at: {entry}")
        if "/Users/" in "\t".join(row.values()) or "/home/" in "\t".join(row.values()):
            errors.append(f"user-named absolute path in ledger: {entry}")
        for response_id in filter(None, (value.strip() for value in row["response_ids"].split(";"))):
            if not re.fullmatch(r"[a-z0-9-]+", response_id):
                errors.append(f"invalid response_id for {entry}: {response_id}")
            elif response_id not in registry:
                errors.append(f"unknown response_id for {entry}: {response_id}")
    for response_id, response in registry.items():
        if not re.fullmatch(r"[a-z0-9-]+", response_id):
            errors.append(f"invalid response registry id: {response_id}")
        if response["state"] not in RESPONSE_STATES:
            errors.append(f"invalid response state: {response_id}: {response['state']}")
        if response["implemented_at"] and not DATE_RE.fullmatch(response["implemented_at"]):
            errors.append(f"invalid response implemented_at: {response_id}")
        if "/Users/" in "\t".join(response.values()) or "/home/" in "\t".join(response.values()):
            errors.append(f"user-named absolute path in response registry: {response_id}")
    if errors:
        for error in errors:
            print(f"FAIL {error}")
        return 1
    print(f"OK: status ledger valid ({len(rows)} rows)")
    return 0


def summary(root: Path) -> int:
    rows = read_ledger(root / "triage/status.tsv")
    registry = read_registry(root / "triage/responses.tsv")
    counts = Counter(row["triage_status"] for row in rows.values())
    response_counts = Counter(row["response_status"] for row in rows.values())
    pending_evaluation = sum(
        response_counts[state] for state in ("legacy_unrecorded", "implemented_unverified")
    )
    registry_counts = Counter(row["state"] for row in registry.values())
    print(f"ledger_rows={len(rows)}")
    print("triage_status=" + ",".join(f"{key}:{counts[key]}" for key in sorted(counts)))
    print("response_status=" + ",".join(f"{key}:{response_counts[key]}" for key in sorted(response_counts)))
    print(f"evaluation_pending_entries={pending_evaluation}")
    print(f"response_registry={len(registry)}")
    if registry_counts:
        print("response_state=" + ",".join(f"{key}:{registry_counts[key]}" for key in sorted(registry_counts)))
    return 0


def update(root: Path, args: argparse.Namespace) -> int:
    ledger_path = root / "triage/status.tsv"
    rows = read_ledger(ledger_path)
    if args.entry not in rows:
        raise SystemExit(f"entry が status ledger にありません。先に sync してください: {args.entry}")
    row = rows[args.entry]
    for field, value in (
        ("triage_status", args.triage_status),
        ("response_status", args.response_status),
        ("response_ids", args.response_ids),
        ("last_triaged", args.last_triaged),
        ("status_updated_at", args.status_updated_at),
        ("status_basis", args.status_basis),
        ("next_action", args.next_action),
    ):
        if value is not None:
            row[field] = value
    if args.status_updated_at is None:
        row["status_updated_at"] = _today()
    write_ledger(ledger_path, rows)
    return validate(root)


def parser() -> argparse.ArgumentParser:
    root_help = "status ledger の保管ルート。省略時は skill の解決規則を使う"
    command = argparse.ArgumentParser(description=__doc__)
    command.add_argument("--root", help=root_help)
    sub = command.add_subparsers(dest="command", required=True)
    sub.add_parser("sync", help="entry と過去レポートから不足行を追加する")
    sub.add_parser("validate", help="entry と ledger の対応・値を検査する")
    sub.add_parser("summary", help="件数だけを表示する")
    update_parser = sub.add_parser("update", help="1 entry の状態を更新する")
    update_parser.add_argument("--entry", required=True)
    update_parser.add_argument("--triage-status", choices=sorted(TRIAGE_STATUSES))
    update_parser.add_argument("--response-status", choices=sorted(RESPONSE_STATUSES))
    update_parser.add_argument("--response-ids")
    update_parser.add_argument("--last-triaged")
    update_parser.add_argument("--status-updated-at")
    update_parser.add_argument("--status-basis")
    update_parser.add_argument("--next-action")
    return command


def main(argv: list[str] | None = None) -> int:
    args = parser().parse_args(argv)
    root = resolve_root(args.root)
    if args.command == "sync":
        return sync(root)
    if args.command == "validate":
        return validate(root)
    if args.command == "summary":
        return summary(root)
    if args.command == "update":
        return update(root, args)
    raise AssertionError(args.command)


if __name__ == "__main__":
    sys.exit(main())
