#!/usr/bin/env python3
"""Build and validate the searchable index for historical trouble entries.

Historical links are analytical evidence, not response-status decisions.  The
optional ``status_updates`` payload is therefore accepted only when a matching
``direct`` link exists in the same payload.
"""

from __future__ import annotations

import argparse
import csv
import importlib.util
import json
import os
import re
import sys
import tempfile
from collections import Counter
from datetime import date
from pathlib import Path
from typing import Iterable


SCRIPT_DIR = Path(__file__).resolve().parent
STATUS_LEDGER_PATH = SCRIPT_DIR / "status_ledger.py"
SPEC = importlib.util.spec_from_file_location("own_trouble_status_ledger", STATUS_LEDGER_PATH)
if SPEC is None or SPEC.loader is None:  # pragma: no cover - installation failure
    raise SystemExit(f"status ledger helper を読み込めません: {STATUS_LEDGER_PATH}")
status_ledger = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(status_ledger)


PATTERN_FIELDS = (
    "pattern_id",
    "title",
    "failure_shape",
    "detection_mode",
    "owner",
    "formalization_state",
    "basis",
)
LINK_FIELDS = (
    "entry",
    "pattern_id",
    "response_id",
    "confidence",
    "basis",
    "analyzed_at",
)
CONFIDENCES = {"direct", "mechanism", "thematic", "unclassified"}
DETECTION_MODES = {"syntax", "runtime_state", "manual_review", "repo_specific"}
FORMALIZATION_STATES = {
    "existing_response",
    "candidate",
    "manual_only",
    "repo_specific",
}
STATUS_TARGETS = {
    "recurred",
    "implemented_unverified",
    "manual_only",
    "repo_specific_pending",
}
PATTERN_ID_RE = re.compile(r"^[a-z0-9]+(?:-[a-z0-9]+)*$")
DATE_RE = re.compile(r"^[0-9]{4}-[0-9]{2}-[0-9]{2}$")
PERSONAL_PATH_RE = re.compile(r"(?:/Users|/home)/[^/\s]+(?:/|$)")
SECTION_RE = re.compile(r"^## ", re.MULTILINE)


class ValidationError(Exception):
    """A payload or persisted index violates the historical-index contract."""


def _fail(messages: Iterable[str]) -> int:
    for message in messages:
        print(message)
    return 1


def _date_valid(value: str) -> bool:
    if not DATE_RE.fullmatch(value):
        return False
    try:
        date.fromisoformat(value)
    except ValueError:
        return False
    return True


def _contains_personal_path(value: str) -> bool:
    return bool(PERSONAL_PATH_RE.search(value))


def _read_tsv(path: Path, fields: tuple[str, ...]) -> list[dict[str, str]]:
    if not path.exists():
        return []
    with path.open(encoding="utf-8", newline="") as handle:
        reader = csv.DictReader(handle, delimiter="\t")
        missing = set(fields) - set(reader.fieldnames or ())
        if missing:
            raise ValidationError(
                f"{path.name} missing columns: {', '.join(sorted(missing))}"
            )
        return [{field: row.get(field, "") for field in fields} for row in reader]


def _tsv_bytes(rows: Iterable[dict[str, str]], fields: tuple[str, ...]) -> bytes:
    import io

    output = io.StringIO(newline="")
    writer = csv.DictWriter(output, fieldnames=fields, delimiter="\t", lineterminator="\n")
    writer.writeheader()
    for row in rows:
        for field in fields:
            value = str(row.get(field, ""))
            if "\t" in value or "\n" in value or "\r" in value:
                raise ValidationError(f"TSV value contains tab/newline: {field}")
        writer.writerow({field: row.get(field, "") for field in fields})
    return output.getvalue().encode("utf-8")


def _frontmatter(text: str) -> tuple[bool, dict[str, str]]:
    """Read useful scalar fields without requiring valid YAML."""

    lines = text.splitlines()
    if not lines or lines[0].strip() != "---":
        return False, {}
    end = next((index for index, line in enumerate(lines[1:], 1) if line.strip() == "---"), None)
    if end is None:
        return False, {}
    result: dict[str, str] = {}
    for line in lines[1:end]:
        match = re.match(r"^([A-Za-z_][A-Za-z0-9_-]*):\s*(.*)$", line)
        if match:
            result[match.group(1)] = match.group(2).strip().strip("'\"")
    return True, result


def _entry_paths(root: Path) -> dict[str, Path]:
    result: dict[str, Path] = {}
    for path in sorted((root / "entries").glob("*/*.md")):
        if status_ledger.ENTRY_NAME_RE.fullmatch(path.name):
            if path.name in result:
                raise ValidationError(f"duplicate entry filename: {path.name}")
            result[path.name] = path
    return result


def _legacy_entries(root: Path) -> set[str]:
    rows = status_ledger.read_ledger(root / "triage/status.tsv")
    return {
        entry
        for entry, row in rows.items()
        if row.get("response_status") == "legacy_unrecorded"
    }


def snapshot_rows(root: Path) -> list[dict[str, object]]:
    paths = _entry_paths(root)
    rows: list[dict[str, object]] = []
    for entry in sorted(_legacy_entries(root)):
        path = paths.get(entry)
        if path is None:
            raise ValidationError(f"legacy entry file missing: {entry}")
        text = path.read_text(encoding="utf-8")
        present, metadata = _frontmatter(text)
        rows.append(
            {
                "entry": entry,
                "date": metadata.get("date", entry[:10]),
                "summary": metadata.get("summary", ""),
                "skills": metadata.get("skills", ""),
                "repo": metadata.get("repo", ""),
                "canon": metadata.get("canon", ""),
                "paths": metadata.get("paths", ""),
                "frontmatter_present": present,
                "seven_sections": len(SECTION_RE.findall(text)) >= 7,
                "text": text,
            }
        )
    return rows


def _validate_pattern_rows(rows: list[dict[str, str]]) -> list[str]:
    errors: list[str] = []
    seen: set[str] = set()
    for number, row in enumerate(rows, 1):
        pattern_id = row.get("pattern_id", "")
        if not PATTERN_ID_RE.fullmatch(pattern_id):
            errors.append(f"invalid pattern id at row {number}: {pattern_id}")
        if pattern_id in seen:
            errors.append(f"duplicate pattern: {pattern_id}")
        seen.add(pattern_id)
        for field in ("title", "failure_shape", "owner", "basis"):
            if not row.get(field, "").strip():
                errors.append(f"empty {field} in pattern: {pattern_id or number}")
        if row.get("detection_mode") not in DETECTION_MODES:
            errors.append(f"invalid detection mode in pattern: {pattern_id or number}")
        if row.get("formalization_state") not in FORMALIZATION_STATES:
            errors.append(f"invalid formalization state in pattern: {pattern_id or number}")
        if any(_contains_personal_path(str(value)) for value in row.values()):
            errors.append(f"absolute personal path in pattern: {pattern_id or number}")
    return errors


def _validate_link_rows(
    rows: list[dict[str, str]],
    *,
    entries: set[str],
    patterns: set[str],
    responses: set[str],
) -> list[str]:
    errors: list[str] = []
    seen: set[tuple[str, str, str]] = set()
    for number, row in enumerate(rows, 1):
        entry = row.get("entry", "")
        pattern_id = row.get("pattern_id", "")
        response_id = row.get("response_id", "")
        key = (entry, pattern_id, response_id)
        if key in seen:
            errors.append(f"duplicate link: {entry} / {pattern_id} / {response_id or '-'}")
        seen.add(key)
        if entry not in entries:
            errors.append(f"unknown entry in link row {number}: {entry}")
        if pattern_id not in patterns:
            errors.append(f"unknown pattern in link row {number}: {pattern_id}")
        if response_id and response_id not in responses:
            errors.append(f"unknown response in link row {number}: {response_id}")
        if row.get("confidence") not in CONFIDENCES:
            errors.append(f"invalid confidence in link row {number}: {row.get('confidence', '')}")
        if not row.get("basis", "").strip():
            errors.append(f"empty basis in link row {number}")
        if not _date_valid(row.get("analyzed_at", "")):
            errors.append(f"invalid analyzed_at in link row {number}: {row.get('analyzed_at', '')}")
        if any(_contains_personal_path(str(value)) for value in row.values()):
            errors.append(f"absolute personal path in link row {number}")
    return errors


def _normalize_rows(raw: object, fields: tuple[str, ...], label: str) -> list[dict[str, str]]:
    if not isinstance(raw, list):
        raise ValidationError(f"{label} must be an array")
    rows: list[dict[str, str]] = []
    for number, item in enumerate(raw, 1):
        if not isinstance(item, dict):
            raise ValidationError(f"{label} row {number} must be an object")
        rows.append({field: str(item.get(field, "")) for field in fields})
    return rows


def _name_set(payload: dict[str, object], key: str) -> tuple[set[str], list[str]]:
    raw = payload.get(key, [])
    if raw is None:
        raw = []
    if not isinstance(raw, list) or not all(isinstance(item, str) for item in raw):
        return set(), [f"{key} must be an array of names"]
    if len(raw) != len(set(raw)):
        return set(raw), [f"duplicate {key} name"]
    return set(raw), []


def _validate_status_updates(
    updates: object,
    *,
    ledger: dict[str, dict[str, str]],
    direct_links: set[tuple[str, str]],
) -> list[str]:
    if updates is None:
        return []
    if not isinstance(updates, list):
        return ["status_updates must be an array"]
    errors: list[str] = []
    seen: set[str] = set()
    for number, update in enumerate(updates, 1):
        if not isinstance(update, dict):
            errors.append(f"status update {number} must be an object")
            continue
        entry = str(update.get("entry", ""))
        response_id = str(update.get("response_id", ""))
        target = str(update.get("response_status", ""))
        basis = str(update.get("status_basis", ""))
        next_action = str(update.get("next_action", ""))
        if entry in seen:
            errors.append(f"duplicate status update: {entry}")
        seen.add(entry)
        if entry not in ledger:
            errors.append(f"unknown entry in status update {number}: {entry}")
        if target not in STATUS_TARGETS:
            errors.append(f"invalid status target in update {number}: {target}")
        if not basis.strip() or not next_action.strip():
            errors.append(f"empty status evidence/action in update {number}")
        if (entry, response_id) not in direct_links:
            errors.append(f"status update lacks matching direct link: {entry} / {response_id or '-'}")
        if target in {"recurred", "implemented_unverified"} and not response_id:
            errors.append(f"status update requires response_id: {entry}")
        if any(_contains_personal_path(value) for value in (basis, next_action)):
            errors.append(f"absolute personal path in status update {number}")
    return errors


def _prepare_status_bytes(
    ledger: dict[str, dict[str, str]], updates: object
) -> bytes | None:
    if not updates:
        return None
    assert isinstance(updates, list)
    changed = {entry: dict(row) for entry, row in ledger.items()}
    for update in updates:
        entry = str(update["entry"])
        row = changed[entry]
        response_id = str(update.get("response_id", ""))
        ids = {item for item in row.get("response_ids", "").split(";") if item}
        if response_id:
            ids.add(response_id)
        row["response_status"] = str(update["response_status"])
        row["response_ids"] = ";".join(sorted(ids))
        row["status_updated_at"] = str(update.get("status_updated_at") or date.today().isoformat())
        row["status_basis"] = str(update["status_basis"])
        row["next_action"] = str(update["next_action"])
    ordered = [changed[entry] for entry in sorted(changed)]
    return _tsv_bytes(ordered, status_ledger.LEDGER_FIELDS)


def _atomic_commit(contents: dict[Path, bytes]) -> None:
    staged: dict[Path, Path] = {}
    old: dict[Path, bytes | None] = {}
    try:
        for target, content in contents.items():
            target.parent.mkdir(parents=True, exist_ok=True)
            old[target] = target.read_bytes() if target.exists() else None
            fd, name = tempfile.mkstemp(prefix=f".{target.name}.", suffix=".tmp", dir=target.parent)
            with os.fdopen(fd, "wb") as handle:
                handle.write(content)
                handle.flush()
                os.fsync(handle.fileno())
            staged[target] = Path(name)
        replaced: list[Path] = []
        try:
            for target, temporary in staged.items():
                os.replace(temporary, target)
                replaced.append(target)
        except OSError:
            for target in reversed(replaced):
                previous = old[target]
                if previous is None:
                    target.unlink(missing_ok=True)
                else:
                    fd, name = tempfile.mkstemp(prefix=f".{target.name}.rollback.", dir=target.parent)
                    with os.fdopen(fd, "wb") as handle:
                        handle.write(previous)
                        handle.flush()
                        os.fsync(handle.fileno())
                    os.replace(name, target)
            raise
    finally:
        for temporary in staged.values():
            temporary.unlink(missing_ok=True)


def _snapshot_bytes(root: Path) -> bytes:
    lines = [
        json.dumps(row, ensure_ascii=False, separators=(",", ":"))
        for row in snapshot_rows(root)
    ]
    return (("\n".join(lines) + "\n") if lines else "").encode("utf-8")


def apply_payload(root: Path, input_path: Path) -> int:
    try:
        payload = json.loads(input_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        return _fail([f"payload read failed: {exc}"])
    if not isinstance(payload, dict):
        return _fail(["payload must be an object"])

    try:
        new_patterns = _normalize_rows(payload.get("patterns"), PATTERN_FIELDS, "patterns")
        new_links = _normalize_rows(payload.get("links"), LINK_FIELDS, "links")
        existing_patterns = _read_tsv(root / "triage/patterns.tsv", PATTERN_FIELDS)
        existing_links = _read_tsv(root / "triage/entry-pattern-links.tsv", LINK_FIELDS)
        ledger = status_ledger.read_ledger(root / "triage/status.tsv")
        registry = status_ledger.read_registry(root / "triage/responses.tsv")
    except (ValidationError, SystemExit) as exc:
        return _fail([str(exc)])

    snapshot_entries = payload.get("snapshot_entries")
    if not isinstance(snapshot_entries, list) or not all(isinstance(item, str) for item in snapshot_entries):
        return _fail(["snapshot_entries must be an array of entry names"])
    if len(snapshot_entries) != len(set(snapshot_entries)):
        return _fail(["duplicate snapshot entry"])
    legacy = _legacy_entries(root)
    snapshot_set = set(snapshot_entries)
    errors: list[str] = []
    if snapshot_set != legacy:
        errors.append(
            f"snapshot mismatch: missing={len(legacy - snapshot_set)} extra={len(snapshot_set - legacy)}"
        )

    replace_entries, replace_entry_errors = _name_set(payload, "replace_entries")
    replace_patterns, replace_pattern_errors = _name_set(payload, "replace_patterns")
    retire_patterns, retire_pattern_errors = _name_set(payload, "retire_patterns")
    errors.extend(replace_entry_errors + replace_pattern_errors + retire_pattern_errors)
    existing_pattern_ids = {row["pattern_id"] for row in existing_patterns}
    if not replace_entries <= snapshot_set:
        errors.append(f"replace_entries outside snapshot: {len(replace_entries - snapshot_set)}")
    if not replace_patterns <= existing_pattern_ids:
        errors.append(f"replace_patterns unknown: {len(replace_patterns - existing_pattern_ids)}")
    if not retire_patterns <= existing_pattern_ids:
        errors.append(f"retire_patterns unknown: {len(retire_patterns - existing_pattern_ids)}")
    if replace_patterns & retire_patterns:
        errors.append("replace_patterns and retire_patterns overlap")

    by_pattern = {row["pattern_id"]: row for row in existing_patterns}
    for row in new_patterns:
        old = by_pattern.get(row["pattern_id"])
        if old is not None and old != row and row["pattern_id"] not in replace_patterns:
            errors.append(f"pattern conflicts with existing definition: {row['pattern_id']}")
        by_pattern[row["pattern_id"]] = row
    merged_links = [row for row in existing_links if row["entry"] not in replace_entries] + new_links
    if retire_patterns & {row["pattern_id"] for row in new_patterns}:
        errors.append("retire_patterns contains a newly supplied pattern")
    for pattern_id in retire_patterns:
        by_pattern.pop(pattern_id, None)
    merged_patterns = [by_pattern[key] for key in sorted(by_pattern)]

    errors.extend(_validate_pattern_rows(merged_patterns))
    errors.extend(
        _validate_link_rows(
            merged_links,
            entries=set(_entry_paths(root)),
            patterns=set(by_pattern),
            responses=set(registry),
        )
    )
    target_entries = replace_entries or snapshot_set
    covered = {row["entry"] for row in new_links}
    missing_coverage = target_entries - covered
    if missing_coverage:
        label = "replacement entries" if replace_entries else "snapshot entries"
        errors.append(f"{label} without links: {len(missing_coverage)}")
    extra_coverage = covered - target_entries
    if extra_coverage:
        label = "replacement target" if replace_entries else "snapshot"
        errors.append(f"links outside {label}: {len(extra_coverage)}")

    direct_links = {
        (row["entry"], row["response_id"])
        for row in new_links
        if row["confidence"] == "direct"
    }
    errors.extend(
        _validate_status_updates(
            payload.get("status_updates"), ledger=ledger, direct_links=direct_links
        )
    )

    report_name = str(payload.get("report_name", ""))
    report_content = str(payload.get("report_content", ""))
    if not re.fullmatch(r"[0-9]{4}-[0-9]{2}-[0-9]{2}-legacy-mining\.md", report_name):
        errors.append(f"invalid report_name: {report_name}")
    if not report_content.strip():
        errors.append("empty report_content")
    if _contains_personal_path(report_content):
        errors.append("absolute personal path in report_content")
    report_path = root / "triage/history" / report_name
    if report_path.exists():
        errors.append(f"history report already exists: {report_name}")
    if errors:
        return _fail(errors)

    try:
        status_bytes = _prepare_status_bytes(ledger, payload.get("status_updates"))
        contents = {
            root / "triage/patterns.tsv": _tsv_bytes(merged_patterns, PATTERN_FIELDS),
            root / "triage/entry-pattern-links.tsv": _tsv_bytes(
                sorted(merged_links, key=lambda row: (row["entry"], row["pattern_id"], row["response_id"])),
                LINK_FIELDS,
            ),
            report_path: report_content.encode("utf-8"),
        }
        if status_bytes is not None:
            contents[root / "triage/status.tsv"] = status_bytes
        _atomic_commit(contents)
    except (OSError, ValidationError) as exc:
        return _fail([f"atomic apply failed: {exc}"])
    print(
        f"applied patterns={len(new_patterns)} links={len(new_links)} "
        f"status_updates={len(payload.get('status_updates') or [])}"
    )
    return 0


def check_batch(root: Path, input_path: Path) -> int:
    """Validate one digest batch without requiring full legacy coverage or writing."""

    try:
        payload = json.loads(input_path.read_text(encoding="utf-8"))
        if not isinstance(payload, dict):
            raise ValidationError("payload must be an object")
        patterns = _normalize_rows(payload.get("patterns"), PATTERN_FIELDS, "patterns")
        links = _normalize_rows(payload.get("links"), LINK_FIELDS, "links")
        registry = status_ledger.read_registry(root / "triage/responses.tsv")
        entries = set(_entry_paths(root))
    except (OSError, json.JSONDecodeError, ValidationError, SystemExit) as exc:
        return _fail([str(exc)])
    snapshot_entries = payload.get("snapshot_entries")
    if not isinstance(snapshot_entries, list) or not all(isinstance(item, str) for item in snapshot_entries):
        return _fail(["snapshot_entries must be an array of entry names"])
    snapshot_set = set(snapshot_entries)
    errors: list[str] = []
    if len(snapshot_entries) != len(snapshot_set):
        errors.append("duplicate snapshot entry")
    errors.extend(_validate_pattern_rows(patterns))
    errors.extend(
        _validate_link_rows(
            links,
            entries=entries,
            patterns={row["pattern_id"] for row in patterns},
            responses=set(registry),
        )
    )
    covered = {row["entry"] for row in links}
    if snapshot_set - covered:
        errors.append(f"snapshot entries without links: {len(snapshot_set - covered)}")
    if covered - snapshot_set:
        errors.append(f"links outside snapshot: {len(covered - snapshot_set)}")
    if errors:
        return _fail(errors)
    print(f"checked entries={len(snapshot_set)} patterns={len(patterns)} links={len(links)}")
    return 0


def validate_index(root: Path) -> int:
    try:
        patterns = _read_tsv(root / "triage/patterns.tsv", PATTERN_FIELDS)
        links = _read_tsv(root / "triage/entry-pattern-links.tsv", LINK_FIELDS)
        registry = status_ledger.read_registry(root / "triage/responses.tsv")
        entries = set(_entry_paths(root))
    except (ValidationError, SystemExit) as exc:
        return _fail([str(exc)])
    errors = _validate_pattern_rows(patterns)
    errors.extend(
        _validate_link_rows(
            links,
            entries=entries,
            patterns={row["pattern_id"] for row in patterns},
            responses=set(registry),
        )
    )
    legacy = _legacy_entries(root)
    covered = {row["entry"] for row in links}
    missing = legacy - covered
    if missing:
        errors.append(f"legacy entries without links: {len(missing)}")
    if errors:
        return _fail(errors)
    print(
        f"valid patterns={len(patterns)} links={len(links)} "
        f"indexed_entries={len(covered)} missing_legacy_entries=0"
    )
    return 0


def print_summary(root: Path) -> int:
    try:
        patterns = _read_tsv(root / "triage/patterns.tsv", PATTERN_FIELDS)
        links = _read_tsv(root / "triage/entry-pattern-links.tsv", LINK_FIELDS)
    except ValidationError as exc:
        return _fail([str(exc)])
    indexed = {row["entry"] for row in links}
    legacy = _legacy_entries(root)
    counts = Counter(row["confidence"] for row in links)
    print(f"patterns={len(patterns)}")
    print(f"links={len(links)}")
    print(f"indexed_entries={len(indexed)}")
    print(f"legacy_entries={len(legacy)}")
    print(f"missing_legacy_entries={len(legacy - indexed)}")
    for confidence in ("direct", "mechanism", "thematic", "unclassified"):
        print(f"confidence_{confidence}={counts[confidence]}")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", help="trouble-log storage root")
    subparsers = parser.add_subparsers(dest="command", required=True)
    snapshot_parser = subparsers.add_parser("snapshot")
    snapshot_parser.add_argument("--output", type=Path)
    subparsers.add_parser("validate")
    subparsers.add_parser("summary")
    apply_parser = subparsers.add_parser("apply")
    apply_parser.add_argument("--input", required=True, type=Path)
    check_parser = subparsers.add_parser("check")
    check_parser.add_argument("--input", required=True, type=Path)
    args = parser.parse_args()
    root = status_ledger.resolve_root(args.root)
    try:
        if args.command == "snapshot":
            content = _snapshot_bytes(root)
            if args.output:
                _atomic_commit({args.output: content})
                print(f"snapshot_entries={len(content.splitlines())} output={args.output}")
            else:
                sys.stdout.buffer.write(content)
            return 0
        if args.command == "apply":
            return apply_payload(root, args.input)
        if args.command == "check":
            return check_batch(root, args.input)
        if args.command == "validate":
            return validate_index(root)
        if args.command == "summary":
            return print_summary(root)
    except (ValidationError, SystemExit) as exc:
        return _fail([str(exc)])
    return 2


if __name__ == "__main__":
    sys.exit(main())
