#!/usr/bin/env python3
"""Measure explicit Claude Code and Codex skill invocations safely.

The two clients persist different JSONL formats.  This module normalises only
messages that can be identified as real user input, matches exact names from an
explicit canonical skill root, and returns aggregate data without retaining
prompt text, paths, session identifiers, or credentials.

No files are written unless the command line caller opts into ``--output`` or
``--append``.  The module is intentionally standard-library-only so it can be
run from a skills checkout without installing anything or contacting a service.
"""

from __future__ import annotations

import argparse
import contextlib
import datetime as _datetime
import importlib.util
import io
import json
import os
import re
import sys
import tempfile
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable, Iterator, Sequence


UTC = _datetime.timezone.utc
SKILL_NAME = r"[A-Za-z0-9][A-Za-z0-9._-]*"
_DOLLAR_TOKEN = re.compile(
    rf"(?<![A-Za-z0-9_.-])\$(?P<name>{SKILL_NAME})(?![A-Za-z0-9_./-])"
)
# A slash token is deliberately stricter than a dollar token.  In particular,
# neither side may be another slash, so ``/skill/SKILL.md`` is not an invocation.
_SLASH_TOKEN = re.compile(
    rf"(?<![A-Za-z0-9_./-])/(?P<name>{SKILL_NAME})(?![A-Za-z0-9_./-])"
)
_COMMAND_NAME = re.compile(
    rf"<command-name>\s*/(?P<name>{SKILL_NAME})(?:\s+[^<]*)?</command-name>",
    re.IGNORECASE,
)
_LABEL = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_.-]*$")

CLAUDE_PROMPT_SOURCES = frozenset({"sdk", "typed"})
CODEX_CONTEXT_KINDS = frozenset(
    {
        "agents_md.instructions",
        "environments.environment_context",
        "plugins.recommendations",
        "skills.selected_skill_instructions",
        "goal.internal_context",
        "permissions.instructions",
        "apps.instructions",
        "host_skills.instructions",
        "multi_agent.usage_hint",
        "multi_agent.mode_instructions",
    }
)


@dataclass(frozen=True)
class Source:
    """One account's local transcript root.

    ``root`` is used only while scanning and is intentionally never rendered in
    a report.  ``kind`` is ``claude`` or ``codex``.
    """

    agent: str
    account: str
    root: Path
    kind: str


@dataclass(frozen=True)
class UsageEvent:
    """A single exact skill match in one message.

    Internal join keys are hidden from ``repr`` so accidental debug output cannot
    disclose session or message identifiers.
    """

    agent: str
    account: str
    skill: str
    timestamp: str
    sidechain: bool = False
    _session_key: str = field(default="", repr=False)
    _message_key: str = field(default="", repr=False)


@dataclass
class ScanStats:
    files_scanned: int = 0
    records_scanned: int = 0
    parse_errors: int = 0
    prompt_messages: int = 0
    matched_messages: int = 0
    ambiguous_messages: int = 0
    first_seen: str | None = None
    last_seen: str | None = None

    def observe(self, timestamp: str) -> None:
        if self.first_seen is None or timestamp < self.first_seen:
            self.first_seen = timestamp
        if self.last_seen is None or timestamp > self.last_seen:
            self.last_seen = timestamp


@dataclass
class ScanResult:
    events: list[UsageEvent] = field(default_factory=list)
    stats: ScanStats = field(default_factory=ScanStats)


def _timestamp(value: object) -> _datetime.datetime | None:
    """Parse an ISO-8601 or Unix timestamp as an aware UTC datetime."""

    if isinstance(value, _datetime.datetime):
        parsed = value
        if parsed.tzinfo is None:
            parsed = parsed.replace(tzinfo=UTC)
        return parsed.astimezone(UTC)
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        number = float(value)
        # Millisecond Unix timestamps are common in metadata; seconds are the
        # normal JSONL representation.  Keep the heuristic bounded to avoid
        # interpreting arbitrary large values as a valid date.
        if abs(number) > 100_000_000_000:
            number /= 1000.0
        try:
            return _datetime.datetime.fromtimestamp(number, tz=UTC)
        except (OverflowError, OSError, ValueError):
            return None
    if not isinstance(value, str):
        return None
    text = value.strip()
    if not text:
        return None
    if text.endswith(("Z", "z")):
        text = text[:-1] + "+00:00"
    try:
        parsed = _datetime.datetime.fromisoformat(text)
    except ValueError:
        return None
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=UTC)
    return parsed.astimezone(UTC)


def _bound(value: object, label: str) -> _datetime.datetime:
    parsed = _timestamp(value)
    if parsed is None:
        raise ValueError(f"{label} must be an ISO-8601 or Unix timestamp")
    return parsed


def _render_timestamp(value: _datetime.datetime) -> str:
    return value.astimezone(UTC).isoformat().replace("+00:00", "Z")


def extract_invocations(text: object, canonical_names: Iterable[str]) -> set[str]:
    """Return exact canonical skill names explicitly mentioned in ``text``.

    Dollar and slash forms are accepted.  Slash forms are boundary-checked to
    exclude file paths, while Claude's command XML is handled separately for
    command records whose ``promptSource`` is not present.
    """

    if not isinstance(text, str) or not text:
        return set()
    names = set(canonical_names)
    found: set[str] = set()
    for pattern in (_DOLLAR_TOKEN, _SLASH_TOKEN, _COMMAND_NAME):
        for match in pattern.finditer(text):
            name = match.group("name")
            if name in names:
                found.add(name)
    return found


def _jsonl_paths(source: Path) -> list[Path]:
    if source.is_file():
        return [source]
    if not source.is_dir():
        return []
    return sorted(path for path in source.rglob("*.jsonl") if path.is_file())


def _records(source: Path, stats: ScanStats) -> Iterator[tuple[Path, int, object]]:
    paths = _jsonl_paths(source)
    stats.files_scanned = len(paths)
    for path in paths:
        try:
            handle = path.open(encoding="utf-8", errors="replace")
        except OSError:
            stats.parse_errors += 1
            continue
        with handle:
            for line_number, line in enumerate(handle, start=1):
                if not line.strip():
                    continue
                stats.records_scanned += 1
                try:
                    record = json.loads(line)
                except (TypeError, ValueError):
                    stats.parse_errors += 1
                    continue
                yield path, line_number, record


def _claude_texts(record: dict) -> tuple[list[str], bool]:
    """Return text and whether a Claude user record is a real prompt."""

    message = record.get("message")
    content = message.get("content") if isinstance(message, dict) else None
    prompt_source = record.get("promptSource")

    if isinstance(content, str):
        if prompt_source in CLAUDE_PROMPT_SOURCES:
            return [content], True
        # Slash commands are persisted as promptSource=null and wrapped in an
        # XML command marker.  Arbitrary injected text with $skill is ignored.
        command_text = " ".join(
            match.group(0) for match in _COMMAND_NAME.finditer(content)
        )
        return ([command_text], True) if command_text else ([], False)

    if not isinstance(content, list):
        return [], False

    text_blocks: list[str] = []
    for block in content:
        if not isinstance(block, dict) or block.get("type") != "text":
            continue
        text = block.get("text")
        if isinstance(text, str):
            text_blocks.append(text)
    if prompt_source in CLAUDE_PROMPT_SOURCES and text_blocks:
        return text_blocks, True
    if prompt_source is None:
        command_texts = [
            text
            for text in text_blocks
            if _COMMAND_NAME.search(text)
        ]
        return command_texts, bool(command_texts)
    return [], False


def scan_claude(
    source: str | Path,
    canonical_names: Iterable[str],
    since: object,
    until: object,
    *,
    agent: str = "claude",
    account: str = "unknown",
) -> ScanResult:
    """Scan one Claude JSONL file or transcript directory."""

    names = set(canonical_names)
    start = _bound(since, "since")
    end = _bound(until, "until")
    if end <= start:
        raise ValueError("until must be after since")
    result = ScanResult()
    seen: set[tuple[str, str, str]] = set()
    for path, line_number, record in _records(Path(source).expanduser(), result.stats):
        if not isinstance(record, dict) or record.get("type") != "user":
            continue
        parsed = _timestamp(record.get("timestamp"))
        if parsed is None:
            result.stats.parse_errors += 1
            continue
        if not start <= parsed < end:
            continue
        texts, actual_prompt = _claude_texts(record)
        if not actual_prompt:
            continue
        timestamp = _render_timestamp(parsed)
        result.stats.prompt_messages += 1
        result.stats.observe(timestamp)
        matched: set[str] = set()
        for text in texts:
            matched.update(extract_invocations(text, names))
        if not matched:
            continue
        result.stats.matched_messages += 1
        session_key = str(record.get("sessionId") or record.get("session_id") or path.name)
        message_key = str(
            record.get("uuid")
            or record.get("promptId")
            or record.get("prompt_id")
            or line_number
        )
        sidechain = bool(record.get("isSidechain") or record.get("is_sidechain"))
        for skill in sorted(matched):
            key = (session_key, message_key, skill)
            if key in seen:
                continue
            seen.add(key)
            result.events.append(
                UsageEvent(
                    agent=agent,
                    account=account,
                    skill=skill,
                    timestamp=timestamp,
                    sidechain=sidechain,
                    _session_key=session_key,
                    _message_key=message_key,
                )
            )
    return result


def _codex_metadata_state(payload: dict) -> str:
    metadata = payload.get("internal_chat_message_metadata_passthrough")
    if not isinstance(metadata, dict) or "content_item_kinds" not in metadata:
        return "ambiguous"
    kinds = metadata.get("content_item_kinds")
    if (
        not isinstance(kinds, list)
        or not kinds
        or not all(isinstance(kind, str) for kind in kinds)
    ):
        return "ambiguous"
    kind_set = set(kinds)
    if kind_set & CODEX_CONTEXT_KINDS:
        return "injected"
    # Repeated user.text entries occur when a turn contains multiple blocks.
    # Unknown mixtures are fail-closed rather than guessed as user input.
    if kind_set == {"user.text"}:
        return "actual"
    return "ambiguous"


def _codex_text(payload: dict) -> str:
    content = payload.get("content")
    if isinstance(content, str):
        return content
    if isinstance(content, list):
        pieces: list[str] = []
        for block in content:
            if not isinstance(block, dict):
                continue
            if block.get("type") not in ("input_text", "text"):
                continue
            text = block.get("text")
            if isinstance(text, str):
                pieces.append(text)
        return "\n".join(pieces)
    message = payload.get("message")
    return message if isinstance(message, str) else ""


def scan_codex(
    source: str | Path,
    canonical_names: Iterable[str],
    since: object,
    until: object,
    *,
    agent: str = "codex",
    account: str = "unknown",
) -> ScanResult:
    """Scan one Codex rollout JSONL file or sessions directory."""

    names = set(canonical_names)
    start = _bound(since, "since")
    end = _bound(until, "until")
    if end <= start:
        raise ValueError("until must be after since")
    result = ScanResult()
    seen: set[tuple[str, str, str]] = set()
    for path, line_number, record in _records(Path(source).expanduser(), result.stats):
        if not isinstance(record, dict):
            continue
        payload = record.get("payload")
        if not isinstance(payload, dict):
            # A small number of early exports put the user_message payload at
            # the top level.  Keep this compatibility path explicit.
            payload = record if record.get("type") == "user_message" else None
        if not isinstance(payload, dict):
            continue
        payload_type = payload.get("type")
        if payload_type not in ("message", "user_message"):
            continue
        if payload_type == "message" and payload.get("role") != "user":
            continue
        parsed = _timestamp(record.get("timestamp")) or _timestamp(payload.get("timestamp"))
        if parsed is None:
            result.stats.parse_errors += 1
            continue
        if not start <= parsed < end:
            continue

        if payload_type == "message":
            state = _codex_metadata_state(payload)
            if state == "ambiguous":
                result.stats.ambiguous_messages += 1
                continue
            if state != "actual":
                continue
        timestamp = _render_timestamp(parsed)
        text = _codex_text(payload)
        result.stats.prompt_messages += 1
        result.stats.observe(timestamp)
        matched = extract_invocations(text, names)
        if not matched:
            continue
        result.stats.matched_messages += 1
        session_key = str(
            payload.get("session_id")
            or payload.get("sessionId")
            or record.get("session_id")
            or record.get("sessionId")
            or path.stem
        )
        message_key = str(
            payload.get("id")
            or payload.get("message_id")
            or record.get("ordinal")
            or line_number
        )
        sidechain = bool(
            payload.get("is_sidechain")
            or payload.get("isSidechain")
            or record.get("is_sidechain")
            or record.get("isSidechain")
        )
        for skill in sorted(matched):
            key = (session_key, message_key, skill)
            if key in seen:
                continue
            seen.add(key)
            result.events.append(
                UsageEvent(
                    agent=agent,
                    account=account,
                    skill=skill,
                    timestamp=timestamp,
                    sidechain=sidechain,
                    _session_key=session_key,
                    _message_key=message_key,
                )
            )
    return result


def canonical_skill_names(root: str | Path) -> list[str]:
    """Return the names of canonical directories that contain ``SKILL.md``."""

    path = Path(root).expanduser()
    if not path.is_dir():
        raise ValueError(f"canonical root is not a directory: {path}")
    excluded = {".git", "docs", "node_modules"}
    names = sorted(
        child.name
        for child in path.iterdir()
        if child.name not in excluded
        and child.is_dir()
        and (child / "SKILL.md").is_file()
    )
    if not names:
        raise ValueError("canonical root contains no skill directories with SKILL.md")
    return names


def _dependency_edges(root: Path, names: set[str]) -> tuple[set[tuple[str, str, str]], int]:
    """Reuse the canonical reference parser without emitting its path output."""

    parser_path = Path(__file__).with_name("skill_lint_refs.py")
    spec = importlib.util.spec_from_file_location("_usage_skill_lint_refs", parser_path)
    if spec is None or spec.loader is None:
        return set(), 1
    module = importlib.util.module_from_spec(spec)
    try:
        spec.loader.exec_module(module)
    except (ImportError, OSError, SyntaxError):
        return set(), 1

    edges: set[tuple[str, str, str]] = set()
    errors = 0
    for source_name in sorted(names):
        markdown = root / source_name / "SKILL.md"
        output = io.StringIO()
        try:
            with contextlib.redirect_stdout(output):
                module.scan(root, markdown)
        except (OSError, UnicodeError, ValueError, AttributeError):
            errors += 1
            continue
        for line in output.getvalue().splitlines():
            fields = line.split("\t")
            if len(fields) != 3 or fields[0] != "cross":
                continue
            target, resource = fields[1], fields[2]
            if target in names:
                edges.add((source_name, target, resource))
    return edges, errors


def _stats_dict(stats: ScanStats) -> dict[str, object]:
    return {
        "files_scanned": stats.files_scanned,
        "records_scanned": stats.records_scanned,
        "parse_errors": stats.parse_errors,
        "prompt_messages": stats.prompt_messages,
        "matched_messages": stats.matched_messages,
        "ambiguous_messages": stats.ambiguous_messages,
        "first_seen": stats.first_seen,
        "last_seen": stats.last_seen,
    }


def build_report(
    canonical_root: str | Path,
    sources: Sequence[Source],
    since: object,
    until: object,
    *,
    generated_at: str | None = None,
) -> dict[str, object]:
    """Build an aggregate report for all canonical skills and account sources."""

    root = Path(canonical_root).expanduser()
    names = canonical_skill_names(root)
    name_set = set(names)
    start = _bound(since, "since")
    end = _bound(until, "until")
    if end <= start:
        raise ValueError("until must be after since")

    if not sources:
        raise ValueError("at least one Claude or Codex source is required")
    results: list[tuple[Source, ScanResult]] = []
    for source in sources:
        if not _LABEL.fullmatch(source.agent) or not _LABEL.fullmatch(source.account):
            raise ValueError("source agent and account labels must be simple stable labels")
        if source.kind == "claude":
            result = scan_claude(
                source.root,
                name_set,
                start,
                end,
                agent=source.agent,
                account=source.account,
            )
        elif source.kind == "codex":
            result = scan_codex(
                source.root,
                name_set,
                start,
                end,
                agent=source.agent,
                account=source.account,
            )
        else:
            raise ValueError(f"unsupported source kind: {source.kind}")
        results.append((source, result))

    edges, dependency_errors = _dependency_edges(root, name_set)
    incoming: dict[str, int] = defaultdict(int)
    for _source_name, target, _resource in edges:
        incoming[target] += 1

    aggregate: dict[str, dict[str, object]] = {
        name: {
            "skill": name,
            "status": "unknown",
            "total_invocations": 0,
            "accounts_with_use": 0,
            "by_account": {},
            "incoming_dependency_count": incoming.get(name, 0),
        }
        for name in names
    }
    internal_accounts: dict[tuple[str, str, str], dict[str, object]] = {}
    for source, result in results:
        account_key = f"{source.agent}:{source.account}"
        for event in result.events:
            row = aggregate[event.skill]
            row["total_invocations"] = int(row["total_invocations"]) + 1
            by_account = row["by_account"]
            assert isinstance(by_account, dict)
            account_row = by_account.setdefault(
                account_key,
                {
                    "invocations": 0,
                    "sessions": 0,
                    "sidechain_invocations": 0,
                    "first_seen": None,
                    "last_seen": None,
                },
            )
            assert isinstance(account_row, dict)
            account_row["invocations"] = int(account_row["invocations"]) + 1
            if event.sidechain:
                account_row["sidechain_invocations"] = (
                    int(account_row["sidechain_invocations"]) + 1
                )
            first_seen = account_row["first_seen"]
            last_seen = account_row["last_seen"]
            if first_seen is None or event.timestamp < first_seen:
                account_row["first_seen"] = event.timestamp
            if last_seen is None or event.timestamp > last_seen:
                account_row["last_seen"] = event.timestamp
            internal_key = (event.skill, account_key, event._session_key)
            internal_accounts.setdefault(internal_key, account_row)

    for name, row in aggregate.items():
        by_account = row["by_account"]
        assert isinstance(by_account, dict)
        for (skill, account_key, _session), account_row in internal_accounts.items():
            if skill != name:
                continue
            # Session count is assigned after all events so duplicate skill hits
            # in a message cannot inflate it.
            account_row["sessions"] = sum(
                1
                for (s, a, _session_key) in internal_accounts
                if s == name and a == account_key
            )
        total = int(row["total_invocations"])
        if total:
            row["status"] = "used"
            row["accounts_with_use"] = sum(
                1
                for item in by_account.values()
                if isinstance(item, dict) and int(item.get("invocations", 0)) > 0
            )
        elif int(row["incoming_dependency_count"]):
            row["status"] = "dependency-only"

    source_rows: list[dict[str, object]] = []
    for source, result in results:
        source_root = Path(source.root).expanduser()
        source_rows.append(
            {
                "agent": source.agent,
                "account": source.account,
                "kind": source.kind,
                "source_available": source_root.is_file() or source_root.is_dir(),
                **_stats_dict(result.stats),
            }
        )
    status_counts = defaultdict(int)
    for row in aggregate.values():
        status_counts[str(row["status"])] += 1

    generated = (
        _render_timestamp(_bound(generated_at, "generated_at"))
        if generated_at
        else _render_timestamp(_datetime.datetime.now(tz=UTC))
    )
    return {
        "schema_version": 1,
        "generated_at": generated,
        "window": {"since": _render_timestamp(start), "until": _render_timestamp(end)},
        "summary": {
            "canonical_skills": len(names),
            "used": status_counts["used"],
            "dependency_only": status_counts["dependency-only"],
            "unknown": status_counts["unknown"],
            # This collector never makes a deletion decision.  Candidate review
            # remains a human-gated operation outside the report.
            "candidate": 0,
        },
        "sources": source_rows,
        "dependency_scan_errors": dependency_errors,
        "skills": [aggregate[name] for name in names],
        "privacy": {
            "raw_prompts_persisted": False,
            "paths_persisted": False,
            "session_ids_persisted": False,
            "credentials_persisted": False,
        },
    }


def _source_spec(value: str, kind: str) -> Source:
    if "=" not in value:
        raise ValueError(f"{kind} source must be ACCOUNT=PATH")
    account, raw_path = value.split("=", 1)
    if not _LABEL.fullmatch(account):
        raise ValueError(f"invalid account label: {account!r}")
    if not raw_path:
        raise ValueError(f"{kind} source path is empty")
    return Source(kind, account, Path(raw_path).expanduser(), kind)


def _write_snapshot(path: Path, rendered: str, *, append: bool) -> None:
    parent = path.parent
    if not parent.is_dir():
        raise ValueError(f"snapshot parent directory does not exist: {parent}")
    if append:
        if path.is_symlink():
            raise ValueError("snapshot path must not be a symlink")
        if path.exists() and not path.is_file():
            raise ValueError("snapshot path is not a regular file")
        with path.open("a", encoding="utf-8") as handle:
            os.fchmod(handle.fileno(), 0o600)
            handle.write(rendered)
            handle.write("\n")
            handle.flush()
            os.fsync(handle.fileno())
        return
    temporary_name: str | None = None
    try:
        with tempfile.NamedTemporaryFile(
            "w", encoding="utf-8", dir=parent, prefix=f".{path.name}.", delete=False
        ) as handle:
            temporary_name = handle.name
            handle.write(rendered)
            handle.write("\n")
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary_name, path)
        temporary_name = None
    finally:
        if temporary_name:
            try:
                os.unlink(temporary_name)
            except OSError:
                pass


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Measure explicit Claude/Codex skill invocations without retaining prompts."
    )
    parser.add_argument("--canonical", required=True, help="canonical skills root")
    parser.add_argument(
        "--claude", action="append", default=[], metavar="ACCOUNT=PATH", help="Claude JSONL root"
    )
    parser.add_argument(
        "--codex", action="append", default=[], metavar="ACCOUNT=PATH", help="Codex JSONL root"
    )
    parser.add_argument("--since", help="inclusive ISO-8601/Unix window start")
    parser.add_argument("--until", help="exclusive ISO-8601/Unix window end (default: now)")
    parser.add_argument("--days", type=float, help="window length ending at --until (default: 30)")
    destination = parser.add_mutually_exclusive_group()
    destination.add_argument("--output", metavar="PATH", help="atomically write one JSON snapshot")
    destination.add_argument("--append", metavar="PATH", help="append one JSON snapshot per line")
    args = parser.parse_args(argv)

    try:
        until = _bound(args.until, "until") if args.until else _datetime.datetime.now(tz=UTC)
        if args.since and args.days is not None:
            raise ValueError("use either --since or --days, not both")
        if args.since:
            since = _bound(args.since, "since")
        else:
            days = 30.0 if args.days is None else args.days
            if days <= 0:
                raise ValueError("--days must be greater than zero")
            since = until - _datetime.timedelta(days=days)
        sources = [_source_spec(value, "claude") for value in args.claude]
        sources.extend(_source_spec(value, "codex") for value in args.codex)
        report = build_report(args.canonical, sources, since, until)
        rendered = json.dumps(report, ensure_ascii=False, indent=2, sort_keys=True)
        if args.output:
            _write_snapshot(Path(args.output).expanduser(), rendered, append=False)
        elif args.append:
            line = json.dumps(report, ensure_ascii=False, separators=(",", ":"), sort_keys=True)
            _write_snapshot(Path(args.append).expanduser(), line, append=True)
        else:
            print(rendered)
        return 0
    except (OSError, ValueError, TypeError) as error:
        print(f"error: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
