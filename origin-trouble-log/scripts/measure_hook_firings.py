#!/usr/bin/env python3
"""Measure what the warn-only shell guards actually did, from agent transcripts.

Every hook response row in `triage/responses.tsv` asks for the same three numbers
before it may leave `implemented_unverified`: 是正率 / 同型再発数 / 誤警告数.
`status_ledger.py` holds the numbers; this script produces two of them.

Where the data comes from. Claude Code persists a PreToolUse hook's stdout into the
session transcript, so past firings are recoverable without adding any collection:

    ~/.claude/projects/<slug>/<session-uuid>.jsonl
      {"type":"attachment", "timestamp":..., "sessionId":..., "isSidechain":...,
       "attachment":{"type":"hook_success","hookName":"PreToolUse:Bash",
                     "toolUseID":"toolu_...","durationMs":96,"exitCode":0,
                     "stdout":"{\\"systemMessage\\":\\"origin-warn-guards: [H2] ...\\"}"}}

Three facts about that shape were measured by hand before this script was written,
and the tests pin all three:

1. One firing writes three attachment records — `hook_success`,
   `hook_additional_context`, `hook_system_message` — carrying the same text. Only
   `hook_success` may be counted, or every number comes out 3x.
2. The literal string also appears in `user`/`assistant` records when the agent
   greps the guard script. Those are not firings.
3. `attachment.toolUseID` joins a firing to the `tool_use` that triggered it. On a
   hand-checked transcript this resolved for 66 of 66 firings.

What is deliberately NOT done here. 是正 is read out of what the transcript recorded,
never by re-running the hook against the following command: the state-reading guards
(H5a/H5b/H5e/H5f/H8) read live git state, so a replay today would judge a command
against a repository that has since moved.

Privacy. Command text is evidence and stays local. `render_report` output is pasted
into the cloud-synced triage report and therefore carries counts only; commands are
emitted solely by `sample`, into a local file the caller names.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import random
import re
import statistics
import sys
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path

GUARD_MARKER = "origin-warn-guards:"
TAG_RE = re.compile(r"\[(H[0-9a-z]+)\]")
FIRING_RECORD = "hook_success"
ECHO_RECORDS = ("hook_additional_context", "hook_system_message")
WRITE_TOOLS = ("Write", "Edit", "NotebookEdit")
# A session that instruments the guard fires it deliberately. Those firings are
# real but say nothing about behaviour, so they are counted and reported apart.
SELF_REFERENTIAL = ("origin_warn_guards", "measure_hook_firings")

DEFAULT_TRANSCRIPTS = Path.home() / ".claude/projects"
DEFAULT_CACHE = Path.home() / ".local/share/origin-warn-guards/cache"

# Guard tag -> response_id in triage/responses.tsv.
# `None` means the tag fires in production but no response row owns it yet; the
# report prints those as UNREGISTERED so a triage cannot silently skip them.
TAG_TO_RESPONSE: dict[str, str | None] = {
    "H1": "hook-h1-pipeline-evidence",
    "H2": "hook-h2-pipeline-status",
    "H3": None,
    "H4": None,
    "H5a": "hook-h5-git-state",
    "H5b": "hook-h5-git-state",
    "H5c": None,
    "H5d": None,
    "H5e": "hook-h5-git-state",
    "H5f": "hook-h5-git-state",
    "H6": "hook-h6-destructive-chain",
    "H7": "hook-h7-check-before-git",
    "H8": "hook-h8-dirty-restore",
    "H9": "metered-cost-preflight",
    "H10": "hook-h10-vacuous-check",
    "H11": "hook-h11-credential-print",
}


def response_id_for(tag: str) -> str | None:
    return TAG_TO_RESPONSE.get(tag)


@dataclass
class TranscriptScan:
    """One transcript's firings plus the denominators they are measured against."""

    path: str = ""
    firings: list[dict] = field(default_factory=list)
    calls: list[dict] = field(default_factory=list)
    echo_records: int = 0

    @property
    def tool_calls(self) -> Counter:
        return Counter(c["tool_name"] for c in self.calls if not c["sidechain"])

    @property
    def sidechain_tool_calls(self) -> Counter:
        return Counter(c["tool_name"] for c in self.calls if c["sidechain"])

    def as_cache(self) -> dict:
        return {"firings": self.firings, "calls": self.calls, "echo_records": self.echo_records}


def _tags(message: str) -> list[str]:
    return sorted(set(TAG_RE.findall(message)))


def _hook_message(attachment: dict) -> str | None:
    """Return the systemMessage of a guard firing, or None if this is not one."""
    raw = attachment.get("stdout")
    if not isinstance(raw, str) or GUARD_MARKER not in raw:
        return None
    try:
        message = json.loads(raw).get("systemMessage")
    except (ValueError, AttributeError):
        return None
    if not isinstance(message, str) or GUARD_MARKER not in message:
        return None
    return message


def _tool_from_hook_name(hook_name: object) -> str | None:
    if isinstance(hook_name, str) and ":" in hook_name:
        return hook_name.split(":", 1)[1] or None
    return None


def parse_transcript(path: Path) -> TranscriptScan:
    """Extract firings and per-tool call counts from one Claude Code transcript."""
    scan = TranscriptScan(path=str(path))
    records = []
    with Path(path).open(encoding="utf-8", errors="replace") as handle:
        for line in handle:
            line = line.strip()
            if not line:
                continue
            try:
                records.append(json.loads(line))
            except ValueError:
                continue

    # Pass 1 — ordered tool calls, and the error flag of each call's result.
    by_id: dict[str, dict] = {}
    sequence: dict[str, list[str]] = {}
    errors: dict[str, bool] = {}
    for record in records:
        if not isinstance(record, dict):
            continue
        kind = record.get("type")
        content = ((record.get("message") or {}).get("content")
                   if isinstance(record.get("message"), dict) else None)
        if kind == "assistant" and isinstance(content, list):
            for block in content:
                if not isinstance(block, dict) or block.get("type") != "tool_use":
                    continue
                tool_name = block.get("name") or "unknown"
                payload = block.get("input") or {}
                argument = payload.get("command")
                if argument is None:
                    argument = payload.get("file_path")
                call = {
                    "id": block.get("id") or "",
                    "tool_name": tool_name,
                    "argument": argument,
                    "timestamp": record.get("timestamp") or "",
                    "session_id": record.get("sessionId") or "",
                    "sidechain": bool(record.get("isSidechain")),
                }
                scan.calls.append(call)
                if call["id"]:
                    by_id[call["id"]] = call
                    sequence.setdefault(tool_name, []).append(call["id"])
        elif kind == "user" and isinstance(content, list):
            for block in content:
                if isinstance(block, dict) and block.get("type") == "tool_result":
                    target = block.get("tool_use_id")
                    if target:
                        errors[target] = bool(block.get("is_error"))

    # Pass 2 — firings. Only hook_success; the other two records repeat the text.
    raw: list[dict] = []
    echoes: list[dict] = []
    tags_by_call: dict[str, list[str]] = {}
    for record in records:
        if not isinstance(record, dict) or record.get("type") != "attachment":
            continue
        attachment = record.get("attachment")
        if not isinstance(attachment, dict):
            continue
        if attachment.get("type") in ECHO_RECORDS:
            echoes.append(attachment)
            continue
        if attachment.get("type") != FIRING_RECORD:
            continue
        # SessionStart の hook_success も同じ marker を運ぶが、tool 呼び出しではない
        # ので per-tool の率に混ぜてはいけない（2026-09-09 実測: 2,773 件中 11 件）。
        if attachment.get("hookEvent") not in (None, "PreToolUse"):
            continue
        message = _hook_message(attachment)
        if message is None:
            continue
        call_id = attachment.get("toolUseID") or ""
        call = by_id.get(call_id, {})
        tool_name = _tool_from_hook_name(attachment.get("hookName")) or call.get("tool_name") or "unknown"
        tags = _tags(message)
        if call_id:
            tags_by_call[call_id] = tags
        raw.append({
            "session_id": record.get("sessionId") or call.get("session_id") or "",
            "timestamp": record.get("timestamp") or call.get("timestamp") or "",
            "tool_use_id": call_id,
            "tool_name": tool_name,
            "tags": tags,
            "duration_ms": attachment.get("durationMs"),
            "command": call.get("argument"),
            "is_error": errors.get(call_id),
            "sidechain": bool(record.get("isSidechain") or call.get("sidechain")),
            "transcript": str(path),
        })

    fired_ids = {f["tool_use_id"] for f in raw if f["tool_use_id"]}
    scan.echo_records = sum(
        1 for a in echoes
        if a.get("hookEvent") in (None, "PreToolUse") and a.get("toolUseID") in fired_ids
    )

    # Pass 3 — the next call of the same tool, and what fired on it.
    position = {tool: {cid: i for i, cid in enumerate(ids)} for tool, ids in sequence.items()}
    for firing in raw:
        call_id = firing["tool_use_id"]
        ids = sequence.get(firing["tool_name"], [])
        index = position.get(firing["tool_name"], {}).get(call_id)
        next_id = ids[index + 1] if index is not None and index + 1 < len(ids) else None
        firing["next_command"] = by_id[next_id]["argument"] if next_id else None
        firing["next_tags"] = tags_by_call.get(next_id, []) if next_id else []
    scan.firings = raw
    return scan


@dataclass
class Stats:
    since: str | None = None
    until: str | None = None
    tags: dict[str, dict] = field(default_factory=dict)
    firings: int = 0
    sidechain_firings: int = 0
    self_referential: int = 0
    duration_median: int | None = None
    duration_max: int | None = None
    denominators: Counter = field(default_factory=Counter)
    echo_records: int = 0
    transcripts: int = 0


def _is_self_referential(firing: dict) -> bool:
    text = firing.get("command") or ""
    return any(marker in text for marker in SELF_REFERENTIAL)


def _in_window(timestamp: str, since: str | None, until: str | None) -> bool:
    day = (timestamp or "")[:10]
    if not day:
        return False
    if since and day < since:
        return False
    if until and day > until:
        return False
    return True


def aggregate(scans, since: str | None = None, until: str | None = None) -> Stats:
    """Roll firings up per tag, measured against calls of the same tool."""
    stats = Stats(since=since, until=until)
    firings: list[dict] = []
    for scan in scans:
        stats.transcripts += 1
        stats.echo_records += scan.echo_records
        firings.extend(f for f in scan.firings if _in_window(f["timestamp"], since, until))
        for call in scan.calls:
            if call["sidechain"] or not _in_window(call["timestamp"], since, until):
                continue
            stats.denominators[call["tool_name"]] += 1

    stats.firings = len(firings)
    stats.sidechain_firings = sum(1 for f in firings if f["sidechain"])
    stats.self_referential = sum(1 for f in firings if _is_self_referential(f))
    durations = [f["duration_ms"] for f in firings if isinstance(f["duration_ms"], (int, float))]
    if durations:
        stats.duration_median = round(statistics.median(durations))
        stats.duration_max = max(durations)

    per_tag: dict[str, list[dict]] = {}
    for firing in firings:
        for tag in firing["tags"]:
            per_tag.setdefault(tag, []).append(firing)

    for tag, hits in sorted(per_tag.items(), key=lambda kv: (-len(kv[1]), kv[0])):
        tools = {h["tool_name"] for h in hits}
        denominator = sum(stats.denominators.get(tool, 0) for tool in tools)
        sessions = Counter(h["session_id"] for h in hits)
        repeated = sum(1 for count in sessions.values() if count > 1)
        immediate = sum(1 for h in hits if tag not in h["next_tags"])
        stats.tags[tag] = {
            "response_id": response_id_for(tag),
            "firings": len(hits),
            "tools": sorted(tools),
            "denominator": denominator,
            "firing_rate": (len(hits) / denominator) if denominator else None,
            "sessions": len(sessions),
            "sessions_repeated": repeated,
            "correction_rate": (1 - repeated / len(sessions)) if sessions else None,
            "immediate_corrections": immediate,
            "immediate_correction_rate": immediate / len(hits) if hits else None,
            "self_referential": sum(1 for h in hits if _is_self_referential(h)),
        }
    return stats


def _pct(value) -> str:
    return "-" if value is None else f"{value * 100:.1f}%"


def render_report(stats: Stats) -> str:
    """Counts only. This text is pasted into the cloud-synced triage report."""
    window = f"{stats.since or '(開始日なし)'} 〜 {stats.until or '(終了日なし)'}"
    lines = [
        f"観測区間: {window}",
        f"transcript {stats.transcripts} 本 / 発火 {stats.firings} 件"
        f"（うち subagent {stats.sidechain_firings} 件、"
        f"hook 自体を計測していたコマンド由来 {stats.self_referential} 件）",
        f"分母（ツール別呼び出し数・subagent 除く）: "
        + (", ".join(f"{k} {v}" for k, v in sorted(stats.denominators.items())) or "なし"),
        f"hook 実行時間: 中央値 {stats.duration_median if stats.duration_median is not None else '-'}ms"
        f" / 最大 {stats.duration_max if stats.duration_max is not None else '-'}ms",
        "",
        "| タグ | response_id | 発火 | 自己参照 | 分母 | 発火率 | セッション | 再発セッション | 是正率 | 直後是正率 |",
        "| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |",
    ]
    for tag, row in stats.tags.items():
        lines.append(
            f"| {tag} | {row['response_id'] or '**未登録**'} | {row['firings']} | "
            f"{row['self_referential']} | "
            f"{row['denominator'] or '-'} | {_pct(row['firing_rate'])} | {row['sessions']} | "
            f"{row['sessions_repeated']} | {_pct(row['correction_rate'])} | "
            f"{_pct(row['immediate_correction_rate'])} |"
        )
    if not stats.tags:
        lines.append("| （区間内の発火 0 件） | - | 0 | 0 | - | - | 0 | 0 | - | - |")
    lines += [
        "",
        "是正率 = そのタグが鳴ったセッションのうち、同一セッション内で二度と鳴らなかった割合。",
        "直後是正率 = 発火のうち、同じツールの次の呼び出しで同じタグが鳴らなかった割合。",
        "自己参照 = 発火元コマンドが guard 本体や本スクリプトを叩いていたもの。意図的な発火なので"
        "行動の指標には数えない。",
        "誤警告は機械では数えない。`sample` で層別無作為標本を取り、判定件数を別に記録する。",
    ]
    return "\n".join(lines)


def sample_firings(scans, tag: str, n: int, seed: int = 0) -> list[dict]:
    """Stratified random sample for false-alarm adjudication. Carries command text."""
    population = [f for scan in scans for f in scan.firings if tag in f["tags"]]
    population.sort(key=lambda f: (f["timestamp"], f["tool_use_id"]))
    if n >= len(population):
        return population
    return random.Random(seed).sample(population, n)


def _cache_path(cache_dir: Path, path: Path) -> Path:
    digest = hashlib.sha1(str(path).encode("utf-8")).hexdigest()
    return cache_dir / f"{digest}.json"


def load_scans(transcripts: Path, cache_dir: Path | None) -> list[TranscriptScan]:
    """Scan every transcript, reusing cached results for files that have not changed."""
    if cache_dir is not None:
        cache_dir.mkdir(parents=True, exist_ok=True)
        try:
            cache_dir.chmod(0o700)
        except OSError:
            pass
    scans: list[TranscriptScan] = []
    for path in sorted(Path(transcripts).rglob("*.jsonl")):
        try:
            stat = path.stat()
        except OSError:
            continue
        entry = _cache_path(cache_dir, path) if cache_dir is not None else None
        if entry is not None and entry.exists():
            try:
                cached = json.loads(entry.read_text(encoding="utf-8"))
                if cached.get("mtime") == stat.st_mtime and cached.get("size") == stat.st_size:
                    scans.append(TranscriptScan(
                        path=str(path),
                        firings=cached["firings"],
                        calls=cached["calls"],
                        echo_records=cached.get("echo_records", 0),
                    ))
                    continue
            except (ValueError, KeyError, OSError):
                pass
        scan = parse_transcript(path)
        if entry is not None:
            payload = {"mtime": stat.st_mtime, "size": stat.st_size, **scan.as_cache()}
            try:
                entry.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
                entry.chmod(0o600)
            except OSError:
                pass
        scans.append(scan)
    return scans


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        description="Measure warn-only guard firings from agent transcripts.",
    )
    parser.add_argument("--transcripts", type=Path, default=DEFAULT_TRANSCRIPTS,
                        help="Claude Code transcript root (default: ~/.claude/projects)")
    parser.add_argument("--cache", type=Path, default=DEFAULT_CACHE,
                        help="local scan cache directory")
    parser.add_argument("--no-cache", action="store_true", help="do not read or write the cache")
    sub = parser.add_subparsers(dest="command", required=True)

    report = sub.add_parser("report", help="print the counts table for a window")
    report.add_argument("--since")
    report.add_argument("--until")

    sample = sub.add_parser("sample", help="write a random sample for false-alarm adjudication")
    sample.add_argument("--tag", required=True)
    sample.add_argument("--n", type=int, default=20)
    sample.add_argument("--seed", type=int, default=0)
    sample.add_argument("--out", type=Path, required=True)
    sample.add_argument("--since")
    sample.add_argument("--until")

    check = sub.add_parser("validate", help="check the one-firing-three-records invariant")
    check.add_argument("--since")
    check.add_argument("--until")

    args = parser.parse_args(argv)
    cache_dir = None if args.no_cache else args.cache
    scans = load_scans(args.transcripts, cache_dir)

    if args.command == "report":
        print(render_report(aggregate(scans, args.since, args.until)))
        return 0

    if args.command == "sample":
        window = [
            TranscriptScan(path=s.path, calls=s.calls, echo_records=s.echo_records, firings=[
                f for f in s.firings if _in_window(f["timestamp"], args.since, args.until)
            ])
            for s in scans
        ]
        chosen = sample_firings(window, tag=args.tag, n=args.n, seed=args.seed)
        population = sum(1 for s in window for f in s.firings if args.tag in f["tags"])
        payload = {
            "tag": args.tag,
            "seed": args.seed,
            "requested": args.n,
            "population": population,
            "firings": chosen,
        }
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        try:
            args.out.chmod(0o600)
        except OSError:
            pass
        print(f"tag={args.tag} population={population} sampled={len(chosen)} -> {args.out}")
        return 0

    stats = aggregate(scans, args.since, args.until)
    print(f"firings={stats.firings} echo_records={stats.echo_records}")
    if stats.echo_records and stats.echo_records != stats.firings * len(ECHO_RECORDS):
        print(
            "WARN: hook_success と echo レコードの比が 1:2 でない。"
            "重複除去の前提が崩れている可能性がある",
            file=sys.stderr,
        )
        return 1
    print("OK: 1 発火 = hook_success 1 件 + echo 2 件")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
