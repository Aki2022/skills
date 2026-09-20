#!/usr/bin/env python3
"""Keep docs/ small enough to be read and honest enough to be trusted.

Two kinds of output, never mixed up:

* ``--fix`` applies **mechanical** repairs whose correctness needs no judgment —
  archive what says it is complete, move narrative out of the routing index,
  normalize status spellings, add missing front matter from git dates.
* ``--report`` writes **candidates** that need a human or agent decision to
  ``docs/log/hygiene-YYYYMMDD.md`` — stale issues, dead references, history
  mixed into guides, non-canonical directories, baseline debt.

Without ``--fix`` the fix section is a dry run: it counts what would change and
touches nothing. Every count is printed, zero included, so "0" is a measured
result and not an unrun check.

No LLM is involved: the tool is deterministic, free, and safe to run across
every governed repository in one pass.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
from datetime import date
from pathlib import Path
from typing import Optional

from validate_repo_docs import (
    INDEX_MAX_BYTES,
    INDEX_MAX_LINE_CHARS,
    ISSUE_STATUSES,
    as_text,
    parse_front_matter,
)

SCRIPTS_DIR = Path(__file__).resolve().parent

# A run of this many *non-blank* lines with no routing link is narrative, not index.
INDEX_PROSE_RUN = 3
# Cells of an over-long table row are cut to this many characters (the full row
# is preserved in the log); the row itself must stay a row.
INDEX_CELL_MAX = 160
# After a bullet's first link, this much description is kept in the index.
INDEX_DESC_KEEP = 200

STALE_DAYS = 30
ABANDONED_DAYS = 60
HISTORY_HEADINGS_MAX = 5
HISTORY_FILE_MAX_BYTES = 60 * 1024

CANONICAL_DOC_DIRS = {"adrs", "specs", "issues", "workstreams", "guides", "log"}
# Directories that hold generated or non-document material and are never edited.
SKIP_DIRS = {"archive", "log", "html", "template", "templates"}

STATUS_ALIASES = {
    "open": "active",
    "in-progress": "in_progress",
    "in progress": "in_progress",
    "inprogress": "in_progress",
    "resolved": "complete",
    "done": "complete",
    "closed": "complete",
    "deferred": "blocked",
    # `archived` on a file that is not under archive/ means someone flipped the
    # status by hand and skipped the move. The move is the fix.
    "archived": "complete",
}

LINK_RE = re.compile(r"\]\(([^)\s]+\.md)(?:#[^)\s]*)?\)")
SELF_LINK_RE = re.compile(r"\[([^\]\s]+\.md)\]\(\1\)")
BARE_PATH_RE = re.compile(r"^[ \t]*[-*+][ \t]+docs/")
HEADING_RE = re.compile(r"^[ \t]{0,3}#{1,6}[ \t]+")
FENCE_RE = re.compile(r"^[ \t]{0,3}(`{3,}|~{3,})")
TABLE_SEP_RE = re.compile(r"^[ \t]*\|?[ \t]*:?-{3,}:?[ \t]*(\|[ \t]*:?-{3,}:?[ \t]*)*\|?[ \t]*$")
DATED_HEADING_RE = re.compile(r"^#{2,4}[ \t].*20\d\d-\d\d-\d\d", re.MULTILINE)
CODE_SPAN_RE = re.compile(r"`([^`\s]+)`")
NPM_RUN_RE = re.compile(r"npm run ([\w:.@/-]+)")
WORKFLOW_RE = re.compile(r"\.github/workflows/([\w.-]+\.ya?ml)")
PLACEHOLDER_CHARS = set("<>{}*$")


# --------------------------------------------------------------------------- git

def git(root: Path, *args: str) -> Optional[str]:
    try:
        proc = subprocess.run(
            ["git", "-C", str(root), *args], capture_output=True, text=True, check=False
        )
    except OSError:
        return None
    if proc.returncode != 0:
        return None
    return proc.stdout


def git_dates(root: Path, path: Path) -> tuple[Optional[str], Optional[str]]:
    """(first_commit_date, last_commit_date) as YYYY-MM-DD, or (None, None)."""
    out = git(root, "log", "--follow", "--format=%cs", "--", str(path.relative_to(root)))
    if not out or not out.strip():
        return None, None
    lines = out.strip().splitlines()
    return lines[-1], lines[0]


def git_branches(root: Path) -> Optional[set[str]]:
    out = git(root, "branch", "-a", "--format=%(refname:short)")
    if out is None:
        return None
    names: set[str] = set()
    for line in out.splitlines():
        line = line.strip()
        if not line or "->" in line:
            continue
        names.add(line)
        for prefix in ("origin/", "remotes/origin/"):
            if line.startswith(prefix):
                names.add(line[len(prefix):])
    return names


# ------------------------------------------------------------------ front matter

def front_matter_span(content: str) -> Optional[tuple[int, int]]:
    """(start, end) offsets of the front-matter block including both fences."""
    if not content.startswith("---\n"):
        return None
    end = content.find("\n---", 4)
    if end == -1:
        return None
    end_line = content.find("\n", end + 1)
    return 0, (len(content) if end_line == -1 else end_line + 1)


def set_scalar(content: str, field: str, value: str) -> str:
    """Set or add a front-matter scalar; the block must already exist."""
    span = front_matter_span(content)
    if span is None:
        raise ValueError("no front matter")
    start, end = span
    block = content[start:end]
    pattern = re.compile(rf"^{re.escape(field)}:[ \t]*.*$", re.MULTILINE)
    if pattern.search(block):
        block = pattern.sub(f"{field}: {value}", block, count=1)
    else:
        # Insert before the closing fence.
        closing = block.rfind("\n---")
        block = block[:closing] + f"\n{field}: {value}" + block[closing:]
    return block + content[end:]


def days_since(value: str, today: date) -> Optional[int]:
    try:
        return (today - date.fromisoformat(value[:10])).days
    except (TypeError, ValueError):
        return None


def rel(root: Path, path: Path) -> str:
    return path.relative_to(root).as_posix()


def docs_files(root: Path, *subdirs: str) -> list[Path]:
    """Active .md files directly under docs/<subdir> (no archive/, no templates)."""
    out: list[Path] = []
    for sub in subdirs:
        directory = root / "docs" / sub
        if directory.is_dir():
            out.extend(sorted(p for p in directory.glob("*.md") if p.is_file()))
    return out


# ------------------------------------------------------------------ A1 archive

def archive_complete(root: Path, fix: bool, result: dict) -> None:
    archived = []
    blocked = []
    targets: list[tuple[str, Path]] = []
    for path in docs_files(root, "issues"):
        fm = parse_front_matter(path) or {}
        if as_text(fm.get("status", "")).strip() == "complete":
            targets.append(("archive_issue.py", path))
    for path in docs_files(root, "workstreams"):
        fm = parse_front_matter(path) or {}
        if as_text(fm.get("status", "")).strip() in ("complete", "archived"):
            targets.append(("archive_workstream.py", path))

    for script, path in targets:
        if not fix:
            archived.append({"file": rel(root, path), "dry_run": True})
            continue
        proc = subprocess.run(
            [sys.executable, str(SCRIPTS_DIR / script), str(path), "--repo", str(root)],
            capture_output=True, text=True, check=False,
        )
        if proc.returncode == 0 and not path.exists():
            archived.append({"file": rel(root, path)})
        else:
            reason = (proc.stderr or proc.stdout).strip().splitlines()
            blocked.append({"file": rel(root, path), "reason": " / ".join(reason[:4])})
    result["A1_archived"] = {"count": len(archived), "items": archived}
    result["A1_archive_blocked"] = {"count": len(blocked), "items": blocked}


# ------------------------------------------------------------ A2 index narrative

def is_routing_line(line: str) -> bool:
    return bool(LINK_RE.search(line)) or bool(BARE_PATH_RE.match(line))


def shorten_table_row(line: str) -> str:
    cells = line.split("|")
    out = []
    for cell in cells:
        stripped = cell.strip()
        if len(stripped) > INDEX_CELL_MAX:
            cut = stripped[:INDEX_CELL_MAX]
            # Never cut inside a markdown link.
            last_open = cut.rfind("[")
            if last_open != -1 and cut.find("](", last_open) == -1:
                cut = cut[:last_open]
            stripped = cut.rstrip() + "…"
        out.append(f" {stripped} " if stripped else cell)
    return "|".join(out).rstrip()


def shorten_bullet(line: str, log_name: str) -> Optional[str]:
    match = LINK_RE.search(line)
    if not match:
        return None
    head = line[: match.end()]
    tail = line[match.end():]
    if len(tail) > INDEX_DESC_KEEP:
        cut = tail[:INDEX_DESC_KEEP]
        space = cut.rfind(" ")
        if space > INDEX_DESC_KEEP // 2:
            cut = cut[:space]
        tail = cut.rstrip() + f" …（全文: {log_name}）"
    short = head + tail
    if len(short) > INDEX_MAX_LINE_CHARS:
        short = short[: INDEX_MAX_LINE_CHARS - 1] + "…"
    return short


def relink_for_log(line: str) -> str:
    """Relative links written for docs/ must still resolve from docs/log/."""
    # `../x` is relative too: from docs/ it meant <repo>/x, so from docs/log/ it is `../../x`.
    return re.sub(r"\]\((?![a-zA-Z][a-zA-Z0-9+.-]*:|#|/)([^)\s]+)\)", r"](../\1)", line)


def split_index_narrative(content: str, log_name: str, today: date) -> tuple[str, list[tuple[str, list[str]]], int]:
    """Return (new_index, [(section_heading, moved_lines)], moved_count)."""
    lines = content.splitlines()
    span = front_matter_span(content)
    fm_lines = 0
    if span is not None:
        fm_lines = content[: span[1]].count("\n")

    moved_sections: list[tuple[str, list[str]]] = []
    out: list[str] = lines[:fm_lines]
    moved_total = 0
    # Front matter scalars are subject to the same line limit (a 12.9 KB
    # `current_focus:` was measured on 2026-09-19); the body pass below never
    # looks at them, so handle them here and keep the key valid YAML.
    for i in range(fm_lines):
        fm_line = out[i]
        key_match = re.match(r"^([A-Za-z_][\w-]*):[ \t]+(.+)$", fm_line)
        if not key_match or len(fm_line) <= INDEX_MAX_LINE_CHARS:
            continue
        key = key_match.group(1)
        out[i] = (
            f"{key}: 経緯は {log_name} へ移動（{today.isoformat()} docs_hygiene）。"
            "現在の状態は各 issue / workstream / guide を読む。"
        )
        moved_sections.append((f"front matter: {key}", [fm_line]))
        moved_total += 1
    section_heading = "(preamble)"
    section_moved: list[str] = []
    section_link_present = False
    section_start_out = len(out)
    in_fence = False
    fence_char = ""
    fence_width = 0
    run: list[str] = []  # candidate prose run (may include blank lines)

    def flush_run(force_keep: bool = False) -> None:
        nonlocal run, moved_total
        if not run:
            return
        nonblank = [l for l in run if l.strip()]
        if not force_keep and len(nonblank) > INDEX_PROSE_RUN:
            section_moved.extend(run)
            moved_total += len(nonblank)
            # keep one blank so neighbouring blocks do not fuse
            if out and out[-1].strip():
                out.append("")
        else:
            out.extend(run)
        run = []

    def close_section() -> None:
        nonlocal section_moved, section_link_present, section_start_out
        flush_run()
        if section_moved:
            moved_sections.append((section_heading, section_moved))
            if not section_link_present:
                link = (
                    f"- 詳細な経緯は [{log_name}]({log_name}) に移動"
                    f"（{today.isoformat()} docs_hygiene）"
                )
                insert_at = section_start_out
                # after the heading line and one blank
                if insert_at < len(out) and HEADING_RE.match(out[insert_at] if insert_at < len(out) else ""):
                    insert_at += 1
                if insert_at < len(out) and out[insert_at].strip() == "":
                    insert_at += 1
                out.insert(insert_at, link)
                out.insert(insert_at + 1, "")
        section_moved = []
        section_link_present = False

    # A wrapped entry (bullet + indented continuation lines) is one logical line:
    # judged, shortened and logged as a unit. Splitting it moved the description
    # and left a bare link, or cut a sentence at the wrap (seen 2026-09-19).
    joined: list[str] = []
    i = fm_lines
    while i < len(lines):
        bare = lines[i].rstrip("\r\n")
        if (
            not in_fence and is_routing_line(bare) and re.match(r"^[ \t]*[-*+][ \t]+", bare)
        ):
            block = [bare]
            j = i + 1
            while (
                j < len(lines)
                and lines[j].strip()
                and re.match(r"^[ \t]+", lines[j])
                and not re.match(r"^[ \t]*[-*+][ \t]+", lines[j])
                and not lines[j].lstrip().startswith("|")
            ):
                block.append(lines[j].rstrip("\r\n"))
                j += 1
            joined.append(" ".join(part.strip() if k else part for k, part in enumerate(block)))
            i = j
            continue
        joined.append(lines[i])
        i += 1

    for line in joined:
        bare = line.rstrip("\r\n")
        if in_fence:
            run_flush_keep = True
            flush_run(force_keep=True)
            out.append(line)
            if re.match(rf"^[ \t]{{0,3}}{re.escape(fence_char)}{{{fence_width},}}[ \t]*$", bare):
                in_fence = False
            continue
        opening = FENCE_RE.match(bare)
        if opening:
            flush_run(force_keep=True)
            in_fence = True
            fence_char = opening.group(1)[0]
            fence_width = len(opening.group(1))
            out.append(line)
            continue
        if HEADING_RE.match(bare):
            close_section()
            section_heading = bare.lstrip("# ").strip() or "(untitled)"
            section_start_out = len(out)
            out.append(line)
            continue
        if f"({log_name})" in bare or "log/index-" in bare:
            section_link_present = True
        if TABLE_SEP_RE.match(bare):
            # Formatter-padded separators (measured: 1,432 chars of dashes) are
            # structure, not content: collapse, do not log, do not count.
            flush_run()
            cells = bare.count("|") - 1 if bare.strip().startswith("|") and bare.strip().endswith("|") else bare.count("|") + 1
            out.append("|" + "|".join([" --- "] * max(cells, 1)) + "|")
            continue
        if len(bare) > INDEX_MAX_LINE_CHARS:
            flush_run()
            section_moved.append(bare)
            moved_total += 1
            if bare.lstrip().startswith("|"):
                out.append(shorten_table_row(bare))
            else:
                short = shorten_bullet(bare, log_name)
                if short is not None:
                    out.append(short)
            continue
        if not bare.strip():
            if run:
                run.append(line)
            else:
                out.append(line)
            continue
        if is_routing_line(bare) or TABLE_SEP_RE.match(bare) or bare.lstrip().startswith("|"):
            flush_run()
            out.append(line)
            continue
        run.append(line)
    close_section()

    new_content = "\n".join(out)
    if content.endswith("\n"):
        new_content += "\n"
    return new_content, moved_sections, moved_total


DESC_CAPS = (200, 120, 80, 40)


def log_suffix(log_name: str) -> str:
    return f" …（全文: {log_name}）"


def shorten_routing_line(line: str, cap: int, log_name: str) -> str:
    """Cut a routing line's description to `cap` chars, keeping every link intact.

    A line the split pass already cut carries the log suffix; it is stripped
    before measuring so the budget pass can shorten it further (before 2026-09-19
    such lines were skipped and the ceiling was reported unmeetable at cap 40
    while every row still had a 200-char tail).
    """
    suffix = log_suffix(log_name)
    if line.endswith(suffix) and not line.lstrip().startswith("|"):
        line = line[: -len(suffix)]
        match = LINK_RE.search(line)
        if match and len(line[match.end():]) <= cap:
            return line + suffix
    if line.lstrip().startswith("|"):
        cells = line.split("|")
        out = []
        for cell in cells:
            stripped = cell.strip()
            if len(stripped) > cap and not LINK_RE.search(stripped):
                stripped = stripped[:cap].rstrip() + "…"
            elif len(stripped) > cap:
                match = LINK_RE.search(stripped)
                head, tail = stripped[: match.end()], stripped[match.end():]
                if len(tail) > cap:
                    tail = tail[:cap].rstrip() + "…"
                stripped = head + tail
            out.append(f" {stripped} " if stripped else cell)
        return "|".join(out).rstrip()
    match = LINK_RE.search(line)
    if not match:
        return line
    head, tail = line[: match.end()], line[match.end():]
    if len(tail) <= cap:
        return line
    cut = tail[:cap]
    space = cut.rfind(" ")
    if space > cap // 2:
        cut = cut[:space]
    return head + cut.rstrip() + log_suffix(log_name)


def fit_index_budget(content: str, log_name: str) -> tuple[str, list[str], Optional[int]]:
    """Progressively shorten routing descriptions until the index fits the ceiling.

    Returns (content, original lines that were shortened, cap used or None when
    the ceiling cannot be met — which means the row count itself is the problem
    and only closing work will fix it).
    """
    if len(content.encode()) <= INDEX_MAX_BYTES:
        return content, [], None
    # `[issues/X.md](issues/X.md)` says the path twice; the id alone routes the
    # same (index_entries matches on the target) and is what the template uses.
    content = SELF_LINK_RE.sub(lambda m: f"[{Path(m.group(1)).stem}]({m.group(1)})", content)
    if len(content.encode()) <= INDEX_MAX_BYTES:
        return content, [], None
    pristine = content.splitlines(keepends=True)
    lines = list(pristine)
    originals: dict[int, str] = {}
    for cap in DESC_CAPS:
        # Always cut from the pristine line: a line already cut at the previous,
        # looser cap carries the log suffix and would otherwise never shrink again.
        for i, line in enumerate(pristine):
            bare = line.rstrip("\n")
            if not is_routing_line(bare) or f"]({log_name})" in bare:
                continue  # the pointer to the log itself is never cut
            short = shorten_routing_line(bare, cap, log_name)
            if short != bare:
                if not bare.endswith(log_suffix(log_name)):
                    originals.setdefault(i, bare)  # split-cut lines are already in the log in full
                lines[i] = short + ("\n" if line.endswith("\n") else "")
        if len("".join(lines).encode()) <= INDEX_MAX_BYTES:
            return "".join(lines), [originals[i] for i in sorted(originals)], cap
    return "".join(lines), [originals[i] for i in sorted(originals)], None


def move_index_narrative(root: Path, fix: bool, today: date, result: dict) -> None:
    index = root / "docs/00_index.md"
    content = index.read_text()
    log_name = f"log/index-{today.strftime('%Y%m')}.md"
    new_content, moved, count = split_index_narrative(content, log_name, today)
    new_content, shortened, cap = fit_index_budget(new_content, log_name)
    if shortened:
        moved.append((f"行の短縮（説明を {cap or DESC_CAPS[-1]} 字に）", shortened))
        count += len(shortened)
    item = {
        "count": count,
        "items": [{"section": h, "lines": len([l for l in ls if l.strip()])} for h, ls in moved],
        "index_bytes_before": len(content.encode()),
        "index_bytes_after": len(new_content.encode()),
        "log": f"docs/{log_name}",
        "description_cap": cap,
        "over_ceiling_after_fix": len(new_content.encode()) > INDEX_MAX_BYTES,
    }
    if item["over_ceiling_after_fix"]:
        item["note"] = (
            "ceiling cannot be met by shortening: too many active rows — close or "
            "archive work (see R1/R2), or record docs/00_index.md in "
            "docs/validator-baseline.txt until then"
        )
    if fix and count:
        new_content = re.sub(r"(^updated_at:[ \t]*)[\d-]+", rf"\g<1>{today.isoformat()}",
                             new_content, count=1, flags=re.MULTILINE)
        log_path = root / "docs" / log_name
        log_path.parent.mkdir(parents=True, exist_ok=True)
        if not log_path.exists():
            log_path.write_text(
                "---\n"
                f"updated_at: {today.isoformat()}\n"
                "kind: hygiene-log\n"
                "---\n\n"
                "# 00_index.md から分離した記述\n\n"
                "docs_hygiene.py が routing index から移した進捗の物語・長い行。"
                "現在の真実ではなく履歴。現在の状態は各 issue / workstream / guide を読む。\n"
            )
        with log_path.open("a") as fh:
            for heading, lines in moved:
                fh.write(f"\n## {heading} — moved {today.isoformat()}\n\n")
                for line in lines:
                    fh.write(relink_for_log(line) + "\n")
        index.write_text(new_content)
    result["A2_index_lines_moved"] = item


# ------------------------------------------------------------ A3 status aliases

def normalize_statuses(root: Path, fix: bool, today: date, result: dict) -> None:
    items = []
    for path in docs_files(root, "issues"):
        fm = parse_front_matter(path)
        if not fm:
            continue
        raw = as_text(fm.get("status", "")).strip()
        key = raw.lower()
        if key in ISSUE_STATUSES or key not in STATUS_ALIASES:
            continue
        new = STATUS_ALIASES[key]
        items.append({"file": rel(root, path), "from": raw, "to": new})
        if fix:
            content = path.read_text()
            content = set_scalar(content, "status", new)
            content = set_scalar(
                content, "hygiene_note",
                f"status normalized from {raw} on {today.isoformat()}",
            )
            path.write_text(content)
    result["A3_status_normalized"] = {"count": len(items), "items": items}


# ---------------------------------------------------------- A4 front matter fill

def fill_front_matter(root: Path, fix: bool, today: date, result: dict) -> None:
    items = []
    docs = root / "docs"
    for path in sorted(docs.glob("**/*.md")):
        parts = path.relative_to(docs).parts
        if path.name == "00_index.md" or any(p in SKIP_DIRS for p in parts[:-1]):
            continue
        content = path.read_text()
        if content.startswith("---"):
            continue
        kind = parts[0] if len(parts) > 1 else ""
        first, last = git_dates(root, path)
        source = "git"
        if not last:
            first = last = today.isoformat()
            source = "today (file not in git history)"
        note = f"front matter added on {today.isoformat()}; dates from {source}"
        if kind == "specs":
            envelope = (
                "---\n"
                f"id: SPEC-{path.stem}\n"
                "status: active\n"
                f"created_at: {first}\n"
                f"updated_at: {last}\n"
                "related_guides: []\n"
                "affected_workstreams: []\n"
                "related_adrs: []\n"
                f"hygiene_note: {note}\n"
                "---\n\n"
            )
        else:
            envelope = (
                "---\n"
                f"updated_at: {last}\n"
                f"hygiene_note: {note}\n"
                "---\n\n"
            )
        items.append({"file": rel(root, path), "kind": kind or "docs", "updated_at": last})
        if fix:
            path.write_text(envelope + content.lstrip("\n"))
    result["A4_front_matter_added"] = {"count": len(items), "items": items}


# ------------------------------------------------------------------- R1..R6

def report_issues(root: Path, today: date, report: dict) -> None:
    branches = git_branches(root)
    stale = []
    untouched = []
    for path in docs_files(root, "issues"):
        fm = parse_front_matter(path) or {}
        status = as_text(fm.get("status", "")).strip().lower()
        if STATUS_ALIASES.get(status, status) == "complete":
            continue
        updated = as_text(fm.get("updated_at", ""))
        age = days_since(updated, today)
        branch = as_text(fm.get("branch", "")).strip()
        if (
            branches is not None and age is not None and age > STALE_DAYS
            and branch and branch not in ("main", "master") and branch not in branches
        ):
            stale.append({"file": rel(root, path), "status": status, "updated_at": updated,
                          "days": age, "branch": branch})
        _first, last = git_dates(root, path)
        last_age = days_since(last, today) if last else age
        if last_age is not None and last_age > ABANDONED_DAYS:
            untouched.append({"file": rel(root, path), "status": status,
                              "last_commit": last or updated, "days": last_age})
    report["R1_stale_branch_gone"] = {
        "count": len(stale), "items": stale,
        "rule": f"not complete, updated_at > {STALE_DAYS}d, recorded branch missing from git branch -a",
        "note": None if branches is not None else "git unavailable — check skipped",
    }
    report["R2_untouched_60d"] = {
        "count": len(untouched), "items": untouched,
        "rule": f"not complete, last commit > {ABANDONED_DAYS}d",
    }


def candidate_paths(text: str) -> set[str]:
    refs: set[str] = set()
    for value in CODE_SPAN_RE.findall(text):
        if "/" not in value or value.startswith(("http", "~", "/", "docs/log", "roles/")):
            continue  # roles/<x>.<y> is an IAM role name, not a path
        if PLACEHOLDER_CHARS & set(value) or "YYYY" in value or value.endswith("/"):
            continue
        if not re.fullmatch(r"[\w.@/-]+\.[A-Za-z0-9]{1,6}", value):
            continue
        refs.add(value[2:] if value.startswith("./") else value)
    return refs


WALK_SKIP = {".git", "node_modules", ".venv", "venv", "__pycache__", ".worktrees", ".claude", "dist", "build"}


def repo_file_suffixes(root: Path) -> set[str]:
    """Every tracked-looking file path, so `src/x.ts` written relative to a package
    directory (the common false positive) is recognised anywhere it actually lives."""
    paths: set[str] = set()
    for base, dirs, files in os.walk(root):
        dirs[:] = [d for d in dirs if d not in WALK_SKIP]
        rel_base = Path(base).relative_to(root).as_posix()
        for name in files:
            paths.add(name if rel_base == "." else f"{rel_base}/{name}")
    return paths


def reference_exists(ref: str, root: Path, doc_dir: Path, suffixes: set[str]) -> bool:
    if (root / ref).exists() or (root / "docs" / ref).exists() or (doc_dir / ref).exists():
        return True
    if ref in suffixes or any(p.endswith("/" + ref) for p in suffixes):
        return True
    # Cross-repository pointers into the shared agent config are common here.
    agents = Path.home() / ".agents"
    return (agents / ref).exists() or (agents / "skills" / ref).exists()


def gitignored(root: Path, refs: list[str]) -> set[str]:
    """Of `refs`, the ones git is told to ignore — absent by design, not stale.

    A guide that says "the bundle lands in `apps/api/.deploy`" or "put your keys in
    `apps/web/.env.local`" is correct precisely because those paths are not committed.
    Reporting them as dead references trains the reader to skim past R3, which is how a
    real stale path gets missed.

    Both `<ref>` and `<ref>/` are asked, because a directory-only pattern
    (`apps/api/.deploy/`) does not match the slashless form when the path is absent:
    git cannot tell a non-existent path is a directory. Returns an empty set outside a
    git repository or when git is unavailable — reporting too much beats hiding a real
    stale path behind a failed subprocess.
    """
    if not refs:
        return set()
    probes = [r for ref in refs for r in (ref, ref + "/")]
    try:
        out = subprocess.run(
            ["git", "check-ignore", "--stdin"], cwd=root, input="\n".join(probes),
            capture_output=True, text=True, check=False,
        )
    except (OSError, ValueError):
        return set()
    # exit 0 = some ignored, 1 = none ignored, 128 = not a repo / git error
    if out.returncode not in (0, 1):
        return set()
    return {line.rstrip("/") for line in out.stdout.splitlines() if line.strip()}


def report_dead_references(root: Path, report: dict) -> None:
    items = []
    suffixes = repo_file_suffixes(root)
    # Monorepos keep scripts in workspace package.json files; a guide saying
    # `npm run build` from apps/api is not dead because the root lacks it.
    scripts: Optional[set[str]] = None
    for pkg in (Path(p) for p in suffixes if p == "package.json" or p.endswith("/package.json")):
        try:
            found = set(json.loads((root / pkg).read_text()).get("scripts", {}).keys())
        except (ValueError, AttributeError, OSError):
            continue
        scripts = (scripts or set()) | found
    for path in docs_files(root, "guides", "specs"):
        text = path.read_text()
        seen: set[str] = set()
        for name in NPM_RUN_RE.findall(text):
            if scripts is not None and name not in scripts and name not in seen:
                seen.add(name)
                items.append({"file": rel(root, path), "ref": f"npm run {name}", "kind": "npm-script"})
        for wf in WORKFLOW_RE.findall(text):
            ref = f".github/workflows/{wf}"
            if not (root / ref).is_file() and ref not in seen:
                seen.add(ref)
                items.append({"file": rel(root, path), "ref": ref, "kind": "workflow"})
        for ref in sorted(candidate_paths(text)):
            if ref in seen or ref.startswith(".github/workflows/"):
                continue
            if reference_exists(ref, root, path.parent, suffixes):
                continue
            seen.add(ref)
            items.append({"file": rel(root, path), "ref": ref, "kind": "path"})
    # Drop the paths git is told to ignore: those are absent on purpose.
    ignored = gitignored(root, sorted({i["ref"] for i in items if i["kind"] == "path"}))
    items = [i for i in items if not (i["kind"] == "path" and i["ref"] in ignored)]
    report["R3_dead_references"] = {
        "count": len(items), "items": items,
        "rule": "npm run <x> not in package.json scripts; .github/workflows/<y> missing; "
                "`path/with.ext` in a code span matching no file in the repo (any directory) "
                "nor under ~/.agents, and not gitignored (absent by design)",
    }


def report_history_in_docs(root: Path, report: dict) -> None:
    items = []
    for path in docs_files(root, "guides", "specs"):
        text = path.read_text()
        dated = len(DATED_HEADING_RE.findall(text))
        size = len(text.encode())
        if dated > HISTORY_HEADINGS_MAX or size > HISTORY_FILE_MAX_BYTES:
            items.append({"file": rel(root, path), "dated_headings": dated, "bytes": size})
    report["R4_history_in_guide_or_spec"] = {
        "count": len(items), "items": items,
        "rule": f"> {HISTORY_HEADINGS_MAX} dated ##/### headings or > {HISTORY_FILE_MAX_BYTES // 1024}KB",
    }


def report_layout(root: Path, report: dict) -> None:
    docs = root / "docs"
    dirs = []
    for child in sorted(docs.iterdir()):
        if child.is_dir() and child.name not in CANONICAL_DOC_DIRS and child.name != "html":
            count = len(list(child.glob("**/*.md")))
            dirs.append({"dir": rel(root, child), "md_files": count})
    names: dict[str, list[str]] = {}
    for path in docs.glob("**/*.md"):
        if path.name.lower() in ("readme.md", "00_index.md"):
            continue
        names.setdefault(path.name, []).append(rel(root, path))
    dups = [{"name": n, "paths": sorted(p)} for n, p in sorted(names.items()) if len(p) > 1]
    report["R5_non_canonical_dirs"] = {
        "count": len(dirs), "items": dirs,
        "rule": "docs/<dir> outside " + ", ".join(sorted(CANONICAL_DOC_DIRS)) + " (html/ is tooling output)",
    }
    report["R5b_duplicate_basenames"] = {"count": len(dups), "items": dups}


def report_baselines(root: Path, report: dict) -> None:
    items = []
    for name in ("validator-baseline.txt", "validator-link-baseline.txt"):
        path = root / "docs" / name
        if path.is_file():
            count = sum(1 for l in path.read_text().splitlines() if l.strip() and not l.startswith("#"))
            items.append({"file": f"docs/{name}", "count": count})
    report["R6_baseline_debt"] = {"count": sum(i["count"] for i in items), "items": items,
                                  "rule": "entries still exempt; the list may only shrink"}


# ----------------------------------------------------------------- rendering

def render_report(result: dict) -> str:
    today = result["today"]
    mode = "fix applied" if result["fix"] else "dry run (nothing written except this report)"
    lines = [
        "---",
        f"updated_at: {today}",
        "kind: hygiene-report",
        "---",
        "",
        f"# docs hygiene report {today}",
        "",
        # No repository path or name: the report lives inside the repository, and
        # vibe-guard's local-info check stopped commits on both the home path and,
        # for one repository, on its own name (2026-09-19/20).
        f"mode: {mode}",
        "",
        "判断が要る候補の一覧。現在の真実ではなく点検結果。0 件も「検査済みで 0」。",
        "",
        "## Fixes (mechanical)",
        "",
        "| check | count |",
        "|---|---|",
    ]
    for key, value in result["fixes"].items():
        lines.append(f"| {key} | {value['count']} |")
    lines += ["", "## Report (needs judgment)", "", "| check | count | rule |", "|---|---|---|"]
    for key, value in result["report"].items():
        lines.append(f"| {key} | {value['count']} | {value.get('rule', '')} |")
    for key, value in result["report"].items():
        if not value["items"]:
            continue
        lines += ["", f"### {key} ({value['count']})", ""]
        for item in value["items"][:60]:
            cells = ", ".join(f"{k}=`{v}`" for k, v in item.items())
            lines.append(f"- {cells}")
        if len(value["items"]) > 60:
            lines.append(f"- … {len(value['items']) - 60} more")
    for key, value in result["fixes"].items():
        if value.get("items"):
            lines += ["", f"### {key} ({value['count']})", ""]
            for item in value["items"][:60]:
                cells = ", ".join(f"{k}=`{v}`" for k, v in item.items())
                lines.append(f"- {cells}")
    lines += [
        "",
        "## Not done",
        "",
        "- guide の内容と code の意味的な食い違いは判定していない（R3 は参照先の存在確認のみ）。",
        "- R1/R2 の issue を閉じるかどうかは人間/agent の判断。",
        "",
    ]
    return "\n".join(lines)


def summary_text(result: dict) -> str:
    out = [f"docs_hygiene: {result['repo']} ({'fix' if result['fix'] else 'dry-run'})"]
    for key, value in result["fixes"].items():
        out.append(f"  {key}: {value['count']}")
    for key, value in result["report"].items():
        out.append(f"  {key}: {value['count']}")
    if result.get("report_path"):
        out.append(f"  report: {result['report_path']}")
    return "\n".join(out)


# ---------------------------------------------------------------------- run

def run(repo: str | Path, fix: bool, report: bool, today: Optional[date] = None) -> dict:
    root = Path(repo).resolve()
    today = today or date.today()
    if not (root / "docs/00_index.md").is_file():
        raise FileNotFoundError(f"{root}: docs/00_index.md not found — not a governed repository")
    result: dict = {
        "repo": str(root), "today": today.isoformat(), "fix": fix,
        "fixes": {}, "report": {}, "report_path": None,
    }
    fixes = result["fixes"]
    # Order matters: envelopes first so statuses can be read, aliases next so
    # `resolved` becomes `complete` before the archive pass, the index last so
    # rows the archive pass removed are not first copied into the log.
    fill_front_matter(root, fix, today, fixes)
    normalize_statuses(root, fix, today, fixes)
    archive_complete(root, fix, fixes)
    move_index_narrative(root, fix, today, fixes)

    rep = result["report"]
    report_issues(root, today, rep)
    report_dead_references(root, rep)
    report_history_in_docs(root, rep)
    report_layout(root, rep)
    report_baselines(root, rep)

    if report:
        log_dir = root / "docs/log"
        log_dir.mkdir(parents=True, exist_ok=True)
        path = log_dir / f"hygiene-{today.strftime('%Y%m%d')}.md"
        path.write_text(render_report(result))
        result["report_path"] = rel(root, path)
    return result


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("repo", nargs="?", default=".", help="repository root (default: cwd)")
    parser.add_argument("--fix", action="store_true", help="apply mechanical fixes")
    parser.add_argument("--report", action="store_true",
                        help="write docs/log/hygiene-YYYYMMDD.md with judgment candidates")
    parser.add_argument("--json", action="store_true", help="print the full result as JSON")
    args = parser.parse_args()
    try:
        result = run(args.repo, fix=args.fix, report=args.report)
    except FileNotFoundError as error:
        print(f"docs_hygiene: {error}", file=sys.stderr)
        sys.exit(2)
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(summary_text(result))


if __name__ == "__main__":
    main()
