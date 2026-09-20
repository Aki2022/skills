#!/usr/bin/env python3
"""Compress docs/00_index.md into the routing digest the SessionStart hook injects.

**Why a digest and not the file.** Measured 2026-09-19 over every session on one
machine: Claude Code delivers a hook's output to the agent inline only up to
about 10,000 CHARACTERS -- the largest ever delivered inline was 8,957, the
smallest ever spilled was 10,019, and that spilled one was this very index
injection. Past the limit the output is written to a file and the agent is handed
a 2 KB preview, so every "read the index first" rule downstream runs against
nothing, and nothing anywhere reports a failure.

**The unit is characters, not bytes.** A Japanese index carries about 1.27 bytes
per character, so a byte budget set at the character limit throws away a quarter
of the room, and a byte ceiling set well above it (the repository ceiling is
32 KB of bytes) does not bound the injection at all: the index measured 30,902
bytes and 24,234 characters on the day this was written, comfortably inside the
ceiling and 2.4x over the limit that decides whether it arrives.

**What the digest keeps, in priority order.** Routing is the index's job, so
every link target survives as long as anything does; descriptions shrink first,
then archived entries drop (they are historical context, not current truth), then
entries drop from the end. Whatever is dropped is counted in the output -- a
digest that quietly omits half the repository would be the same silent failure in
a smaller package. The full file is named on every run.

Link extraction here is deliberately simpler than `validate_repo_docs`: this
module only summarizes, never rewrites, so a missed exotic link costs a routing
line rather than a broken document.
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

# Sits under the measured inline maximum (8,957) with room for the two header
# lines the hook prints around it.
DIGEST_MAX_CHARS = 8000

# What the budget buys, richest first: (description cap, cap for archived rows,
# keep archived rows at all, rows kept per section).
#
# The order encodes one priority, measured against the real index: current work
# WITH its one line of meaning beats a complete list of paths without any. The
# first run spent all 8,000 characters on 94 bare paths, 42 of them archived --
# a session opens on current work, and archived documents stay named in the full
# file and one directory listing away. Whatever a rung drops is counted in the
# output; a digest that quietly omitted half the repository would be the same
# silent failure in a smaller package.
LADDER = (
    (200, 200, True, None),
    (120, 120, True, None),
    (80, 80, True, None),
    (80, 0, True, None),
    (40, 0, True, None),
    (120, 0, False, None),
    (80, 0, False, None),
    (40, 0, False, None),
    (0, 0, False, None),
    (0, 0, False, 24),
    (0, 0, False, 12),
    (0, 0, False, 6),
    (0, 0, False, 3),
    (0, 0, False, 1),
)

FOCUS_CAP = 240
POLICY_CAP = 320

LINK_RE = re.compile(r"\[[^\]\n]*\]\((?!https?://|mailto:)([^)\s]+\.md)(?:#[^)\s]*)?\)")
INLINE_LINK_RE = re.compile(r"\[([^\]\n]*)\]\([^)\s]*(?:\s+\"[^\"]*\")?\)")
HEADING_RE = re.compile(r"^[ \t]{0,3}#{1,6}[ \t]+(.*?)[ \t]*$")
FENCE_RE = re.compile(r"^[ \t]{0,3}(`{3,}|~{3,})")
BARE_ROW_RE = re.compile(r"^[ \t]*[-*+][ \t]+(docs/[\w./-]+\.md)")

HEADER = "Repository documentation index — routing digest. Full file: docs/00_index.md"
PATHS_NOTE = "Paths below are relative to docs/."


def _clean(text: str) -> str:
    """Strip the markup a routing line does not need to route."""
    text = INLINE_LINK_RE.sub(r"\1", text)          # [label](url) -> label
    text = re.sub(r"[*`]+", "", text)
    # docs_hygiene.py appends "…（全文: log/index-YYYYMM.md）" to every row it
    # shortens; the digest names the full file once, so the repeat is budget.
    text = re.sub(r"\s*…?\s*（全文[:：][^）]*）\s*$", "…", text)
    text = re.sub(r"^[\s:—–-]+", "", text)
    return re.sub(r"[ \t]+", " ", text).strip()


def _cut(text: str, cap: int) -> str:
    if cap <= 0 or not text:
        return ""
    if len(text) <= cap:
        return text
    return text[:cap].rstrip() + "…"


def _targets(line: str) -> list[str]:
    found = [match.group(1) for match in LINK_RE.finditer(line)]
    bare = BARE_ROW_RE.match(line)
    if bare:
        found.append(bare.group(1))
    out: list[str] = []
    for target in found:
        target = target.split("#")[0]
        if target.startswith("docs/"):
            target = target[len("docs/"):]
        if target and target not in out:
            out.append(target)
    return out


class _Row:
    __slots__ = ("targets", "desc")

    def __init__(self, targets: list[str], desc: str) -> None:
        self.targets = targets
        self.desc = desc

    @property
    def archived(self) -> bool:
        return all("archive/" in target for target in self.targets)

    def render(self, cap: int, archive_cap: int) -> str:
        head = "- " + " / ".join(self.targets)
        desc = _cut(self.desc, archive_cap if self.archived else cap)
        return f"{head} — {desc}" if desc else head


class _Section:
    __slots__ = ("title", "rows", "prose")

    def __init__(self, title: str) -> None:
        self.title = title
        self.rows: list[_Row] = []
        self.prose: list[str] = []


def _front_matter(lines: list[str]) -> tuple[list[str], int]:
    """`(digest lines for the front matter, index of the first body line)`."""
    if not lines or lines[0].strip() != "---":
        return [], 0
    end = next((n for n in range(1, len(lines)) if lines[n].strip() == "---"), None)
    if end is None:
        return [], 0
    out = []
    for line in lines[1:end]:
        key = re.match(r"^([A-Za-z_][\w-]*):[ \t]*(.*)$", line)
        if not key:
            continue
        name, value = key.group(1), _clean(key.group(2))
        if name == "updated_at" and value:
            out.append(f"updated_at: {value}")
        elif name == "current_focus" and value:
            out.append(f"current_focus: {_cut(value, FOCUS_CAP)}")
    return out, end + 1


def _parse(text: str) -> tuple[list[str], list[_Section]]:
    lines = text.splitlines()
    front, start = _front_matter(lines)
    sections: list[_Section] = []
    current = _Section("")
    in_fence = False
    fence = ""
    for line in lines[start:]:
        if in_fence:
            if re.match(rf"^[ \t]{{0,3}}{re.escape(fence)}{{3,}}[ \t]*$", line):
                in_fence = False
            continue
        opening = FENCE_RE.match(line)
        if opening:
            in_fence = True
            fence = opening.group(1)[0]
            continue
        heading = HEADING_RE.match(line)
        if heading:
            if current.rows or current.prose:
                sections.append(current)
            current = _Section(heading.group(1))
            continue
        if not line.strip():
            continue
        targets = _targets(line)
        if targets:
            current.rows.append(_Row(targets, _clean(LINK_RE.sub("", line))))
        else:
            current.prose.append(_clean(line))
    if current.rows or current.prose:
        sections.append(current)
    return front, sections


def _render(front: list[str], sections: list[_Section], cap: int, archive_cap: int,
            keep_archived: bool, per_section: int | None) -> str:
    out = [HEADER, PATHS_NOTE]
    out.extend(front)
    dropped = 0
    archived_dropped = 0
    seen: set[str] = set()
    for section in sections:
        body: list[str] = []
        if section.rows:
            kept = 0
            for row in section.rows:
                fresh = [target for target in row.targets if target not in seen]
                if not fresh:
                    continue
                if not keep_archived and row.archived:
                    archived_dropped += 1
                    continue
                if per_section is not None and kept >= per_section:
                    dropped += 1
                    continue
                seen.update(fresh)
                body.append(_Row(fresh, row.desc).render(cap, archive_cap))
                kept += 1
        elif section.prose:
            # A section with no links is policy (how to read), not progress.
            body.append(_cut(" ".join(section.prose), POLICY_CAP))
        if not body:
            continue
        if section.title:
            out.append("")
            out.append(f"## {section.title}")
        out.extend(body)
    tail = []
    if archived_dropped:
        tail.append(f"archive 済みの {archived_dropped} 件は省略")
    if dropped:
        tail.append(f"ほか {dropped} 件を省略")
    if tail:
        out.append("")
        out.append("（" + "、".join(tail) + "。全文は docs/00_index.md）")
    return "\n".join(out).rstrip() + "\n"


def build_digest(text: str, max_chars: int = DIGEST_MAX_CHARS) -> str:
    """Return a routing digest of `text` that is at most `max_chars` characters."""
    front, sections = _parse(text or "")
    rendered = ""
    for cap, archive_cap, keep_archived, per_section in LADDER:
        rendered = _render(front, sections, cap, archive_cap,
                           keep_archived=keep_archived, per_section=per_section)
        if len(rendered) <= max_chars:
            return rendered
    return rendered[: max_chars - 1].rstrip() + "…"


def main() -> int:
    if len(sys.argv) != 2:
        print("usage: index_digest.py <path to docs/00_index.md>", file=sys.stderr)
        return 2
    try:
        text = Path(sys.argv[1]).read_text()
    except OSError as error:
        print(f"index_digest: {error}", file=sys.stderr)
        return 1
    sys.stdout.write(build_digest(text))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
