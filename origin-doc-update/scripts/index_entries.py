"""Shared helpers for editing docs/00_index.md active-reference lines.

00_index.template.md ships the bare-path form (`- docs/issues/X.md — ...`), but real
indexes commonly grow into markdown-link form (`- [X](issues/X.md) — ...`). Archiving
must handle both, otherwise the completed entry silently stays listed as active and
has to be deleted by hand every time.
"""

import re

__all__ = ["find_index_entry_lines", "normalize_entry_id", "remove_index_entry"]


def normalize_entry_id(value: str) -> str:
    """Normalize an id, filename, or path to the stem used by an index link."""
    value = value.strip().replace("\\", "/").rstrip("/")
    value = value.rsplit("/", 1)[-1]
    return value[:-3] if value.endswith(".md") else value


def _section_bodies(content: str) -> list[tuple[int, int]]:
    """Return ranges between headings, excluding heading lines themselves.

    Markdown headings inside fenced code are ignored. If there is no heading, the
    whole input remains one compatible fixture body.
    """
    headings: list[tuple[int, int]] = []
    offset = 0
    fence_char = None
    fence_width = 0
    for line in content.splitlines(keepends=True):
        bare = line.rstrip("\r\n")
        if fence_char is not None:
            if re.match(
                rf"^[ \t]{{0,3}}{re.escape(fence_char)}{{{fence_width},}}[ \t]*$",
                bare,
            ):
                fence_char = None
                fence_width = 0
        else:
            opening = re.match(r"^[ \t]{0,3}(`{3,}|~{3,})", bare)
            if opening:
                fence_char = opening.group(1)[0]
                fence_width = len(opening.group(1))
            elif re.match(r"^[ \t]{0,3}#{1,6}[ \t]+", bare):
                headings.append((offset, offset + len(line)))
        offset += len(line)

    if not headings:
        return [(0, len(content))]
    return [
        (heading_end, headings[index + 1][0] if index + 1 < len(headings) else len(content))
        for index, (_heading_start, heading_end) in enumerate(headings)
    ]


def _editable_ranges(content: str) -> list[tuple[int, int]]:
    """Return heading-body ranges with fenced code removed from edit scope."""
    ranges: list[tuple[int, int]] = []
    for section_start, section_end in _section_bodies(content):
        segment_start = section_start
        offset = section_start
        fence_char = None
        fence_width = 0
        for line in content[section_start:section_end].splitlines(keepends=True):
            bare = line.rstrip("\r\n")
            if fence_char is not None:
                if re.match(
                    rf"^[ \t]{{0,3}}{re.escape(fence_char)}{{{fence_width},}}[ \t]*$",
                    bare,
                ):
                    segment_start = offset + len(line)
                    fence_char = None
                    fence_width = 0
            else:
                opening = re.match(r"^[ \t]{0,3}(`{3,}|~{3,})", bare)
                if opening:
                    if segment_start < offset:
                        ranges.append((segment_start, offset))
                    fence_char = opening.group(1)[0]
                    fence_width = len(opening.group(1))
            offset += len(line)
        if fence_char is None and segment_start < section_end:
            ranges.append((segment_start, section_end))
    return ranges


def _entry_line_pattern(rel_dir: str, entry_id: str) -> re.Pattern:
    """Match a list line whose *link target* (or bare path) is this entry's file.

    Deliberately anchored on the `.md` path so that a mere prose mention of the id in
    another entry's description is left alone.
    """
    ident = re.escape(normalize_entry_id(entry_id))
    # docs/issues -> issues, so both "docs/issues/X.md" and "](issues/X.md)" match.
    leaf = re.escape(rel_dir.split("/")[-1])
    target = rf"(?:{re.escape(rel_dir)}|{leaf})/{ident}\.md"
    # A wrapped entry continues on indented lines that carry no bullet of their
    # own. Matching only the bullet line left those behind as an orphan paragraph
    # under the section heading — the index still validates, so nothing
    # downstream disagrees and only a human notices the fragment. The
    # continuation clause stops at the next bullet or any unindented line.
    #
    # 2026-08-26: a bare `[^\n]*` before the alternation let the target link
    # match *anywhere* in the bullet, so a DIFFERENT entry's bullet whose prose
    # cites the archived one as a link (not just an id mention) got deleted
    # whole. `(?:(?!\]\()[^\n])*` walks forward one char at a time but refuses
    # to step past the start of the bullet's first `](` — so the alternation
    # can only match against that first link (or bare text before it), never a
    # later one. This is the link-form counterpart of the existing "prose
    # mention is left alone" guarantee below, which only ever covered bare ids.
    return re.compile(
        rf"^[ \t]*-[ \t]+(?:(?!\]\()[^\n])*(?:\]\([^)\n]*{target}\)|(?<![\w./]){target}(?![\w.]))[^\n]*\n?"
        # 2026-08-18: a plain (backtracking) `[ \t]+` here let the quantifier
        # retry with fewer whitespace chars whenever a lookahead placed right
        # after it rejected the greedy match. For a *nested* sibling bullet
        # (e.g. "  - [B](...)") that retry lands on a 1-space split where the
        # next char is a plain space rather than `-`/`*`/`+`, so a lookahead
        # sitting at that post-`[ \t]+` position gets satisfied against the
        # wrong offset and the trailing `[^\n]*` swallows the sibling bullet
        # as continuation prose — cascading into every following nested
        # sibling and wiping the block.
        #
        # A first fix used `++` (possessive) to forbid that retry, but `++`
        # is Python 3.11+ only (`re.error: multiple repeat` on 3.9/3.10) and
        # the dependency was invisible from the call site — nothing declares
        # a 3.11 floor, so `archive_issue.py` would die at import on an older
        # interpreter. The portable fix instead moves the "is this a bullet
        # line?" check to a fixed anchor that never moves: a lookahead
        # evaluated *before* any indentation is consumed, which scans
        # `[ \t]*` internally to ask "does *some* amount of leading
        # whitespace on this line reach a bullet marker?" That question's
        # answer does not depend on how the following `[ \t]+` later
        # backtracks — the assertion was already resolved at a fixed offset
        # — so the plain backtracking quantifier can no longer defeat it, and
        # no possessive quantifier (or any 3.11+ syntax) is needed.
        rf"(?:(?![ \t]*[-*+][ \t])[ \t]+[^\n]*\n?)*",
        re.MULTILINE,
    )


def _table_row_pattern(rel_dir: str, entry_id: str) -> re.Pattern:
    """Match a table row whose first cell points at the target entry."""
    ident = re.escape(normalize_entry_id(entry_id))
    leaf = re.escape(rel_dir.split("/")[-1])
    target = rf"(?:{re.escape(rel_dir)}|{leaf})/{ident}\.md"
    return re.compile(
        rf"^[ \t]*\|[^\n|]*(?:\]\([^\)\n|]*{target}\)|(?<![\w./]){target}(?![\w.]))[^\n|]*\|[^\n]*\n?",
        re.MULTILINE,
    )


def find_index_entry_lines(content: str, rel_dir: str, entry_id: str) -> list[tuple[int, str]]:
    """Return ``(1-based line number, first line)`` for every target row/bullet."""
    patterns = (_entry_line_pattern(rel_dir, entry_id), _table_row_pattern(rel_dir, entry_id))
    found: set[tuple[int, str]] = set()
    for body_start, body_end in _editable_ranges(content):
        body = content[body_start:body_end]
        for pattern in patterns:
            for match in pattern.finditer(body):
                line_number = content.count("\n", 0, body_start + match.start()) + 1
                first_line = match.group(0).splitlines()[0].rstrip("\r")
                found.add((line_number, first_line))
    return sorted(found)


def remove_index_entry(content: str, rel_dir: str, entry_id: str):
    """Drop every index list line pointing at ``<rel_dir>/<entry_id>.md``.

    Returns ``(new_content, removed_count)``. An entry can legitimately appear in more
    than one section (e.g. both Current Focus and Active Issues), so all are removed.
    """
    ranges = _editable_ranges(content)
    pieces: list[str] = []
    cursor = 0
    removed_total = 0
    for start, end in ranges:
        pieces.append(content[cursor:start])
        body = content[start:end]
        updated, removed = _entry_line_pattern(rel_dir, entry_id).subn("", body)
        updated, removed_rows = _table_row_pattern(rel_dir, entry_id).subn("", updated)
        pieces.append(updated)
        removed_total += removed + removed_rows
        cursor = end
    pieces.append(content[cursor:])
    return "".join(pieces), removed_total
