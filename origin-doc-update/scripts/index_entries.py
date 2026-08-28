"""Shared helpers for editing docs/00_index.md active-reference lines.

00_index.template.md ships the bare-path form (`- docs/issues/X.md — ...`), but real
indexes commonly grow into markdown-link form (`- [X](issues/X.md) — ...`). Archiving
must handle both, otherwise the completed entry silently stays listed as active and
has to be deleted by hand every time.
"""

import re

__all__ = ["remove_index_entry"]


def _entry_line_pattern(rel_dir: str, entry_id: str) -> re.Pattern:
    """Match a list line whose *link target* (or bare path) is this entry's file.

    Deliberately anchored on the `.md` path so that a mere prose mention of the id in
    another entry's description is left alone.
    """
    ident = re.escape(entry_id)
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


def remove_index_entry(content: str, rel_dir: str, entry_id: str):
    """Drop every index list line pointing at ``<rel_dir>/<entry_id>.md``.

    Returns ``(new_content, removed_count)``. An entry can legitimately appear in more
    than one section (e.g. both Current Focus and Active Issues), so all are removed.
    """
    pattern = _entry_line_pattern(rel_dir, entry_id)
    new_content, removed = pattern.subn("", content)

    # Indexes also grow markdown *table* rows (`| [X](issues/X.md) | ... |`).
    # The list-line pattern missed those twice, printed "not found in
    # docs/00_index.md", and left the archived entry listed as active.
    ident = re.escape(entry_id)
    leaf = re.escape(rel_dir.split("/")[-1])
    target = rf"(?:{re.escape(rel_dir)}|{leaf})/{ident}\.md"
    table_pattern = re.compile(rf"^[ \t]*\|[^\n]*{target}[^\n]*\n?", re.MULTILINE)
    new_content, removed_rows = table_pattern.subn("", new_content)

    return new_content, removed + removed_rows
