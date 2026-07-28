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
    return re.compile(
        rf"^[ \t]*-[ \t]+[^\n]*(?:\]\([^)\n]*{target}\)|(?<![\w./]){target}(?![\w.]))[^\n]*\n?",
        re.MULTILINE,
    )


def remove_index_entry(content: str, rel_dir: str, entry_id: str):
    """Drop every index list line pointing at ``<rel_dir>/<entry_id>.md``.

    Returns ``(new_content, removed_count)``. An entry can legitimately appear in more
    than one section (e.g. both Current Focus and Active Issues), so all are removed.
    """
    pattern = _entry_line_pattern(rel_dir, entry_id)
    new_content, removed = pattern.subn("", content)
    return new_content, removed
