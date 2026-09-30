"""Repository policy for completed entries in docs/00_index.md."""

import re

from index_entries import (
    _editable_ranges,
    _entry_line_pattern,
    _table_row_pattern,
    find_index_entry_lines,
    remove_index_entry,
)


def read_policy(content: str) -> str:
    """Return the archive row policy; absence preserves the historic default."""
    front = re.match(r"\A---\r?\n(.*?)^---[ \t]*$", content, re.DOTALL | re.MULTILINE)
    if front is None:
        return "remove"
    values = re.findall(r"^archive_index_rows:[ \t]*(.*?)[ \t]*$", front.group(1), re.MULTILINE)
    if not values:
        return "remove"
    if len(values) != 1 or values[0] not in {"remove", "completed"}:
        raise ValueError("archive_index_rows must be 'remove' or 'completed' in docs/00_index.md front matter")
    return values[0]


def update_rows(content: str, rel_dir: str, entry_id: str, mode: str):
    """Remove active rows or carry one into the completed section.

    Returns updated content and the original row locations. All matching active
    rows are removed before inserting one completed entry, so a focus row cannot
    leave a duplicate active route after archive.
    """
    target_lines = find_index_entry_lines(content, rel_dir, entry_id)
    if mode == "keep":
        return content, [], 0
    updated, removed = remove_index_entry(content, rel_dir, entry_id)
    if removed != len(target_lines):
        raise ValueError(f"index target count mismatch: removed={removed}, reported={len(target_lines)}")
    if mode == "remove" or not target_lines:
        return updated, target_lines, 0
    if mode != "completed":
        raise ValueError(f"unsupported archive index policy: {mode}")

    heading = re.search(r"^## Completed \(archive\)[ \t]*$", updated, re.MULTILINE)
    if heading is None:
        raise ValueError("archive_index_rows: completed requires '## Completed (archive)' in docs/00_index.md")

    row = None
    for start, end in _editable_ranges(content):
        body = content[start:end]
        match = _entry_line_pattern(rel_dir, entry_id).search(body)
        if match:
            row = match.group(0).rstrip("\n")
            break
        match = _table_row_pattern(rel_dir, entry_id).search(body)
        if match:
            cells = [part.strip() for part in match.group(0).strip().strip("|").split("|")]
            row = "- " + " — ".join(cell for cell in cells if cell)
            break
    if row is None:
        raise ValueError("index row disappeared during archive preparation")

    # Handle Markdown links and the template's plain-path row form alike.
    old = f"{rel_dir}/{entry_id}.md"
    new = f"{rel_dir}/archive/{entry_id}.md"
    row = row.replace(old, new)
    insertion = "\n\n" + row + "\n"
    updated = updated[:heading.end()] + insertion + updated[heading.end():]
    return updated, target_lines, 1
