"""Rewrite relative links when a document moves into `archive/`.

Archiving puts a document one directory deeper. Two sets of links break, in
opposite directions:

- the links **inside** the moved document are one `../` short;
- the links **to** it from every other document point at a path that no longer
  exists.

Neither was rewritten before 2026-09-18, and nothing failed at the time: a
broken link in a docs tree breaks no test, and the validator's link check only
arrived on 2026-09-07. One session measured seven separate breakages from its
own archive moves and repaired each by hand.

Link extraction is NOT re-implemented here. `validate_repo_docs.iter_link_destinations`
already walks inline links with balanced parentheses, and `strip_code` already
decides what is a code span and therefore not a link. A second walk here would
be free to drift from the one the validator enforces -- so this module rewrites
through the same functions, and `strip_code` blanks to the same length precisely
so its offsets address the original text.
"""

from __future__ import annotations

import posixpath
from pathlib import Path
from typing import Callable, Optional

from validate_repo_docs import iter_link_destinations, link_target, strip_code


def rewrite_destinations(text: str, rewrite: Callable[[str], Optional[str]]) -> str:
    """Apply `rewrite` to each link destination, splicing results into `text`.

    `rewrite` receives the raw destination (which may carry an anchor, a query
    or a title) and returns a replacement, or None to leave it alone. Spans come
    from the code-stripped text, so a link inside an inline code span is never
    offered -- that is how `[x](../placeholder.md)` stays an opt-out.
    """
    edits: list[tuple[int, int, str]] = []
    for start, end, raw in iter_link_destinations(strip_code(text)):
        replacement = rewrite(raw)
        if replacement is not None and replacement != raw:
            edits.append((start, end, replacement))
    if not edits:
        return text
    out = text
    for start, end, replacement in reversed(edits):
        out = out[:start] + replacement + out[end:]
    return out


def _split_destination(raw: str) -> Optional[tuple[str, str, str]]:
    """`(prefix, path, suffix)` for a rewritable relative destination, else None.

    The path is what `link_target` resolves; the prefix and suffix carry the
    angle brackets, anchor, query and title back into the rewritten text, so an
    anchor is not silently dropped by a move.
    """
    target = link_target(raw)
    if target is None or posixpath.isabs(target):
        return None
    index = raw.find(target)
    if index == -1:
        # `link_target` percent-decodes; an encoded destination is left alone
        # rather than rewritten into a shape the author did not write.
        return None
    return raw[:index], target, raw[index + len(target) :]


def reanchor(text: str, *, old_dir: str, new_dir: str) -> str:
    """Rewrite the moved document's own relative links for its new directory."""

    def rewrite(raw: str) -> Optional[str]:
        parts = _split_destination(raw)
        if parts is None:
            return None
        prefix, target, suffix = parts
        resolved = posixpath.normpath(posixpath.join(old_dir, target))
        updated = posixpath.relpath(resolved, new_dir)
        return f"{prefix}{updated}{suffix}"

    return rewrite_destinations(text, rewrite)


def repoint(text: str, *, referrer_dir: str, old_path: str, new_path: str) -> str:
    """Repoint a referring document's links from `old_path` to `new_path`.

    Only links that actually resolve to `old_path` are touched, so a link to a
    different document in the same directory is left alone.
    """
    old_path = posixpath.normpath(old_path)

    def rewrite(raw: str) -> Optional[str]:
        parts = _split_destination(raw)
        if parts is None:
            return None
        prefix, target, suffix = parts
        resolved = posixpath.normpath(posixpath.join(referrer_dir, target))
        if resolved != old_path:
            return None
        updated = posixpath.relpath(new_path, referrer_dir)
        return f"{prefix}{updated}{suffix}"

    return rewrite_destinations(text, rewrite)


def plan_link_updates(
    repo: str,
    old_rel: str,
    new_rel: str,
    content: str,
    exclude: Optional[set[str]] = None,
) -> tuple[str, list[tuple[str, str, str]]]:
    """`(reanchored_content, [(path, original, rewritten), ...])` for one move.

    `old_rel` / `new_rel` are repo-relative POSIX paths. The returned referrer
    list covers every `docs/**/*.md` whose links resolve to `old_rel`, except the
    moved document itself -- its own links are handled by the reanchor, and
    touching it twice would fight the archive transaction.

    `exclude` names repo-relative paths to leave alone. The callers pass
    `docs/00_index.md`, which owns its own entry handling: the archive scripts
    *remove* the active row entirely, and repointing the same line here installed
    after that removal and undid it. The existing suite caught exactly that.
    """
    skip = {posixpath.normpath(path) for path in (exclude or set())}
    import os

    old_dir = posixpath.dirname(old_rel)
    new_dir = posixpath.dirname(new_rel)
    reanchored = reanchor(content, old_dir=old_dir, new_dir=new_dir)

    updates: list[tuple[str, str, str]] = []
    docs_root = os.path.join(repo, "docs")
    for dirpath, _dirnames, filenames in os.walk(docs_root):
        for name in sorted(filenames):
            if not name.endswith(".md"):
                continue
            absolute = os.path.join(dirpath, name)
            rel = posixpath.normpath(os.path.relpath(absolute, repo).replace(os.sep, "/"))
            if rel in (old_rel, new_rel) or rel in skip:
                continue
            try:
                original = Path(absolute).read_text()
            except (OSError, UnicodeDecodeError):
                continue
            rewritten = repoint(
                original,
                referrer_dir=posixpath.dirname(rel),
                old_path=old_rel,
                new_path=new_rel,
            )
            if rewritten != original:
                updates.append((absolute, original, rewritten))
    return reanchored, updates
