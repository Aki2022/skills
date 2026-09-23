"""Issue ownership and the routing lists derived from it (SPEC-doc-governance, Workstream Model).

An issue file in docs/issues/ declares `workstream: WS-...` (owned) or `workstream: none`
(standalone); that field is the only source of truth. The owning workstream's split-issue
list and the index's Active Issues / Active Workstreams rows are generated from it between
marker comments, never hand-written. Frontmatter parsing is passed in by the caller so this
module has no import cycle with validate_repo_docs.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Optional

# Files created on/after this date must carry the new fields; older ones only warn
# (SPEC: existing issues are never force-migrated). The day after the merge: anything
# created up to the merge day came from a create_issue.py that had no ownership flags.
ROLLOUT_DATE = "2026-09-24"

PRIORITIES = ("high", "medium", "low")
DUE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
ISSUE_ID_RE = re.compile(r"ISSUE-\d{8}-[A-Za-z0-9-]+")
WS_ID_RE = re.compile(r"^WS-\d{8}-[A-Za-z0-9-]+$")

SPLIT = "split-issues"
ACTIVE_ISSUES = "active-issues"
UNASSIGNED = "unassigned-issues"
ACTIVE_WORKSTREAMS = "active-workstreams"
HEADINGS = {
    SPLIT: "Split Issues",
    ACTIVE_ISSUES: "Active Issues",
    UNASSIGNED: "Unassigned Issues",
    ACTIVE_WORKSTREAMS: "Active Workstreams",
}
ROW_DESC_MAX = 200


def is_post_rollout(created_at: str) -> bool:
    return bool(created_at) and created_at[:10] >= ROLLOUT_DATE


def valid_due(value: str) -> bool:
    return value == "none" or bool(DUE_RE.match(value))


def scalar(front_matter: dict, key: str) -> str:
    value = front_matter.get(key, "")
    return value.split("#")[0].strip() if isinstance(value, str) else ""


# ------------------------------------------------------------------ generated blocks

def block_begin(name: str) -> str:
    return f"<!-- own-doc-update:generated {name} begin -->"


def block_end(name: str) -> str:
    return f"<!-- own-doc-update:generated {name} end -->"


def _block_span(content: str, name: str) -> Optional[tuple[int, int]]:
    begin = content.find(block_begin(name))
    if begin == -1:
        return None
    end = content.find(block_end(name), begin)
    if end == -1:
        return None
    return begin, end + len(block_end(name))


def read_block(content: str, name: str) -> Optional[list[str]]:
    """Lines between the markers, or None when the block does not exist."""
    span = _block_span(content, name)
    if span is None:
        return None
    inner = content[span[0] + len(block_begin(name)):span[1] - len(block_end(name))]
    return [line for line in inner.strip("\n").splitlines() if line.strip()]


def heading_span(content: str, prefix: str) -> Optional[tuple[int, int]]:
    """(start of the heading line, end of its section) for `## <prefix>...`, matched by prefix."""
    match = re.search(rf"^##[ \t]+{re.escape(prefix)}[^\n]*$", content, re.MULTILINE)
    if not match:
        return None
    nxt = re.search(r"^##[ \t]", content[match.end():], re.MULTILINE)
    end = match.end() + nxt.start() if nxt else len(content)
    return match.start(), end


def write_block(content: str, name: str, lines: list[str], heading: str,
                after_heading: Optional[str] = None) -> str:
    """Replace the generated block, creating it under `## <heading>` (prefix match) if absent.

    A missing heading is created at the end of the `after_heading` section when that exists,
    otherwise at the end of the file.
    """
    body = "\n".join(lines)
    block = f"{block_begin(name)}\n{body}\n{block_end(name)}" if lines else f"{block_begin(name)}\n{block_end(name)}"
    span = _block_span(content, name)
    if span is not None:
        return content[:span[0]] + block + content[span[1]:]
    section = heading_span(content, heading)
    if section is None:
        anchor = heading_span(content, after_heading) if after_heading else None
        if anchor is not None:
            head, tail = content[:anchor[1]].rstrip("\n"), content[anchor[1]:]
            return f"{head}\n\n## {heading}\n\n{block}\n\n{tail}"
        sep = "" if content.endswith("\n") else "\n"
        return f"{content}{sep}\n## {heading}\n\n{block}\n"
    heading_end = content.index("\n", section[0]) + 1 if "\n" in content[section[0]:] else len(content)
    return content[:heading_end] + "\n" + block + "\n" + content[heading_end:]


def section_outside_block(content: str, heading: str, name: str) -> str:
    """Text of the section named by `heading` with the generated block removed."""
    section = heading_span(content, heading)
    if section is None:
        return ""
    text = content[section[0]:section[1]]
    span = _block_span(text, name)
    if span is not None:
        text = text[:span[0]] + text[span[1]:]
    return text


# ------------------------------------------------------------------ repository model

@dataclass
class Doc:
    doc_id: str
    rel: str
    title: str
    created: str
    fm: dict
    content: str
    owner: Optional[str] = None  # issues only: WS id, "none", or None when undeclared

    @property
    def priority(self) -> str:
        return scalar(self.fm, "priority")

    @property
    def due(self) -> str:
        return scalar(self.fm, "due")


@dataclass
class Model:
    issues: dict[str, Doc] = field(default_factory=dict)
    archived_issue_ids: set[str] = field(default_factory=set)
    workstreams: dict[str, Doc] = field(default_factory=dict)
    archived_ws_ids: set[str] = field(default_factory=set)

    def owned_by(self, ws_id: str) -> list[Doc]:
        return sorted((d for d in self.issues.values() if d.owner == ws_id), key=lambda d: d.doc_id)

    def unassigned(self) -> list[Doc]:
        """Issues whose route cannot be derived: owner not an active WS, or a new file with none declared."""
        out = []
        for doc in self.issues.values():
            if doc.owner is None:
                if is_post_rollout(doc.created):
                    out.append(doc)
            elif doc.owner != "none" and doc.owner not in self.workstreams:
                out.append(doc)
        return sorted(out, key=lambda d: d.doc_id)

    def standalone(self) -> list[Doc]:
        return sorted((d for d in self.issues.values() if d.owner == "none"), key=lambda d: d.doc_id)


def _title(content: str, fallback: str) -> str:
    match = re.search(r"^#[ \t]+(.+)$", content, re.MULTILINE)
    return match.group(1).strip() if match else fallback


def load_model(root: Path, parse_front_matter: Callable[[Path], Optional[dict]]) -> Model:
    model = Model()
    for kind, folder, active, archived in (
        ("issue", root / "docs/issues", model.issues, model.archived_issue_ids),
        ("ws", root / "docs/workstreams", model.workstreams, model.archived_ws_ids),
    ):
        for path in sorted((folder / "archive").glob("*.md")) if (folder / "archive").is_dir() else []:
            archived.add(path.stem)
        if not folder.is_dir():
            continue
        for path in sorted(folder.glob("*.md")):
            fm = parse_front_matter(path) or {}
            content = path.read_text()
            doc = Doc(
                doc_id=path.stem,
                rel=path.relative_to(root).as_posix(),
                title=_title(content, path.stem),
                created=scalar(fm, "created_at")[:10],
                fm=fm,
                content=content,
            )
            if kind == "issue":
                owner = scalar(fm, "workstream")
                doc.owner = owner or None
            active[path.stem] = doc
    return model


# ------------------------------------------------------------------ rows

def _row(link: str, label: str, parts: list[str]) -> str:
    desc = " · ".join(p for p in parts if p)
    if len(desc) > ROW_DESC_MAX:
        desc = desc[: ROW_DESC_MAX - 1] + "…"
    return f"- [{label}]({link}) — {desc}"


def split_rows(model: Model, ws_id: str, exclude: set[str] = frozenset()) -> list[str]:
    return [
        _row(f"../issues/{d.doc_id}.md", d.doc_id, [d.title, f"due {d.due}" if d.due and d.due != "none" else ""])
        for d in model.owned_by(ws_id)
        if d.doc_id not in exclude
    ]


def active_issue_rows(model: Model, exclude: set[str] = frozenset()) -> list[str]:
    return [
        _row(f"issues/{d.doc_id}.md", d.doc_id, [d.title, d.priority, f"due {d.due}" if d.due else ""])
        for d in model.standalone()
        if d.doc_id not in exclude
    ]


def unassigned_rows(model: Model) -> list[str]:
    rows = []
    for d in model.unassigned():
        reason = f"workstream {d.owner} is not active" if d.owner else "workstream not declared"
        rows.append(_row(f"issues/{d.doc_id}.md", d.doc_id, [d.title, reason]))
    return rows


def earliest_due(model: Model, ws: Doc) -> tuple[str, str]:
    """(date, source id) — the nearest dated `due` of the WS and its owned issues, or ("", "")."""
    candidates = [(ws.due, ws.doc_id)] + [(d.due, d.doc_id) for d in model.owned_by(ws.doc_id)]
    dated = sorted((due, src) for due, src in candidates if DUE_RE.match(due or ""))
    if dated:
        return dated[0]
    return ("none", ws.doc_id) if ws.due == "none" else ("", "")


def workstream_rows(model: Model) -> list[str]:
    rows = []
    for ws in sorted(model.workstreams.values(), key=lambda d: d.doc_id):
        if not ws.priority and not ws.due:
            continue  # legacy workstream: still routed by its hand-written row
        due, source = earliest_due(model, ws)
        due_part = ""
        if due:
            due_part = f"due {due}" if source == ws.doc_id else f"due {due} ({source})"
        rows.append(_row(f"workstreams/{ws.doc_id}.md", ws.doc_id, [ws.title, ws.priority, due_part]))
    return rows


def generated_ws_ids(model: Model) -> set[str]:
    return {ws.doc_id for ws in model.workstreams.values() if ws.priority or ws.due}


def ids_in(lines: list[str]) -> set[str]:
    found = set()
    for line in lines:
        found.update(ISSUE_ID_RE.findall(line))
    return found


# ------------------------------------------------------------------ regeneration and reachability

def _drop_hand_rows(content: str, heading: str, block_name: str, rel_dir: str, ids: list[str]) -> tuple[str, int]:
    """Remove rows for `ids` from the `heading` section, outside its generated block."""
    from index_entries import remove_index_entry

    section = heading_span(content, heading)
    if section is None or not ids:
        return content, 0
    text = content[section[0]:section[1]]
    span = _block_span(text, block_name)
    parts = [(text, True)] if span is None else [
        (text[:span[0]], True), (text[span[0]:span[1]], False), (text[span[1]:], True)
    ]
    removed = 0
    out = []
    for part, editable in parts:
        if editable:
            for entry_id in ids:
                part, count = remove_index_entry(part, rel_dir, entry_id)
                removed += count
        out.append(part)
    return content[:section[0]] + "".join(out) + content[section[1]:], removed


def regenerate(root: Path, parse_front_matter: Callable[[Path], Optional[dict]]) -> list[dict]:
    """Rewrite every generated block from front matter and drop the hand rows they replace.

    Ownership is never rewritten here: an issue whose route cannot be derived is listed under
    Unassigned Issues with its declared owner left as it is. Returns one item per changed file.
    """
    model = load_model(root, parse_front_matter)
    changes: list[dict] = []
    for ws in model.workstreams.values():
        if not model.owned_by(ws.doc_id) and read_block(ws.content, SPLIT) is None:
            continue
        updated = write_block(ws.content, SPLIT, split_rows(model, ws.doc_id), HEADINGS[SPLIT])
        if updated != ws.content:
            (root / ws.rel).write_text(updated)
            changes.append({"file": ws.rel, "list": SPLIT})

    index_path = root / "docs/00_index.md"
    if not index_path.is_file():
        return changes
    original = index_path.read_text()
    text = original
    removed = 0
    rows = active_issue_rows(model)
    if rows or read_block(text, ACTIVE_ISSUES) is not None:
        text = write_block(text, ACTIVE_ISSUES, rows, HEADINGS[ACTIVE_ISSUES])
    declared = sorted(d.doc_id for d in model.issues.values() if d.owner is not None)
    text, count = _drop_hand_rows(text, HEADINGS[ACTIVE_ISSUES], ACTIVE_ISSUES, "docs/issues", declared)
    removed += count
    rows = unassigned_rows(model)
    if rows or read_block(text, UNASSIGNED) is not None:
        text = write_block(text, UNASSIGNED, rows, HEADINGS[UNASSIGNED], after_heading=HEADINGS[ACTIVE_ISSUES])
    rows = workstream_rows(model)
    if rows or read_block(text, ACTIVE_WORKSTREAMS) is not None:
        text = write_block(text, ACTIVE_WORKSTREAMS, rows, HEADINGS[ACTIVE_WORKSTREAMS])
    text, count = _drop_hand_rows(
        text, HEADINGS[ACTIVE_WORKSTREAMS], ACTIVE_WORKSTREAMS, "docs/workstreams", sorted(generated_ws_ids(model))
    )
    removed += count
    if text != original:
        index_path.write_text(text)
        changes.append({"file": "docs/00_index.md", "hand_rows_removed": removed})
    return changes


def reachability(root: Path, parse_front_matter: Callable[[Path], Optional[dict]]) -> dict:
    """Where every active issue is routed from; an orphan is routed from nowhere."""
    from index_entries import find_index_entry_lines

    model = load_model(root, parse_front_matter)
    index_path = root / "docs/00_index.md"
    index = index_path.read_text() if index_path.is_file() else ""
    span = heading_span(index, HEADINGS[ACTIVE_ISSUES])
    active_text = index[span[0]:span[1]] if span else ""
    listed_by_ws = set()
    for ws in model.workstreams.values():
        listed_by_ws |= ids_in(read_block(ws.content, SPLIT) or [])
    unassigned_ids = {d.doc_id for d in model.unassigned()}
    shown_unassigned = ids_in(read_block(index, UNASSIGNED) or [])
    summary = {"owned_listed": 0, "standalone_listed": 0, "legacy_listed": 0, "misrouted": 0,
               "unassigned": 0, "orphan": 0}
    orphans = []
    for doc in model.issues.values():
        in_index = bool(find_index_entry_lines(active_text, "docs/issues", doc.doc_id))
        if doc.doc_id in unassigned_ids:
            summary["unassigned"] += 1
        elif doc.doc_id in listed_by_ws:
            summary["owned_listed"] += 1
        elif in_index and doc.owner == "none":
            summary["standalone_listed"] += 1
        elif in_index and doc.owner is None:
            summary["legacy_listed"] += 1
        elif in_index:
            summary["misrouted"] += 1  # owned, but listed by the index instead of its workstream
        elif doc.doc_id not in shown_unassigned:
            summary["orphan"] += 1
            orphans.append(doc.rel)
    return {"summary": summary, "orphans": orphans, "unassigned": sorted(unassigned_ids)}
