#!/usr/bin/env python3
"""Validate the origin-doc-update repository structure and v2 lifecycle contracts."""
from __future__ import annotations

import argparse
import os
import re
import sys
from pathlib import Path
from typing import Optional


FrontMatter = dict[str, object]


def parse_inline_list(value: str) -> list[str]:
    value = value.strip()
    if value == "[]":
        return []
    if not (value.startswith("[") and value.endswith("]")):
        return []
    return [
        item.strip().strip('"').strip("'")
        for item in value[1:-1].split(",")
        if item.strip()
    ]


KEY_LINE = re.compile(r"^[A-Za-z0-9_]+:")

# Reserved result key holding the names of keys written in flow style. Reading it
# through flow_style_keys() keeps callers from having to know the spelling.
FLOW_STYLE_KEYS = "__flow_style_keys__"


def flow_style_keys(front_matter: Optional[FrontMatter]) -> list[str]:
    """Keys whose value was a bracket list opened on the line after the key."""
    if not front_matter:
        return []
    keys = front_matter.get(FLOW_STYLE_KEYS, [])
    return list(keys) if isinstance(keys, list) else []


def read_flow_continuation(lines: list[str], start: int) -> tuple[int, list[str]]:
    """Read a bracket list that opens on `lines[start]`, i.e. below its key.

    Returns the number of lines consumed (0 when there is no such list) and the
    parsed entries. This form is not valid input for the line-based block-list
    reader below, so without this it would silently parse as an empty list.
    """
    if start >= len(lines) or not lines[start].strip().startswith("["):
        return 0, []

    buffer: list[str] = []
    depth = 0
    for offset in range(start, len(lines)):
        line = lines[offset]
        if offset > start and KEY_LINE.match(line):
            return 0, []  # unterminated — not a flow list after all
        buffer.append(line.strip())
        depth += line.count("[") - line.count("]")
        if depth <= 0:
            return offset - start + 1, parse_inline_list(" ".join(buffer))
    return 0, []


def parse_front_matter(path: str | Path) -> Optional[FrontMatter]:
    """Parse the small YAML subset used by origin-doc-update templates."""
    try:
        content = Path(path).read_text()
    except OSError:
        return None

    if not content.startswith("---"):
        return {}

    end = content.find("\n---", 3)
    if end == -1:
        return None

    result: FrontMatter = {}
    flow_keys: list[str] = []
    current_list: Optional[str] = None
    lines = content[3:end].strip().splitlines()
    index = 0
    while index < len(lines):
        line = lines[index]
        index += 1
        if KEY_LINE.match(line):
            key, _, raw = line.partition(":")
            raw = raw.strip()
            current_list = None
            if raw.startswith("[") and raw.endswith("]"):
                result[key] = parse_inline_list(raw)
            elif not raw:
                result[key] = []
                current_list = key
                consumed, items = read_flow_continuation(lines, index)
                if consumed:
                    result[key] = items
                    flow_keys.append(key)
                    current_list = None
                    index += consumed
            else:
                result[key] = raw.strip('"').strip("'")
        elif current_list and re.match(r"^[ \t]+-[ \t]+", line):
            item = re.sub(r"^[ \t]+-[ \t]+", "", line).strip()
            cast = result[current_list]
            if isinstance(cast, list):
                cast.append(item.strip('"').strip("'"))
    if flow_keys:
        result[FLOW_STYLE_KEYS] = flow_keys
    return result


def as_list(value: object) -> list[str]:
    if isinstance(value, list):
        return [str(item) for item in value]
    if isinstance(value, str) and value:
        return parse_inline_list(value)
    return []


def as_text(value: object) -> str:
    return value if isinstance(value, str) else ""


def validate_guide_impact(
    label: str, impact: str, guides: list[str], reason: str, errors: list[str]
) -> None:
    if impact not in ("required", "none"):
        errors.append(f"{label}: guide_impact must be 'required' or 'none'")
    elif impact == "required" and not guides:
        errors.append(f"{label}: guide_impact is required but related_guides is empty")
    elif impact == "none" and not reason:
        errors.append(f"{label}: guide_impact_reason is required when guide_impact is none")


def validate_guide_sources(
    label: str,
    source_id: str,
    source_key: str,
    guide_refs: list[str],
    id_locations: dict[str, Path],
    root: Path,
    errors: list[str],
) -> None:
    for reference in guide_refs:
        guide_path = id_locations.get(reference)
        if guide_path is None:
            candidate = root / reference
            guide_path = candidate if candidate.is_file() else None
        if guide_path is None:
            errors.append(f"{label}: related guide not found: {reference}")
            continue
        guide_fm = parse_front_matter(guide_path)
        if not guide_fm or source_id not in as_list(guide_fm.get(source_key, [])):
            errors.append(
                f"{label}: {reference} must list {source_key}: [{source_id}]"
            )


def parse_workstream_issue_blocks(content: str) -> list[tuple[str, dict[str, str]]]:
    """Read compact issue metadata under `### ISSUE-*` headings."""
    matches = list(re.finditer(r"^### (ISSUE-[A-Za-z0-9-]+).*$", content, re.MULTILINE))
    blocks: list[tuple[str, dict[str, str]]] = []
    for index, match in enumerate(matches):
        end = matches[index + 1].start() if index + 1 < len(matches) else len(content)
        body = content[match.end():end]
        metadata: dict[str, str] = {}
        for item in re.finditer(r"^- ([a-z_]+):[ \t]*(.*?)$", body, re.MULTILINE):
            metadata[item.group(1)] = item.group(2).strip().strip('"').strip("'")
        blocks.append((match.group(1), metadata))
    return blocks


def parse_issue_queue_table_ids(content: str) -> set[str]:
    """Read ISSUE-* ids out of the Issue Queue markdown table."""
    section = re.search(
        r"^## Issue Queue\s*$(.*?)(?=^## |\Z)", content, re.MULTILINE | re.DOTALL
    )
    if not section:
        return set()
    ids: set[str] = set()
    for line in section.group(1).splitlines():
        if not line.lstrip().startswith("|"):
            continue
        cells = [cell.strip() for cell in line.strip().strip("|").split("|")]
        if not cells:
            continue
        match = re.search(r"\b(ISSUE-[A-Za-z0-9-]+)", cells[0])
        if match:
            ids.add(match.group(1))
    return ids


def validate_repo(repo: str | Path) -> tuple[list[str], list[str]]:
    root = Path(repo).resolve()
    errors: list[str] = []
    warnings: list[str] = []

    required_dirs = (
        "docs/adrs",
        "docs/specs",
        "docs/issues",
        "docs/issues/archive",
        "docs/workstreams",
        "docs/workstreams/archive",
        "docs/guides",
    )
    for directory in required_dirs:
        if not (root / directory).is_dir():
            errors.append(f"Missing required directory: {directory}")

    index_path = root / "docs/00_index.md"
    if not index_path.is_file():
        errors.append("Missing docs/00_index.md")
    else:
        fm = parse_front_matter(index_path)
        if fm is None:
            errors.append("docs/00_index.md: broken front matter")
        elif "updated_at" not in fm:
            warnings.append("docs/00_index.md: missing updated_at in front matter")

    id_locations: dict[str, Path] = {}
    for path in sorted((root / "docs").glob("**/*.md")):
        fm = parse_front_matter(path)
        if fm:
            doc_id = as_text(fm.get("id", ""))
            if doc_id:
                # Overwriting on collision let an archived namesake win, because
                # sorted() puts `archive/` after the live file — so a later
                # lookup checked the archived copy and reported on the wrong
                # document. Prefer the live one and say the id is duplicated.
                previous = id_locations.get(doc_id)
                if previous is not None:
                    errors.append(
                        f"{path.relative_to(root)}: duplicate id '{doc_id}', also in "
                        f"{previous.relative_to(root)}"
                    )
                    if "archive" in path.parts:
                        continue
                id_locations[doc_id] = path
            for key in flow_style_keys(fm):
                errors.append(
                    f"{path.relative_to(root)}: front matter '{key}' opens a bracket "
                    "list on the line below the key; rewrite it in block style "
                    "(one '- item' per line)"
                )

    archive_dirs = (
        root / "docs/issues/archive",
        root / "docs/workstreams/archive",
    )
    for archive_dir in archive_dirs:
        if not archive_dir.is_dir():
            continue
        for path in sorted(archive_dir.glob("*.md")):
            fm = parse_front_matter(path)
            if fm is None:
                errors.append(f"{path.relative_to(root)}: broken front matter")
            elif fm and as_text(fm.get("status", "")) not in ("archived", "archive", ""):
                warnings.append(
                    f"{path.relative_to(root)}: status is '{fm.get('status')}', expected 'archived'"
                )

    adrs_dir = root / "docs/adrs"
    if adrs_dir.is_dir():
        valid_statuses = {"proposed", "accepted", "rejected", "superseded"}
        # Only ADR files, so a README explaining the directory is not an error.
        for path in sorted(adrs_dir.glob("ADR-*.md")):
            rel = str(path.relative_to(root))
            fm = parse_front_matter(path)
            if fm is None:
                errors.append(f"{rel}: broken front matter")
                continue
            adr_id = as_text(fm.get("id", ""))
            if not adr_id.startswith("ADR-"):
                errors.append(f"{rel}: id must start with ADR-")
            scope = as_text(fm.get("scope", ""))
            if scope not in ("spec", "development"):
                errors.append(f"{rel}: scope must be 'spec' or 'development'")
            status = as_text(fm.get("status", ""))
            if status not in valid_statuses:
                errors.append(
                    f"{rel}: status must be proposed, accepted, rejected, or superseded"
                )
            if status == "superseded" and not as_text(fm.get("superseded_by", "")):
                errors.append(f"{rel}: superseded_by is required when status is superseded")
            for key in flow_style_keys(fm):
                errors.append(
                    f"{rel}: front matter '{key}' opens a bracket list on the line below the key; "
                    "rewrite it in block style (one '- item' per line)"
                )

    branch_owners: dict[str, list[str]] = {}
    issues_dir = root / "docs/issues"
    if issues_dir.is_dir():
        for path in sorted(issues_dir.glob("*.md")):
            rel = str(path.relative_to(root))
            fm = parse_front_matter(path)
            if fm is None:
                errors.append(f"{rel}: broken front matter")
                continue
            branch = as_text(fm.get("branch", "")).strip()
            if not branch:
                warnings.append(f"{rel}: missing branch (no branch recorded to resume/clean up)")
            elif branch not in ("main", "master"):
                # Shared trunk branches are never deleted by cleanup, so multiple
                # direct-to-main issues sharing them is not an ownership conflict.
                branch_owners.setdefault(branch, []).append(path.name)
            if as_text(fm.get("schema_version", "")) != "2":
                # Same silent exemption as workstreams: without the version, every
                # guide-impact check below is skipped and the file passes clean.
                errors.append(f"{rel}: schema_version must be 2")
            else:
                guide_refs = as_list(fm.get("related_guides", []))
                guide_impact = as_text(fm.get("guide_impact", ""))
                validate_guide_impact(
                    rel,
                    guide_impact,
                    guide_refs,
                    as_text(fm.get("guide_impact_reason", "")),
                    errors,
                )
                if guide_impact == "required":
                    validate_guide_sources(
                        rel,
                        as_text(fm.get("id", "")),
                        "source_issues",
                        guide_refs,
                        id_locations,
                        root,
                        errors,
                    )

    for branch, owners in branch_owners.items():
        if len(owners) > 1:
            warnings.append(
                f"branch '{branch}' is recorded on multiple active issues: {', '.join(owners)}"
            )

    workstreams_dir = root / "docs/workstreams"
    if workstreams_dir.is_dir():
        for path in sorted(workstreams_dir.glob("*.md")):
            rel = str(path.relative_to(root))
            fm = parse_front_matter(path)
            if fm is None:
                errors.append(f"{rel}: broken front matter")
                continue
            if as_text(fm.get("schema_version", "")) != "2":
                # Skipping quietly made the whole contract below opt-in by the
                # document under test: a workstream with no envelope, no gates
                # and no issue blocks passed clean, so origin-goal's rule that
                # the envelope must be recorded before starting had no
                # mechanical check left.
                errors.append(f"{rel}: schema_version must be 2")
                continue
            if not as_text(fm.get("human_boundary_confirmed_at", "")):
                errors.append(f"{rel}: human_boundary_confirmed_at is required")
            if not as_text(fm.get("next_human_gate", "")):
                errors.append(f"{rel}: next_human_gate is required")
            content = path.read_text()
            for heading in ("## Authorization Envelope", "## Human Gates", "## Issue Queue"):
                if heading not in content:
                    errors.append(f"{rel}: missing section '{heading}'")
            issue_blocks = parse_workstream_issue_blocks(content)
            if not issue_blocks:
                errors.append(f"{rel}: no embedded ISSUE-* blocks found")
            # The Issue Queue table is what a human reads and what the loop reads
            # to find the next pending issue, but no gate looked at it. A table
            # row whose block was never written — or whose block a table reformat
            # ate — was invisible, so committed scope could archive as complete.
            tabled = parse_issue_queue_table_ids(content)
            blocked = {issue_id for issue_id, _metadata in issue_blocks}
            for issue_id in sorted(tabled - blocked):
                errors.append(
                    f"{rel}: {issue_id} is listed in the Issue Queue table but has no "
                    "'### ' block, so its status is never validated"
                )
            for issue_id in sorted(blocked - tabled):
                errors.append(
                    f"{rel}: {issue_id} has a '### ' block but is missing from the "
                    "Issue Queue table"
                )
            for issue_id, metadata in issue_blocks:
                issue_status = metadata.get("status", "")
                if issue_status not in ("pending", "in_progress", "blocked", "complete"):
                    errors.append(
                        f"{rel}#{issue_id}: status must be pending, in_progress, blocked, or complete"
                    )
                guide_refs = parse_inline_list(metadata.get("related_guides", "[]"))
                guide_impact = metadata.get("guide_impact", "")
                validate_guide_impact(
                    f"{rel}#{issue_id}",
                    guide_impact,
                    guide_refs,
                    metadata.get("guide_impact_reason", ""),
                    errors,
                )
                if guide_impact == "required" and issue_status == "complete":
                    validate_guide_sources(
                        f"{rel}#{issue_id}",
                        as_text(fm.get("id", "")),
                        "source_workstreams",
                        guide_refs,
                        id_locations,
                        root,
                        errors,
                    )

    guides_dir = root / "docs/guides"
    if guides_dir.is_dir():
        for path in sorted(guides_dir.glob("*.md")):
            rel = str(path.relative_to(root))
            fm = parse_front_matter(path)
            if fm is None:
                errors.append(f"{rel}: broken front matter")
                continue
            for key in ("source_issues", "source_workstreams"):
                for reference in as_list(fm.get(key, [])):
                    ref_path = root / reference
                    if reference in id_locations or ref_path.is_file():
                        continue
                    warnings.append(f"{rel}: {key} ref not found: {reference}")

    return errors, warnings


def main() -> None:
    parser = argparse.ArgumentParser(description="Validate origin-doc-update structure.")
    parser.add_argument("repo", nargs="?", default=".", help="Repository root (default: cwd)")
    args = parser.parse_args()

    # Say which repository was read. The default is the current directory, and a
    # closeout reached through an orchestrator stands in a *different* repository —
    # so the orchestrator's own docs could validate clean and be reported as the
    # target's result, with nothing to tell the two apart.
    print(f"validated: {Path(args.repo).resolve()}")
    errors, warnings = validate_repo(args.repo)
    if errors:
        print("ERRORS:")
        for error in errors:
            print(f"  ✗ {error}")
    if warnings:
        print("WARNINGS:")
        for warning in warnings:
            print(f"  ⚠ {warning}")
    if not errors and not warnings:
        print("OK: all checks passed.")
    sys.exit(1 if errors else 0)


if __name__ == "__main__":
    main()
