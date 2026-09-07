#!/usr/bin/env python3
"""Deterministic preflight, validation, and record backfill for project notes.

The script deliberately does not create or rewrite a project note.  Semantic project
content is reviewed separately; this helper owns the mechanical record/list contract.
"""

from __future__ import annotations

import argparse
import os
import re
import stat
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Sequence

try:
    import yaml
except ImportError as exc:  # pragma: no cover - environment failure
    raise SystemExit("PyYAML is required to validate project frontmatter") from exc


class ProjectUpdateError(Exception):
    """Base error with a user-facing, safe message."""


class ConflictError(ProjectUpdateError):
    """An existing association or file prevents a safe operation."""


class ValidationError(ProjectUpdateError):
    """The selected files do not satisfy the project contract."""


_BACKLINK_RE = re.compile(
    r"^>\s*project:\s*\[project_(?P<label>[^\]]+)\]\(\.\./project/project_(?P<link>[^)]+)\.md\)\s*$"
)
_PROJECT_LINE_RE = re.compile(r"^\s*project\s*:")
_LOCAL_PATH_RE = re.compile(r"(?:^|[\s(])/(?:Users|private|Volumes)/|file://")
_REQUIRED_HEADINGS = (
    "overview",
    "people",
    "stakeholders",
    "next actions",
    "open issues / risks",
    "minutes",
    "related notes",
    "documents",
)


@dataclass(frozen=True)
class Note:
    path: Path
    text: str
    lines: tuple[str, ...]
    close_index: int
    data: dict
    newline: str

    @property
    def body(self) -> str:
        return "".join(self.lines[self.close_index + 1 :])


@dataclass(frozen=True)
class Change:
    path: Path
    before: str
    after: str


def _parse_note_text(path: Path, text: str) -> Note:
    lines = tuple(text.splitlines(keepends=True))
    if not lines or lines[0].strip() != "---":
        raise ValidationError(f"{path.name}: YAML frontmatter opening delimiter is missing")

    close_index = next(
        (index for index in range(1, len(lines)) if lines[index].strip() == "---"),
        None,
    )
    if close_index is None:
        raise ValidationError(f"{path.name}: YAML frontmatter closing delimiter is missing")

    yaml_text = "".join(lines[1:close_index])
    try:
        data = yaml.safe_load(yaml_text) or {}
    except yaml.YAMLError as exc:
        raise ValidationError(f"{path.name}: YAML parse failed: {exc}") from exc
    if not isinstance(data, dict):
        raise ValidationError(f"{path.name}: YAML frontmatter must be a mapping")

    newline = "\r\n" if "\r\n" in lines[0] else "\n"
    return Note(path, text, lines, close_index, data, newline)


def parse_note(path: Path) -> Note:
    try:
        text = path.read_text(encoding="utf-8")
    except OSError as exc:
        raise ValidationError(f"{path.name}: cannot read file: {exc}") from exc
    return _parse_note_text(path, text)


def _safe_key(value: str) -> str:
    key = value.strip()
    if not key or key in {".", ".."} or "/" in key or "\\" in key or "\x00" in key:
        raise ValidationError("project key must be non-empty and cannot contain path separators")
    if any(ord(char) < 0x20 for char in key):
        raise ValidationError("project key cannot contain control characters")
    if any(char in key for char in "[]()"):
        raise ValidationError("project key cannot contain Markdown link delimiters")
    return key


def _yaml_scalar(value: str) -> str:
    dumped = yaml.safe_dump(value, allow_unicode=True, default_flow_style=True)
    return dumped.splitlines()[0]


def _project_value(note: Note) -> str | None:
    value = note.data.get("project")
    if value is None or str(value).strip() == "":
        return None
    return str(value).strip()


def _backlink_names(text: str) -> tuple[list[str], bool]:
    names: list[str] = []
    malformed = False
    for line in text.splitlines():
        if not line.lstrip().startswith("> project:"):
            continue
        match = _BACKLINK_RE.match(line)
        if match is None or match.group("label") != match.group("link"):
            malformed = True
            continue
        names.append(match.group("label"))
    return names, malformed


def _assert_no_other_project(note: Note, key: str) -> list[str]:
    value = _project_value(note)
    if value is not None and value != key:
        raise ConflictError(f"{note.path.name}: existing YAML project={value!r} differs from {key!r}")

    names, malformed = _backlink_names(note.text)
    if malformed:
        raise ConflictError(f"{note.path.name}: an existing project backlink is malformed")
    other_names = sorted(set(name for name in names if name != key))
    if other_names:
        raise ConflictError(
            f"{note.path.name}: existing project backlink points to {', '.join(other_names)!r}"
        )
    return names


def _add_project_key(note: Note, key: str) -> str:
    if _project_value(note) == key:
        return note.text

    lines = note.lines
    fm_lines = list(lines[1 : note.close_index])
    scalar = _yaml_scalar(key)
    for index, line in enumerate(fm_lines):
        if _PROJECT_LINE_RE.match(line):
            fm_lines[index] = f"project: {scalar}{note.newline}"
            yaml_text = "".join(fm_lines)
            return lines[0] + yaml_text + lines[note.close_index] + "".join(lines[note.close_index + 1 :])

    yaml_text = "".join(fm_lines)
    if yaml_text and not yaml_text.endswith(note.newline):
        yaml_text += note.newline
    yaml_text += f"project: {scalar}{note.newline}"
    return lines[0] + yaml_text + lines[note.close_index] + "".join(lines[note.close_index + 1 :])


def _add_record_backlink(note: Note, key: str) -> str:
    expected = f"> project: [project_{key}](../project/project_{key}.md)"
    names, malformed = _backlink_names(note.text)
    if malformed:
        raise ConflictError(f"{note.path.name}: an existing project backlink is malformed")
    if any(name != key for name in names):
        raise ConflictError(f"{note.path.name}: another project backlink already exists")
    if names.count(key) > 1:
        raise ValidationError(f"{note.path.name}: duplicate project backlinks")
    if names:
        body_lines = note.body.splitlines()
        first_nonempty = next((line for line in body_lines if line.strip()), None)
        if first_nonempty != expected:
            raise ValidationError(f"{note.path.name}: project backlink is not immediately after frontmatter")
        return note.text

    body = note.body
    if body and not body.startswith(note.newline):
        body = note.newline + body
    closing = note.lines[note.close_index]
    if not closing.endswith(note.newline):
        closing += note.newline
    return (
        "".join(note.lines[: note.close_index])
        + closing
        + expected
        + note.newline
        + body
    )


def prepare_record(path: Path, key: str) -> Change:
    note = parse_note(path)
    _assert_no_other_project(note, key)
    text = _add_project_key(note, key)
    text = _add_record_backlink(_parse_note_text(path, text), key)
    if text == note.text:
        return Change(path, note.text, note.text)
    return Change(path, note.text, text)


def _resolve_under(root: Path, value: str, label: str) -> Path:
    candidate = Path(value)
    if not candidate.is_absolute():
        candidate = root / candidate
    if candidate.is_symlink():
        raise ValidationError(f"{label} must not be a symlink")
    try:
        resolved = candidate.resolve()
        resolved.relative_to(root.resolve())
    except ValueError as exc:
        raise ValidationError(f"{label} must be inside the selected vault") from exc
    return resolved


def record_paths(vault: Path, values: Sequence[str]) -> list[Path]:
    if not values:
        raise ValidationError("at least one --record path is required")
    record_root = (vault / "record").resolve()
    paths: list[Path] = []
    seen: set[Path] = set()
    for value in values:
        path = _resolve_under(vault, value, "record path")
        try:
            path.relative_to(record_root)
        except ValueError as exc:
            raise ValidationError("record paths must be inside vault/record") from exc
        if path.suffix.lower() != ".md":
            raise ValidationError(f"{path.name}: record path must be a Markdown file")
        if not path.is_file():
            raise ValidationError(f"{path.name}: record file does not exist")
        if path not in seen:
            paths.append(path)
            seen.add(path)
    return paths


def _table_cells(line: str) -> list[str] | None:
    stripped = line.strip()
    if not stripped.startswith("|") or not stripped.endswith("|"):
        return None
    return [cell.strip().replace("\\_", "_") for cell in stripped[1:-1].split("|")]


def _list_table(text: str) -> tuple[dict[str, int], list[tuple[int, list[str]]]]:
    lines = text.splitlines()
    header_index = None
    headers: list[str] = []
    for index, line in enumerate(lines):
        cells = _table_cells(line)
        if cells and "name" in cells and "client" in cells:
            header_index = index
            headers = cells
            break
    if header_index is None:
        raise ValidationError("list_project.md: project table header is missing")
    header_map = {name: index for index, name in enumerate(headers)}
    rows: list[tuple[int, list[str]]] = []
    for index in range(header_index + 2, len(lines)):
        cells = _table_cells(lines[index])
        if cells is None or not cells or all(set(cell) <= {"-", " "} for cell in cells):
            continue
        if len(cells) > header_map["name"]:
            rows.append((index, cells))
    return header_map, rows


def list_rows(vault: Path, key: str) -> tuple[list[str], str | None]:
    path = vault / "setting" / "list" / "list_project.md"
    if not path.is_file():
        raise ValidationError("setting/list/list_project.md does not exist")
    if path.is_symlink():
        raise ValidationError("setting/list/list_project.md must not be a symlink")
    text = path.read_text(encoding="utf-8")
    header_map, rows = _list_table(text)
    names: list[str] = []
    matching: list[list[str]] = []
    for _, cells in rows:
        name = cells[header_map["name"]].strip()
        names.append(name)
        if name == key:
            matching.append(cells)
    if len(matching) > 1:
        raise ValidationError(f"list_project.md: project {key!r} appears more than once")
    client = None
    if matching:
        cells = matching[0]
        client_index = header_map.get("client")
        if client_index is not None and len(cells) > client_index:
            client = cells[client_index].strip()
    return names, client


def _check_local_paths(path: Path, text: str) -> None:
    if _LOCAL_PATH_RE.search(text):
        raise ValidationError(f"{path.name}: local absolute path or file:// URL detected")


def _validate_project_note(path: Path, key: str) -> Note:
    if path.is_symlink():
        raise ValidationError(f"{path.name}: project note must not be a symlink")
    note = parse_note(path)
    value = _project_value(note)
    if value != key:
        raise ValidationError(f"{path.name}: YAML project must be {key!r}")
    if not str(note.data.get("client", "")).strip():
        raise ValidationError(f"{path.name}: YAML client is required")
    headings = {match.group(1).strip() for match in re.finditer(r"^##\s+(.+?)\s*$", note.text, re.MULTILINE)}
    missing = [heading for heading in _REQUIRED_HEADINGS if heading not in headings]
    if missing:
        raise ValidationError(f"{path.name}: required headings missing: {', '.join(missing)}")
    _check_local_paths(path, note.text)
    return note


def _validate_record(path: Path, key: str) -> None:
    note = parse_note(path)
    names = _assert_no_other_project(note, key)
    if _project_value(note) != key:
        raise ValidationError(f"{path.name}: YAML project is missing or does not match {key!r}")
    expected = f"> project: [project_{key}](../project/project_{key}.md)"
    if names != [key]:
        raise ValidationError(f"{path.name}: exactly one project backlink is required")
    first_nonempty = next((line for line in note.body.splitlines() if line.strip()), None)
    if first_nonempty != expected:
        raise ValidationError(f"{path.name}: project backlink is not immediately after frontmatter")
    _check_local_paths(path, note.text)


def _print_path(vault: Path, path: Path) -> str:
    try:
        return str(path.relative_to(vault))
    except ValueError:
        return path.name


def _preflight(args: argparse.Namespace, paths: list[Path]) -> tuple[Path, list[Change]]:
    key = _safe_key(args.project_key)
    project_path = vault_project_path(args.vault, key)
    list_names, list_client = list_rows(args.vault, key)

    if args.mode == "bootstrap":
        if project_path.exists():
            raise ConflictError(f"project note already exists: {project_path.name}")
        if key in list_names:
            raise ConflictError(f"list_project.md already contains project {key!r}")
        for path in paths:
            note = parse_note(path)
            names = _assert_no_other_project(note, key)
            if _project_value(note) == key or key in names:
                raise ConflictError(
                    f"{path.name}: project {key!r} is already present without a project note"
                )
            _check_local_paths(path, note.text)
        print(f"PREFLIGHT_OK bootstrap project={key} records={len(paths)}")
        return project_path, []

    if not project_path.is_file():
        raise ValidationError(f"project note does not exist: {project_path.name}")
    project_note = _validate_project_note(project_path, key)
    project_client = str(project_note.data.get("client", "")).strip()
    if list_client is None:
        raise ConflictError(f"list_project.md has no row for project {key!r}")
    if list_client != project_client and not args.allow_list_client_mismatch:
        raise ConflictError(
            f"list_project.md client {list_client!r} differs from project note client {project_client!r}; "
            "confirm explicitly before continuing"
        )

    changes: list[Change] = []
    for path in paths:
        if args.mode == "validate":
            _validate_record(path, key)
        else:
            changes.append(prepare_record(path, key))
    if args.mode == "validate":
        print(f"VALID_OK project={key} records={len(paths)}")
    return project_path, changes


def vault_project_path(vault: Path, key: str) -> Path:
    return vault / "project" / f"project_{key}.md"


def _write_batch(vault: Path, changes: Iterable[Change]) -> None:
    changes = [change for change in changes if change.before != change.after]
    if not changes:
        return
    originals: dict[Path, bytes] = {}
    modes: dict[Path, int] = {}
    temp_paths: list[Path] = []
    applied: list[Path] = []
    try:
        for change in changes:
            originals[change.path] = change.path.read_bytes()
            modes[change.path] = stat.S_IMODE(change.path.stat().st_mode)
        for change in changes:
            with tempfile.NamedTemporaryFile(
                mode="wb", dir=change.path.parent, prefix=f".{change.path.name}.project-update-", delete=False
            ) as handle:
                handle.write(change.after.encode("utf-8"))
                temp_path = Path(handle.name)
            temp_paths.append(temp_path)
            os.chmod(temp_path, modes[change.path])
            os.replace(temp_path, change.path)
            applied.append(change.path)
    except Exception:
        for path in applied:
            path.write_bytes(originals[path])
            os.chmod(path, modes[path])
        raise
    finally:
        for temp_path in temp_paths:
            try:
                temp_path.unlink()
            except FileNotFoundError:
                pass


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--vault", type=Path, required=True, help="Obsidian data vault root")
    parser.add_argument("--project-key", required=True, help="Canonical project key")
    parser.add_argument(
        "--record", action="append", default=[], help="Record path relative to vault or absolute inside vault/record"
    )
    parser.add_argument(
        "--mode", choices=("bootstrap", "backfill", "validate"), default="bootstrap"
    )
    parser.add_argument(
        "--apply", action="store_true", help="Apply backfill changes; default is a dry-run"
    )
    parser.add_argument(
        "--allow-list-client-mismatch",
        action="store_true",
        help="Continue after an explicitly reviewed list/project client mismatch",
    )
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        args.vault = args.vault.expanduser().resolve()
        if not args.vault.is_dir():
            raise ValidationError("selected vault directory does not exist")
        paths = record_paths(args.vault, args.record)
        _, changes = _preflight(args, paths)
        if args.mode != "backfill":
            return 0

        changed = [change for change in changes if change.before != change.after]
        if not changed:
            print("NOOP no record changes required")
            return 0
        for change in changed:
            print(f"CHANGE {_print_path(args.vault, change.path)}")
        if not args.apply:
            print("DRY_RUN no files written (pass --apply after review)")
            return 0
        _write_batch(args.vault, changed)
        print(f"APPLIED records={len(changed)}")
        return 0
    except ConflictError as exc:
        print(f"CONFLICT {exc}", file=sys.stderr)
        return 2
    except (ProjectUpdateError, OSError) as exc:
        print(f"ERROR {exc}", file=sys.stderr)
        return 3


if __name__ == "__main__":
    raise SystemExit(main())
