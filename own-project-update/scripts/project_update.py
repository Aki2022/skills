#!/usr/bin/env python3
"""Deterministic preflight, validation, and record backfill for project notes.

The script deliberately does not create or rewrite the narrative of a project note.  Semantic
project content is reviewed separately; this helper owns the mechanical record/list contract.

Two subcommands extend it for SPEC-document-publish / SPEC-project-association:

* ``render``  rewrites only the ``generated:associated-notes`` region of project notes.
* ``attach-document``  publishes a generated document's text as a document note, associates it
  with projects through frontmatter, and renders the affected project notes in the same batch.

The original flat ``--mode bootstrap|backfill|validate`` interface is unchanged.
"""

from __future__ import annotations

import argparse
import datetime
import difflib
import filecmp
import hashlib
import html
import json
import os
import re
import stat
import subprocess
import sys
import tempfile
import unicodedata
import urllib.parse
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
    created: bool = False
    # True when `before` was read byte-exact: the write refuses if the file changed since planning.
    verify_before: bool = False


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


def _read_exact(path: Path) -> str:
    """Read a file without newline translation, so CRLF notes survive a rewrite."""
    with path.open(encoding="utf-8", newline="") as handle:
        return handle.read()


def parse_note(path: Path, exact: bool = False) -> Note:
    try:
        text = _read_exact(path) if exact else path.read_text(encoding="utf-8")
    except (OSError, UnicodeDecodeError) as exc:
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
    # width を広げないと長い値が折り返され、先頭行だけを使うここで値が欠ける。
    dumped = yaml.safe_dump(value, allow_unicode=True, default_flow_style=True, width=10**6)
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
    # 書く経路の検査は読む経路（_validate_record）と同じ強さにする。
    # 2026-09-29 以前は backfill だけがこの検査を通らず、validate が拒否する入力を
    # 受け入れて書き込んでいた。読むモードより書くモードが緩いのは「成功に見えて誤る」側。
    _check_local_paths(path, note.text)
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
    changes = [change for change in changes if change.before != change.after or change.created]
    if not changes:
        return
    originals: dict[Path, bytes | None] = {}
    modes: dict[Path, int] = {}
    temp_paths: list[Path] = []
    applied: list[Path] = []
    try:
        for change in changes:
            if change.created:
                if change.path.exists() or change.path.is_symlink():
                    raise ConflictError(f"{change.path.name}: file appeared before it could be created")
                originals[change.path] = None
                modes[change.path] = 0o644
            else:
                originals[change.path] = change.path.read_bytes()
                modes[change.path] = stat.S_IMODE(change.path.stat().st_mode)
                if change.verify_before and originals[change.path] != change.before.encode("utf-8"):
                    raise ConflictError(
                        f"{change.path.name}: file changed since it was read; nothing was written, rerun"
                    )
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
            original = originals[path]
            if original is None:
                path.unlink(missing_ok=True)
            else:
                path.write_bytes(original)
                os.chmod(path, modes[path])
        raise
    finally:
        for temp_path in temp_paths:
            try:
                temp_path.unlink()
            except FileNotFoundError:
                pass


# ---------------------------------------------------------------------------
# render / attach-document（SPEC-project-association・SPEC-document-publish）
# ---------------------------------------------------------------------------

_BEGIN_MARK = "<!-- generated:associated-notes begin -->"
_END_MARK = "<!-- generated:associated-notes end -->"
_REGION_NOTE = "<!-- render が毎回書き直す領域。手で編集しない（SPEC-project-association） -->"
_EMPTY_REGION_LINE = "_関連ノートなし_"
# render が関連ノートを集めない最上位ディレクトリ。project は project ノート自身
# （自分の key を `project:` に持つ）、setting は一覧・テンプレートで、どちらも関連ノートではない。
_RENDER_SKIP_DIRS = frozenset({"project", "setting"})
# document note の kind にできない名前（議事録・project・設定の置き場）。
_RESERVED_KINDS = frozenset({"project", "setting", "record"})
_KIND_RE = re.compile(r"^[a-z][a-z0-9_-]*$")
_REPO_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._-]*$")
_ISO_DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
_DATE_PREFIX_RE = re.compile(r"^(\d{4})-?(\d{2})-?(\d{2})(?!\d)")
# 取り込み経路（document note の素材・title・hint、render が読む summary）用。legacy の
# _LOCAL_PATH_RE は行頭・空白・`(` の直後しか見ないので、クォート直後・~/・ドライブ文字・
# クラウド同期フォルダの実体パスを通してしまう。公開 vault に載る入口なので広く拒否する。
# 入口の best-effort 検査。確定判定は vibe-guard の scan-text（pre-commit・夜間 vault_doctor）が持つ。
# 目的は「普通に起きる形」を vault に書く前に止めること: ユーザー名を含むホーム配下、マウント接頭辞つき、
# クラウド同期フォルダ、Windows/UNC/WSL、環境変数、URL エンコード・HTML 実体・JSON エスケープ・改行分断。
_NOT_AFTER_WORD = r"(?<![A-Za-z0-9_])"
_INGEST_LOCAL_RE = re.compile(
    # macOS: /Users/<name> は大文字小文字を区別し、どの前置でも拾う（/System/Volumes/Data/Users/…・/cygdrive/c/Users/…）。
    # 小文字の /users/ は REST ルートなので拾わない。/Users/{id} のようなプレースホルダと /Users/Shared は通す。
    r"/Users/(?!Shared\b)[^/\s{}<>)\]\"'`:]+|"
    # Linux: /home/<name>。英数字の直後（URL の path）は除く。automount 系は前置つきでも拾う。
    rf"(?<![A-Za-z0-9_./\-])/home/[^/\s{{}}<>)\]\"'`:]+|/(?:export|net/[^/\s]+)/home/|@[\w.-]+/home/|/volume\d+/homes?/|"
    r"(?<![A-Za-z0-9_./\-])/(?:root|Volumes)/\S|(?<![A-Za-z0-9_./\-])/mnt/[a-z]/|"
    r"(?<![A-Za-z0-9_./\-])/(?:private/(?:var|tmp|etc)|var/folders)/|"
    # ホーム直下: ~/ は ~/.config|.local|.cache 以外を拒否、~user/ も拒否。
    r"(?<![A-Za-z0-9_])~/(?!\.(?:config|local|cache)\b)[A-Za-z.]|(?<![A-Za-z0-9_~])~[a-z][\w.-]*/|"
    r"\$\{?HOME\}?\b|%(?:USERPROFILE|APPDATA|LOCALAPPDATA|HOMEPATH)%|"
    # file: / smb: / UNC / WSL / Windows のドライブ付き
    rf"{_NOT_AFTER_WORD}file:/|smb://|\\\\[A-Za-z0-9.$_-]+\\[A-Za-z0-9$_]|"
    r"(?<![A-Za-z0-9_])[A-Za-z]:(?:\\[^\\/\n:]+[\\/]|[\\/](?i:Users|Windows)[\\/])|"
    r"(?i:Library/CloudStorage|Library/Mobile Documents|com~apple~CloudDocs|/My Drive/|マイドライブ/)|"
    rf"{_NOT_AFTER_WORD}GoogleDrive-[^/\s]*@|OneDrive - |共有ドライブ/|Google ドライブ/|iCloud Drive/|"
    r"Dropbox \(|(?<![A-Za-z0-9_])CloudStorage/|sharepoint\.com/personal/"
)
# [^\S\n] は改行以外の空白。\s*\n\s* の入れ子は、空白だけの行が続くと 3 乗時間になる。
_LINEBREAK_AT_SLASH_RE = re.compile(r"[^\S\n]*\n[^\S\n]*(?=[/\\])|(?<=[/\\])[^\S\n]*\n[^\S\n]*")


def _ingest_variants(text: str) -> list[str]:
    """The text and its decoded/normalised forms, so an encoded or wrapped path is still seen."""
    base = unicodedata.normalize("NFKC", text)
    forms = [text, base]
    current = base
    for _ in range(2):
        current = html.unescape(urllib.parse.unquote(current))
        forms.append(current)
    out: list[str] = []
    for form in forms:
        form = form.replace("\\/", "/")
        out.append(form)
        out.append(form.replace("\\\\", "\\"))  # JSON・コード中の二重バックスラッシュ
        out.append(re.sub(r"/{2,}", "/", _LINEBREAK_AT_SLASH_RE.sub("", form)))
    return out


def _has_local_path(text: str) -> bool:
    return any(_INGEST_LOCAL_RE.search(form) for form in _ingest_variants(text))


_FM_KEY_RE = re.compile(r"^([A-Za-z_][\w-]*)\s*:")
_BULLET_RE = re.compile(r"^\s*(?:[-*+]\s+|#+\s+|\d+[.)]\s+)")
_HINT_RE = re.compile(r"^([A-Za-z_][\w-]*)=(.*)$")
_ARTIFACT_LINE_RE = re.compile(r"^\[([^\]]+)\]\(<?\1>?\)$")


@dataclass(frozen=True)
class _Row:
    kind: str
    date: str
    title: str
    link: str
    summary: str


def _read_frontmatter(path: Path, text: str) -> Note | None:
    """Return the parsed note, or None when it has no frontmatter to speak of.

    A note that declares `project:` but cannot be read — YAML that does not parse, or a
    frontmatter that is never closed — would silently vanish from every project's view, so
    that case stops the whole run instead of being skipped.
    """
    text = text.lstrip("\ufeff")
    lines = text.splitlines(keepends=True)
    if not lines or lines[0].strip() != "---":
        return None
    close = next((i for i in range(1, len(lines)) if lines[i].strip() == "---"), None)
    if close is None:
        if any(_PROJECT_LINE_RE.match(line) and not line[:1].isspace() for line in lines[1:]):
            raise ValidationError(f"{path.name}: frontmatter is never closed but declares project")
        return None
    try:
        return _parse_note_text(path, text)
    except ValidationError:
        raw = "".join(lines[1:close])
        if any(_PROJECT_LINE_RE.match(line) for line in raw.splitlines()):
            raise ValidationError(
                f"{path.name}: frontmatter does not parse but declares project; fix it before render"
            ) from None
        return None


def _projects_of(note: Note) -> list[str]:
    """Association keys of a note. A legacy scalar `project:` counts as a one-element list."""
    value = note.data.get("project")
    if value is None:
        return []
    items = value if isinstance(value, list) else [value]
    keys: list[str] = []
    for item in items:
        if not isinstance(item, str):
            raise ValidationError(f"{note.path.name}: project must be a string or a list of strings")
        key = item.strip()
        if key and key not in keys:
            keys.append(key)
    return keys


_LOOSE_DATE_RE = re.compile(r"^(\d{4})[-/.]?(\d{1,2})[-/.]?(\d{1,2})$")


def _fmt_date(value: object) -> str:
    if isinstance(value, (datetime.datetime, datetime.date)):
        return value.isoformat()[:10]
    if isinstance(value, (str, int)) and not isinstance(value, bool):
        text = str(value).strip()
        match = _LOOSE_DATE_RE.match(text[:10]) if len(text) <= 10 else None
        if match is None and _ISO_DATE_RE.match(text[:10]):
            return text[:10]
        if match:
            try:
                return datetime.date(int(match.group(1)), int(match.group(2)), int(match.group(3))).isoformat()
            except ValueError:
                return text
        return text
    return ""


def _note_date(note: Note) -> str:
    for field in ("date", "created"):
        found = _fmt_date(note.data.get(field))
        if found:
            return found
    match = _DATE_PREFIX_RE.match(note.path.name)
    return f"{match.group(1)}-{match.group(2)}-{match.group(3)}" if match else ""


def _cell(value: object) -> str:
    return " ".join(str(value).split()).replace("|", "\\|")


def _link_target(vault: Path, path: Path) -> str:
    rel = os.path.relpath(path, vault / "project").replace(os.sep, "/")
    return f"<{rel}>" if re.search(r"[\s()]", rel) else rel


def _collect_associations(vault: Path, overlay: dict[Path, str]) -> dict[str, list[tuple[_Row, str]]]:
    """Group every associated note by project key, reading each note once."""
    paths: list[Path] = []
    for dirpath, dirnames, filenames in os.walk(vault):
        top = Path(dirpath) == vault
        dirnames[:] = sorted(
            d for d in dirnames if not d.startswith(".") and not (top and d in _RENDER_SKIP_DIRS)
        )
        for filename in sorted(filenames):
            if filename.endswith(".md") and not filename.startswith("."):
                paths.append(Path(dirpath) / filename)
    known = set(paths)
    for extra in sorted(overlay):
        rel_parts = extra.relative_to(vault).parts
        if extra not in known and rel_parts[0] not in _RENDER_SKIP_DIRS and not rel_parts[0].startswith("."):
            paths.append(extra)

    grouped: dict[str, list[tuple[_Row, str]]] = {}
    for path in paths:
        if path.is_symlink():
            continue
        text = overlay.get(path)
        if text is None:
            try:
                text = _read_exact(path)
            except (OSError, UnicodeDecodeError) as exc:
                raise ValidationError(f"{path.name}: cannot read file: {exc}") from exc
        note = _read_frontmatter(path, text)
        if note is None:
            continue
        keys = _projects_of(note)
        if not keys:
            continue
        rel = path.relative_to(vault)
        kind = str(note.data.get("kind") or rel.parts[0]).strip()
        title = str(note.data.get("title") or path.stem).strip()
        summary = " ".join(str(note.data.get("summary") or "").split())
        for label, value in (("title", title), ("summary", summary)):
            if _has_local_path(value):
                raise ValidationError(f"{path.name}: {label} contains a local absolute path or file:// URL")
        row = _Row(kind, _note_date(note), title, _link_target(vault, path), summary)
        for key in keys:
            grouped.setdefault(key, []).append((row, rel.as_posix()))
    return grouped


def _region_lines(rows: list[_Row]) -> list[str]:
    lines = [_BEGIN_MARK, _REGION_NOTE, ""]
    if not rows:
        lines.append(_EMPTY_REGION_LINE)
    else:
        ordered = sorted(rows, key=lambda r: r.link)
        ordered.sort(key=lambda r: r.date, reverse=True)
        ordered.sort(key=lambda r: r.kind)
        lines += ["| kind | date | note | summary |", "| --- | --- | --- | --- |"]
        for row in ordered:
            title = _cell(row.title).replace("[", "\\[").replace("]", "\\]")
            lines.append(
                f"| {_cell(row.kind)} | {_cell(row.date)} | [{title}]({row.link}) | {_cell(row.summary)} |"
            )
    lines.append(_END_MARK)
    return lines


def _with_region(note: Note, region: list[str] | None) -> str:
    """Return the note text with the associated-notes region rewritten (or created)."""
    nl = note.newline
    begins = [i for i, line in enumerate(note.lines) if line.strip() == _BEGIN_MARK]
    ends = [i for i, line in enumerate(note.lines) if line.strip() == _END_MARK]
    if len(begins) != len(ends) or len(begins) > 1 or (begins and ends[0] < begins[0]):
        raise ValidationError(
            f"{note.path.name}: associated-notes markers are broken "
            f"(begin={len(begins)}, end={len(ends)}); fix them by hand before render"
        )
    if begins:
        if region is None:
            region = _region_lines([])
        replaced = (
            list(note.lines[: begins[0]])
            + [line + nl for line in region]
            + list(note.lines[ends[0] + 1 :])
        )
        return "".join(replaced)
    if region is None:
        return note.text
    text = note.text if note.text.endswith(nl) else note.text + nl
    return text + nl + nl.join(region) + nl


def _project_note_targets(vault: Path, keys: Sequence[str] | None) -> list[tuple[str, Path]]:
    if keys:
        targets = []
        for raw in dict.fromkeys(keys):
            key = _safe_key(raw)
            path = vault_project_path(vault, key)
            if not path.is_file() or path.is_symlink():
                raise ValidationError(f"project note does not exist: {path.name}")
            targets.append((key, path))
        return targets
    targets = []
    for path in sorted((vault / "project").glob("project_*.md")):
        if path.is_symlink() or not path.is_file():
            continue
        targets.append((path.name[len("project_") : -len(".md")], path))
    return targets


def plan_render(
    vault: Path, keys: Sequence[str] | None, overlay: dict[Path, str] | None = None
) -> tuple[list[Change], list[tuple[str, int, bool]], list[tuple[str, str]]]:
    """Compute the render result without writing anything.

    Returns (changes, per-project (key, notes, changed), unknown (key, note) references).
    All inputs are read before any output is decided, so a failure leaves nothing half-done.
    """
    overlay = overlay or {}
    grouped = _collect_associations(vault, overlay)
    targets = _project_note_targets(vault, keys)
    changes: list[Change] = []
    summary: list[tuple[str, int, bool]] = []
    for key, path in targets:
        note = parse_note(path, exact=True)
        if _project_value(note) != key:
            raise ValidationError(f"{path.name}: YAML project must be {key!r}")
        rows = [row for row, _ in grouped.get(key, [])]
        has_region = any(line.strip() == _BEGIN_MARK for line in note.lines)
        after = _with_region(note, _region_lines(rows) if (rows or has_region) else None)
        changed = after != note.text
        if changed:
            changes.append(Change(path, note.text, after, verify_before=True))
        summary.append((key, len(rows), changed))
    existing = {key for key, _ in _project_note_targets(vault, None)}
    unknown = sorted(
        (key, rel) for key, items in grouped.items() if key not in existing for _, rel in items
    )
    return changes, summary, unknown


def _report_render(
    summary: list[tuple[str, int, bool]], unknown: list[tuple[str, str]]
) -> None:
    for key, count, changed in summary:
        print(f"RENDER project={key} notes={count} {'CHANGE' if changed else 'NOOP'}")
    for key, rel in unknown:
        print(f"WARN unknown project key {key!r} referenced by {rel}", file=sys.stderr)


def _cmd_render(args: argparse.Namespace) -> int:
    changes, summary, unknown = plan_render(args.vault, args.project_key)
    _report_render(summary, unknown)
    if not changes:
        print("NOOP no render changes required")
        return 0
    if not args.apply:
        print("DRY_RUN no files written (pass --apply after review)")
        return 0
    _write_batch(args.vault, changes)
    print(f"APPLIED files={len(changes)}")
    return 0


# ---- document note ---------------------------------------------------------------


class _NotARepo(Exception):
    """`git` ran but the path is not inside a repository."""


def _git_output(args: list[str], cwd: Path) -> str:
    result = subprocess.run(
        ["git", *args], cwd=cwd, capture_output=True, text=True, timeout=15, check=False
    )
    if result.returncode != 0:
        raise _NotARepo(result.stderr.strip())
    return result.stdout


def _main_worktree(top: Path) -> Path | None:
    for line in _git_output(["worktree", "list", "--porcelain"], top).splitlines():
        if line.startswith("worktree "):
            return Path(line[len("worktree ") :])
    return None


def _resolve_artifact(artifact: Path) -> tuple[Path | None, str]:
    """Resolve the original deliverable to a path that outlives a task worktree.

    Returns (target, "") or (None, reason).  A deliverable inside a linked git worktree is
    mapped to the same relative path in the main checkout, because a worktree is removed
    after the work is merged and would leave the symlink dangling.
    """
    path = Path(os.path.abspath(artifact.expanduser()))
    if not path.is_file():
        return None, "artifact file does not exist"
    if path.suffix.lower() in {"", ".md"}:
        return None, "artifact must be a non-Markdown file"
    try:
        top = Path(_git_output(["rev-parse", "--show-toplevel"], path.parent).strip())
    except _NotARepo:
        return path, ""
    except (FileNotFoundError, subprocess.SubprocessError):
        return None, "git is not available to resolve the main checkout"
    try:
        main = _main_worktree(top)
    except (_NotARepo, FileNotFoundError, subprocess.SubprocessError):
        return None, "cannot list git worktrees to find the main checkout"
    if main is None:
        return None, "cannot find the main checkout"
    if top.resolve() == main.resolve():
        return path, ""
    try:
        relative = path.resolve().relative_to(top.resolve())
    except ValueError:
        return None, "artifact is outside its repository"
    candidate = main / relative
    if not candidate.is_file():
        return None, "artifact is not present in the main checkout"
    if not filecmp.cmp(path, candidate, shallow=False):
        return None, "artifact differs from the main checkout's copy (not merged yet); rerun after it lands"
    return candidate, ""


def _sha256(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def _normalize_content(text: str) -> str:
    return text.strip("\n") + "\n"


def _backlink_line(key: str) -> str:
    return f"> project: [project_{key}](../project/project_{key}.md)"


def _split_body(note: Note) -> tuple[list[str], str]:
    """Split a document note body into its leading backlink keys and the rest."""
    lines = note.body.splitlines()
    names: list[str] = []
    index = 0
    while index < len(lines):
        if not lines[index].strip():
            index += 1
            continue
        match = _BACKLINK_RE.match(lines[index])
        if match is None or match.group("label") != match.group("link"):
            break
        names.append(match.group("label"))
        index += 1
    return names, "\n".join(lines[index:])


def _fm_block(key: str, value: object) -> list[str]:
    if isinstance(value, list):
        return [f"{key}:\n"] + [f"  - {_yaml_scalar(str(item))}\n" for item in value]
    if key == "date":
        return [f"date: {value}\n"]  # bare ISO date so YAML reads it back as a date
    return [f"{key}: {_yaml_scalar(str(value))}\n"]


def _replace_fm_key(fm_lines: list[str], key: str, block: list[str] | None) -> list[str]:
    out: list[str] = []
    replaced = False
    index = 0
    while index < len(fm_lines):
        match = _FM_KEY_RE.match(fm_lines[index])
        if match and match.group(1) == key:
            end = index + 1
            while end < len(fm_lines) and (
                fm_lines[end][:1] in (" ", "\t") or fm_lines[end].startswith("-")
            ):
                end += 1
            if block is not None and not replaced:
                out += block
            replaced = True
            index = end
            continue
        out.append(fm_lines[index])
        index += 1
    if not replaced and block is not None:
        if out and not out[-1].endswith("\n"):
            out[-1] += "\n"
        out += block
    return out


def _compose(fm_lines: list[str], projects: list[str], content: str, nl: str = "\n") -> str:
    head = "---" + nl + "".join(line.rstrip("\r\n") + nl for line in fm_lines) + "---" + nl
    if projects:
        head += "".join(_backlink_line(key) + nl for key in projects) + nl
    return head + content.replace("\n", nl)


def _summary_line(digest: str) -> str:
    for line in digest.splitlines():
        stripped = _BULLET_RE.sub("", line).strip()
        if stripped:
            return " ".join(stripped.split())
    raise ValidationError("digest has no text to use as summary")


def _read_material(path: Path | None, label: str, required: bool) -> str:
    if path is None:
        if required:
            raise ValidationError(f"--{label}-file is required")
        return ""
    try:
        text = path.expanduser().read_text(encoding="utf-8")
    except (OSError, UnicodeDecodeError) as exc:
        raise ValidationError(f"{label}: cannot read {path.name}: {exc}") from exc
    if required and not text.strip():
        raise ValidationError(f"{label}: {path.name} is empty")
    if _has_local_path(text):
        raise ValidationError(f"{label}: local absolute path or file:// URL detected in {path.name}")
    return text


def _check_source_path(value: str) -> str:
    text = value.strip()
    parts = text.replace("\\", "/").split("/")
    if (
        not text
        or text.startswith(("/", "~"))
        or re.match(r"^[A-Za-z]:", text)
        or ".." in parts
        or "\\" in text
        or any(ord(char) < 0x20 for char in text)
        or _has_local_path(text)
    ):
        raise ValidationError("--source-path must be a repository-relative path without local paths")
    return text


def _document_date(name: str, explicit: str | None) -> str:
    if explicit:
        try:
            return datetime.date.fromisoformat(explicit).isoformat()
        except ValueError as exc:
            raise ValidationError("--date must be YYYY-MM-DD") from exc
    match = _DATE_PREFIX_RE.match(name)
    if match:
        try:
            return datetime.date(int(match.group(1)), int(match.group(2)), int(match.group(3))).isoformat()
        except ValueError:
            pass
    raise ValidationError("name has no yyyymmdd prefix; pass --date YYYY-MM-DD")


def _now() -> str:
    return datetime.datetime.now(datetime.timezone.utc).replace(microsecond=0).isoformat()


def _content_diff(old: str, new: str) -> str:
    diff = list(
        difflib.unified_diff(
            old.splitlines(), new.splitlines(), "vault (edited)", "regenerated", lineterm="", n=1
        )
    )
    shown = diff[:40]
    if len(diff) > len(shown):
        shown.append(f"... {len(diff) - len(shown)} more diff lines")
    return "\n".join(shown)


def _document_content(digest: str, outline: str, notes: str, artifact_name: str | None) -> str:
    parts: list[str] = []
    if artifact_name:
        parts.append(_artifact_line(artifact_name))
    parts.append("## digest\n\n" + digest.strip())
    if outline.strip():
        parts.append("## outline\n\n" + outline.strip())
    if notes.strip():
        parts.append("## speaker notes\n\n" + notes.strip())
    return "\n\n".join(parts) + "\n"


def _artifact_line(artifact_name: str) -> str:
    target = f"<{artifact_name}>" if re.search(r"[\s()]", artifact_name) else artifact_name
    return f"[{artifact_name}]({target})"


def _leading_artifact_lines(content: str) -> list[str]:
    found: list[str] = []
    for line in content.strip("\n").splitlines():
        if not line.strip():
            continue
        if not _ARTIFACT_LINE_RE.match(line):
            break
        found.append(line)
    return found


def _has_artifact_line(content: str, artifact_name: str) -> bool:
    return _leading_artifact_lines(content) == [_artifact_line(artifact_name)]


def _without_artifact_lines(content: str) -> str:
    drop = len(_leading_artifact_lines(content))
    lines = content.strip("\n").splitlines()
    seen = 0
    out: list[str] = []
    for line in lines:
        if seen < drop and _ARTIFACT_LINE_RE.match(line):
            seen += 1
            continue
        out.append(line)
    return "\n".join(out).lstrip("\n")


def _install_link(link: Path, target: str) -> None:
    temp_link = link.parent / f".{link.name}.project-update-link"
    try:
        temp_link.unlink(missing_ok=True)
        os.symlink(target, temp_link)
        os.replace(temp_link, link)
    finally:
        try:
            temp_link.unlink(missing_ok=True)
        except OSError:
            pass


def _projects_listing(vault: Path, note_path: Path) -> list[str]:
    """Projects whose generated region still links to this note (e.g. after a hand-removed association)."""
    needle = f"]({_link_target(vault, note_path)})"
    keys: list[str] = []
    for key, path in _project_note_targets(vault, None):
        try:
            if needle in _read_exact(path):
                keys.append(key)
        except (OSError, UnicodeDecodeError) as exc:
            raise ValidationError(f"{path.name}: cannot read file: {exc}") from exc
    return keys


def _cmd_attach_document(args: argparse.Namespace) -> int:
    vault: Path = args.vault
    kind = args.kind.strip()
    if not _KIND_RE.match(kind) or kind in _RESERVED_KINDS:
        raise ValidationError(f"kind {kind!r} is not allowed")
    kind_dir = vault / kind
    if kind_dir.is_symlink() or not kind_dir.is_dir():
        raise ValidationError(f"kind directory vault/{kind} does not exist; create it explicitly first")
    try:
        name = _safe_key(args.name)
    except ValidationError as exc:
        raise ValidationError(str(exc).replace("project key", "document name")) from None
    if name.startswith(".") or name.lower().endswith(".md") or any(char in name for char in "#%|"):
        raise ValidationError("document name must not start with '.', end with .md, or contain # % |")
    if not _REPO_RE.match(args.source_repo.strip()):
        raise ValidationError("--source-repo must be a repository name, not a path")
    source_repo = args.source_repo.strip()
    source_path = _check_source_path(args.source_path)
    date = _document_date(name, args.date)

    digest = _read_material(args.digest_file, "digest", required=True)
    outline = _read_material(args.outline_file, "outline", required=False)
    speaker_notes = _read_material(args.notes_file, "notes", required=False)
    title = (args.title or name).strip()
    for label, value in [("title", title)] + [("hint", h) for h in args.hint]:
        if _has_local_path(value) or "\n" in value:
            raise ValidationError(f"{label}: local path or multi-line value is not allowed")
    hints = []
    for raw in args.hint:
        match = _HINT_RE.match(raw)
        if match is None:
            raise ValidationError(f"--hint must be KEY=VALUE, got {raw!r}")
        hints.append(raw)
    summary = _summary_line(digest)

    manual = [_safe_key(key) for key in dict.fromkeys(args.project)]
    for key in manual:
        if not vault_project_path(vault, key).is_file():
            raise ValidationError(f"project note does not exist for {key!r}; refusing to associate")

    note_path = kind_dir / f"{name}.md"
    if note_path.is_symlink():
        raise ValidationError(f"{note_path.name}: document note must not be a symlink")
    existing = parse_note(note_path, exact=True) if note_path.exists() else None
    source_hash = _sha256(
        json.dumps(
            {"digest": digest.strip(), "outline": outline.strip(), "speaker_notes": speaker_notes.strip()},
            ensure_ascii=False,
            sort_keys=True,
        )
    )

    current: list[str] = []
    old_content = ""
    if existing is not None:
        for field, expected in (("kind", kind), ("source_repo", source_repo), ("source_path", source_path)):
            found = existing.data.get(field)
            if found is not None and str(found) != expected:
                raise ConflictError(
                    f"{note_path.name}: existing note has {field}={found!r}, not {expected!r}; "
                    "refusing to overwrite another document"
                )
        current = _projects_of(existing)
        _, old_content = _split_body(existing)
        stored = existing.data.get("content_hash")
        if not stored:
            raise ConflictError(
                f"{note_path.name}: no content_hash recorded, so human edits cannot be told apart; "
                "resolve by hand"
            )
        if _sha256(_normalize_content(old_content)) != str(stored):
            regenerated = _document_content(digest, outline, speaker_notes, None)
            raise ConflictError(
                f"{note_path.name}: the note was edited after publish; not overwriting "
                "(diff against what the current source would generate, artifact line omitted)\n"
                + _content_diff(old_content, regenerated)
            )

    projects = list(dict.fromkeys(current + manual))
    if existing is None and not projects and not args.no_project:
        undecided = True
    else:
        undecided = False

    # --- assess seam (SPEC-project-association layer B is not implemented yet) ---
    print(f"PREFLIGHT_OK attach-document kind={kind} name={name}")
    for hint in hints:
        print(f"HINT {hint}")
    if manual:
        print("ASSESS skipped (manual project given)")
    elif undecided or not projects:
        print("ASSESS unavailable (no proposals)")
    if undecided:
        message = "association undecided: pass --project KEY (repeatable) or --no-project"
        if args.apply:
            raise ConflictError(message)
        print(f"ASSOCIATION undecided ({message})")

    # --- artifact link ---
    link_path: Path | None = None
    link_target = ""
    link_status = ""
    artifact_name: str | None = None
    if args.artifact is not None:
        target, reason = _resolve_artifact(args.artifact)
        if target is None:
            print(f"ARTIFACT skipped ({reason})")
        else:
            candidate = kind_dir / f"{name}{target.suffix}"
            relative = os.path.relpath(target, kind_dir)
            if candidate.is_symlink():
                link_status = "same" if os.readlink(candidate) == relative else "replace"
            elif candidate.exists():
                print("ARTIFACT skipped (a regular file already exists at the artifact path)")
            else:
                link_status = "create"
            if link_status:
                link_path, link_target, artifact_name = candidate, relative, candidate.name
    if artifact_name is None and existing is not None and old_content.strip():
        # 再生成しても、まだ有効な artifact link 行は落とさない
        kept = _ARTIFACT_LINE_RE.match(old_content.strip().splitlines()[0])
        if kept and (kind_dir / kept.group(1)).is_symlink():
            artifact_name = kept.group(1)

    # --- create the symlink first: a failure then costs only the link, never a dangling body line ---
    previous_target: str | None = None
    link_installed = False
    if args.apply and link_path is not None and link_status in ("create", "replace"):
        try:
            if link_status == "replace":
                previous_target = os.readlink(link_path)
            _install_link(link_path, link_target)
            link_installed = True
            print(f"ARTIFACT linked {link_path.name} -> {link_target}")
        except OSError as exc:
            print(f"ARTIFACT skipped (could not create symlink: {exc})")
            link_path, link_status, artifact_name = None, "", None

    def undo_link() -> None:
        if not link_installed or link_path is None:
            return
        try:
            if previous_target is None:
                link_path.unlink(missing_ok=True)
            else:
                _install_link(link_path, previous_target)
        except OSError:
            pass

    # --- compose the note, render, and write one batch ---
    try:
        if existing is None:
            content = _document_content(digest, outline, speaker_notes, artifact_name)
            fm: list[str] = []
            fm += _fm_block("title", title) + _fm_block("kind", kind)
            if projects:
                fm += _fm_block("project", projects)
                fm += _fm_block("project_source", "manual")
            fm += _fm_block("date", date) + _fm_block("summary", summary)
            fm += _fm_block("source_repo", source_repo) + _fm_block("source_path", source_path)
            fm += _fm_block("source_hash", source_hash)
            fm += _fm_block("content_hash", _sha256(_normalize_content(content)))
            fm += _fm_block("published_at", _now())
            after = _compose(fm, projects, content)
            action = "CREATE"
        else:
            same_source = str(existing.data.get("source_hash")) == source_hash
            fm_lines = list(existing.lines[1 : existing.close_index])
            content = old_content
            if not same_source:
                content = _document_content(digest, outline, speaker_notes, artifact_name)
            elif artifact_name and not _has_artifact_line(old_content, artifact_name):
                content = _artifact_line(artifact_name) + "\n\n" + _without_artifact_lines(old_content)
            if content is not old_content:
                updates = [("content_hash", _sha256(_normalize_content(content)))]
                if not same_source:
                    updates = [
                        ("summary", summary),
                        ("source_hash", source_hash),
                        *updates,
                        ("published_at", _now()),
                    ]
                for key, value in updates:
                    fm_lines = _replace_fm_key(fm_lines, key, _fm_block(key, value))
            if projects != current:
                fm_lines = _replace_fm_key(fm_lines, "project", _fm_block("project", projects))
            if manual:
                fm_lines = _replace_fm_key(fm_lines, "project_source", _fm_block("project_source", "manual"))
            after = _compose(fm_lines, projects, _normalize_content(content), existing.newline)
            action = "NOOP" if after == existing.text else "UPDATE"

        changes: list[Change] = []
        overlay: dict[Path, str] = {}
        if action == "CREATE":
            changes.append(Change(note_path, "", after, created=True))
            overlay[note_path] = after
        elif action == "UPDATE":
            assert existing is not None
            changes.append(Change(note_path, existing.text, after, verify_before=True))
            overlay[note_path] = after
        print(f"DOCUMENT {action} {note_path.relative_to(vault).as_posix()}")

        if link_path is not None and not link_installed:
            verb = {"create": "would link", "replace": "would relink", "same": "unchanged"}[link_status]
            print(f"ARTIFACT {verb} {link_path.name} -> {link_target}")

        # 紐付け先と、まだこのノートを載せている project（手で外した所属）を同じ batch で render する
        render_keys = list(dict.fromkeys(projects + _projects_listing(vault, note_path)))
        if render_keys:
            render_changes, summary_rows, unknown = plan_render(vault, render_keys, overlay)
            _report_render(summary_rows, unknown)
            changes += render_changes

        if not changes and link_status in ("", "same"):
            print("NOOP no changes required")
            return 0
        if not args.apply:
            print("DRY_RUN no files written (pass --apply after review)")
            return 0
        _write_batch(vault, changes)
    except BaseException:
        undo_link()
        raise
    print(f"APPLIED files={len(changes)}")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=__doc__,
        epilog="Subcommands: `render` and `attach-document` (first argument; see `<subcommand> --help`).",
    )
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


def build_subcommand_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="project_update.py", description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)

    render = sub.add_parser("render", help="rewrite the generated associated-notes region of project notes")
    render.add_argument("--vault", type=Path, required=True, help="Obsidian data vault root")
    render.add_argument(
        "--project-key", action="append", default=[], help="Limit to this project (repeatable); default all"
    )
    render.add_argument("--apply", action="store_true", help="Write changes; default is a dry-run")

    attach = sub.add_parser("attach-document", help="publish a generated document as a document note")
    attach.add_argument("--vault", type=Path, required=True, help="Obsidian data vault root")
    attach.add_argument("--kind", required=True, help="Document kind = vault directory name (e.g. presentation)")
    attach.add_argument("--name", required=True, help="Document name = note file stem (yyyymmdd_...)")
    attach.add_argument("--title", help="Note title (default: name)")
    attach.add_argument("--date", help="YYYY-MM-DD (default: yyyymmdd prefix of the name)")
    attach.add_argument("--source-repo", required=True, help="Repository that owns the source")
    attach.add_argument("--source-path", required=True, help="Repository-relative path of the source")
    attach.add_argument("--digest-file", type=Path, required=True, help="Key-message digest (first line becomes summary)")
    attach.add_argument("--outline-file", type=Path, help="Outline text")
    attach.add_argument("--notes-file", type=Path, help="Speaker notes text")
    attach.add_argument("--hint", action="append", default=[], help="KEY=VALUE hint forwarded to assess (repeatable)")
    group = attach.add_mutually_exclusive_group()
    group.add_argument("--project", action="append", default=[], help="Confirmed project key (repeatable)")
    group.add_argument("--no-project", action="store_true", help="Publish without associating a project")
    attach.add_argument("--artifact", type=Path, help="Original deliverable (e.g. pptx) to link next to the note")
    attach.add_argument("--apply", action="store_true", help="Write changes; default is a dry-run")
    return parser


_SUBCOMMANDS = {"render": _cmd_render, "attach-document": _cmd_attach_document}


def _run_guarded(action) -> int:
    try:
        return action()
    except ConflictError as exc:
        print(f"CONFLICT {exc}", file=sys.stderr)
        return 2
    except (ProjectUpdateError, OSError) as exc:
        print(f"ERROR {exc}", file=sys.stderr)
        return 3


def _main_subcommand(argv: Sequence[str]) -> int:
    args = build_subcommand_parser().parse_args(argv)

    def action() -> int:
        args.vault = args.vault.expanduser().resolve()
        if not args.vault.is_dir():
            raise ValidationError("selected vault directory does not exist")
        return _SUBCOMMANDS[args.command](args)

    return _run_guarded(action)


def main(argv: Sequence[str] | None = None) -> int:
    raw = list(sys.argv[1:] if argv is None else argv)
    if raw and raw[0] in _SUBCOMMANDS:
        return _main_subcommand(raw)
    args = build_parser().parse_args(raw)
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
