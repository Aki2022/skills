"""migrate: 既存 vault を新契約へ 1 回で移す（dry-run → 差分 → 人間 → 適用）。

SPEC-project-association の「配置と移行」。夜間 job では実行しない。実行主体は人間が見たプランだけ:
`--apply` は、dry-run が出した PLAN_ID を `--plan-id` で渡したときだけ書く（見た差分と書く内容を一致させる）。

やること（すべて決定論。LLM を使わない）:
- record: `project` を常にリストにし、`project_source: legacy` を付ける（既に出所があれば保つ）。vault 内のどの `.md` でも、
  project を持つものは対象（record/ 直下だけではない）。project ノートの旧 3 節（minutes / related notes / documents）に載っている
  record・資料は、その project に属する証拠として `legacy` で紐付ける（既存の所属は保持して追加）。minutes の topics は、
  record に summary が無いときだけ `summary` へ移す。summary にならなかった topics・余分な列・record と食い違う日付の行は、行ごと残す。
- project ノート: 旧 3 節（見出しの括弧書き・`###` も含む）を取り除き、表現できない内容（引用メモ・外部リンク・related notes の説明文・
  実体の無い record の行・人が書いた HTML コメント・コードブロック）は `## migrated notes` に原文のまま残す。generated 領域は render が作る。
  `list_project.md` の partner / status が frontmatter に無い（または空）なら移す。
- template_project.md: 旧 3 節を外し、`partner:`・`scope:` と generated 領域を足す。旧 3 節のコメントは捨てるが `TEMPLATE_COMMENTS_DROPPED` と報告する。
  project ノートのコメントは、template の旧 3 節にあるものと同じ文面のときだけ捨てる（それ以外は人が書いたものとして残す）。
- そのうえで render（generated 領域・list_project.md・任意で Raycast）の結果を同じ batch に含める。適用後に render が no-op になる。

モジュールは project_update から呼ばれ、そのモジュール自身を `cmd_migrate(args, module)` で受け取る。
"""

from __future__ import annotations

import difflib
import hashlib
import re
import urllib.parse
from dataclasses import dataclass, field
from pathlib import Path

P = None  # project_update モジュール（cmd_migrate が設定する）

_HEADING_RE = re.compile(r"^(#{2,6})\s+(.+?)\s*$")
_FENCE_RE = re.compile(r"^\s{0,3}(`{3,}|~{3,})")
# `[text](target "title")` / `[text](<target with spaces>)` / ファイル名に `(1)` を含む target
_LINK_RE = re.compile(r"\[([^\]]*)\]\(\s*(?:<([^>]*)>|((?:[^()\s]|\([^()]*\))*))[^)]*\)")
_LIST_ITEM_RE = re.compile(r"^\s*[-*]\s+(.*)$")
_TARGETS = ("minutes", "related notes", "documents")
_MIGRATED = "migrated notes"
_BOM = "\ufeff"


class _Raw(str):
    """コードブロックやコメントの中の 1 行。空行の圧縮などの整形をしない。"""


@dataclass
class _Assoc:
    path: Path
    key: str
    topics: str | None
    source: str  # minutes / related notes / documents
    row: str = ""  # minutes の元の行（残すときに使う）
    date: str = ""
    extra: bool = False  # date / minutes / topics の外に列がある
    keep: bool = False  # 行ごと `## migrated notes` に残す


@dataclass
class _Report:
    lines: list[str] = field(default_factory=list)

    def add(self, line: str) -> None:
        self.lines.append(line)


@dataclass
class _Plan:
    changes: list = field(default_factory=list)
    report: _Report = field(default_factory=_Report)
    counts: dict[str, int] = field(default_factory=dict)
    template_state: str = "NOOP"


# ---- helpers ----------------------------------------------------------------------------


def _target_of(match: re.Match[str]) -> str:
    return match.group(2) if match.group(2) is not None else (match.group(3) or "")


def _resolve(vault: Path, base: Path, target: str) -> Path | None:
    target = target.strip()
    target = target.split("#", 1)[0].split("?", 1)[0]  # アンカーとクエリは付いていても同じファイル
    target = urllib.parse.unquote(target.strip())
    if not target or "://" in target or target.startswith("mailto:"):
        return None
    path = (base / target).resolve()
    try:
        relative = path.relative_to(vault.resolve())
    except ValueError:
        return None
    if path.suffix.lower() != ".md" or not path.is_file():
        return None
    if relative.parts[0] in ("project", "setting"):
        return None  # project ノート・設定は関連ノートではない
    return path


def _fenced_lines(lines: list[str]) -> list[bool]:
    """Which lines sit inside a fenced code block (the fence lines included)."""
    flags = [False] * len(lines)
    fence: tuple[str, int] | None = None
    for index, line in enumerate(lines):
        match = _FENCE_RE.match(line)
        if fence is None:
            if match:
                fence = (match.group(1)[0], len(match.group(1)))
                flags[index] = True
        else:
            flags[index] = True
            stripped = line.strip()
            if match and set(stripped) == {fence[0]} and len(stripped) >= fence[1]:
                fence = None
    return flags


def _headings(lines: list[str], fenced: list[bool]) -> list[tuple[int, int, str]]:
    out = []
    for index, line in enumerate(lines):
        if fenced[index]:
            continue
        match = _HEADING_RE.match(line.rstrip("\r\n"))
        if match:
            out.append((index, len(match.group(1)), match.group(2).strip()))
    return out


def _old_title(title: str) -> str | None:
    base = re.sub(r"\s*[（(][^）)]*[）)]\s*$", "", title).strip().lower()
    return base if base in _TARGETS else None


def _section_ends(lines: list[str], heads: list[tuple[int, int, str]], n: int) -> int:
    index, level, _ = heads[n]
    return next((h[0] for h in heads[n + 1 :] if h[1] <= level), len(lines))


def _old_sections(lines: list[str], fenced: list[bool]) -> list[tuple[str, int, int]]:
    """(normalised title, start, end) of the old minutes / related notes / documents sections.

    `## minutes (2026)`・`### minutes`・大文字小文字違いも旧節。別の旧節の中に入れ子になっているものは親の一部として扱う。
    """
    heads = _headings(lines, fenced)
    out: list[tuple[str, int, int]] = []
    covered = 0
    for n, (index, _level, title) in enumerate(heads):
        normal = _old_title(title)
        if normal is None or index < covered:
            continue
        end = _section_ends(lines, heads, n)
        out.append((normal, index, end))
        covered = end
    return out


def _find_section(lines: list[str], fenced: list[bool], title: str) -> tuple[int, int] | None:
    heads = _headings(lines, fenced)
    for n, (index, _level, text) in enumerate(heads):
        if text.lower() == title:
            return index, _section_ends(lines, heads, n)
    return None


def _is_empty_row(cells: list[str]) -> bool:
    return all(not c.strip() for c in cells)


def _is_separator(cells: list[str]) -> bool:
    return all(set(c.strip()) <= {"-", ":", " "} for c in cells)


def _norm_comment(text: str) -> str:
    return " ".join(text.split())


@dataclass
class _Classified:
    assoc: list[_Assoc] = field(default_factory=list)
    preserve: list = field(default_factory=list)  # str / _Raw / _Assoc（行ごと残すかは後で決める）
    dropped: int = 0
    comments: list[str] = field(default_factory=list)


def _classify_section(
    vault: Path, note, key, title: str, body: list[str], fenced: list[bool], rec_ok, related: str = "keep",
    known_comments: set[str] | None = None,
) -> _Classified:
    """Classify one old section. `known_comments=None` drops every comment (the template itself)."""
    out = _Classified()
    body = list(body)
    i = 0
    while i < len(body):
        text = body[i].rstrip("\r\n")
        stripped = text.strip()
        if fenced[i]:
            out.preserve.append(_Raw(text))
            i += 1
            continue
        if stripped == P._BEGIN_MARK:
            end = next((k for k in range(i + 1, len(body)) if not fenced[k] and body[k].strip() == P._END_MARK), None)
            if end is None:
                raise P.ConflictError(f"{note.path.name}: a generated-region begin marker without an end is inside an old section; fix it by hand")
            i = end + 1  # 既存の generated 領域は render が作り直す
            continue
        if stripped == P._END_MARK:
            raise P.ConflictError(f"{note.path.name}: a generated-region end marker without a begin is inside an old section; fix it by hand")
        if not stripped:
            out.preserve.append("")
            i += 1
            continue
        if stripped.startswith("<!--"):
            first = body[i].split("<!--", 1)[1]
            if "-->" in first:
                end, tail = i, first.split("-->", 1)[1]
            else:
                end = next((k for k in range(i + 1, len(body)) if "-->" in body[k]), None)
                tail = body[end].split("-->", 1)[1] if end is not None else ""
            if end is None:
                out.preserve.extend(_Raw(b.rstrip("\r\n")) for b in body[i:])  # 閉じていないコメントは原文のまま
                break
            lines = [b.rstrip("\r\n") for b in body[i : end + 1]]
            lines[-1] = lines[-1][: len(lines[-1]) - len(tail.rstrip("\r\n"))] if tail.rstrip("\r\n") else lines[-1]
            comment = _norm_comment(" ".join(lines))
            if known_comments is None or comment in known_comments:
                out.comments.append(comment)
                if tail.strip():
                    body[end] = tail
                    i = end
                else:
                    i = end + 1
                continue
            out.preserve.extend(_Raw(b.rstrip("\r\n")) for b in body[i : end + 1])
            i = end + 1
            continue
        if re.fullmatch(r"[-*]", stripped):
            out.dropped += 1  # 中身の無い箇条書き（テンプレの placeholder）
            i += 1
            continue
        cells = P._table_cells(text)
        if cells is not None:
            i += 1
            if _is_separator(cells) or _is_empty_row(cells):
                continue
            if [c.lower() for c in cells[:3]] == ["date", "minutes", "topics"]:
                continue
            link = _LINK_RE.search(cells[1]) if title == "minutes" and len(cells) >= 2 else None
            target = _resolve(vault, note.path.parent, _target_of(link)) if link else None
            if target is not None and rec_ok(target):
                assoc = _Assoc(
                    target, key, cells[2].strip() if len(cells) > 2 else "", "minutes",
                    row=text, date=cells[0].strip(), extra=any(c.strip() for c in cells[3:]),
                )
                out.assoc.append(assoc)
                out.preserve.append(assoc)
            else:
                out.preserve.append(text)
            continue
        i += 1
        item = _LIST_ITEM_RE.match(text)
        if item and title in ("related notes", "documents"):
            link = _LINK_RE.search(item.group(1))
            if link and not _target_of(link).strip():
                out.dropped += 1  # `[local]()` のような空リンクのテンプレ placeholder。情報が無い
                continue
            target = _resolve(vault, note.path.parent, _target_of(link)) if link else None
            if title == "related notes" and related == "keep":
                out.preserve.append(text)  # 言及は所属ではない。既定では紐付けず原文のまま残す
                continue
            if target is not None and rec_ok(target):
                out.assoc.append(_Assoc(target, key, None, title))
                rest = item.group(1)[link.end() :].lstrip(" —–-:：\t").strip()
                if rest:
                    out.preserve.append(text)  # 説明文は record の summary ではない。原文のまま残す
                continue
        out.preserve.append(text)
    return out


def _resolved(items: list) -> list[str]:
    """The lines to write: kept rows resolved, plain blank runs collapsed, leading/trailing blanks trimmed."""
    lines: list[str] = []
    for item in items:
        if isinstance(item, _Assoc):
            if item.keep:
                lines.append(item.row)
            continue
        if type(item) is str and not item.strip():
            if lines and lines[-1] != "":
                lines.append("")
            continue
        lines.append(item)
    while lines and type(lines[-1]) is str and lines[-1] == "":
        lines.pop()
    return lines


def _has_content(items: list) -> bool:
    return any(isinstance(i, _Assoc) or i.strip() or type(i) is _Raw for i in items)


def _rebuild_project_note(note, removed: list[tuple[int, int]], block: list[str], existing: tuple[int, int] | None = None) -> list[str]:
    nl = note.newline
    lines = list(note.lines)
    out: list[str] = []
    skip: set[int] = set()
    for start, end in removed:
        skip.update(range(start, end))
    lead = False
    if existing is not None and block:
        at = existing[1]
        while at - 1 > existing[0] and not lines[at - 1].strip():
            at -= 1
        insert_at, lead = at, True
        block = block[2:]  # 既にある `## migrated notes` の末尾へ足す（見出しを重ねない）
    else:
        insert_at = removed[0][0] if removed else None
    ends = {end for _, end in removed}
    for index, line in enumerate(lines):
        if index == insert_at and block:
            if lead:
                out.append(nl)
            out.extend(line_ + nl for line_ in block)
            out.append(nl)
        if index in skip:
            continue
        if out and index in ends and out[-1].strip():
            out.append(nl)  # 取り除いた節の次の見出しの前に空行を 1 つ
        out.append(line)
    if insert_at == len(lines) and block:
        if lead:
            out.append(nl)
        out.extend(line_ + nl for line_ in block)
    return out


def _migrated_block(preserved: dict[str, list]) -> list[str]:
    block: list[str] = []
    for title in _TARGETS:
        lines = _resolved(preserved.get(title, []))
        if lines:
            if not block:
                block += [f"## {_MIGRATED}", ""]
            block += [f"### from {title}", ""] + lines + [""]
    return block[:-1] if block else block


def _body_without_any_backlink(note) -> str:
    """The text after the frontmatter with every strict backlink line removed, wherever it sits.

    実 vault には backlink が本文の途中にある record がある。先頭に正規のブロックを足して元を残すと二重になるので、
    本文中の backlink 行は取り除き、先頭のブロックに 1 本化する（machine-derived な行で、手書きの本文ではない）。
    """
    kept = []
    for line in note.lines[note.close_index + 1 :]:
        match = P._BACKLINK_RE.match(line.strip())
        if match is not None and P._backlink_label(match) == match.group("link"):
            continue
        kept.append(line)
    text = "".join(kept)
    return text.lstrip("\r\n") if text.strip() else text


def _legacy_record_text(note, new_projects: list[str], source: str | None, summary: str | None, bom: bool = False) -> str:
    nl = note.newline
    fm_lines = list(note.lines[1 : note.close_index])
    for name in ("project", "project_source"):
        if any(P._YAML_COMMENT_RE.search(line) for line in P._fm_key_block(fm_lines, name)):
            raise P.ConflictError(f"{note.path.name}: the {name!r} frontmatter has a YAML comment; edit it by hand before migrating")
    fm_lines = P._replace_fm_key(fm_lines, "project", P._fm_block("project", new_projects))
    if source is not None:
        fm_lines = P._replace_fm_key(fm_lines, "project_source", P._fm_block("project_source", source))
    if summary is not None:
        fm_lines = P._replace_fm_key(fm_lines, "summary", P._fm_block("summary", summary))
    rest = _body_without_any_backlink(note)
    block = "".join(P._backlink_line(k) + nl for k in new_projects)
    separator = nl if rest and not rest.startswith(nl) else ""
    closing = note.lines[note.close_index]
    if not closing.endswith(nl):
        closing += nl
    text = note.lines[0] + "".join(line.rstrip("\r\n") + nl for line in fm_lines) + closing + block + separator + rest
    written = P._parse_note_text(note.path, text)
    keys = sum(1 for line in text.splitlines()[1 : written.close_index] if re.match(r"^[\"']?project[\"']?\s*:", line))
    if P._projects_of(written) != new_projects or keys != 1:
        raise P.ConflictError(f"{note.path.name}: refusing to write a record whose `project` key would be duplicated or misread")
    return _BOM + text if bom else text


def _clean_topics(raw: str) -> str:
    return " ".join(P._normalize_line_separators(raw).split())


# ---- the plan ---------------------------------------------------------------------------


def _read_note_text(path: Path) -> str:
    text = P._read_exact(path)
    if text.startswith(_BOM):
        raise P.ValidationError(f"{path.name}: starts with a BOM; remove the BOM before migrating")
    return text


def build_plan(vault: Path, raycast_script: Path | None, accept_list_changes: bool, related: str = "keep") -> _Plan:
    plan = _Plan()
    overlay: dict[Path, str] = {}
    originals: dict[Path, str] = {}
    report = plan.report
    resolved_vault = vault.resolve()

    # 1. project を持ちうる note（vault 内のどの .md でも。record/ 直下だけではない）
    records: dict[Path, object] = {}
    raw_of: dict[Path, str] = {}
    for path in P._walk_markdown(vault):
        if path.is_symlink():
            continue
        try:
            text = P._read_exact(path)
        except (OSError, UnicodeDecodeError):
            report.add(f"WARN {path.relative_to(vault).as_posix()}: not readable as UTF-8; left alone")
            continue
        note = P._read_frontmatter(path, text)
        if note is None:
            continue
        records[path.resolve()] = note
        raw_of[path.resolve()] = text

    frontmatter_cache: dict[Path, bool] = {}

    def rec_ok(target: Path) -> bool:
        """実在する vault 内 md で、frontmatter を読めるものだけ紐付けできる（無いものは行を消さずに残す）。"""
        if target not in frontmatter_cache:
            try:
                frontmatter_cache[target] = P._read_frontmatter(target, P._read_exact(target)) is not None
            except (OSError, UnicodeDecodeError):
                frontmatter_cache[target] = False
        return frontmatter_cache[target]

    # 2. template の旧 3 節（コメントは project ノートのコメントを見分ける物差しにもなる）
    template = vault / "setting" / "template" / "template_project.md"
    tmpl = None
    known_comments: set[str] = set()
    if template.is_file():
        _read_note_text(template)
        tnote = P.parse_note(template, exact=True)
        tlines = list(tnote.lines)
        tfenced = _fenced_lines(tlines)
        tsecs = _old_sections(tlines, tfenced)
        manual = False
        tcomments: list[str] = []
        for title, start, end in tsecs:
            cl = _classify_section(vault, tnote, None, title, tlines[start + 1 : end], tfenced[start + 1 : end], rec_ok, known_comments=None)
            if cl.assoc or _has_content(cl.preserve):
                manual = True
            tcomments.extend(cl.comments)
        known_comments = set(tcomments)
        tmpl = (tnote, tlines, tsecs, manual, tcomments)

    # 3. project ノートの旧 3 節
    assocs: list[_Assoc] = []
    list_rows: dict[str, dict[str, str]] = {}
    list_path = vault / "setting" / "list" / "list_project.md"
    if list_path.is_file():
        try:
            header_map, rows_ = P._list_table(P._read_exact(list_path))
            for _, cells in rows_:
                def cell(name: str, cells=cells) -> str:
                    index = header_map.get(name)
                    return re.sub(r"\\([_|*`\[\]()<>#\\])", r"\1", cells[index].strip()) if index is not None and index < len(cells) else ""
                if cell("name"):
                    list_rows[cell("name")] = {"partner": cell("partner"), "status": cell("status")}
        except (P.ValidationError, OSError, UnicodeDecodeError):
            pass

    infos: list[dict] = []
    dropped_total = comments_total = missing = 0
    for key, path in P._project_note_targets(vault, None):
        _read_note_text(path)
        note = P.parse_note(path, exact=True)
        lines = list(note.lines)
        fenced = _fenced_lines(lines)
        secs = _old_sections(lines, fenced)
        preserved: dict[str, list] = {}
        for title, start, end in secs:
            cl = _classify_section(vault, note, key, title, lines[start + 1 : end], fenced[start + 1 : end], rec_ok, related, known_comments)
            assocs.extend(cl.assoc)
            dropped_total += cl.dropped
            comments_total += len(cl.comments)
            preserved.setdefault(title, []).extend(cl.preserve)
            if title == "minutes":
                for item in cl.preserve:
                    if type(item) is str and item.startswith("|"):
                        missing += 1
                        report.add(f"UNLINKED_ROW {key}: row kept in migrated notes (record not found, not a note, or the link is not readable): {item.strip()[:80]}")
        infos.append({"key": key, "path": path, "note": note, "lines": lines, "fenced": fenced, "secs": secs, "preserved": preserved})

    # 4. record の更新（既存の所属を保持して追加・legacy・summary）
    by_record: dict[Path, list[_Assoc]] = {}
    for assoc in assocs:
        by_record.setdefault(assoc.path.resolve(), []).append(assoc)
    touched = set(by_record) | {p for p, n in records.items() if P._projects_of(n)}
    updated_records = documents = kept_summaries = kept_rows = 0
    for rpath in sorted(touched):
        note = records.get(rpath)
        raw = raw_of.get(rpath)
        if note is None:
            try:
                raw = P._read_exact(rpath)
            except (OSError, UnicodeDecodeError):
                continue
            note = P._read_frontmatter(rpath, raw)
            if note is None:
                report.add(f"WARN {rpath.relative_to(resolved_vault).as_posix()}: no frontmatter; cannot be associated, left alone")
                continue
        projects = P._projects_of(note)
        names, malformed = P._backlink_names(note.text)
        if malformed:
            raise P.ConflictError(f"{note.path.name}: an existing project backlink is malformed; fix it by hand")
        if len(set(names)) != len(names):
            raise P.ValidationError(f"{note.path.name}: duplicate project backlinks")
        if set(names) != set(projects):
            # 旧運用では backlink だけある・project だけある record がある。どちらも所属の証拠なので和集合を legacy で保つ。
            # 推測を含むので、dry-run に必ず出して人間が見る。
            repaired = projects + [n for n in names if n not in projects]
            if set(names) < set(projects):
                report.add(f"ADDED_BACKLINK {note.path.name}: backlink added for {sorted(set(projects) - set(names))!r}")
            else:
                report.add(f"REPAIRED_RECORD {note.path.name}: project {projects!r} + backlinks {names!r} -> {repaired!r}")
            projects = repaired
        elif names and len(P._leading_backlink_lines(note)) != len(names):
            report.add(f"MOVED_BACKLINK {note.path.name}: backlink moved to the top of the body")
        new_projects = list(projects)
        mine = by_record.get(rpath, [])
        for assoc in mine:
            if assoc.key not in new_projects:
                new_projects.append(assoc.key)
        source = None if str(note.data.get("project_source") or "").strip() else "legacy"
        summary = None
        existing = " ".join(str(note.data.get("summary") or "").split())
        topics = next((_clean_topics(a.topics) for a in mine if a.source == "minutes" and a.topics), "")
        if topics and not existing:
            if P._has_local_path(topics):
                report.add(f"SUMMARY_SKIPPED {rpath.name}: topics contain a local path")
            else:
                summary = topics
        elif topics and existing and existing != topics:
            kept_summaries += 1
            report.add(f"SUMMARY_KEPT {rpath.name}")
        final_summary = existing or summary or ""
        for assoc in mine:
            if assoc.source != "minutes":
                continue
            reasons = []
            topic = _clean_topics(assoc.topics or "")
            if assoc.extra:
                reasons.append("extra columns")
            if assoc.date and assoc.date != P._note_date(note):
                reasons.append("date differs from the record")
            if topic and topic != final_summary:
                reasons.append("topic is not the summary")
            if reasons:
                assoc.keep = True
                kept_rows += 1
                report.add(f"ROW_KEPT {assoc.key}: {rpath.name}: " + ", ".join(reasons))
        if (
            new_projects == P._projects_of(note)
            and isinstance(note.data.get("project"), list)
            and source is None
            and summary is None
            and set(names) == set(new_projects)
            and len(P._leading_backlink_lines(note)) == len(names)
        ):
            continue
        text = _legacy_record_text(note, new_projects, source, summary, bom=bool(raw and raw.startswith(_BOM)))
        if raw is not None and text != raw:
            overlay[rpath] = text
            originals[rpath] = raw
            updated_records += 1
            if rpath.relative_to(resolved_vault).parts[0] != "record":
                documents += 1

    # 5. project ノート本体（行ごと残すものが決まってから組み立てる）
    project_notes = len(infos)
    for info in infos:
        key, path, note, lines = info["key"], info["path"], info["note"], info["lines"]
        removed = [(start, end) for _, start, end in info["secs"]]
        block = _migrated_block(info["preserved"])
        existing = _find_section(lines, info["fenced"], _MIGRATED)
        new_lines = _rebuild_project_note(note, removed, block, existing) if removed else lines
        text = "".join(new_lines)
        # frontmatter: list_project.md にだけある partner / status を移す（frontmatter に無い、または空のとき）
        row = list_rows.get(key, {})
        extra = {
            k: v
            for k, v in (("partner", row.get("partner", "")), ("status", row.get("status", "")))
            if v and note.data.get(k) in (None, "")
        }
        if extra:
            parsed = P._parse_note_text(path, text)
            fm_lines = list(parsed.lines[1 : parsed.close_index])
            for name, value in extra.items():
                fm_lines = P._replace_fm_key(fm_lines, name, P._fm_block(name, value))
            nl = parsed.newline
            text = parsed.lines[0] + "".join(l.rstrip("\r\n") + nl for l in fm_lines) + "".join(parsed.lines[parsed.close_index :])
        if text != note.text:
            overlay[path] = text
            originals[path] = note.text

    # 6. template
    template_state = "NOOP"
    if tmpl is not None:
        tnote, tlines, tsecs, manual, tcomments = tmpl
        if manual:
            template_state = "MANUAL"
            report.add("TEMPLATE_MANUAL template_project.md has content in the old sections; edit it by hand")
        else:
            lines2 = _rebuild_project_note(tnote, [(a, b) for _, a, b in tsecs], []) if tsecs else tlines
            text = "".join(lines2)
            parsed = P._parse_note_text(template, text)
            fm_lines = list(parsed.lines[1 : parsed.close_index])
            for name in ("partner", "scope"):
                if name not in parsed.data:
                    fm_lines = P._replace_fm_key(fm_lines, name, [f"{name}:\n"])
            nl = parsed.newline
            text = parsed.lines[0] + "".join(l.rstrip("\r\n") + nl for l in fm_lines) + "".join(parsed.lines[parsed.close_index :])
            region_note = P._parse_note_text(template, text)
            if not any(l.strip() == P._BEGIN_MARK for l in region_note.lines):
                text = P._with_region(region_note, P._region_lines([]))
            if text != tnote.text:
                overlay[template] = text
                originals[template] = tnote.text
                template_state = "CHANGE"
                for comment in tcomments:
                    report.add(f"TEMPLATE_COMMENTS_DROPPED template_project.md: {comment[:100]}")

    # 7. render（領域・list・Raycast）を、計画後の内容（overlay）に対して
    grouped = P._collect_associations(vault, overlay)
    region_changes, _, _ = P._plan_regions(vault, None, grouped, overlay)
    final: dict[Path, str] = dict(overlay)
    for change in region_changes:
        final[change.path] = change.after
        originals.setdefault(change.path, change.before)
    rows = P._project_rows(vault, grouped, overlay)
    list_change, list_changed = P._plan_list(vault, rows, accept_list_changes)
    created: set[Path] = set()
    if list_change is not None:
        final[list_change.path] = list_change.after
        originals.setdefault(list_change.path, list_change.before)
        if list_change.created:
            created.add(list_change.path)
    if raycast_script is not None:
        raycast_change, _ = P._plan_raycast(raycast_script.expanduser(), rows)
        if raycast_change is not None:
            final[raycast_change.path] = raycast_change.after
            originals.setdefault(raycast_change.path, raycast_change.before)

    for path in sorted(final):
        before = originals.get(path, "")
        if final[path] != before or path in created:
            plan.changes.append(
                P.Change(path, before, final[path], created=path in created, verify_before=path not in created)
            )
    plan.counts = {
        "project_notes": project_notes,
        "records_to_update": updated_records,
        "documents_to_associate": documents,
        "summaries_kept": kept_summaries,
        "rows_kept_in_migrated_notes": missing + kept_rows,
        "placeholders_dropped": dropped_total,
        "comments_dropped": comments_total,
        "list_rows": len(rows),
    }
    plan.template_state = template_state
    return plan


def _plan_id(vault: Path, changes) -> str:
    digest = hashlib.sha256()
    for change in sorted(changes, key=lambda c: str(c.path)):
        digest.update(str(change.path.relative_to(vault) if change.path.is_relative_to(vault) else change.path.name).encode())
        digest.update(hashlib.sha256(change.before.encode()).digest())
        digest.update(hashlib.sha256(change.after.encode()).digest())
    return digest.hexdigest()[:16]


def _check_diff_dir(vault: Path, directory: Path) -> Path:
    """The diff directory must be outside the vault (private diffs, and a dry-run writes nothing in the vault)."""
    target = directory.expanduser().resolve()
    root = vault.resolve()
    if target == root or root in target.parents or target in root.parents:
        raise P.ValidationError("--diff-dir must be a directory of its own outside the vault")
    if target.exists():
        if not target.is_dir():
            raise P.ValidationError("--diff-dir is not a directory")
        if any(target.iterdir()) and not (target / "SUMMARY.md").is_file():
            raise P.ValidationError("--diff-dir is not empty and does not look like a previous migrate diff directory")
    return target


def _unified(rel: Path, before: str, after: str) -> str:
    out = []
    for line in difflib.unified_diff(
        before.splitlines(keepends=True), after.splitlines(keepends=True), f"a/{rel.as_posix()}", f"b/{rel.as_posix()}"
    ):
        out.append(line if line.endswith("\n") else line + "\n\\ No newline at end of file\n")
    return "".join(out)


def _write_diffs(vault: Path, changes, directory: Path, plan: _Plan) -> None:
    directory.mkdir(parents=True, exist_ok=True)
    for stale in directory.rglob("*.diff"):
        stale.unlink()  # 前回の dry-run の diff を残さない（今のプランに無いファイルの diff が混ざらないように）
    for change in changes:
        try:
            rel = change.path.relative_to(vault)
        except ValueError:
            rel = Path("outside") / change.path.name
        target = directory / rel.parent / f"{rel.name}.diff"
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(_unified(rel, change.before, change.after).encode("utf-8"))
    summary = ["# migrate dry-run summary", ""] + [f"- {k}: {v}" for k, v in plan.counts.items()]
    summary += ["", "## files", ""] + [f"- {c.path.relative_to(vault).as_posix() if c.path.is_relative_to(vault) else c.path.name}" for c in changes]
    summary += ["", "## notes", ""] + [f"- {line}" for line in plan.report.lines]
    (directory / "SUMMARY.md").write_text("\n".join(summary) + "\n", encoding="utf-8")


def cmd_migrate(args, module) -> int:
    global P
    P = module
    vault: Path = args.vault
    if args.plan_id and not args.apply:
        raise P.ValidationError("--plan-id is only used together with --apply")
    diff_dir = _check_diff_dir(vault, args.diff_dir) if args.diff_dir is not None else None
    plan = build_plan(vault, args.raycast_script, args.accept_list_changes, args.related_notes)
    changes = plan.changes
    for line in plan.report.lines:
        print(line)
    print("MIGRATE " + " ".join(f"{k}={v}" for k, v in plan.counts.items()) + f" template={plan.template_state}")
    for change in changes:
        print(f"CHANGE {change.path.relative_to(vault).as_posix() if change.path.is_relative_to(vault) else change.path.name}")
    if diff_dir is not None:
        _write_diffs(vault, changes, diff_dir, plan)
    if not changes:
        print("NOOP no migration changes required")
        return 0
    pid = _plan_id(vault, changes)
    print(f"PLAN_ID {pid}")
    if not args.apply:
        print("DRY_RUN no files written (review the diff, then pass --apply --plan-id <PLAN_ID>)")
        return 0
    if not args.plan_id:
        raise P.ValidationError("--apply needs --plan-id from the dry-run you reviewed")
    if args.plan_id != pid:
        raise P.ConflictError("the plan changed since the reviewed dry-run; run the dry-run again and review it")
    P._write_batch(vault, changes)
    print(f"APPLIED files={len(changes)}")
    return 0
