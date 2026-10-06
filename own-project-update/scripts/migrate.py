"""migrate: 既存 vault を新契約へ 1 回で移す（dry-run → 差分 → 人間 → 適用）。

SPEC-project-association の「配置と移行」。夜間 job では実行しない。実行主体は人間が見たプランだけ:
`--apply` は、dry-run が出した PLAN_ID を `--plan-id` で渡したときだけ書く（見た差分と書く内容を一致させる）。

やること（すべて決定論。LLM を使わない）:
- record: `project` を常にリストにし、`project_source: legacy` を付ける（既に出所があれば保つ）。
  project ノートの旧 3 節（minutes / related notes / documents）に載っている record・資料は、その project に属する証拠として
  `legacy` で紐付ける（既存の所属は保持して追加）。minutes の topics は、record に summary が無いときだけ `summary` へ移す。
- project ノート: 旧 3 節を取り除き、表現できない内容（引用メモ・外部リンク・related notes の説明文・実体の無い record の行）は
  `## migrated notes` に原文のまま残す。generated 領域は render が作る。`list_project.md` の partner / status が frontmatter に無ければ移す。
- template_project.md: 旧 3 節を外し、`partner:`・`scope:` と generated 領域を足す。
- そのうえで render（generated 領域・list_project.md・任意で Raycast）の結果を同じ batch に含める。適用後に render が no-op になる。

モジュールは project_update から呼ばれ、そのモジュール自身を `cmd_migrate(args, module)` で受け取る。
"""

from __future__ import annotations

import difflib
import hashlib
import re
import sys
import urllib.parse
from dataclasses import dataclass, field
from pathlib import Path

P = None  # project_update モジュール（cmd_migrate が設定する）

_HEADING_RE = re.compile(r"^##\s+(.+?)\s*$")
_LINK_RE = re.compile(r"\[([^\]]*)\]\(([^)]*)\)")
_LIST_ITEM_RE = re.compile(r"^\s*[-*]\s+(.*)$")
_TARGETS = ("minutes", "related notes", "documents")


@dataclass
class _Assoc:
    path: Path
    key: str
    topics: str | None
    source: str  # minutes / related notes / documents


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


# ---- helpers ----------------------------------------------------------------------------


def _resolve(vault: Path, base: Path, target: str) -> Path | None:
    target = urllib.parse.unquote(target.strip())
    if not target or "://" in target or target.startswith(("#", "mailto:")):
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


def _sections(lines: list[str]) -> list[tuple[str, int, int]]:
    heads = [(i, m.group(1).strip()) for i, line in enumerate(lines) if (m := _HEADING_RE.match(line))]
    out = []
    for n, (index, title) in enumerate(heads):
        end = heads[n + 1][0] if n + 1 < len(heads) else len(lines)
        out.append((title, index, end))
    return out


def _is_empty_row(cells: list[str]) -> bool:
    return all(not c.strip() for c in cells)


def _is_separator(cells: list[str]) -> bool:
    return all(set(c.strip()) <= {"-", ":", " "} for c in cells)


def _classify_section(
    vault: Path, note, title: str, body: list[str], rec_ok, related: str = "keep"
) -> tuple[list[_Assoc], list[str], int]:
    """(associations, lines to preserve verbatim, dropped placeholder count) for one old section."""
    assoc: list[_Assoc] = []
    preserve: list[str] = []
    dropped = 0
    in_comment = False
    key = P._project_value(note)
    for raw in body:
        text = raw.rstrip("\r\n")
        stripped = text.strip()
        if in_comment:
            in_comment = "-->" not in stripped
            continue
        if not stripped:
            continue
        if stripped.startswith("<!--"):
            in_comment = "-->" not in stripped
            continue
        if re.fullmatch(r"[-*]", stripped):
            dropped += 1  # 中身の無い箇条書き（テンプレの placeholder）
            continue
        cells = P._table_cells(text)
        if cells is not None:
            if _is_separator(cells) or _is_empty_row(cells):
                continue
            if [c.lower() for c in cells[:3]] == ["date", "minutes", "topics"]:
                continue
            link = _LINK_RE.search(cells[1]) if title == "minutes" and len(cells) >= 3 else None
            target = _resolve(vault, note.path.parent, link.group(2)) if link else None
            if target is not None and rec_ok(target):
                assoc.append(_Assoc(target, key, cells[2].strip(), "minutes"))
            else:
                preserve.append(text)
            continue
        item = _LIST_ITEM_RE.match(text)
        if item and title in ("related notes", "documents"):
            link = _LINK_RE.search(item.group(1))
            if link and not link.group(2).strip():
                dropped += 1  # `[local]()` のような空リンクのテンプレ placeholder。情報が無い
                continue
            target = _resolve(vault, note.path.parent, link.group(2)) if link else None
            if title == "related notes" and related == "keep":
                preserve.append(text)  # 言及は所属ではない。既定では紐付けず原文のまま残す
                continue
            if target is not None and rec_ok(target):
                assoc.append(_Assoc(target, key, None, title))
                rest = item.group(1)[link.end() :].lstrip(" —–-:：\t").strip()
                if rest:
                    preserve.append(text)  # 説明文は record の summary ではない。原文のまま残す
                continue
        preserve.append(text)
    return assoc, preserve, dropped


def _rebuild_project_note(note, removed: list[tuple[int, int]], block: list[str]) -> list[str]:
    nl = note.newline
    lines = list(note.lines)
    out: list[str] = []
    skip: set[int] = set()
    for start, end in removed:
        skip.update(range(start, end))
    first = removed[0][0] if removed else None
    for index, line in enumerate(lines):
        if index == first and block:
            out.extend(line_ + nl for line_ in block)
            out.append(nl)
        if index in skip:
            continue
        if out and index in {end for _, end in removed} and out[-1].strip():
            out.append(nl)  # 取り除いた節の次の見出しの前に空行を 1 つ
        out.append(line)
    return out


def _migrated_block(preserved: dict[str, list[str]]) -> list[str]:
    block: list[str] = []
    for title in _TARGETS:
        if preserved.get(title):
            if not block:
                block += ["## migrated notes", ""]
            block += [f"### from {title}", ""] + preserved[title] + [""]
    return block[:-1] if block else block


def _body_without_any_backlink(note) -> str:
    """The text after the frontmatter with every strict backlink line removed, wherever it sits.

    実 vault には backlink が本文の途中にある record がある。先頭に正規のブロックを足して元を残すと二重になるので、
    本文中の backlink 行は取り除き、先頭のブロックに 1 本化する（machine-derived な行で、手書きの本文ではない）。
    """
    kept = []
    for line in note.lines[note.close_index + 1 :]:
        match = P._BACKLINK_RE.match(line.strip())
        if match is not None and match.group("label") == match.group("link"):
            continue
        kept.append(line)
    text = "".join(kept)
    return text.lstrip("\r\n") if text.strip() else text


def _legacy_record_text(note, new_projects: list[str], source: str | None, summary: str | None) -> str:
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
    return text


def _clean_topics(raw: str) -> str:
    return " ".join(P._normalize_line_separators(raw).replace("\\|", "|").split())


# ---- the plan ---------------------------------------------------------------------------


def build_plan(vault: Path, raycast_script: Path | None, accept_list_changes: bool, related: str = "keep") -> _Plan:
    plan = _Plan()
    overlay: dict[Path, str] = {}
    originals: dict[Path, str] = {}
    report = plan.report

    # 1. records（vault/record/*.md）
    records: dict[Path, object] = {}
    for path in sorted((vault / "record").glob("*.md")):
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

    frontmatter_cache: dict[Path, bool] = {}

    def rec_ok(target: Path) -> bool:
        """実在する vault 内 md で、frontmatter を読めるものだけ紐付けできる（無いものは行を消さずに残す）。"""
        if target not in frontmatter_cache:
            try:
                frontmatter_cache[target] = P._read_frontmatter(target, P._read_exact(target)) is not None
            except (OSError, UnicodeDecodeError):
                frontmatter_cache[target] = False
        return frontmatter_cache[target]

    # 2. project ノートの旧 3 節
    assocs: list[_Assoc] = []
    project_texts: dict[Path, tuple[object, str]] = {}
    list_rows: dict[str, dict[str, str]] = {}
    list_path = vault / "setting" / "list" / "list_project.md"
    if list_path.is_file():
        try:
            header_map, rows_ = P._list_table(P._read_exact(list_path))
            for _, cells in rows_:
                def cell(name: str, cells=cells) -> str:
                    index = header_map.get(name)
                    return cells[index].strip().replace("\\", "") if index is not None and index < len(cells) else ""
                if cell("name"):
                    list_rows[cell("name")] = {"partner": cell("partner"), "status": cell("status")}
        except (P.ValidationError, OSError, UnicodeDecodeError):
            pass

    project_notes = 0
    missing = dropped_total = 0
    for key, path in P._project_note_targets(vault, None):
        project_notes += 1
        note = P.parse_note(path, exact=True)
        lines = list(note.lines)
        secs = [(t, a, b) for t, a, b in _sections(lines) if t.lower() in _TARGETS]
        preserved: dict[str, list[str]] = {}
        removed: list[tuple[int, int]] = []
        for title, start, end in secs:
            found, keep, dropped = _classify_section(vault, note, title.lower(), lines[start + 1 : end], rec_ok, related)
            assocs.extend(found)
            dropped_total += dropped
            if keep:
                preserved.setdefault(title.lower(), []).extend(keep)
            removed.append((start, end))
            for line in keep:
                if title.lower() == "minutes" and line.startswith("|"):
                    missing += 1
                    report.add(f"MISSING_RECORD {key}: row kept in migrated notes: {line.strip()[:80]}")
        new_lines = _rebuild_project_note(note, removed, _migrated_block(preserved)) if removed else lines
        text = "".join(new_lines)
        # frontmatter: list_project.md にだけある partner / status を移す
        row = list_rows.get(key, {})
        extra = {k: v for k, v in (("partner", row.get("partner", "")), ("status", row.get("status", ""))) if v and k not in note.data}
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
        project_texts[path] = (note, text)

    # 3. record の更新（既存の所属を保持して追加・legacy・summary）
    by_record: dict[Path, list[_Assoc]] = {}
    for assoc in assocs:
        by_record.setdefault(assoc.path.resolve(), []).append(assoc)
    touched = set(by_record) | {p for p, n in records.items() if P._projects_of(n)}
    updated_records = documents = kept_summaries = 0
    for rpath in sorted(touched):
        note = records.get(rpath)
        if note is None:
            try:
                text0 = P._read_exact(rpath)
            except (OSError, UnicodeDecodeError):
                continue
            note = P._read_frontmatter(rpath, text0)
            if note is None:
                report.add(f"WARN {rpath.relative_to(vault.resolve()).as_posix()}: no frontmatter; cannot be associated, left alone")
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
            report.add(f"REPAIRED_RECORD {note.path.name}: project {projects!r} + backlinks {names!r} -> {repaired!r}")
            projects = repaired
        elif names and len(P._leading_backlink_lines(note)) != len(names):
            report.add(f"MOVED_BACKLINK {note.path.name}: backlink moved to the top of the body")
        new_projects = list(projects)
        for assoc in by_record.get(rpath, []):
            if assoc.key not in new_projects:
                new_projects.append(assoc.key)
        source = None if str(note.data.get("project_source") or "").strip() else "legacy"
        summary = None
        existing = " ".join(str(note.data.get("summary") or "").split())
        topics = next((_clean_topics(a.topics) for a in by_record.get(rpath, []) if a.source == "minutes" and a.topics), "")
        if topics and not existing:
            if P._has_local_path(topics):
                report.add(f"SUMMARY_SKIPPED {rpath.name}: topics contain a local path")
            else:
                summary = topics
        elif topics and existing and existing != topics:
            kept_summaries += 1
            report.add(f"SUMMARY_KEPT {rpath.name}")
        if new_projects == P._projects_of(note) and source is None and summary is None and set(names) == set(new_projects):
            continue
        text = _legacy_record_text(note, new_projects, source, summary)
        if text != note.text:
            overlay[rpath] = text
            originals[rpath] = note.text
            updated_records += 1
            if rpath.relative_to(vault.resolve()).parts[0] != "record":
                documents += 1

    # 4. template
    template_state = "NOOP"
    template = vault / "setting" / "template" / "template_project.md"
    if template.is_file():
        tnote = P.parse_note(template, exact=True)
        tlines = list(tnote.lines)
        tsecs = [(t, a, b) for t, a, b in _sections(tlines) if t.lower() in _TARGETS]
        manual = False
        for title, start, end in tsecs:
            found, keep, _ = _classify_section(vault, tnote, title.lower(), tlines[start + 1 : end], rec_ok)
            if found or keep:
                manual = True
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

    # 5. render（領域・list・Raycast）を、計画後の内容（overlay）に対して
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
        "rows_kept_in_migrated_notes": missing,
        "placeholders_dropped": dropped_total,
        "list_rows": len(rows),
    }
    plan.template_state = template_state  # type: ignore[attr-defined]
    return plan


def _plan_id(vault: Path, changes) -> str:
    digest = hashlib.sha256()
    for change in sorted(changes, key=lambda c: str(c.path)):
        digest.update(str(change.path.relative_to(vault) if change.path.is_relative_to(vault) else change.path.name).encode())
        digest.update(hashlib.sha256(change.before.encode()).digest())
        digest.update(hashlib.sha256(change.after.encode()).digest())
    return digest.hexdigest()[:16]


def _write_diffs(vault: Path, changes, directory: Path, plan: _Plan) -> None:
    directory.mkdir(parents=True, exist_ok=True)
    for change in changes:
        try:
            rel = change.path.relative_to(vault)
        except ValueError:
            rel = Path("outside") / change.path.name
        target = directory / rel.parent / f"{rel.name}.diff"
        target.parent.mkdir(parents=True, exist_ok=True)
        diff = difflib.unified_diff(
            change.before.splitlines(), change.after.splitlines(), f"a/{rel.as_posix()}", f"b/{rel.as_posix()}", lineterm=""
        )
        target.write_text("\n".join(diff) + "\n", encoding="utf-8")
    summary = ["# migrate dry-run summary", ""] + [f"- {k}: {v}" for k, v in plan.counts.items()]
    summary += ["", "## files", ""] + [f"- {c.path.relative_to(vault).as_posix() if c.path.is_relative_to(vault) else c.path.name}" for c in changes]
    summary += ["", "## notes", ""] + [f"- {line}" for line in plan.report.lines]
    (directory / "SUMMARY.md").write_text("\n".join(summary) + "\n", encoding="utf-8")


def cmd_migrate(args, module) -> int:
    global P
    P = module
    vault: Path = args.vault
    plan = build_plan(vault, args.raycast_script, args.accept_list_changes, args.related_notes)
    changes = plan.changes
    for line in plan.report.lines:
        print(line)
    print("MIGRATE " + " ".join(f"{k}={v}" for k, v in plan.counts.items()) + f" template={getattr(plan, 'template_state', 'NOOP')}")
    for change in changes:
        print(f"CHANGE {change.path.relative_to(vault).as_posix() if change.path.is_relative_to(vault) else change.path.name}")
    if args.diff_dir is not None:
        _write_diffs(vault, changes, args.diff_dir.expanduser(), plan)
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
        raise P.ConflictError(f"the plan changed since the reviewed dry-run (now {pid}); review again")
    P._write_batch(vault, changes)
    print(f"APPLIED files={len(changes)}")
    return 0
