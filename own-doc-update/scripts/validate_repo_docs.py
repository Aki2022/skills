#!/usr/bin/env python3
"""Validate the own-doc-update repository structure and v2 lifecycle contracts."""
from __future__ import annotations

import argparse
import os
import re
import sys
from pathlib import Path
from typing import Optional


FrontMatter = dict[str, object]

# docs/00_index.md is the routing index the SessionStart hook injects -- as a
# digest, not as the file: measured 2026-09-19, Claude Code delivers a hook's
# output to the agent inline only up to about 10,000 CHARACTERS (largest ever
# delivered inline 8,957; smallest ever spilled 10,019, and that one was an index
# injection), so index_digest.py compresses it to fit and the agent opens the file
# when it needs the rest.
#
# This ceiling is therefore a readability bound on the file, not the injection
# bound it was first written as. That earlier reading was wrong twice over: it
# named a ~19 KB spill point and claimed 32 KB sat under it, and the index passed
# the ceiling at 24,234 characters while still spilling every session. What the
# ceiling does say is that an index past it has stopped being a routing index and
# become a progress dashboard -- the same thing the line-length and prose-ratio
# rules below measure. docs_hygiene.py moves narrative out to docs/log/ to meet it.
INDEX_MAX_BYTES = 32 * 1024
# Headroom the index must keep below the ceiling. Measured 2026-09-22: an index
# sat at 32,688 of 32,768 bytes -- 80 bytes free -- and docs_hygiene.py --fix
# reported nothing to do, because its budget pass both starts and stops at the
# ceiling itself. So the file stabilizes one row below red: the next row anyone
# adds turns the default branch red, that person shortens it back to just under,
# and the cycle repeats (three observed rounds: 32,229 -> 33,036 -> 32,688).
# Reserving headroom gives the fix pass somewhere to aim and turns "already red"
# into "getting close", which is a state a session can act on without urgency.
INDEX_HEADROOM_BYTES = 2 * 1024
INDEX_TARGET_BYTES = INDEX_MAX_BYTES - INDEX_HEADROOM_BYTES
INDEX_MAX_LINE_CHARS = 500
INDEX_MAX_PROSE_RATIO = 0.5
DOC_SOFT_MAX_BYTES = 32 * 1024

# The issue template ships `status: active`, so `active` is canonical for a
# standalone issue even though embedded workstream blocks use the four-value set.
ISSUE_STATUSES = ("active", "pending", "in_progress", "blocked", "complete")


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


# Reserved result key holding the names of keys declared more than once. The
# parser applies each declaration in turn, so the last one wins and the earlier
# ones are dropped without a word — recording the names is what lets the caller
# report that instead of validating a value the file does not appear to set.
DUPLICATE_KEYS = "__duplicate_keys__"


def flow_style_keys(front_matter: Optional[FrontMatter]) -> list[str]:
    """Keys whose value was a bracket list opened on the line after the key."""
    if not front_matter:
        return []
    keys = front_matter.get(FLOW_STYLE_KEYS, [])
    return list(keys) if isinstance(keys, list) else []


def duplicate_keys(front_matter: Optional[FrontMatter]) -> list[str]:
    """Keys declared more than once, in the order they were first declared."""
    if not front_matter:
        return []
    keys = front_matter.get(DUPLICATE_KEYS, [])
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
    """Parse the small YAML subset used by own-doc-update templates."""
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
    duplicates: list[str] = []
    current_list: Optional[str] = None
    lines = content[3:end].strip().splitlines()
    index = 0
    while index < len(lines):
        line = lines[index]
        index += 1
        if KEY_LINE.match(line):
            key, _, raw = line.partition(":")
            raw = raw.strip()
            if key in result and key not in duplicates:
                duplicates.append(key)
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
    if duplicates:
        result[DUPLICATE_KEYS] = duplicates
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


def parse_workstream_issue_blocks(content: str) -> list[tuple[str, dict[str, str], str]]:
    """Read compact issue metadata, plus the raw block body, under `### ISSUE-*` headings.

    Only a heading that is exactly the issue id (`### ISSUE-...` with nothing after it)
    is an issue block. Prose headings that merely mention an id
    (e.g. `### ISSUE-07 の切り出し`) used to be parsed as blocks and made
    archiving fail twice with errors pointing at a block that did not exist.
    """
    matches = list(re.finditer(r"^### (ISSUE-[A-Za-z0-9-]+)[ \t]*$", content, re.MULTILINE))
    blocks: list[tuple[str, dict[str, str]]] = []
    for index, match in enumerate(matches):
        end = matches[index + 1].start() if index + 1 < len(matches) else len(content)
        body = content[match.end():end]
        metadata: dict[str, str] = {}
        for item in re.finditer(r"^- ([a-z_]+):[ \t]*(.*?)$", body, re.MULTILINE):
            metadata[item.group(1)] = item.group(2).strip().strip('"').strip("'")
        blocks.append((match.group(1), metadata, body))
    return blocks


def extract_section(content: str, heading: str) -> str | None:
    """Return the text between `heading` and the next same-level heading."""
    match = re.search(
        rf"^{re.escape(heading)}\s*$(.*?)(?=^## |\Z)",
        content,
        re.MULTILINE | re.DOTALL,
    )
    return match.group(1) if match else None


# own-ws-drain treats a missing runnability record as `gated` and stops the
# whole run; own-goal-run refuses to start on an unrecorded envelope and cannot
# verify an empty acceptance. These checks keep those safe defaults from firing
# where no human ever set a gate.
RUNNABILITY_RE = re.compile(r"^(ready|gated on \S.*)$")
VERIFY_LINE_RE = re.compile(
    r"^-[ \t]*verify:[ \t]*(machine|human-review)[ \t]*[—–:-][ \t]*\S", re.MULTILINE
)
REQUIRED_ENVELOPE_BULLETS = ("Autonomous actions allowed", "Confirm first")


def validate_envelope_bullets(rel: str, content: str, errors: list[str]) -> None:
    envelope = extract_section(content, "## Authorization Envelope")
    if envelope is None:
        return  # the missing-section error is already reported
    for bullet in REQUIRED_ENVELOPE_BULLETS:
        match = re.search(rf"^- {re.escape(bullet)}:(.*)$", envelope, re.MULTILINE)
        if match is None:
            errors.append(f"{rel}: Authorization Envelope must record '- {bullet}: ...'")
            continue
        if match.group(1).strip():
            continue
        rest = envelope[match.end():]
        continuation = False
        for line in rest.splitlines():
            if not line.strip():
                continue
            if line.startswith((" ", "\t")):
                continuation = True
            break
        if not continuation:
            errors.append(
                f"{rel}: '- {bullet}:' is empty — an unrecorded boundary becomes a "
                "question gate that stops autonomous runs"
            )


def validate_acceptance_verify(label: str, body: str, errors: list[str]) -> None:
    if not VERIFY_LINE_RE.search(body):
        errors.append(
            f"{label}: Acceptance needs '- verify: machine — <command and expected "
            "result>' or '- verify: human-review — <review gate>' — an agent cannot "
            "self-verify an unstated acceptance"
        )


# `runnability` と Acceptance の `verify:` は、既存文書がまだ書かれていなかった頃の
# 規約に対して後から足された。既存の corpus 全体に一斉に当たるため、導入したリポジトリ
# では 100% の文書が落ちうる（実測: あるリポジトリで issue 134/139・workstream 19/19）。
#
# そこを一括バックフィルで埋めると、**中身を知らないまま「それらしい検証手順」を書く**
# ことになり、この規則が防ごうとしているものそのものを作る。緑になるが嘘が増える。
#
# したがって既存分は「債務」として明示的に列挙して逃がし、新しい文書には最初から
# 効かせる。逃がしたものは減る一方になるよう、リストが古びたら落ちる。
LINK_BASELINE_RELPATH = "docs/validator-link-baseline.txt"
PLACEMENT_BASELINE_RELPATH = "docs/validator-placement-baseline.txt"

# `docs/specs/` に置くと持ち分を外れる見出し。spec は intent・requirements・design policy を
# 持ち、work unit（受入条件・人間ゲート・既存への影響・レビュー結果）と実装状態は持たない。
# 判定基準は「何が起きたらこの文が間違いになるか」——作業が完了したら間違いになる文は
# work unit の持ち分である（ADR-20260922-doc-layer-ownership-and-placement-check、seedion）。
#
# **日本語の同義を必ず含める。** 最初の走査を英語だけで書いたとき、同じ形の見出しを
# 半分近く見落とした（16KB と数えたものが実際は 29KB だった）。
PLACEMENT_FORBIDDEN_IN_SPEC = (
    "Acceptance Criteria",
    "受入条件",
    "受け入れ条件",
    "Impact on Existing System",
    "既存記述への影響",
    "既存への影響",
    "レビュー結果",
    "Next Question",
)
# 見出しが実装の状態を名乗るなら、それは現況＝guide の持ち分である。
PLACEMENT_STATUS_IN_HEADING = re.compile(r"実装(完了|済み?|中)|実装は段階")

# 直近の validate_repo() が「何件を検査し、何件を分岐で外したか」。
# 2026-09-20 の実測: validator は 477 件中 56 件の作業単位を条件式で飛ばしていたが、
# 飛ばしたことをどこにも出していなかった。バグは検出できなかったのではなく
# 報告されなかった。根拠: SPEC-justified-nonconformance F4。
LAST_COVERAGE: dict[str, dict[str, int]] = {}

# Only the `](` matters: the link text may contain brackets, or a whole nested
# image, and anchoring on `[` would skip the outer destination entirely.
_LINK_OPEN = re.compile(r"\]\(")
_MAX_DESTINATION = 2048
_URI_SCHEME = re.compile(r"^[A-Za-z][A-Za-z0-9+.\-]*:")
# Same-length backtick delimiters, within one line. Line-scoped on purpose: a
# span cannot contain a blank line, and a pattern that crosses lines lets one
# unpaired backtick delete every link until the next one.
_INLINE_CODE = re.compile(r"(`+)([^\n]+?)\1")


def strip_code(text: str) -> str:
    """Blank out inline code spans, and nothing else.

    Earlier versions also tracked fenced blocks, indented code blocks,
    blockquotes and HTML comments, so that examples inside them would not be
    read as references. Six review rounds each found a construct where that
    tracking opened a block it never closed and silently deleted every link to
    the end of the document — the exact failure this check exists to prevent.

    It was then measured. Across 26 repositories using this convention (1,924
    documents), removing the whole block layer changed the finding count by
    **+5, in 2 of them** — and never lost a finding it previously had. The two
    repositories originally sampled changed by zero. Roughly 150 lines of
    parsing were buying five loud, baselineable false positives, at the price of
    an unbounded silent-loss surface.

    The five are real document shapes, not accidents: a `](` inside a
    `gcloud --format` string in a shell example, a markdown table row quoted
    inside a fence, and an `![](...)` inside an example data record.

    What remains cannot swallow more than one line, and the residual blind spot
    is one unpaired backtick hiding the links after it *on that line*.

    The consequence for authors: a markdown link inside a fenced example — or
    inside an HTML comment — is reported. Write an unresolvable path as an
    inline code span, or record it in `docs/validator-link-baseline.txt`. That
    is a loud, fixable cost; the alternative was a quiet one.
    """
    # Blank to the SAME LENGTH, not to one space: `archive_links` rewrites links
    # in place and needs the offsets from this text to address the original.
    return "\n".join(
        _INLINE_CODE.sub(lambda match: " " * len(match.group(0)), line)
        for line in text.splitlines()
    )


def iter_link_destinations(text: str):
    """Yield `(start, end, raw)` for every inline link, parens balanced.

    The span is over the text handed in. `strip_code` blanks to the same length,
    so a span taken from stripped text addresses the original -- which is what
    `archive_links` rewrites through.

    `[a](../x/note(1).md)` is one link with a destination containing
    parentheses, not a truncated one — CommonMark allows balanced pairs.
    """
    for match in _LINK_OPEN.finditer(text):
        index = match.end()
        depth = 1
        chars: list[str] = []
        closed = False
        while index < len(text):
            char = text[index]
            if char == "\\" and index + 1 < len(text):
                chars.append(text[index + 1])
                index += 2
                continue
            if char == "\n":
                break
            if char == "(":
                depth += 1
            elif char == ")":
                depth -= 1
                if depth == 0:
                    closed = True
                    break
            chars.append(char)
            index += 1
            if len(chars) > _MAX_DESTINATION:
                break
        if closed:
            yield match.end(), index, "".join(chars)


def link_target(raw: str) -> Optional[str]:
    """The path a destination points at, or None when it is not a path."""
    from urllib.parse import unquote

    target = raw.strip()
    if target.startswith("<"):
        closing = target.find(">")
        target = target[1:closing] if closing != -1 else target[1:]
    else:
        # A bare destination cannot contain unescaped whitespace, so anything
        # after the first space is a title, not part of the path.
        target = target.split()[0] if target.split() else ""
    for separator in ("#", "?"):
        target = target.split(separator, 1)[0]
    target = target.strip()
    if not target or _URI_SCHEME.match(target):
        return None
    return unquote(target)


def resolves_case_sensitively(base: Path, target: str) -> bool:
    """Whether `target` names an existing entry, matching case exactly.

    `Path.exists()` answers what the filesystem thinks, so a link that differs
    only in case passes on macOS and fails on Linux. A cross-repo check whose
    verdict depends on the machine that ran it is worse than no check.
    """
    current = base
    for part in target.split("/"):
        if part in ("", "."):
            continue
        if part == "..":
            current = current.parent
            continue
        try:
            if part not in os.listdir(current):
                return False
        except OSError:
            return False
        current = current / part
    return current.exists()


def load_link_baseline(root: Path) -> set[tuple[str, Optional[str]]]:
    """Existing link rot recorded as debt, as (file, target-or-None) entries.

    `path<TAB>target` exempts one link and is the form to prefer: a bare `path`
    exempts the whole file, so a link broken there tomorrow would never be
    reported either.
    """
    path = root / LINK_BASELINE_RELPATH
    if not path.is_file():
        return set()
    entries: set[tuple[str, Optional[str]]] = set()
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.split("#", 1)[0].strip()
        if not line:
            continue
        parts = line.split("\t", 1) if "\t" in line else line.split(None, 1)
        entries.add((parts[0], parts[1].strip() if len(parts) == 2 else None))
    return entries



def load_placement_baseline(root: Path) -> set[tuple[str, str]]:
    """既知の誤配置を debt として記録したもの。`path<TAB>heading` の entry 単位。

    file 単位の免除は用意しない——それを許すと、そのファイルに明日足された
    誤配置も黙って通る。link baseline と同じ理由である。
    """
    path = root / PLACEMENT_BASELINE_RELPATH
    if not path.is_file():
        return set()
    entries: set[tuple[str, str]] = set()
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.split("#", 1)[0].rstrip()
        if not line.strip():
            continue
        if "\t" not in line:
            continue
        head, _, tail = line.partition("\t")
        entries.add((head.strip(), tail.strip()))
    return entries


def validate_spec_placement(
    root: Path, errors: list[str], warnings: list[str]
) -> None:
    """`docs/specs/` の見出しが spec の持ち分に収まっているか。

    見つけたものは error。既知分は `docs/validator-placement-baseline.txt` に
    `path<TAB>heading` で記録して免除できるが、**そのリストは縮むだけ**である
    ——解決した行を残すと別の error になる。
    """
    specs_dir = root / "docs/specs"
    if not specs_dir.is_dir():
        return
    baseline = load_placement_baseline(root)
    seen: set[tuple[str, str]] = set()
    for path in sorted(specs_dir.glob("*.md")):
        rel = str(path.relative_to(root))
        try:
            lines = path.read_text(encoding="utf-8").splitlines()
        except OSError as exc:  # pragma: no cover
            errors.append(f"{rel}: unreadable while checking placement ({exc})")
            continue
        fenced = False
        for line in lines:
            if line.lstrip().startswith("```"):
                fenced = not fenced
                continue
            if fenced:
                continue
            match = re.match(r"^#{2,6} (.+?)\s*$", line)
            if not match:
                continue
            heading = match.group(1)
            reason = None
            if any(word in heading for word in PLACEMENT_FORBIDDEN_IN_SPEC):
                reason = "work unit の持ち分（受入条件・ゲート・既存への影響・レビュー結果）"
            elif PLACEMENT_STATUS_IN_HEADING.search(heading):
                reason = "実装状態を名乗る見出しは現況＝guide の持ち分"
            if reason is None:
                continue
            seen.add((rel, heading))
            if (rel, heading) in baseline:
                continue
            errors.append(
                f"{rel}: spec に置けない見出し: '{heading}' — {reason}。"
                f"移すか見出しを改めるか、{PLACEMENT_BASELINE_RELPATH} へ "
                f"'{rel}\t{heading}' を記録すること"
            )
    for rel, heading in sorted(baseline - seen):
        errors.append(
            f"{PLACEMENT_BASELINE_RELPATH}: '{rel}' -> '{heading}' は既に解決している — "
            "行を消すこと（このリストは縮むだけ）"
        )
    if baseline:
        warnings.append(
            f"{PLACEMENT_BASELINE_RELPATH}: {len(baseline)} 件の誤配置がまだ免除されている"
        )


def validate_relative_links(
    root: Path, errors: list[str], warnings: list[str]
) -> None:
    """Every relative link under docs/ must resolve to something that exists.

    Archiving is the routine operation that breaks these: it moves a file one
    level deeper and leaves every referrer pointing at the old path. Measured in
    a single session on one repository, archiving an issue broke 11 links and
    archiving a workstream broke 5 more — and this validator stayed exit 0
    through all of it, because nothing here ever resolved a link. Breakage is
    invisible until a human or an agent actually follows one.

    The rule carries no knowledge of any repository. A path that genuinely
    cannot resolve — a placeholder, or a file outside the repository — is not a
    link and should not be written as one; write it as code instead, which this
    check ignores.

    Not covered, deliberately: reference-style links (`[x][ref]`), HTML `<a
    href>` and `<img src>`. They are rare in this convention's documents and
    each needs a different resolver.

    Known blind spots, both silent, both bounded to what they can reach:

    * an unpaired backtick hides the links after it on that line
      (the span match is lazy so that a link *between* two spans survives;
      `tests/test_doc_update.py` pins that);
    * a baseline entry for a target that appears more than once in a file
      exempts every occurrence of it, including one added later.

    Links inside fenced code blocks and HTML comments are **reported**, not
    ignored. Tracking them was tried and abandoned: across 26 repositories it
    cost 5 loud false positives in 2 of them while repeatedly introducing
    unbounded silent loss, and it never found a link this does not.

    An adopting repository can record existing rot in
    `docs/validator-link-baseline.txt`, one entry per line, preferably as
    `path<TAB>target` so that only the known-broken link is exempt. The list
    only shrinks: an entry that is now resolvable, or whose file is missing, is
    an error.
    """
    baseline = load_link_baseline(root)
    broken: dict[str, list[str]] = {}

    for path in sorted(root.glob("docs/**/*.md")):
        rel = path.relative_to(root).as_posix()
        try:
            text = path.read_text(encoding="utf-8")
        except OSError as exc:
            errors.append(f"{rel}: unreadable while checking links ({exc})")
            continue
        for _start, _end, raw in iter_link_destinations(strip_code(text)):
            target = link_target(raw)
            if target is None:
                continue
            base = root if target.startswith("/") else path.parent
            if not resolves_case_sensitively(base, target.lstrip("/")):
                broken.setdefault(rel, []).append(target)

    exempt_files = {rel for rel, target in baseline if target is None}
    exempt_links = {(rel, target) for rel, target in baseline if target is not None}

    for rel in sorted(broken):
        for target in broken[rel]:
            if rel in exempt_files or (rel, target) in exempt_links:
                continue
            errors.append(f"{rel}: broken relative link: {target}")

    for rel, target in sorted(baseline, key=lambda entry: (entry[0], entry[1] or "")):
        if not (root / rel).is_file():
            errors.append(
                f"{LINK_BASELINE_RELPATH}: '{rel}' is listed but does not exist — "
                "remove the stale entry"
            )
        elif target is None:
            if rel not in broken:
                errors.append(
                    f"{LINK_BASELINE_RELPATH}: '{rel}' no longer has broken links — "
                    "remove it from the baseline so it cannot regress"
                )
        elif target not in broken.get(rel, []):
            errors.append(
                f"{LINK_BASELINE_RELPATH}: '{rel}' -> '{target}' is no longer broken — "
                "remove it from the baseline so it cannot regress"
            )
    if baseline:
        warnings.append(
            f"{LINK_BASELINE_RELPATH}: {len(baseline)} file(s) still exempt from "
            "link resolution"
        )


BASELINE_RELPATH = "docs/validator-baseline.txt"
BASELINED_CHECKS = "runnability / Acceptance の verify:"


def load_baseline(root: Path) -> set[str]:
    """Repo-relative paths exempted from the newer checks. Absent file = no exemptions."""
    path = root / BASELINE_RELPATH
    if not path.is_file():
        return set()
    entries: set[str] = set()
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.split("#", 1)[0].strip()
        if line:
            entries.add(line)
    return entries


def route_baselined(
    rel: str,
    baseline: set[str],
    deferred: dict[str, list[str]],
    errors: list[str],
    produce,
) -> None:
    """Send the newer checks' errors to `deferred` when the file is baselined."""
    collected: list[str] = []
    produce(collected)
    if rel in baseline:
        deferred.setdefault(rel, []).extend(collected)
    else:
        errors.extend(collected)


def validate_baseline_freshness(
    root: Path,
    baseline: set[str],
    deferred: dict[str, list[str]],
    errors: list[str],
    warnings: list[str],
) -> None:
    """Keep the list shrinking: stale or already-compliant entries fail."""
    for rel in sorted(baseline):
        if not (root / rel).is_file():
            errors.append(
                f"{BASELINE_RELPATH}: '{rel}' is listed but does not exist — "
                "remove the stale entry"
            )
        elif not deferred.get(rel):
            errors.append(
                f"{BASELINE_RELPATH}: '{rel}' now satisfies {BASELINED_CHECKS} — "
                "remove it from the baseline so it cannot regress"
            )
    if baseline:
        warnings.append(
            f"{BASELINE_RELPATH}: {len(baseline)} file(s) still exempt from "
            f"{BASELINED_CHECKS}"
        )


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


INDEX_LINK_RE = re.compile(r"\]\([^)\s]+\.md(?:#[^)\s]*)?\)|^[ \t]*[-*+][ \t]+docs/")
# Words that state progress in a routing row; the frontmatter `status` owns that.
INDEX_STATUS_WORD_RE = re.compile(r"✅|完了|済み|未着手|着手済|進行中|blocked|in_progress|\bdone\b|\bWIP\b")


def section_body(body: str, heading: str) -> str:
    """Text of `## <heading>` up to the next `## `, comments stripped ('' when absent)."""
    match = re.search(rf"^##[ \t]+{re.escape(heading)}[ \t]*$", body, re.MULTILINE)
    if not match:
        return ""
    rest = body[match.end():]
    nxt = re.search(r"^##[ \t]", rest, re.MULTILINE)
    section = rest[: nxt.start()] if nxt else rest
    return re.sub(r"<!--.*?-->", "", section, flags=re.DOTALL)


def section_is_empty(body: str, heading: str) -> bool:
    """True when `## <heading>` exists and holds only blanks/comments up to the next `## `."""
    match = re.search(rf"^##[ \t]+{re.escape(heading)}[ \t]*$", body, re.MULTILINE)
    if not match:
        return False
    rest = body[match.end():]
    nxt = re.search(r"^##[ \t]", rest, re.MULTILINE)
    section = rest[: nxt.start()] if nxt else rest
    section = re.sub(r"<!--.*?-->", "", section, flags=re.DOTALL)
    return not section.strip()


def validate_index_size(index_path: Path, errors: list[str], warnings: list[str]) -> None:
    """The routing index must stay small enough to be read whole.

    Measured 2026-09-18: four repositories carried indexes of 34–291 KB, and a
    reader handed one of those was handed a progress dashboard, not a route.
    The injection itself is bounded by index_digest.py (see the note on
    INDEX_MAX_BYTES), so this is a bound on the file a human or an agent opens.
    Long lines are the same defect one row at a time; a low share of routing
    lines says the file has become a dashboard.
    """
    content = index_path.read_text()
    size = len(content.encode())
    if size > INDEX_MAX_BYTES:
        errors.append(
            f"docs/00_index.md: {size // 1024} KB exceeds the {INDEX_MAX_BYTES // 1024} KB "
            "ceiling — move narrative to docs/log/ (docs_hygiene.py --fix does this)"
        )
    elif size > INDEX_TARGET_BYTES:
        warnings.append(
            f"docs/00_index.md: {INDEX_MAX_BYTES - size} bytes from the "
            f"{INDEX_MAX_BYTES // 1024} KB ceiling — the next row added turns the branch "
            "red; run docs_hygiene.py --fix now (it shortens to "
            f"{INDEX_TARGET_BYTES // 1024} KB), or close and archive active work"
        )
    long_lines = [
        n for n, line in enumerate(content.splitlines(), 1) if len(line) > INDEX_MAX_LINE_CHARS
    ]
    if long_lines:
        shown = ", ".join(str(n) for n in long_lines[:5])
        errors.append(
            f"docs/00_index.md: {len(long_lines)} line(s) over {INDEX_MAX_LINE_CHARS} chars "
            f"(lines {shown}{', …' if len(long_lines) > 5 else ''}) — an index row is a link "
            "and one line of routing, not a status report"
        )
    body = content
    end = content.find("\n---", 3) if content.startswith("---") else -1
    if end != -1:
        body = content[end + 4:]
    text_lines = [
        line for line in body.splitlines()
        if line.strip() and not re.match(r"^[ \t]{0,3}#{1,6}[ \t]+", line)
        and not re.match(r"^[ \t]*\|?[ \t]*:?-{3,}", line)
    ]
    status_rows = [
        n for n, line in enumerate(content.splitlines(), 1)
        if INDEX_LINK_RE.search(line) and INDEX_STATUS_WORD_RE.search(line.split("](", 1)[-1])
    ]
    if status_rows:
        shown = ", ".join(str(n) for n in status_rows[:5])
        warnings.append(
            f"docs/00_index.md: {len(status_rows)} row(s) carry status words (lines {shown}"
            f"{', …' if len(status_rows) > 5 else ''}) — the frontmatter status is the only status; "
            "an index row is a link and one line of routing"
        )
    if len(text_lines) >= 10:
        routing = sum(1 for line in text_lines if INDEX_LINK_RE.search(line))
        ratio = 1 - routing / len(text_lines)
        if ratio > INDEX_MAX_PROSE_RATIO:
            warnings.append(
                f"docs/00_index.md: {ratio:.0%} of body lines carry no routing link — "
                "the index is turning into a dashboard; keep prose in docs/log/"
            )


REVIEW_EVERY_DAYS = 14


def validate_review_recency(root: Path, warnings: list[str]) -> None:
    """A repository must have been reviewed (docs_hygiene.py --review) within 14 days.

    Closing work keeps the execution layer honest; nothing else re-weights the
    current-truth layer as the repository grows. A warning, not an error: the
    review is a judgment the validator cannot perform, only demand.
    """
    import datetime as _dt
    log_dir = root / "docs/log"
    newest = None
    if log_dir.is_dir():
        for path in log_dir.glob("review-*.md"):
            m = re.fullmatch(r"review-(\d{8})\.md", path.name)
            if m:
                try:
                    d = _dt.datetime.strptime(m.group(1), "%Y%m%d").date()
                except ValueError:
                    continue
                newest = d if newest is None or d > newest else newest
    today = _dt.date.today()
    if newest is None:
        warnings.append(
            "docs review: never run — python3 docs_hygiene.py <repo> --review re-weights "
            "guides/specs/ADRs by use and lists demote/split candidates"
        )
    elif (today - newest).days > REVIEW_EVERY_DAYS:
        warnings.append(
            f"docs review: last {newest.isoformat()} ({(today - newest).days} days ago) — "
            f"run docs_hygiene.py --review (every {REVIEW_EVERY_DAYS} days)"
        )


def validate_repo(repo: str | Path) -> tuple[list[str], list[str]]:
    root = Path(repo).resolve()
    errors: list[str] = []
    warnings: list[str] = []
    coverage = {"next_actions": {"total": 0, "checked": 0, "skipped": 0}}
    LAST_COVERAGE.clear()
    LAST_COVERAGE.update(coverage)
    baseline = load_baseline(root)
    deferred: dict[str, list[str]] = {}

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
        route_baselined(
            "docs/00_index.md", baseline, deferred, errors,
            lambda sink: validate_index_size(index_path, sink, warnings),
        )

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
            for key in duplicate_keys(fm):
                errors.append(
                    f"{path.relative_to(root)}: front matter '{key}' is declared more "
                    "than once; only the last declaration is read, so merge them into "
                    "one"
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

    specs_dir = root / "docs/specs"
    if specs_dir.is_dir():
        valid_statuses = {"draft", "active", "superseded"}
        # Only the top level: archive/ keeps history verbatim and template/ holds
        # placeholder text that is not a spec.
        for path in sorted(specs_dir.glob("*.md")):
            rel = str(path.relative_to(root))
            fm = parse_front_matter(path)
            if fm is None:
                errors.append(f"{rel}: broken front matter")
                continue
            if not fm:
                # Predates the spec envelope. A warning rather than an error so
                # the legacy population stays visible without blocking, but
                # dropping the envelope from a managed spec cannot go unsaid.
                warnings.append(
                    f"{rel}: no front matter — not validated as a managed spec"
                )
                continue
            spec_id = as_text(fm.get("id", ""))
            if not spec_id.startswith("SPEC-"):
                errors.append(f"{rel}: id must start with SPEC-")
            for key in ("created_at", "updated_at"):
                value = as_text(fm.get(key, ""))
                if key == "created_at" and not value:
                    # references/spec.template.md omits it; only grill's does not.
                    continue
                if not value:
                    errors.append(f"{rel}: {key} is required")
                elif not re.fullmatch(r"\d{4}-\d{2}-\d{2}", value):
                    errors.append(f"{rel}: {key} must be YYYY-MM-DD, got '{value}'")
            # Both spec templates ship `status: draft # draft | active | superseded`,
            # and the front-matter reader does not strip YAML comments — so without
            # this the check would reject a spec created straight from the template.
            status = as_text(fm.get("status", "")).split("#")[0].strip()
            if status not in valid_statuses:
                errors.append(f"{rel}: status must be draft, active, or superseded")
            if status == "superseded" and not as_text(fm.get("superseded_by", "")):
                errors.append(
                    f"{rel}: superseded_by is required when status is superseded"
                )
            for key in ("related_guides", "affected_workstreams", "related_adrs"):
                for reference in as_list(fm.get(key, [])):
                    if reference in id_locations or (root / reference).is_file():
                        continue
                    warnings.append(f"{rel}: {key} ref not found: {reference}")

    branch_owners: dict[str, list[str]] = {}
    issues_dir = root / "docs/issues"
    if issues_dir.is_dir():
        for path in sorted(issues_dir.glob("*.md")):
            rel = str(path.relative_to(root))
            fm = parse_front_matter(path)
            if fm is None:
                errors.append(f"{rel}: broken front matter")
                continue
            status = as_text(fm.get("status", "")).split("#")[0].strip()
            def _status_checks(sink, _s=status, _rel=rel):
                if _s not in ISSUE_STATUSES:
                    sink.append(
                        f"{_rel}: status must be one of {', '.join(ISSUE_STATUSES)}, got '{_s}' "
                        "(docs_hygiene.py --fix normalizes known aliases)"
                    )
                elif _s == "complete":
                    # archive_issue.py / archive_workstream.py exclude this message
                    # from their pre-move gate; it names the move that clears it.
                    sink.append(
                        f"{_rel}: complete but not archived — run archive_issue.py "
                        "(or docs_hygiene.py --fix)"
                    )
            route_baselined(rel, baseline, deferred, errors, _status_checks)
            created = as_text(fm.get("created_at", ""))[:10]
            updated = as_text(fm.get("updated_at", ""))[:10]
            _open_unit = status in ISSUE_STATUSES and status != "complete"
            if _open_unit:
                coverage["next_actions"]["total"] += 1
                # 免除は廃止した。分岐で外れる作業単位は無い。
                coverage["next_actions"]["checked"] += 1
            if _open_unit:
                # The next session resumes from `## Next Actions`; an empty one
                # hands it nothing and it rescans the repository instead.
                #
                # There is no exemption. `updated_at == created_at` used to exempt
                # a fresh file, but that condition is a hand-maintained value: an
                # issue whose updated_at was never touched stayed exempt forever
                # even while being worked on (measured 2026-09-20: 56 of 477 open
                # work units across 13 repositories, unreported for over two weeks).
                # create_issue.py now requires --next-action instead, so its output
                # validates without a dummy step and the exemption is not needed.
                body_text = path.read_text()
                route_baselined(
                    rel, baseline, deferred, errors,
                    lambda sink, _rel=rel, _b=body_text: sink.append(
                        f"{_rel}: '## Next Actions' is empty — an open issue must name the "
                        "very next step (or what unblocks it) so the next session can resume"
                    ) if section_is_empty(_b, "Next Actions") else None,
                )
            if status in ISSUE_STATUSES and status != "complete" and updated != created:
                # The snapshot contract: a resuming session reads the first line of
                # Current Status as "as of <date> — <state>". Prose without a date
                # cannot be told apart from last month's prose.
                first = next(
                    (l.strip() for l in section_body(path.read_text(), "Current Status").splitlines() if l.strip()),
                    "",
                )
                if first and not re.match(r"^(as of|As of)\s+\d{4}-\d{2}-\d{2}", first):
                    warnings.append(
                        f"{rel}: Current Status should open with 'as of YYYY-MM-DD — <one sentence>' "
                        "so the next session knows how fresh the snapshot is"
                    )
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
                body = path.read_text()
                route_baselined(
                    rel, baseline, deferred, errors,
                    lambda sink: validate_acceptance_verify(rel, body, sink),
                )
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
                # and no issue blocks passed clean, so own-goal-run's rule that
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
            validate_envelope_bullets(rel, content, errors)
            issue_blocks = parse_workstream_issue_blocks(content)
            if not issue_blocks:
                errors.append(f"{rel}: no embedded ISSUE-* blocks found")
            # The Issue Queue table is what a human reads and what the loop reads
            # to find the next pending issue, but no gate looked at it. A table
            # row whose block was never written — or whose block a table reformat
            # ate — was invisible, so committed scope could archive as complete.
            tabled = parse_issue_queue_table_ids(content)
            blocked = {issue_id for issue_id, _metadata, _body in issue_blocks}
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
            for issue_id, metadata, block_body in issue_blocks:
                issue_status = metadata.get("status", "")
                if issue_status not in ("pending", "in_progress", "blocked", "complete"):
                    errors.append(
                        f"{rel}#{issue_id}: status must be pending, in_progress, blocked, or complete"
                    )
                def _newer_checks(sink, _m=metadata, _b=block_body, _i=issue_id):
                    if not RUNNABILITY_RE.fullmatch(_m.get("runnability", "")):
                        sink.append(
                            f"{rel}#{_i}: runnability must be 'ready' or "
                            "'gated on <reason>' — executors treat a missing record as "
                            "gated and stop"
                        )
                    validate_acceptance_verify(f"{rel}#{_i}", _b, sink)

                route_baselined(rel, baseline, deferred, errors, _newer_checks)
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

    # A guide or spec is opened whole by the session that needs it, so one file
    # over the index ceiling costs more than the entire routing budget.
    for sub in ("guides", "specs"):
        directory = root / "docs" / sub
        if directory.is_dir():
            for path in sorted(directory.glob("*.md")):
                size = path.stat().st_size
                if size > DOC_SOFT_MAX_BYTES:
                    warnings.append(
                        f"{path.relative_to(root)}: {size // 1024} KB — a guide/spec is read whole; "
                        f"over {DOC_SOFT_MAX_BYTES // 1024} KB split it by task (docs_hygiene.py --review lists sections)"
                    )

    guides_dir = root / "docs/guides"
    if guides_dir.is_dir():
        for path in sorted(guides_dir.glob("*.md")):
            rel = str(path.relative_to(root))
            fm = parse_front_matter(path)
            if fm is None:
                errors.append(f"{rel}: broken front matter")
                continue
            if not fm:
                route_baselined(
                    rel, baseline, deferred, errors,
                    lambda sink, _rel=rel: sink.append(
                        f"{_rel}: no front matter — a guide without updated_at cannot be "
                        "dated or filtered (docs_hygiene.py --fix adds one from git)"
                    ),
                )
                continue
            for key in ("source_issues", "source_workstreams"):
                for reference in as_list(fm.get(key, [])):
                    ref_path = root / reference
                    if reference in id_locations or ref_path.is_file():
                        continue
                    warnings.append(f"{rel}: {key} ref not found: {reference}")

    validate_ownership(root, errors, warnings)
    validate_relative_links(root, errors, warnings)
    validate_spec_placement(root, errors, warnings)
    validate_baseline_freshness(root, baseline, deferred, errors, warnings)
    validate_review_recency(root, warnings)

    return errors, warnings


def validate_ownership(root: Path, errors: list[str], warnings: list[str]) -> None:
    """Issue ownership and routing (SPEC-doc-governance, Workstream Model).

    Every message starts with the path of the file whose change causes it, because the
    pre-commit hook blocks only errors on staged files: a new issue without a route is the
    issue's error, a hand-deleted split-list row is the workstream's, an owned issue left
    in Active Issues is the index's. Files created before ownership.ROLLOUT_DATE only warn.
    """
    scripts_dir = str(Path(__file__).resolve().parent)
    if scripts_dir not in sys.path:
        sys.path.insert(0, scripts_dir)
    import ownership as own
    from index_entries import find_index_entry_lines

    model = own.load_model(root, parse_front_matter)
    index_rel = "docs/00_index.md"
    index_path = root / index_rel
    index_text = index_path.read_text() if index_path.is_file() else ""
    active_span = own.heading_span(index_text, own.HEADINGS[own.ACTIVE_ISSUES])
    active_text = index_text[active_span[0]:active_span[1]] if active_span else ""

    def in_active_issues(issue_id: str) -> bool:
        return bool(find_index_entry_lines(active_text, "docs/issues", issue_id))

    def need(is_new: bool, message: str) -> None:
        (errors if is_new else warnings).append(message)

    for doc in model.issues.values():
        new = own.is_post_rollout(doc.created)
        rel = doc.rel
        if doc.owner is None:
            need(new, f"{rel}: workstream is not declared — set `workstream: WS-...` (owned) "
                      "or `workstream: none` (standalone)")
            if not new and not in_active_issues(doc.doc_id):
                warnings.append(f"{rel}: not routed from Active Issues or any workstream")
            continue
        if doc.due and not own.valid_due(doc.due):
            errors.append(f"{rel}: due must be YYYY-MM-DD or none, got '{doc.due}'")
        if doc.owner != "none":
            if doc.owner not in model.workstreams:
                state = "archived" if doc.owner in model.archived_ws_ids else "missing"
                errors.append(
                    f"{rel}: workstream {doc.owner} is not an active workstream ({state}) — "
                    "reassign it; docs_hygiene.py --fix lists it under Unassigned Issues"
                )
            continue
        if not doc.priority:
            need(new, f"{rel}: priority is required for a standalone issue (high|medium|low)")
        elif doc.priority not in own.PRIORITIES:
            errors.append(f"{rel}: priority must be high, medium or low, got '{doc.priority}'")
        if not doc.due:
            need(new, f"{rel}: due is required for a standalone issue (YYYY-MM-DD or none)")
        if not in_active_issues(doc.doc_id):
            errors.append(
                f"{rel}: standalone issue is not routed from Active Issues in {index_rel} "
                "(create_issue.py writes it; docs_hygiene.py --fix regenerates it)"
            )

    for ws in model.workstreams.values():
        new = own.is_post_rollout(ws.created)
        rel = ws.rel
        if not ws.priority:
            need(new, f"{rel}: priority is required for a workstream (high|medium|low)")
        elif ws.priority not in own.PRIORITIES:
            errors.append(f"{rel}: priority must be high, medium or low, got '{ws.priority}'")
        if not ws.due:
            need(new, f"{rel}: due is required for a workstream (YYYY-MM-DD or none)")
        elif not own.valid_due(ws.due):
            errors.append(f"{rel}: due must be YYYY-MM-DD or none, got '{ws.due}'")

        owned = {d.doc_id for d in model.owned_by(ws.doc_id)}
        block = own.read_block(ws.content, own.SPLIT)
        if block is None:
            if owned:
                errors.append(
                    f"{rel}: issues declaring this workstream have no Split Issues list: "
                    f"{', '.join(sorted(owned))} (docs_hygiene.py --fix generates it)"
                )
            continue
        listed = own.ids_in(block)
        for issue_id in sorted(owned - listed):
            errors.append(f"{rel}: {issue_id} declares this workstream but is missing from Split Issues")
        for issue_id in sorted(listed - owned):
            if issue_id in model.issues:
                reason = f"declares {model.issues[issue_id].owner or 'no workstream'}"
            elif issue_id in model.archived_issue_ids:
                reason = "is archived"
            else:
                reason = "does not exist"
            errors.append(f"{rel}: Split Issues lists {issue_id}, which {reason}")
        if not (owned ^ listed) and block != own.split_rows(model, ws.doc_id):
            warnings.append(f"{rel}: generated Split Issues list is stale — run docs_hygiene.py --fix")

    for doc in model.issues.values():
        if doc.owner in model.workstreams and in_active_issues(doc.doc_id):
            errors.append(
                f"{index_rel}: {doc.doc_id} belongs to {doc.owner} — route it from that "
                "workstream's Split Issues, not Active Issues (docs_hygiene.py --fix removes the row)"
            )

    for name, rows in (
        (own.ACTIVE_ISSUES, own.active_issue_rows(model)),
        (own.UNASSIGNED, own.unassigned_rows(model)),
        (own.ACTIVE_WORKSTREAMS, own.workstream_rows(model)),
    ):
        block = own.read_block(index_text, name)
        if block is not None and block != rows:
            warnings.append(f"{index_rel}: generated {name} list is stale — run docs_hygiene.py --fix")


def main() -> None:
    parser = argparse.ArgumentParser(description="Validate own-doc-update structure.")
    parser.add_argument("repo", nargs="?", default=".", help="Repository root (default: cwd)")
    args = parser.parse_args()

    # Say which repository was read. The default is the current directory, and a
    # closeout reached through an orchestrator stands in a *different* repository —
    # so the orchestrator's own docs could validate clean and be reported as the
    # target's result, with nothing to tell the two apart.
    print(f"validated: {Path(args.repo).resolve()}")
    errors, warnings = validate_repo(args.repo)

    # 何件を検査し、何件を分岐で外したかを常に言う。緑であることと、
    # 見るべきものを見たことは別物で、区別できないと「常に緑」に気づけない。
    na = LAST_COVERAGE.get("next_actions", {})
    if na.get("total"):
        line = (f"coverage: Next Actions {na['checked']}/{na['total']} 件を検査"
                f"（分岐で除外 {na['skipped']}）")
        print(f"  ! {line}" if na["skipped"] else f"  {line}")
    if errors:
        print("ERRORS:")
        for error in errors:
            print(f"  ✗ {error}")
        if any("broken relative link" in error for error in errors):
            # The escape hatch is useless if nobody meeting the errors knows it
            # exists; a check that cannot be brought to green stops being read.
            print(
                f"  note: existing link rot can be recorded as debt in "
                f"{LINK_BASELINE_RELPATH} (one 'path<TAB>target' per line); "
                "the list only shrinks."
            )
    if warnings:
        print("WARNINGS:")
        for warning in warnings:
            print(f"  ⚠ {warning}")
    if not errors and not warnings:
        print("OK: all checks passed.")
    sys.exit(1 if errors else 0)


if __name__ == "__main__":
    main()
