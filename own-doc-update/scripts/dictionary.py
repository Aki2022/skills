#!/usr/bin/env python3
"""The docs dictionary: the one source for official names, fixed values and development terms.

`docs/dictionary/<domain>.json`, one file per domain, so a reader loads only the domain it needs. Each entry has
the correct text, the variants that must not appear, and what it means. A term carries meaning and boundary only:
no link to the spec it came from and no state (SPEC-deliverable-production-layers, Requirement 10), so unknown keys
are refused rather than ignored. A value declares `expires` or `evergreen: true`; there is no implicit lifetime.

The same reader serves two checks: variants in current docs (`check-docs`, also run by validate_repo_docs.py), and
a rendered-or-static HTML deliverable (`check-html`): variants, `data-dict` anchors whose text differs from the
dictionary, and expired values. Standard library only, so every governed repository can run it.

Exit codes: 0 checked and clean, 1 findings, 2 nothing was checked (no dictionary, no documents, no HTML) or the
dictionary itself is malformed. Nothing checked is never green.
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import re
import sys
from html.parser import HTMLParser
from pathlib import Path

DICTIONARY_DIR = Path("docs/dictionary")
SCHEMA_VERSION = 1
KINDS = ("term", "name", "value")
FILE_KEYS = {"schema_version", "domain", "summary", "entries"}
ENTRY_KEYS = {"id", "kind", "text", "meaning", "boundary", "variants", "as_of", "expires", "evergreen"}
REQUIRED = {"term": ("meaning", "boundary"), "name": ("meaning",), "value": ("meaning",)}
VALUE_ONLY = ("as_of", "expires", "evergreen")
ID_RE = re.compile(r"^(term|name|value)-[a-z0-9]+(?:-[a-z0-9]+)*$")
DOMAIN_RE = re.compile(r"^[a-z0-9]+(?:-[a-z0-9]+)*$")
DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
# History is not current truth: decision records, narrative logs and archives keep the words of their time.
SKIP_DOC_PARTS = {"archive"}
SKIP_DOC_TOPS = {"adrs", "log"}
INLINE_CODE = re.compile(r"(`+)(?:(?!\1).)+?\1")
ANCHOR_ATTR = "data-dict"


# ---- reading and shape -------------------------------------------------------------------------------------------

def read_files(root: Path) -> tuple[dict[str, object], list[str]]:
    """{file name: parsed JSON} under docs/dictionary/, and errors for files that are not JSON."""
    folder = Path(root) / DICTIONARY_DIR
    data, errors = {}, []
    if not folder.is_dir():
        return data, errors
    for path in sorted(folder.glob("*.json")):
        try:
            data[path.name] = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, ValueError) as exc:
            errors.append(f"{DICTIONARY_DIR / path.name}: not readable JSON ({exc})")
    return data, errors


def _date_ok(value) -> bool:
    if not isinstance(value, str) or not DATE_RE.match(value):
        return False
    try:
        dt.date.fromisoformat(value)
    except ValueError:
        return False
    return True


def shape_errors(files: dict[str, object]) -> list[str]:
    """Every way the dictionary breaks its shape, each prefixed with its file. Empty when the shape holds."""
    errors: list[str] = []
    seen_ids: dict[str, str] = {}
    texts: dict[str, str] = {}
    variants: list[tuple[str, str, str]] = []  # (file, id, variant)
    for name, data in sorted(files.items()):
        where = str(DICTIONARY_DIR / name)
        if not isinstance(data, dict):
            errors.append(f"{where}: top level must be an object")
            continue
        for key in sorted(set(data) - FILE_KEYS):
            errors.append(f"{where}: unknown key {key!r}")
        if data.get("schema_version") != SCHEMA_VERSION:
            errors.append(f"{where}: schema_version must be {SCHEMA_VERSION}")
        stem = Path(name).stem
        domain = data.get("domain")
        if not isinstance(domain, str) or not DOMAIN_RE.match(domain):
            errors.append(f"{where}: domain must be a lower-case slug")
        elif domain != stem:
            errors.append(f"{where}: domain {domain!r} must equal the file name {stem!r}")
        if not isinstance(data.get("summary"), str) or not data["summary"].strip():
            errors.append(f"{where}: 'summary' must be a non-empty string")
        entries = data.get("entries")
        if not isinstance(entries, list) or not entries:
            errors.append(f"{where}: 'entries' must be a non-empty list (an empty dictionary checks nothing)")
            continue
        for i, entry in enumerate(entries):
            if not isinstance(entry, dict):
                errors.append(f"{where}: entries[{i}] must be an object")
                continue
            eid = entry.get("id") if isinstance(entry.get("id"), str) else f"entries[{i}]"
            at = f"{where}: {eid}"
            for key in sorted(set(entry) - ENTRY_KEYS):
                errors.append(f"{at}: unknown key {key!r} (an entry has meaning and boundary only, no source or state)")
            kind = entry.get("kind")
            if kind not in KINDS:
                errors.append(f"{at}: kind must be one of {list(KINDS)}")
            m = ID_RE.match(eid)
            if not m:
                errors.append(f"{at}: id must look like <kind>-<slug> (lower-case, hyphenated)")
            elif kind in KINDS and m.group(1) != kind:
                errors.append(f"{at}: id prefix {m.group(1)!r} does not match kind {kind!r}")
            if eid in seen_ids:
                errors.append(f"{at}: duplicate id {eid} (also in {seen_ids[eid]})")
            seen_ids.setdefault(eid, where)
            text = entry.get("text")
            if not isinstance(text, str) or not text.strip():
                errors.append(f"{at}: 'text' must be a non-empty string")
            else:
                texts.setdefault(text, eid)
            for key in REQUIRED.get(kind, ()):
                if not isinstance(entry.get(key), str) or not entry[key].strip():
                    errors.append(f"{at}: kind {kind} needs {key!r}")
            vs = entry.get("variants", [])
            if not isinstance(vs, list) or not all(isinstance(v, str) and v.strip() for v in vs):
                errors.append(f"{at}: 'variants' must be a list of non-empty strings")
            else:
                if len(set(vs)) != len(vs):
                    errors.append(f"{at}: 'variants' has duplicates")
                variants += [(where, eid, v) for v in vs]
            if kind == "value":
                has_exp, has_ever = "expires" in entry, entry.get("evergreen") is True
                if has_exp == has_ever or ("evergreen" in entry and not has_ever):
                    errors.append(f"{at}: kind value needs exactly one of 'expires' or 'evergreen: true'")
                for key in ("as_of", "expires"):
                    if key in entry and not _date_ok(entry[key]):
                        errors.append(f"{at}: {key} must be YYYY-MM-DD")
            else:
                for key in VALUE_ONLY:
                    if key in entry:
                        errors.append(f"{at}: {key!r} belongs to kind value only")
    for where, eid, v in variants:
        owner = texts.get(v)
        if owner is not None:
            errors.append(f"{where}: {eid}: variant {v!r} is the correct text of {owner}")
    return errors


def entries_of(files: dict[str, object]) -> list[dict]:
    return [e for _, d in sorted(files.items()) for e in d["entries"]]


def load(root: Path) -> tuple[list[dict], int, list[str]]:
    """(entries, file count, errors). Entries are only returned when the shape holds."""
    files, errors = read_files(root)
    errors += shape_errors(files)
    return ([] if errors else entries_of(files)), len(files), errors


def expiry(entry: dict):
    return dt.date.fromisoformat(entry["expires"]) if entry.get("kind") == "value" and entry.get("expires") else None


def expired(entries: list[dict], date: dt.date) -> list[dict]:
    return [e for e in entries if (x := expiry(e)) is not None and x < date]


# ---- variant matching --------------------------------------------------------------------------------------------

def _pattern(variant: str) -> re.Pattern:
    # An ASCII variant matches as a word ("Layer" must not fire inside "Layered"); Japanese has no word boundary.
    if variant.isascii():
        return re.compile(r"(?<![A-Za-z0-9_])" + re.escape(variant) + r"(?![A-Za-z0-9_])")
    return re.compile(re.escape(variant))


def variant_hits(entries: list[dict], text: str):
    """(entry, variant) for each variant present in text."""
    for e in entries:
        for v in e.get("variants", []):
            if _pattern(v).search(text):
                yield e, v


# ---- docs --------------------------------------------------------------------------------------------------------

def current_docs(root: Path) -> list[Path]:
    docs = Path(root) / "docs"
    out = []
    for path in sorted(docs.rglob("*.md")):
        rel = path.relative_to(docs)
        if rel.parts[0] in SKIP_DOC_TOPS or SKIP_DOC_PARTS & set(rel.parts[:-1]):
            continue
        out.append(path)
    return out


def docs_findings(root: Path, entries: list[dict]) -> tuple[list[str], int]:
    """(findings, documents read). Inline code spans are not read, so a variant can be quoted to be named."""
    findings, paths = [], current_docs(root)
    for path in paths:
        rel = path.relative_to(root)
        for n, line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
            plain = INLINE_CODE.sub(" ", line)
            for e, v in variant_hits(entries, plain):
                findings.append(f"{rel}:{n}: 「{v}」は {e['id']} のゆれ（正しい表記は「{e['text']}」）")
    return findings, len(paths)


# ---- html --------------------------------------------------------------------------------------------------------

class _Text(HTMLParser):
    """Visible text (script, style and template contents excluded) and the text of each data-dict anchor."""

    SKIP = {"script", "style", "template", "noscript"}
    VOID = {"area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "source", "track", "wbr"}

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.parts: list[str] = []
        self.stack: list[tuple[str, int | None]] = []  # (tag, index into anchors or None)
        self.anchors: list[list] = []  # [id, text parts]
        self.skip = 0

    def handle_starttag(self, tag, attrs):
        if tag in self.VOID:
            return
        if tag in self.SKIP:
            self.skip += 1
        aid = dict(attrs).get(ANCHOR_ATTR)
        idx = None
        if aid is not None:
            self.anchors.append([aid, []])
            idx = len(self.anchors) - 1
        self.stack.append((tag, idx))

    def handle_endtag(self, tag):
        while self.stack:
            t, _ = self.stack.pop()
            if t in self.SKIP:
                self.skip -= 1
            if t == tag:
                break

    def handle_data(self, data):
        if self.skip:
            return
        self.parts.append(data)
        for _, idx in self.stack:
            if idx is not None:
                self.anchors[idx][1].append(data)


def _norm(text: str) -> str:
    return " ".join(text.split())


def html_findings(path: Path, entries: list[dict], date: dt.date) -> tuple[list[str], int]:
    """(findings, anchors read) for one HTML file."""
    parser = _Text()
    parser.feed(Path(path).read_text(encoding="utf-8"))
    parser.close()
    by_id = {e["id"]: e for e in entries}
    text = _norm(" ".join(parser.parts))
    out = []
    for e, v in variant_hits(entries, text):
        out.append(f"{path}: DICT_VARIANT 「{v}」は {e['id']} のゆれ（正しい表記は「{e['text']}」）")
    stale = {e["id"]: e for e in expired(entries, date)}
    for aid, parts in parser.anchors:
        shown = _norm("".join(parts))
        e = by_id.get(aid)
        if e is None:
            out.append(f"{path}: DICT_UNKNOWN {ANCHOR_ATTR}={aid!r} は辞書に無い")
            continue
        if shown != _norm(e["text"]):
            # The shown text is not echoed for values: only which entry differs.
            out.append(f"{path}: DICT_VALUE_MISMATCH {aid} の表示が辞書の「{e['text']}」と違う")
        if aid in stale:
            out.append(f"{path}: DICT_EXPIRED {aid} の期限 {e['expires']} が過ぎている（{date} 時点）")
    anchored = {aid for aid, _ in parser.anchors}
    for aid, e in sorted(stale.items()):
        if aid not in anchored and _norm(e["text"]) in text:
            out.append(f"{path}: DICT_EXPIRED {aid} の値「{e['text']}」が印の外に出ている（期限 {e['expires']}、{date} 時点）")
    return out, len(parser.anchors)


# ---- cli ---------------------------------------------------------------------------------------------------------

def _load_or_exit(root: Path) -> tuple[list[dict], int]:
    entries, nfiles, errors = load(root)
    if nfiles == 0 and not errors:
        print(f"not checked: no dictionary files under {DICTIONARY_DIR}/ in {root}")
        sys.exit(2)
    if errors:
        print("ERRORS:")
        for err in errors:
            print(f"  ✗ {err}")
        sys.exit(2)
    return entries, nfiles


def _finish(findings: list[str], summary: str) -> None:
    for f in findings:
        print(f"  ✗ {f}")
    print(summary)
    sys.exit(1 if findings else 0)


def main(argv=None) -> None:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    sub = parser.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("validate", help="check the dictionary's shape")
    p.add_argument("repo")
    p = sub.add_parser("check-docs", help="variants in current docs")
    p.add_argument("repo")
    p = sub.add_parser("check-html", help="variants, anchors and expired values in HTML deliverables")
    p.add_argument("repo")
    p.add_argument("html", nargs="*")
    p.add_argument("--date", default=dt.date.today().isoformat(), help="the day expiry is judged on (YYYY-MM-DD)")
    args = parser.parse_args(argv)
    root = Path(args.repo)
    entries, nfiles = _load_or_exit(root)
    if args.cmd == "validate":
        _finish([], f"checked: dictionary {len(entries)} entries in {nfiles} files")
    elif args.cmd == "check-docs":
        findings, ndocs = docs_findings(root, entries)
        if ndocs == 0:
            print("not checked: no current docs to read (docs/**/*.md outside adrs, log and archive)")
            sys.exit(2)
        _finish(findings, f"checked: docs {ndocs} files against {len(entries)} entries")
    else:
        if not args.html:
            print("not checked: no HTML files given")
            sys.exit(2)
        date = dt.date.fromisoformat(args.date)
        findings, nanchors = [], 0
        for h in args.html:
            f, n = html_findings(Path(h), entries, date)
            findings += f
            nanchors += n
        _finish(findings, f"checked: html {len(args.html)} files, {nanchors} anchors against {len(entries)} entries")


if __name__ == "__main__":
    main()
