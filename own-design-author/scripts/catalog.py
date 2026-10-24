#!/usr/bin/env python3
"""Extract the pinned DADS catalog; identity changes require explicit renames.

No classification, network fetch, or model inference occurs here. A match key is
JSON [document, [[heading level, heading text, sibling occurrence], ...]].
"""
from __future__ import annotations

import argparse
import copy
import hashlib
import json
from pathlib import Path
import re
import sys
import zipfile

import yaml

SOURCE_URL = "https://design.digital.go.jp/dads/"
TARGET = re.compile(r"(foundations|components)/[^/]+/index\.md")
SUPPLEMENTS = (
    ("logos", "ロゴ", "ロゴの利用方針と参照するアセットを定義する。"),
    ("image-tone", "画像トーン", "画像のトーンと選定方針を定義する。"),
    ("writing", "ライティング", "文章の口調と表記方針を定義する。"),
    ("motion", "モーション", "動きの方針と動きを減らす設定への対応を定義する。"),
    ("media-profile", "媒体プロファイル", "媒体ごとのブランド値と適用範囲を定義する。"),
    ("asset-ledger", "アセット台帳", "使用するアセットの出所と利用条件を記録する。"),
)


def _match_key(document, headings):
    return json.dumps([document, headings], ensure_ascii=False, separators=(",", ":"))


def _sections(markdown, document):
    lines = markdown.splitlines()
    if not lines or lines[0].strip() != "---":
        raise ValueError(f"missing front matter: {document}")
    try:
        front_end = lines.index("---", 1)
    except ValueError as exc:
        raise ValueError(f"unterminated front matter: {document}") from exc
    metadata = yaml.safe_load("\n".join(lines[1:front_end]))
    if not isinstance(metadata, dict) or not all(isinstance(metadata.get(k), str) and metadata[k] for k in ("title", "source_url")):
        raise ValueError(f"front matter needs title and source_url: {document}")
    positions, h2, h2_counts, h3_counts = [], None, {}, {}
    fence = None
    for index in range(front_end + 1, len(lines)):
        line = lines[index]
        if fence:
            if re.fullmatch(r" {0,3}" + re.escape(fence[0]) + "{" + str(fence[1]) + r",}\s*", line):
                fence = None
            continue
        opening = re.match(r"^ {0,3}(`{3,}|~{3,})(.*)$", line)
        if opening and (opening[1][0] != "`" or "`" not in opening[2]):
            fence = (opening[1][0], len(opening[1]))
            continue
        heading = re.match(r"^ {0,3}(#{1,6})(?:[ \t]+(.*?)|[ \t]*)$", line)
        if not heading:
            continue
        level = len(heading[1])
        if level == 1:
            h2, h3_counts = None, {}
        if level not in (2, 3):
            continue
        title = re.sub(r"[ \t]+#+[ \t]*$", "", heading[2] or "").strip()
        if not title:
            raise ValueError(f"empty heading: {document}:{index + 1}")
        if level == 2:
            h2_counts[title] = h2_counts.get(title, 0) + 1
            h2, h3_counts = [2, title, h2_counts[title]], {}
            path = [h2]
        else:
            h3_counts[title] = h3_counts.get(title, 0) + 1
            path = ([h2] if h2 else []) + [[3, title, h3_counts[title]]]
        positions.append((index, level, title, path))
    rows = []
    for number, (start, level, title, path) in enumerate(positions):
        end = next((p[0] for p in positions[number + 1:] if p[1] <= level), len(lines))
        rows.append({
            "match_key": _match_key(document, path), "document": document,
            "document_title": metadata["title"], "title": title,
            "headings": [part[1] for part in path],
            "text": "\n".join(lines[start + 1:end]).strip(),
            "source_url": metadata["source_url"],
            "evidence_types": ["token", "doc", "rule"], "classification": None,
        })
    return rows


def _ledger_copy(ledger):
    ledger = copy.deepcopy(ledger) if ledger is not None else {"schema_version": 1, "next_id": 1, "entries": []}
    if ledger.get("schema_version") != 1 or type(ledger.get("next_id")) is not int or not isinstance(ledger.get("entries"), list):
        raise ValueError("invalid ledger schema")
    ids, keys, maximum = set(), set(), 0
    historical_keys = set()
    for entry in ledger["entries"]:
        if not isinstance(entry, dict):
            raise ValueError("invalid ledger entry")
        identity, key = entry.get("id"), entry.get("match_key")
        match = re.fullmatch(r"dads-([0-9]{6,})", identity or "")
        if not match or not isinstance(key, str) or not key or type(entry.get("retired")) is not bool:
            raise ValueError("invalid ledger entry")
        if identity in ids or key in keys:
            raise ValueError("duplicate ledger id or match key")
        history = entry.get("previous_match_keys", [])
        if not isinstance(history, list) or any(not isinstance(k, str) or not k for k in history) or len(history) != len(set(history)):
            raise ValueError("invalid ledger match key history")
        if set(history) & (historical_keys | keys) or key in historical_keys or key in history:
            raise ValueError("ambiguous ledger match key history")
        historical_keys.update(history)
        ids.add(identity)
        keys.add(key)
        maximum = max(maximum, int(match[1]))
    if ledger["next_id"] <= maximum:
        raise ValueError("ledger next_id would reuse an issued id")
    return ledger


def generate_catalog(zip_path, version, ledger=None, previous=None, renames=None):
    """Return (catalog, updated ledger, delta), leaving caller inputs untouched.

    renames maps exact old match keys to exact new match keys. No similarity
    heuristic is used. Removed entries stay retired in the append-only id ledger.
    """
    if not isinstance(version, str) or not version.strip():
        raise ValueError("version must be a nonempty string")
    archive = Path(zip_path).read_bytes()
    rows, documents = [], {"foundations": 0, "components": 0}
    with zipfile.ZipFile(zip_path) as source:
        names = [name for name in source.namelist() if TARGET.fullmatch(name)]
        if not names or len(names) != len(set(names)):
            raise ValueError("archive has no target documents or duplicate document paths")
        for document in sorted(names):
            documents[document.split("/")[0]] += 1
            rows.extend(_sections(source.read(document).decode("utf-8-sig"), document))
    if not rows:
        raise ValueError("archive target documents contain no h2/h3 items")
    for slug, title, text in SUPPLEMENTS:
        document = f"supplements/{slug}"
        rows.append({"match_key": _match_key(document, [[2, title, 1]]),
                     "document": document, "document_title": "補遺", "title": title,
                     "headings": [title], "text": text, "source_url": None,
                     "evidence_types": ["token", "doc", "rule"], "classification": None})
    keys = {row["match_key"] for row in rows}
    if len(keys) != len(rows):
        raise ValueError("ambiguous duplicate match keys")
    updated = _ledger_copy(ledger)
    entries = {entry["match_key"]: entry for entry in updated["entries"]}
    historical = {key: entry for entry in updated["entries"] for key in entry.get("previous_match_keys", [])}
    old_keys = {key for key, entry in entries.items() if not entry["retired"]}
    if previous is not None:
        if previous.get("schema_version") != 1 or not isinstance(previous.get("items"), list):
            raise ValueError("invalid previous catalog")
        old_keys = set()
        for row in previous["items"]:
            key = row["match_key"]
            if key in old_keys or key not in entries or row["id"] != entries[key]["id"]:
                raise ValueError("previous catalog and ledger disagree")
            old_keys.add(key)
    renames = {} if renames is None else renames
    if not isinstance(renames, dict) or len(set(renames.values())) != len(renames):
        raise ValueError("renames must be an unambiguous old-key to new-key mapping")
    for old, new in renames.items():
        if (old not in old_keys or old in keys or new not in keys or new in entries
                or (new in historical and historical[new] is not entries[old])):
            raise ValueError("rename needs a disappeared old key and an unissued present new key")
    changed = []
    for old, new in sorted(renames.items()):
        entry = entries.pop(old)
        history = entry.setdefault("previous_match_keys", [])
        if new in history:
            history.remove(new)
        history.append(old)
        entry["match_key"] = new
        entries[new] = entry
        changed.append({"id": entry["id"], "from": old, "to": new})
    previous_items = {item["id"]: item for item in (previous or {}).get("items", [])}
    classification_fields = ("document", "title", "headings", "text", "source_url")
    new_ids = []
    for row in rows:
        key = row["match_key"]
        if key not in entries:
            if key in historical:
                raise ValueError("historical match key requires an explicit rename")
            identity = f"dads-{updated['next_id']:06d}"
            updated["next_id"] += 1
            entry = {"id": identity, "match_key": key, "retired": False}
            updated["entries"].append(entry)
            entries[key] = entry
            new_ids.append(identity)
        row["id"] = entries[key]["id"]
        old_item = previous_items.get(row["id"])
        if old_item is not None and all(row.get(field) == old_item.get(field) for field in classification_fields):
            row["classification"] = copy.deepcopy(old_item.get("classification"))
    for key, entry in entries.items():
        entry["retired"] = key not in keys
    catalog = {
        "schema_version": 1, "version": version,
        "source": {"url": SOURCE_URL, "archive_sha256": hashlib.sha256(archive).hexdigest(), "documents": documents},
        "attribution": "出典：デジタル庁デザインシステムウェブサイト https://design.digital.go.jp/dads/ の Markdown を抽出・加工し、補遺を追加して作成。デジタル庁が作成したカタログではない。",
        "edited": True, "items": rows,
    }
    delta = {"new_ids": new_ids,
             "disappeared_keys": sorted(old_keys - keys - set(renames)),
             "changed_keys": changed}
    return catalog, updated, delta


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)
    generate = subparsers.add_parser("generate")
    for name in ("zip", "version", "output", "ledger"):
        generate.add_argument(f"--{name}", required=True)
    generate.add_argument("--previous")
    generate.add_argument("--renames", help="JSON object mapping exact old match keys to new keys")
    args = parser.parse_args(argv)
    load = lambda path: json.loads(Path(path).read_text(encoding="utf-8")) if path else None
    try:
        catalog, ledger, delta = generate_catalog(args.zip, args.version,
            ledger=load(args.ledger) if Path(args.ledger).exists() else None,
            previous=load(args.previous), renames=load(args.renames))
        for destination, value in ((args.output, catalog), (args.ledger, ledger)):
            Path(destination).write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    except (ValueError, KeyError, TypeError, OSError, zipfile.BadZipFile, yaml.YAMLError) as exc:
        print(f"catalog error: {exc}", file=sys.stderr)
        return 2
    print(json.dumps(delta, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
