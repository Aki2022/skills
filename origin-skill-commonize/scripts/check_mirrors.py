#!/usr/bin/env python3
"""check_mirrors.py — 第三者 skill ミラーの同値・鮮度検査（決定論・LLM 不使用）。

使い方:
  python3 check_mirrors.py [skills_root] [--max-age-days N]

検査の範囲（ここに無いものは見ない）:
  parity   : mirrors.yaml の各 dir が local_copy と**バイト同一**か（全ファイルの相対パスと内容の
             sha256）。local_copy が無い entry は SKIP（built-in 等は比較不可）。不一致は FAIL（exit 1）。
  freshness: fetched_at から max_age_days を超えていれば WARN（exit 0）。**上流の最新版は見ない**
             （ネットワークに出ない）ので、上流が更新されたかは分からない。空振りを含む警告として扱う。
exit: 0 = OK/WARN のみ, 1 = FAIL あり, 2 = 台帳が無い・読めない
PyYAML に依存しない（環境に無いため）。台帳は「- key: value」のフラットな list のみを想定する。
"""
from __future__ import annotations

import argparse
import hashlib
import os
import sys
from datetime import date
from pathlib import Path


def parse_manifest(text: str) -> list[dict]:
    entries: list[dict] = []
    cur: dict | None = None
    in_list = False
    for raw in text.splitlines():
        line = raw.split(" #", 1)[0].rstrip() if not raw.lstrip().startswith("#") else ""
        if not line.strip():
            continue
        if line.strip() == "mirrors:":
            in_list = True
            continue
        if not in_list:
            continue
        s = line.strip()
        if s.startswith("- "):
            cur = {}
            entries.append(cur)
            s = s[2:]
        if cur is None or ":" not in s:
            continue
        k, v = s.split(":", 1)
        cur[k.strip()] = v.strip()
    return entries


def tree_digest(root: Path) -> str:
    h = hashlib.sha256()
    for p in sorted(x for x in root.rglob("*") if x.is_file()):
        rel = p.relative_to(root).as_posix()
        h.update(rel.encode())
        h.update(b"\0")
        h.update(p.read_bytes())
        h.update(b"\0")
    return h.hexdigest()


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("root", nargs="?", default=str(Path.home() / ".agents" / "skills"))
    ap.add_argument("--max-age-days", type=int, default=90)
    a = ap.parse_args(argv)
    root = Path(a.root).expanduser()
    manifest = root / "mirrors.yaml"
    print(f"scope: parity = local_copy とのバイト同一（local_copy 無しは SKIP）/ "
          f"freshness = fetched_at から {a.max_age_days} 日で WARN。上流の最新版は見ない。")
    if not manifest.is_file():
        print(f"ERROR manifest not found: {manifest}")
        return 2
    entries = parse_manifest(manifest.read_text())
    if not entries:
        print(f"ERROR manifest has no mirrors: {manifest}")
        return 2
    fail = False
    for e in entries:
        d = e.get("dir", "?")
        target = root / d
        if not target.is_dir():
            print(f"FAIL parity {d}: mirror dir missing: {target}")
            fail = True
            continue
        lc = e.get("local_copy")
        if lc:
            lcp = Path(os.path.expanduser(lc))
            if not lcp.is_dir():
                print(f"SKIP parity {d}: local_copy not present: {lcp}")
            elif tree_digest(target) != tree_digest(lcp):
                print(f"FAIL parity {d}: differs from local_copy {lcp}")
                fail = True
            else:
                print(f"OK parity {d}")
        else:
            print(f"SKIP parity {d}: no local_copy (built-in 等・比較不可)")
        fa = e.get("fetched_at")
        try:
            age = (date.today() - date.fromisoformat(fa)).days if fa else None
        except ValueError:
            age = None
        if age is None:
            print(f"WARN stale {d}: fetched_at missing or unparsable")
        elif age > a.max_age_days:
            print(f"WARN stale {d}: fetched {age} days ago (> {a.max_age_days})")
        else:
            print(f"OK fresh {d}: {age} days")
    print("RESULT:", "FAIL" if fail else "OK")
    return 1 if fail else 0


if __name__ == "__main__":
    sys.exit(main())
