#!/usr/bin/env python3
"""Material Symbols SVG を vocab_map.json から一括取得して icons_std/ にキャッシュする。

⚠️ ネットワークアクセスを伴う。リポジトリ規約に従い、実行前に人間の確認を取ること
（デッキごとに1回・数百KB・Apache 2.0 LICENSE を同梱保存する）。

標準バリアントは wght300×opsz48（references/material-icons.md）。既定の wght400×opsz24 は
拡大表示で太く黒く見えるため使わない（2026-08-27 人間ラダー選定）。

Usage:
    python3 fetch_material_icons.py <vocab_map.json> <icons_dir> [--weight wght300] [--opsz 48]
    python3 fetch_material_icons.py --symbols group,payments <icons_dir>

既に存在するファイルはスキップ（冪等）。取得後は recolor_svg.py で4色版を生成すること。
"""
import argparse
import json
import sys
import urllib.request
from pathlib import Path

BASE = "https://raw.githubusercontent.com/google/material-design-icons/master"
LICENSE_URL = f"{BASE}/LICENSE"


def fetch(url: str) -> bytes:
    with urllib.request.urlopen(url, timeout=30) as r:  # noqa: S310 — 固定ホストのraw取得
        return r.read()


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("vocab_map", nargs="?", type=Path, help="vocab_map.json のパス")
    ap.add_argument("icons_dir", type=Path)
    ap.add_argument("--symbols", help="vocab_map の代わりに symbol 名をカンマ区切りで指定")
    ap.add_argument("--weight", default="wght300", help="wght100〜wght700。既定 wght300")
    ap.add_argument("--opsz", default="48", choices=["20", "24", "40", "48"])
    args = ap.parse_args()

    if args.symbols:
        symbols = [s.strip() for s in args.symbols.split(",") if s.strip()]
    elif args.vocab_map:
        vocab = json.loads(args.vocab_map.read_text(encoding="utf-8"))["vocab"]
        symbols = sorted({v["symbol"] for v in vocab.values()})
    else:
        ap.error("vocab_map.json か --symbols のどちらかを指定する")

    args.icons_dir.mkdir(parents=True, exist_ok=True)
    wpart = "" if args.weight == "wght400" else f"_{args.weight}"
    fetched = skipped = 0
    for sym in symbols:
        fname = f"{sym}{wpart}_{args.opsz}px.svg"
        out = args.icons_dir / fname
        if out.exists():
            skipped += 1
            continue
        url = f"{BASE}/symbols/web/{sym}/materialsymbolsoutlined/{fname}"
        try:
            out.write_bytes(fetch(url))
            fetched += 1
            print(f"fetched {fname}")
        except Exception as e:
            print(f"FAIL {sym}: {url} ({e})", file=sys.stderr)

    lic = args.icons_dir / "LICENSE"
    if not lic.exists():
        lic.write_bytes(fetch(LICENSE_URL))
        print("fetched LICENSE (Apache 2.0)")
    print(f"done: fetched={fetched} skipped={skipped} dir={args.icons_dir}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
