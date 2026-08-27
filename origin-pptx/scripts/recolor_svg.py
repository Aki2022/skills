#!/usr/bin/env python3
"""Material Symbols SVG を tokens 色に再着色する。

image_gen 産アイコンの normalize_icons.py（輝度→アルファ変換）と違い、
SVG は fill 属性の書き換えだけで決定的に再着色できる。

Usage:
    python3 recolor_svg.py <icons_dir> [--colors 404040,44546A,C00000,FFFFFF]

<icons_dir>/*.svg（既に *_c<hex>.svg のものは除外）ごとに
<icons_dir>/<name>_c<hex>.svg を色数ぶん生成する（冪等・上書き）。
"""
import argparse
import re
import sys
from pathlib import Path

DEFAULT_COLORS = ["404040", "44546A", "C00000", "FFFFFF"]


def recolor(svg_text: str, hex_color: str) -> str:
    # Material Symbols は fill 無指定（= black）の <path> のみ。
    # 既存 fill があれば置換、無ければ <svg> ルートに fill を付与する。
    if re.search(r'fill="[^"]*"', svg_text):
        return re.sub(r'fill="[^"]*"', f'fill="#{hex_color}"', svg_text)
    return svg_text.replace("<svg ", f'<svg fill="#{hex_color}" ', 1)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("icons_dir", type=Path)
    ap.add_argument("--colors", default=",".join(DEFAULT_COLORS))
    args = ap.parse_args()

    colors = [c.strip().lstrip("#").upper() for c in args.colors.split(",") if c.strip()]
    sources = [
        p for p in sorted(args.icons_dir.glob("*.svg"))
        if not re.search(r"_c[0-9A-Fa-f]{6}\.svg$", p.name)
    ]
    if not sources:
        print(f"no source svg in {args.icons_dir}", file=sys.stderr)
        return 1

    count = 0
    for src in sources:
        text = src.read_text(encoding="utf-8")
        for color in colors:
            out = src.with_name(f"{src.stem}_c{color}.svg")
            out.write_text(recolor(text, color), encoding="utf-8")
            count += 1
    print(f"recolored: {len(sources)} sources x {len(colors)} colors -> {count} files")
    return 0


if __name__ == "__main__":
    sys.exit(main())
