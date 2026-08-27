#!/usr/bin/env python3
"""Material Symbols SVG → 透過PNG（決定的）。

soffice の SVG→PNG は不透明白背景で出力される（2026-08-26 実測）ため、
黒アイコンを白背景でレンダし、輝度→アルファ変換＋単色着色で透過PNGを作る。
入力はクリーンなベクタレンダなので、image_gen 産と違い変換は完全に決定的。

Usage:
    python3 svg_to_png.py <src.svg ...> --out <dir> [--size 384] [--colors 404040,FFFFFF]

出力: <dir>/<stem>_c<hex>_<size>.png
"""
import argparse
import subprocess
import sys
import tempfile
from pathlib import Path

from PIL import Image

DEFAULT_COLORS = ["404040", "44546A", "C00000", "FFFFFF"]


def render_black(src: Path, size: int, workdir: Path) -> Image.Image:
    scaled = workdir / f"{src.stem}_scaled.svg"
    text = src.read_text(encoding="utf-8")
    text = text.replace('height="24"', f'height="{size}"').replace(
        'width="24"', f'width="{size}"'
    )
    scaled.write_text(text, encoding="utf-8")
    subprocess.run(
        ["soffice", "--headless", "--convert-to", "png", "--outdir", str(workdir), str(scaled)],
        check=True,
        capture_output=True,
    )
    png = workdir / f"{scaled.stem}.png"
    return Image.open(png).convert("L")


def colorize(lum: Image.Image, hex_color: str) -> Image.Image:
    r, g, b = (int(hex_color[i : i + 2], 16) for i in (0, 2, 4))
    alpha = lum.point(lambda v: 255 - v)
    out = Image.new("RGBA", lum.size, (r, g, b, 0))
    out.putalpha(alpha)
    return out


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("sources", nargs="+", type=Path)
    ap.add_argument("--out", required=True, type=Path)
    ap.add_argument("--size", type=int, default=384)
    ap.add_argument("--colors", default=",".join(DEFAULT_COLORS))
    args = ap.parse_args()

    colors = [c.strip().lstrip("#").upper() for c in args.colors.split(",") if c.strip()]
    args.out.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as td:
        workdir = Path(td)
        for src in args.sources:
            lum = render_black(src, args.size, workdir)
            for color in colors:
                out = args.out / f"{src.stem}_c{color}_{args.size}.png"
                colorize(lum, color).save(out)
                print(f"wrote {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
