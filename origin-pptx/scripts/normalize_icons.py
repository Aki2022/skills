#!/usr/bin/env python3
"""
normalize_icons.py — image_gen 産のラインアイコンを PPTX 埋め込み用に正規化する。

なぜ必須か（2026-07-10 実証・2ラウンド費やした最大の学び）:
  image_gen のラインアイコンは「白の不透明背景・細く淡い線」で出てくるのが常態。
  そのまま埋め込むと (1) 淡色/濃色セルの上で白い四角（bounding box）が露出し、
  (2) 線が薄いため白/淡色の上でも視認不良になる。透過だけでは (2) が残る。
  → **線画は白背景を透過にした上で、単色へ再着色**（輝度→α変換）するのが定型。
    v3 (2026-07-13) の既定色はグレー原則に従い #404040。ポジ/ネガの意味を持つアイコンのみ
    --color 44546A / --color C00000 を明示する。色付きイラスト（人物など）は再着色せず
    白背景の透過のみ行う。

使い方:
  python3 normalize_icons.py <assets_dir> [--color 404040] [--illustrations name1.png,name2.png]
  - <assets_dir> 内の icon_*.png をその場で正規化（元は _opaque_backup/ に退避）
  - --color は再着色hex（既定 404040 = tokens.json gray.heading）
  - --illustrations に列挙したファイルは「色付きイラスト」扱い（再着色せず透過のみ）

補足:
  - 濃色（強調塗り）セルに載せるアイコンは、濃グレー線が背景に沈むため
    ビルド側で小さな白チップ（角丸）を敷いてから重ねる（deck_helpers 参照）。
  - キャンバスが正方形でない場合は先に正方形へ白パディング（macOS: sips --padToHeightWidth）
    しておくと、imgPath(..., aspect=1, ...) で歪まない。
"""
import sys, os, glob, argparse
from PIL import Image
import numpy as np

DEFAULT_COLOR = "404040"  # tokens.json v3 gray.heading

def normalize(path, is_illustration, color):
    im = Image.open(path)
    if im.mode in ("RGBA", "LA", "P"):
        # 冪等性の担保: 既に透過化済みの画像を白地に合成してから処理する。
        # .convert("RGB")直行だと透明部が黒扱いになり全面ベタ塗りに壊れる（2026-07-14実証: 2回目の
        # normalize で17アイコンが単色四角になった）
        rgba = im.convert("RGBA")
        base = Image.new("RGBA", rgba.size, (255, 255, 255, 255))
        base.alpha_composite(rgba)
        im = base
    im = im.convert("RGB")
    a = np.array(im).astype(int)
    L = 0.299 * a[:, :, 0] + 0.587 * a[:, :, 1] + 0.114 * a[:, :, 2]
    out = np.zeros((a.shape[0], a.shape[1], 4), dtype="uint8")
    if is_illustration:
        # 色付きイラスト: 色は保持し、白背景のみ透過
        out[:, :, 0], out[:, :, 1], out[:, :, 2] = a[:, :, 0], a[:, :, 1], a[:, :, 2]
        out[:, :, 3] = np.where(L > 240, 0, 255).astype("uint8")
    else:
        # ライン画: 単色化＋輝度→α（暗い線ほど不透明、白は透明）
        alpha = np.clip((250.0 - L) / 14.0 * 255.0, 0, 255)
        out[:, :, 0], out[:, :, 1], out[:, :, 2] = color
        out[:, :, 3] = alpha.astype("uint8")
    Image.fromarray(out, "RGBA").save(path)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("assets_dir")
    ap.add_argument("--color", default=DEFAULT_COLOR, help="recolor hex for line icons (default 404040)")
    ap.add_argument("--illustrations", default="", help="comma-separated filenames to treat as colored illustrations")
    args = ap.parse_args()
    hexstr = args.color.lstrip("#")
    color = tuple(int(hexstr[i : i + 2], 16) for i in (0, 2, 4))
    illos = {s.strip() for s in args.illustrations.split(",") if s.strip()}
    backup = os.path.join(args.assets_dir, "_opaque_backup")
    os.makedirs(backup, exist_ok=True)
    files = sorted(glob.glob(os.path.join(args.assets_dir, "icon_*.png")))
    if not files:
        print("no icon_*.png found in", args.assets_dir); return
    for f in files:
        name = os.path.basename(f)
        bk = os.path.join(backup, name)
        if not os.path.exists(bk):
            Image.open(f).save(bk)  # keep original once
        normalize(f, name in illos, color)
        print("normalized", name, "(illustration)" if name in illos else f"(line->#{hexstr})")
    print(f"done: {len(files)} icons")

if __name__ == "__main__":
    main()
