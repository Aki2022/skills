#!/usr/bin/env python3
"""check_script_parity.py — 二重コピーの collect_codex_images.py が sha256 で一致するかを見る。

段階移行のあいだ、次の2ファイルは**バイト一致**でなければならない
（片方だけ直してもう片方を直し忘れる、を防ぐ）:

  - origin-image-gen/scripts/collect_codex_images.py
  - origin-pptx/scripts/collect_codex_images.py

scope: 本スクリプトが見るのは2つのコピーがバイト一致かどうかだけ。どちらの内容が
正しい（新しい／意図通り）かは判定しない。

exit: 0 一致 / 1 不一致（両方の sha256 を表示） / 2 どちらかのファイルが存在しない

使い方:
  check_script_parity.py [--image-gen <path>] [--pptx <path>]
  （省略時は正典2ファイルの既定パスを使う）
"""
from __future__ import annotations

import argparse
import hashlib
import sys
from pathlib import Path

SCOPE_LINE = (
    "scope: この検査は2つのコピーがバイト一致かのみを見る。どちらが正しいかは判定しない。"
)

# origin-image-gen/scripts/check_script_parity.py から見た正典パス
_SCRIPTS_DIR = Path(__file__).resolve().parent
_ORIGIN_IMAGE_GEN_DIR = _SCRIPTS_DIR.parent
_REPO_ROOT = _ORIGIN_IMAGE_GEN_DIR.parent

DEFAULT_IMAGE_GEN_COPY = _SCRIPTS_DIR / "collect_codex_images.py"
DEFAULT_PPTX_COPY = _REPO_ROOT / "origin-pptx" / "scripts" / "collect_codex_images.py"


def sha256_of(path: Path) -> str:
    h = hashlib.sha256()
    h.update(path.read_bytes())
    return h.hexdigest()


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--image-gen", default=str(DEFAULT_IMAGE_GEN_COPY),
                     help="origin-image-gen 側 collect_codex_images.py のパス")
    ap.add_argument("--pptx", default=str(DEFAULT_PPTX_COPY),
                     help="origin-pptx 側 collect_codex_images.py のパス")
    args = ap.parse_args(argv)

    print(SCOPE_LINE)

    image_gen_path = Path(args.image_gen)
    pptx_path = Path(args.pptx)

    missing = [str(p) for p in (image_gen_path, pptx_path) if not p.is_file()]
    if missing:
        print(f"NG: ファイルが存在しない: {missing}")
        return 2

    image_gen_hash = sha256_of(image_gen_path)
    pptx_hash = sha256_of(pptx_path)

    print(f"  {image_gen_path}: {image_gen_hash}")
    print(f"  {pptx_path}: {pptx_hash}")

    if image_gen_hash == pptx_hash:
        print("OK: sha256 が一致")
        return 0

    print("NG（不一致）: sha256 がズレている")
    return 1


if __name__ == "__main__":
    sys.exit(main())
