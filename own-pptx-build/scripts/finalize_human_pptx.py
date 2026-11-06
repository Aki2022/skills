#!/usr/bin/env python3
"""
finalize_human_pptx.py — 人間が PowerPoint で編集した pptx から、納品用コピーを作る。

PowerPoint で保存し直すと (1) docProps/core.xml に最終更新者名が入り、(2) レイアウト・マスター・
テーマに Arial 等が戻る（2026-10-19 豊橋デッキで実測。提出 PDF の全ページに ArialMT が埋め込まれた）。
本スクリプトは入力を変更せず、次を順に通して出力へ書く。どれかが失敗したら非 0 で止まる:

  1. コピー（入力と同じパスへは書かない）
  2. docProps/core.xml の lastModifiedBy を空にし、--title があれば dc:title を設定
  3. set_fonts.py でフォントを標準（BIZ UDPGothic）へ統一
  4. sanitize_pptx.py --check
  5. xsd_validate_pptx.py（全部品 XSD 検証）

本文（人間の編集内容）には触れない。誤字・番号体裁の校正は人間か別工程で行う。

使い方:
  python3 finalize_human_pptx.py <入力.pptx> <出力.pptx> [--title "文書タイトル"]
"""
import argparse
import re
import subprocess
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

HERE = Path(__file__).resolve().parent


def patch_core(src: Path, dst: Path, title: str | None) -> None:
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "docProps/core.xml":
                x = data.decode("utf-8")
                x = re.sub(
                    r"<cp:lastModifiedBy>.*?</cp:lastModifiedBy>|<cp:lastModifiedBy\s*/>",
                    "<cp:lastModifiedBy></cp:lastModifiedBy>",
                    x,
                    flags=re.S,
                )
                if title is not None:
                    t = f"<dc:title>{escape(title)}</dc:title>"
                    if re.search(r"<dc:title>.*?</dc:title>|<dc:title\s*/>", x, flags=re.S):
                        x = re.sub(r"<dc:title>.*?</dc:title>|<dc:title\s*/>", t, x, flags=re.S)
                    else:
                        x = x.replace("</cp:coreProperties>", t + "</cp:coreProperties>")
                data = x.encode("utf-8")
            zout.writestr(item, data)


def step(label: str, cmd: list[str]) -> None:
    r = subprocess.run(cmd, capture_output=True, text=True)
    if r.returncode != 0:
        sys.stderr.write(f"[{label}] failed (exit {r.returncode})\n{r.stdout}{r.stderr}")
        raise SystemExit(1)
    print(f"[{label}] ok")


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("src")
    ap.add_argument("dst")
    ap.add_argument("--title")
    a = ap.parse_args()
    src, dst = Path(a.src), Path(a.dst)
    if src.resolve() == dst.resolve():
        raise SystemExit("入力と出力が同じパスです。人間の編集版を上書きしないため別名で指定してください (same path)")
    try:
        patch_core(src, dst, a.title)
    except (zipfile.BadZipFile, OSError) as e:
        dst.unlink(missing_ok=True)
        raise SystemExit(f"[copy] failed: {e}")
    py = sys.executable
    try:
        step("fonts", [py, str(HERE / "set_fonts.py"), str(dst)])
        step("sanitize-check", [py, str(HERE / "sanitize_pptx.py"), "--check", str(dst)])
        step("xsd", [py, str(HERE / "xsd_validate_pptx.py"), str(dst)])
    except SystemExit:
        dst.unlink(missing_ok=True)  # 検証に落ちた出力は残さない
        raise


if __name__ == "__main__":
    main()
