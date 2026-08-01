#!/usr/bin/env python3
"""
set_fonts.py — pptx全体のフォントを標準フォント（BIZ UDPGothic）に統一する（ビルド後の必須工程）。

標準フォントは BIZ UDPGothic 一律（2026-07-28 ユーザー決定。mac/share 2プロファイル制は
「Hiragino Sans W4」がPowerPoint実機で正しく解決されず廃止。BIZ UDPGothicは
Windows 10+標準搭載・プロポーショナル・LibreOfficeでも正しく描画されるため、
ビルド〜④検証〜最終成果物まで単一フォントで通せる）。

なぜ後処理も必要か: ビルド（fontFace指定）だけでは、テンプレ注入（inject_template.py）で
持ち込まれるレイアウト・マスター・テーマ（fontScheme）に旧フォントが残る。本スクリプトは
パッケージ内の全XMLの `<a:latin>`/`<a:ea>`/`<a:cs>` を書き換えて統一する。冪等。

使い方:
  python3 set_fonts.py <pptx>                    # 標準（BIZ UDPGothic に統一）
  python3 set_fonts.py <pptx> --ea "..." [--latin "..."]  # 例外的な手動指定
"""
import re
import shutil
import sys
import tempfile
import zipfile

DEFAULT_FONT = "BIZ UDPGothic"


def _reject_dtd(data: bytes, name: str) -> None:
    head = data[:4096]
    if b"<!DOCTYPE" in head or b"<!ENTITY" in head:
        raise SystemExit(f"refusing to parse {name}: DTD/ENTITY declaration found")


def rewrite(xml: str, ea: str, latin: str) -> str:
    # typeface以外の属性（panose/pitchFamily/charset）は旧フォントの記述なので落とす。
    # buFont（箇条書きマーカー用）はタグが別なので対象外＝維持される。
    xml = re.sub(r'<a:latin\b[^>]*?/>', f'<a:latin typeface="{latin}"/>', xml)
    xml = re.sub(r'<a:ea\b[^>]*?/>', f'<a:ea typeface="{ea}"/>', xml)
    xml = re.sub(r'<a:cs\b[^>]*?/>', f'<a:cs typeface="{ea}"/>', xml)
    return xml


def main():
    opts = {}
    args = []
    argv = sys.argv[1:]
    i = 0
    while i < len(argv):
        if argv[i].startswith("--"):
            opts[argv[i]] = argv[i + 1] if i + 1 < len(argv) else None
            i += 2
        else:
            args.append(argv[i])
            i += 1
    if len(args) != 1:
        print(__doc__)
        sys.exit(2)
    path = args[0]

    ea = opts.get("--ea") or DEFAULT_FONT
    latin = opts.get("--latin") or ea

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
    tmp.close()
    n = 0
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp.name, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("ppt/") and item.filename.endswith(".xml"):
                _reject_dtd(data, item.filename)
                new = rewrite(data.decode("utf-8"), ea, latin).encode("utf-8")
                if new != data:
                    n += 1
                data = new
            zout.writestr(item, data)
    shutil.move(tmp.name, path)
    print(f"fonts set: ea={ea!r} latin={latin!r} ({n} parts rewritten)")


if __name__ == "__main__":
    main()
