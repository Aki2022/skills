#!/usr/bin/env python3
"""
collect_codex_images.py — codex exec が生成した画像を、ログのsession idから決定的に回収する。

なぜ必要か（2026-07-14 実証）: 並列セッションに「generated_images直下の最新を拾え」と
指示すると他セッションの生成物を掴む競合が起きる（11枚中6枚がシャッフルされた実失敗）。
正しい回収 = ログの `session id:` → `$CODEX_HOME/generated_images/<session-id>/` 配下のみを
mtime昇順で並べる（タスクファイルの生成順と一致する）。

使い方:
  python3 collect_codex_images.py <codexログ> <出力名1> <出力名2> ...
  例: python3 collect_codex_images.py process/icons_1.log \
        assets/icon_factory.png assets/icon_rocket.png assets/icon_launchpad.png
生成枚数と出力名の数が一致しない場合はエラー終了する（黙って間引かない）。
"""
import glob
import os
import shutil
import sys


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(2)
    log, outs = sys.argv[1], sys.argv[2:]
    sid = None
    for line in open(log, errors="ignore"):
        if "session id:" in line:
            sid = line.split()[-1].strip()
            break
    if not sid:
        raise SystemExit(f"session id not found in {log}（codexがstdin待ちでハングした可能性。"
                         "起動時に < /dev/null を付けたか確認）")
    codex = os.environ.get("CODEX_HOME", os.path.expanduser("~/.codex"))
    files = sorted(glob.glob(f"{codex}/generated_images/{sid}/*.png"), key=os.path.getmtime)
    if len(files) != len(outs):
        raise SystemExit(f"count mismatch: session {sid} has {len(files)} images, "
                         f"but {len(outs)} output names given")
    for src, dst in zip(files, outs):
        os.makedirs(os.path.dirname(dst) or ".", exist_ok=True)
        shutil.copy(src, dst)
        print(f"{os.path.basename(src)[:16]}... -> {dst}")


if __name__ == "__main__":
    main()
