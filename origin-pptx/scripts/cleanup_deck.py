#!/usr/bin/env python3
"""
cleanup_deck.py — ⑤人間最終承認後のデッキクリーンアップ（成果物ライフサイクルの終端処理）。

## タイミング（重要）

**実行してよいのは「⑤で人間が最終承認した後」だけ**。②〜④の修正サイクル中に中間生成物を
消すと、修正のたびに再生成のオーバーヘッド（imagegen再実行・再検証）が発生する。
修正サイクル中は何も消さない。承認＝クローズ宣言をトリガーに1回だけ実行する。

## 何を消して何を残すか（役割の寿命で判定）

削除（役割終了・再生成可能）:
  - process/mockup_*.png    … ②承認〜⑤承認までの検証用仕様書。⑤以降の正はスクリプト側
  - process/preview-*.png / *.pdf … ④レンダリングの一時物
  - process/*.log / _gen_dirs_before.txt … 実行残骸
  - process/node_modules/   … npm install で再生成可能
  - process/assets/_opaque_backup/ … 正規化前バックアップ（正規化済みが正）
  - generated/ / process/icons_task_*/ … codexが勝手に作る残骸ディレクトリ

保持（修正時の再構築に必要 or 正典）:
  - input/・process/outline.md・slides/・build_deck.js・deck_helpers.js・
    mockup_task*.md・icons_task*.md・builder_brief.md・feedback.json・package*.json
  - process/assets/*.png    … build_deck.js が参照。消すと修正リビルドでアイコンが消える
                              （imagegen非決定的なので再生成すると絵が変わる。最終pptxから
                              抽出復元も可能だが、残す方が安全で安い）
  - process/output.pptx     … 上書き前 cmp ガード（人間編集の検知）の照合元

使い方:
  python3 cleanup_deck.py <デッキディレクトリ>            # dry-run（削除予定を表示するだけ）
  python3 cleanup_deck.py <デッキディレクトリ> --apply    # 実削除
"""
import glob
import os
import shutil
import sys


def targets(deck: str):
    p = os.path.join(deck, "process")
    out = []
    for pat in [
        f"{p}/mockup_*.png",
        f"{p}/anchor_*.png",  # ②でcodexに渡すためのアンカーコピー（正典はスキル側）
        f"{p}/preview-*.png",
        f"{p}/*.pdf",
        f"{p}/*.log",
        f"{p}/_gen_dirs_before.txt",
        f"{p}/*backup*.pptx",
        f"{deck}/~$*.pptx",
    ]:
        out.extend(sorted(glob.glob(pat)))
    for d in [
        f"{p}/node_modules",
        f"{p}/assets/_opaque_backup",
        f"{deck}/generated",
        *sorted(glob.glob(f"{p}/icons_task_*/")),
    ]:
        if os.path.isdir(d):
            out.append(d.rstrip("/"))
    return out


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    apply_ = "--apply" in sys.argv
    if len(args) != 1:
        print(__doc__)
        sys.exit(2)
    deck = args[0].rstrip("/")
    if not os.path.isdir(os.path.join(deck, "process")):
        raise SystemExit(f"not a deck directory (no process/): {deck}")

    items = targets(deck)
    total = 0
    for it in items:
        size = 0
        if os.path.isdir(it):
            for root, _, files in os.walk(it):
                size += sum(os.path.getsize(os.path.join(root, f)) for f in files)
        else:
            size = os.path.getsize(it)
        total += size
        print(f"{'DELETE' if apply_ else 'dry-run'}: {it} ({size/1e6:.1f}MB)")
        if apply_:
            (shutil.rmtree if os.path.isdir(it) else os.remove)(it)
    print(f"{'freed' if apply_ else 'would free'}: {total/1e6:.1f}MB ({len(items)} items)")
    if not apply_:
        print("実削除は --apply を付ける（⑤の人間最終承認後のみ実行すること）")


if __name__ == "__main__":
    main()
