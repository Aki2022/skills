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
  - process/mockup_task*.md / icons_task*.md 等 … codexへ渡したタスクファイル（再生成可能）
  - process/icon_NN_*.png / image_N_*.png / persona_*.png
        … codex が --cd 配下へ直接落とす残骸。正本は assets/ の正規化済みPNG
  - process/anchors/ / mockups_v1/ / *_tmp/ … 複製・退避・作業ディレクトリ
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
        f"{p}/fin-*.png",  # ⑤最終レンダリング
        f"{p}/contact_sheet.png",  # ⑤人間レビュー用コンタクトシート
        f"{p}/v[0-9]*-*.png",  # ビルダーが自己検証で作る verifyNN-NN.png 等
        f"{p}/verify*.png",
        f"{p}/*check-*.png",
        f"{p}/*.pdf",
        f"{p}/*.log",
        f"{p}/_gen_dirs_before.txt",
        f"{p}/*backup*.pptx",
        f"{deck}/~$*.pptx",
        # codex が --cd 配下へ直接落とす説明的ファイル名のPNG。
        # SAVE-PATH GOTCHA は「指定パスに保存されない」だが、逆に
        # ワークスペースへ勝手に保存されることもある（2026-08-05 実測: 25枚残った）。
        # 正規化済みの正本は assets/ にあるので process/ 直下のものは残骸。
        f"{p}/icon_[0-9]*.png",
        f"{p}/image_[0-9]*.png",
        f"{p}/persona_*.png",
        f"{p}/mockup_task*.md",
        f"{p}/icons_task*.md",
        f"{p}/assets_task*.md",
        f"{p}/persona_task*.md",
    ]:
        out.extend(sorted(glob.glob(pat)))
    for d in [
        f"{p}/node_modules",
        f"{p}/assets/_opaque_backup",
        f"{p}/anchors",  # スキルから複製したスタイル見本（正典はスキル側）
        f"{p}/mockups_v1",  # 差し替え前モックアップの退避先
        f"{deck}/generated",
        *sorted(glob.glob(f"{p}/icons_task_*/")),
        *sorted(glob.glob(f"{p}/*_tmp/")),  # normalize等の作業ディレクトリ
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
