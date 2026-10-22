---
name: own-infoarch-author
description: "repo 直下 design/info_arch/ に置く情報設計（対象者・課題・メッセージ・主張・用語・対象）の正典を作り、決定論で検査する。L0 は info_arch/ 自体の整合、L1 は成果物テキストの主張（数値・修飾・語彙）と用語（prh 形式）。「情報設計を作って」「IA を整えて」「主張を登録して」「数値の表記を検査」「用語の表記ゆれを検査」「info_arch を検査」で使う。成果物の生成は工場 skill（own-web-design 等）、見た目は own-design-author。"
---

# own-infoarch-author — 情報設計の正典と決定論ハーネス

`design/info_arch/` は repo の情報設計の唯一の正典で、`design/` の兄弟モジュール。この skill はスキーマ・
プリセット・テンプレート・検査を持ち、**情報設計の中身（どの主張を言うか）は持たない**。中身は各 repo の
`info_arch.yaml` にあり、承認はオーナーの署名だけが表す。

## 重要ルール

- **承認を書かない。** `info_arch.yaml` の `state` は `empty` / `proposed` だけで、`approved` はスキーマで失格。
  承認台帳 `design/approvals.yaml` はオーナーが署名して書く（runner 統合前は存在しない扱い）。
- **制作中に `info_arch.yaml` とチャネル定義を書き換えない。** 不整合は「成果物の修正」「IA の不備」
  「前提の変更」に分けて報告する。
- **値はコードに持たない。** 数値の正規化表・非主張パターン表・語彙・プリセットは `references/` の JSON。
  誤検出は表に行を足し、足した行には赤と緑のフィクスチャを付けて直す。表に無いことを合格の理由にしない。
- **対象 0 件は失格。** チャネル 0 件・outputs に一致するファイル 0 件・テキスト 0 件は終了コード 2。
- **抽出アダプタの無い形式は未カバー**（終了コード 1、`uncovered` に列挙）。合格とは報告しない。

## 使い方

```sh
infoarch-author check --root . --level l1   # L0 → L1。終了コード 0 / 1（未カバー等）/ 2（失格）
infoarch-author check --root . --level l0   # info_arch/ だけ
infoarch-author gen-terms --root . --write  # terms.prh.yml の entity 生成区間を再生成
infoarch-author normalize "最大1万円(税込)"   # 主張候補トークンと正規化値を見る
```

`scripts/infoarch-author` を PATH に置き、PyYAML と jsonschema のある Python を `DESIGN_AUTHOR_PYTHON` で
選ぶ（自動インストールしない。無ければ終了コード 2）。repo の設定にローカルパスを書かない。

新しい repo では `templates/design/info_arch/` を写し、`ledger.next` から id を発行して原子を `proposed` で起こす。

## 参照

- `references/contract.md` — ディレクトリ契約、L0 / L1 の検査一覧と code、prh 形式の対応範囲、runner との境界。
- `references/schema/` — `info_arch.yaml` と `channels/*.yaml` の JSON Schema（版 1、手書き・出典付き）。
- `references/presets.json` — プリセットごとの必須ブロックと必須役割。
- `references/numeric.json` / `references/non-claim-patterns.json` / `references/claim-lexicon.json` — 数値正規化・非主張・語彙の表。
- `fixtures/red-cases.json` — 必須の赤ケース。自己テストは `python3 -m pytest -q <skill>/scripts/tests`。
