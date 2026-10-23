---
name: own-infoarch-author
description: "repo 直下 design/info_arch/ の「伝え方の SSOT」（対象者・課題・メッセージ・証拠・固定値・正式名称・声）を読み書きし、Web サイト・LP・チラシ・ポスター・資料を作るときに、誰に・何を・どの順で・どう話すかをそろえる。「情報設計を作って」「伝え方をそろえて」「SSOT から LP の構成を作って」「メッセージを整理して」「info_arch を検査して」で使う。見た目は own-design-author と工場 skill（own-web-design 等）。"
---

# own-infoarch-author — 伝え方の SSOT と構造化アウトライン

`design/info_arch/` は、repo の伝え方の唯一の正典。視覚の `design/` の兄弟モジュールで、スタイリングではなく
「誰に・何を・どの順で・どれだけ強く・どんな話し方で伝えるか」を持つ。
この skill は、SSOT から成果物ごとの**構造化アウトライン**を作り（作る側）、SSOT とアウトラインの整合を検査する（保守する側）。
工場 skill はアウトラインを受け取って描画するだけで、情報設計の判断はしない。

## 重要ルール

- **承認を書かない。** `info_arch.yaml` の `state` は `empty` / `proposed` だけ。承認は、このファイルの変更をオーナーがレビューした
  PR で表す（CODEOWNERS とブランチ保護）。AI は `info_arch.yaml` の変更を自律マージしない。
- **制作中に SSOT を書き換えない。** 足りない・合わないと分かったら、成果物を直すか、SSOT の変更を別の PR で提案する。
- **固定値は参照で書く。** アウトラインの文案では、数値・連絡先・免責などは `{{fct-NNNN}}` で参照し、値を直接書かない。
- **対象 0 件は失格。** チャネル 0 件・アウトライン 0 件は終了コード 2。

## 使い方

```sh
infoarch-author check --root .              # SSOT・チャネル・アウトラインの決定論の検査。終了コード 0 / 1（警告）/ 2（失格）
infoarch-author impact --root . --id fct-0001  # その原子を参照するアウトラインと成果物を列挙
```

`scripts/infoarch-author` を PATH に置き、PyYAML と jsonschema のある Python を `DESIGN_AUTHOR_PYTHON` で選ぶ
（自動インストールしない。無ければ終了コード 2）。repo の設定にローカルパスを書かない。
新しい repo では `templates/design/info_arch/` を写す。

## 参照

- `references/contract.md` — ディレクトリ・原子・アウトラインの契約、検査の code 一覧、期限の扱い。
- `references/schema/` — `info_arch.yaml`（版 2）・`channels/*.yaml`（版 2）・`outlines/*.yaml`（版 1）の JSON Schema。
- `references/expiry.json` — 証拠・固定値の期限の初期値。
- `fixtures/red-cases.json` — 必須の赤ケース。自己テストは `python3 -m pytest -q <skill>/scripts/tests`。
