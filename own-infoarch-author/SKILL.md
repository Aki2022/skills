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

## 作る：依頼からアウトラインまでの 6 段

依頼を受けたら、次の 6 段を順に行い、`design/info_arch/outlines/<成果物 id>.yaml` を書く。各段では、その段のファイルと、作る媒体のファイルだけを読む。

| 段 | 読むファイル | アウトラインに書くもの |
| --- | --- | --- |
| 1 依頼を固める | `references/steps/01-brief.md` | brief（目的・行動・媒体・必須事項） |
| 2 対象者を選ぶ | `references/steps/02-audience.md` | audience |
| 3 選ぶ・優先順・削る | `references/steps/03-select-and-cut.md` | sections の message と proofs、brief.cut |
| 4 構造を決める | `references/steps/04-structure.md` と媒体のファイル | pages・role |
| 5 文案を書く | `references/steps/05-copy.md` | copy（固定値は `{{fct-NNNN}}`） |
| 6 見直して渡す | `references/steps/06-review.md` | `check` を緑にして工場へ渡す |

媒体のファイル：LP は `references/media/lp.md`、複数ページのサイトは `references/media/site.md`、チラシ・ポスターは
`references/media/flyer-poster.md`、資料は `references/media/deck.md`、アプリの文言は `references/media/app-copy.md`。

- 小さな依頼（単発のチラシなど）は `references/light-process.md` に従って簡略にしてよい。段 1 と段 6 は省かない。
- 新しい repo で SSOT が空なら、先に `references/bootstrap.md` に従って既存の資料から下書きを起こし、オーナーの PR 承認を待つ。
- 例：`references/examples/design/info_arch/` に、同じ SSOT から対象者の違う 2 つの成果物（LP と資料）を作ったアウトラインがある。

ノウハウの数値に「出典なし・目安」とあるものは、検査に使わない。

## 使い方

```sh
infoarch-author check --root .              # SSOT・チャネル・アウトラインの決定論の検査。終了コード 0 / 1（警告）/ 2（失格）
infoarch-author impact --root . --id fct-0001  # その原子を参照するアウトラインと成果物を列挙
infoarch-author dom --root . --outline lp-xxx # 描画物をブラウザで開き、アンカー・順序・固定値の表示を確かめる
```

`scripts/infoarch-author` を PATH に置き、PyYAML と jsonschema のある Python を `DESIGN_AUTHOR_PYTHON` で選ぶ
（自動インストールしない。無ければ終了コード 2）。repo の設定にローカルパスを書かない。
新しい repo では `templates/design/info_arch/` を写す。

## 参照

- `references/contract.md` — ディレクトリ・原子・アウトラインの契約、検査の code 一覧、期限の扱い、描画物の確認（工場が埋めるアンカー）。
- `references/schema/` — `info_arch.yaml`（版 2）・`channels/*.yaml`（版 2）・`outlines/*.yaml`（版 1）の JSON Schema。
- `references/expiry.json` — 証拠・固定値の期限の初期値。
- `fixtures/red-cases.json` — 必須の赤ケース。自己テストは `python3 -m pytest -q <skill>/scripts/tests`。
