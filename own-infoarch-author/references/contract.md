# design/info_arch/ の契約

実装元: biz_ops `WS-20261022-information-architecture-standardization` ISSUE-01a（runner 非依存の先行分）。
方針の正典は `SPEC-information-architecture-standardization`。

## ディレクトリ

- `info_arch.yaml`: 原子 5 種（`audiences` `problems` `messages` `claims` `entities`）と id 台帳（`ledger.next`・`tombstones`）。
  用語（`trm-`）の番号は台帳が持つが、規則は `terms.prh.yml` にある。
- `terms.prh.yml`: prh 形式。`# >>> own-infoarch-author:generated` から `# <<< own-infoarch-author:generated` までは
  entity の `name` / `forbidden` から `gen-terms` が作る生成区間で、手で直すと L0 失格。区間の外は手書き規則。
- `channels/<media>.yaml`: `media`（ファイル名と一致）・`preset`・`outputs`（repo 相対 glob）・`ship`・`ship_command`
  （`ship: true` なら必須）・`messages`・`required_claims`・`order`・`viewports`。
- `INFO_ARCH.md`: 散文。原子を id で参照するだけで値を持たない。

## id

種別接頭辞（`aud` `prb` `msg` `clm` `trm` `ent`）＋4 桁以上の連番。`ledger.next` 未満でなければ「台帳が発行していない」で失格。
墓標の id は生きた原子にできない。墓標の先が無いと失格。墓標の id を参照すると失格（先の id を案内する）。

## 終了コード

0 合格（info を含む）/ 1 警告（未カバー）/ 2 失格。JSON 報告に評価日・環境（Python 版・OS）・件数・対象数・未カバーを載せる。

## L0（info_arch/ だけを見る）

| code | 内容 |
| --- | --- |
| IA_MISSING / IA_SCHEMA | ファイルが無い・スキーマ違反（`state: approved` を含む） |
| IA_NO_CHANNELS / IA_CHANNEL_NAME | チャネル 0 件・media とファイル名の不一致 |
| IA_ID_NOT_ISSUED / IA_ID_TOMBSTONED / IA_TOMBSTONE_TARGET | 台帳と墓標 |
| IA_REF_BROKEN | 参照切れ（墓標なら先を案内） |
| IA_EMPTY_REFERENCED | チャネルが空欄の原子を使う |
| IA_RENDITION_RAW_NUMBER | rendition に生の数値（主張は `{{clm-…}}` で参照） |
| IA_CLAIM_EXPIRED / IA_CLAIM_SCOPE / IA_CLAIM_UNIT | 期限切れ・対象者が適用範囲外・単位が正規単位に無い |
| IA_PRESET_BLOCK / IA_PRESET_ROLE | プリセットの必須ブロックが空・チャネルに必須役割のメッセージが無い |
| IA_TERMS_MISSING / IA_TERMS_MARKERS / IA_TERMS_GENERATED_DRIFT / IA_TERMS_INVALID / IA_TERMS_SPEC | 用語ファイル |
| IA_CLAIMS_UNREFERENCED（info） | どこからも使われていない proposed の主張。終了コードに影響しない |

## L1（各チャネルの outputs のテキスト）

抽出アダプタ: HTML（ブロックごとのテキスト・`alt` `title` `aria-label` `aria-description` `placeholder`・
submit の `value`・`meta` の description / og / twitter。`script` `style` `template` は除く）と
Markdown / Quarto（行・Pandoc span `[…]{data-ia-claim=…}`・fenced div `::: {data-ia-claim=…}`・front matter の
title / subtitle / description。コードは除く）。それ以外は未カバー。

| code | 内容 |
| --- | --- |
| IA_L1_NO_TARGETS | outputs に一致するファイル 0 件、または全体のテキスト 0 件 |
| IA_L1_UNCOVERED_FORMAT（warn） | 抽出アダプタの無い形式（TSX・pptx・PDF 等） |
| IA_L1_UNANCHORED_NUMBER | 主張アンカーの外にある主張候補の数値 |
| IA_L1_UNKNOWN_CLAIM / IA_L1_CLAIM_NO_VALUE | アンカーの主張が無い・主張に正規化値が無い |
| IA_L1_CLAIM_VALUE_MISMATCH / IA_L1_CLAIM_QUALIFIER_MISMATCH | 値・単位の不一致、修飾（qualifier）の不一致 |
| IA_L1_LEXICON_HIT | 主張アンカーの外の語彙ヒット |
| IA_L1_TERM | prh 規則のヒット（全テキスト） |

数値: NFKC → dash と wave の畳み込み → 非主張パターンを空白で塗る → トークン化。算用数字は単位が無くても候補。
漢数字は単位が続き、かつ 十 / 百 / 千 / 万 / 億 / 兆 を含む時だけ候補（「一緒」「千葉」「一人」は候補にしない。
「三社」のような単独の漢数字＋単位は取りこぼす既知の限界）。範囲（`10〜20社`）は値の組。割は % に換算する。

語彙は閉じたリストなので、言い換えは素通りする。素通りを「主張が無い」とは報告しない（判定台帳の screened が受け持つ）。

## prh 形式の対応範囲

トップレベル `version: 1` と `rules`。規則は `expected`・`pattern` / `patterns`（文字列はリテラル、`/本体/フラグ` は正規表現で
フラグは g i m u）・`regexpMustEmpty`（`$N`）・`options.wordBoundary`・`specs`。`imports`・`patterns` の省略・名前付きグループ・
他のフラグ・未知のキーは失格にする（本家の挙動を近似して黙って違う結果を出さないため）。`specs` は本家と同じく
読み込み時に全件実行し、通らなければ失格。ヒットは置換結果がマッチと異なる箇所。

entity の生成規則は `expected: name` と、大小文字・全角英数・半角カナ・長音の字形・空白挿入を吸収する正規表現
（英数で始まる / 終わる名前は前後に英数が続く時は当てない）と、`forbidden` の各表記（正表記の一部なら正表記の中では当てない）。

## runner との境界（ISSUE-01 で統合）

この版は単体 CLI。次は `own-design-author` の runner に載せる時に足す: 署名付き承認台帳（`ia:` キー・名前空間 `ia-approval`・
JCS ハッシュ・対象者ダイジェスト）と遅延承認（「参照された」）、署名付き免除、判定台帳（未アンカーの残り span）、
`.resolved/` への IA の解決、成果物マニフェスト、`ship_command` と pre-push の予測マージ木、doctor、5 区分報告。
outputs の宣言を `design.yaml` とチャネル定義のどちらで持つかも統合時に runner の契約に合わせる。
