# design/info_arch/ の契約

この版は runner 非依存の先行分（単体 CLI）。方針は利用側 repo の spec `information-architecture-standardization` に従う。

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
墓標の id は生きた原子にできない。墓標の先（null は削除）が無いと失格。墓標の id も台帳が発行した番号でなければ失格。墓標の id を参照すると失格（先の id を案内する）。

## 終了コード

0 合格（info を含む）/ 1 警告（未カバー）/ 2 失格。JSON 報告に評価日・環境（Python 版・OS）・件数・対象数・未カバーを載せる。

## L0（info_arch/ だけを見る）

| code | 内容 |
| --- | --- |
| IA_MISSING / IA_YAML / IA_SCHEMA | ファイルが無い・YAML として読めない（重複キーを含む。日付は文字列として読む）・スキーマ違反（`state: approved`、outputs の絶対パスと `..` を含む） |
| IA_NO_CHANNELS / IA_CHANNEL_NAME / IA_CHANNEL_FILE | チャネル 0 件・media とファイル名の不一致・`channels/` の `*.yaml` 以外のファイルやサブディレクトリ（黙って無視しない） |
| IA_ID_NOT_ISSUED / IA_ID_TOMBSTONED / IA_TOMBSTONE_TARGET | 台帳と墓標 |
| IA_REF_BROKEN | 参照切れ（墓標なら先を案内） |
| IA_EMPTY_REFERENCED | チャネル、または proposed の原子が空欄の原子を参照する（audience・evidence_claims・addresses・supports・rendition の `{{clm-…}}`・scope.audiences の全参照。参照の表は `l0.atom_refs` の 1 か所） |
| IA_RENDITION_RAW_NUMBER / IA_RENDITION_BAD_PLACEHOLDER | rendition に生の数値（主張は `{{clm-…}}` で参照）・`{{clm-NNNN}}` の形でない `{{…}}` |
| IA_CLAIM_EXPIRED / IA_CLAIM_SCOPE / IA_CLAIM_UNIT | 期限切れ・対象者が適用範囲外・単位が正規単位に無い |
| IA_PRESET_BLOCK / IA_PRESET_ROLE | プリセットの必須ブロックが空・チャネルに必須役割のメッセージが無い |
| IA_TERMS_MISSING / IA_TERMS_MARKERS / IA_TERMS_GENERATED_DRIFT / IA_TERMS_INVALID / IA_TERMS_SPEC | 用語ファイル |
| IA_CLAIMS_UNREFERENCED（info） | どこからも使われていない proposed の主張。終了コードに影響しない |

## L1（各チャネルの outputs のテキスト）

抽出の単位はブロック（段落・見出し・リスト項目・セル・属性値）で、ブロックの全文と各文字の主張アンカーを持つ。
修飾語はアンカーの外側も含むブロック全文から読む（`約<span …>100社</span>` は approx）。改行・`<br>` を挟んだ CJK 文字は
つなぐ（`業界\n初` は「業界初」）。数値と語彙は正規化後の全文、prh は正規化前の全文に当てる。

- HTML: ブロックごとのテキスト（HTML5 の終了タグ省略 `li` `p` `td` `tr` `dt` `dd` `option` は暗黙に閉じる）、
  `alt` `title` `aria-label` `aria-description` `placeholder` `label`、submit の `value`、`meta` の description / og / twitter
  （image:alt を含む）。`script` `style` `template` は除く。JSON-LD は未カバーとして報告する。
- Markdown / Quarto: 段落・見出し・リスト項目・表の行・引用、Pandoc span `[…]{data-ia-claim=…}`、fenced div
  `::: {data-ia-claim=…}`、生 HTML の `alt` 等の属性、front matter（YAML として読む）の title / subtitle / description。
  コードは除く。Markdown 内の生 HTML の `data-ia-*` は読まない（未カバーとして報告し、IA_L1_ANCHOR_UNPARSED で失格にする。Pandoc 属性で書く）。生 HTML の `alt` 等は大文字小文字を問わず読む。front matter の数値の値も読む。HTML の `template` 等に付いたアンカーも検証する。Quarto のショートコード `{{< … >}}` とインラインコード `` `{r} …` `` は計算される値なので未カバー。
- それ以外の形式（TSX・pptx・PDF 等）は未カバー。CSS の `content` は抽出しない（既知の限界）。

| code | 内容 |
| --- | --- |
| IA_L1_NO_TARGETS | チャネルの outputs に一致するファイル 0 件、ファイル単位のテキスト 0 件（描画後にしか文言が無いページ）、全体のテキスト 0 件 |
| IA_L1_DECODE | UTF-8 として読めない成果物 |
| IA_L1_UNCOVERED_FORMAT / IA_L1_UNCOVERED_CONSTRUCT（warn） | 抽出アダプタの無い形式・計算される値や構造化データ |
| IA_L1_UNANCHORED_NUMBER | 有効な主張アンカー（存在し proposed）の外にある主張候補の数値 |
| IA_L1_UNKNOWN_CLAIM / IA_L1_CLAIM_EMPTY / IA_L1_CLAIM_NO_VALUE | アンカーの主張が無い・空欄・正規化値が無い。本文を持たないアンカー（画像だけの要素・空の span・空の div）も検証する。無効なアンカーの中身は未アンカーとして扱う |
| IA_L1_ANCHOR_MALFORMED | アンカー属性の値が `clm-NNNN` などちょうど 1 つの id の形でない（`clm-0001_old`・id を 2 つ並べる・空）、同じ要素に同じアンカー属性が複数、未知の `data-ia-*` 属性 |
| IA_L1_ANCHOR_UNPARSED | Markdown の `data-ia-*` の記述（大文字小文字を問わない。コードは除く）が、読めたアンカー属性より多い。行をまたぐ span・リンク属性・生 HTML などは読まないので、推測せず失格にする |
| IA_L1_REQUIRED_CLAIM_MISSING | チャネルの `required_claims` のアンカーが outputs に無い |
| IA_L1_CLAIM_VALUE_MISMATCH / IA_L1_CLAIM_QUALIFIER_MISMATCH | 値・単位の不一致、修飾（qualifier）の不一致 |
| IA_L1_LEXICON_HIT | 有効な主張アンカーの外の語彙ヒット |
| IA_L1_TERM | prh 規則のヒット（全テキスト） |

数値: NFKC → dash と wave の畳み込み → 非主張パターンを空白で塗る → トークン化。算用数字は単位が無くても候補。
漢数字は単位が続き、かつ 十 / 百 / 千 / 万 / 億 / 兆 を含む時だけ候補（「一緒」「千葉」「一人」は候補にしない）。
単独の漢数字でも 割 / 倍 / 位 が続けば候補（「九割」）。「数万人」「何百社」は値の分からない数として候補になり、
どの登録値とも一致しない。「三社」のような単独の漢数字＋その他の単位は取りこぼす既知の限界。
英字に直接つながる数字（B2B・5G・Windows11）は識別子として候補にしない。英字と空白を挟む数字（Microsoft 365）は
候補に残る（安全側の誤検出。製品名なら主張アンカーで囲む）。範囲（`10〜20社`）は値の組。割は % に換算する。
算用数字＋十 / 百 / 千（`1万2千円`）は一つの値。前置の修飾は重ねて読み（「最大約」）、熟語の一部（「契約」の約）は外す。

語彙は閉じたリストなので、言い換えは素通りする。素通りを「主張が無い」とは報告しない（判定台帳の screened が受け持つ）。

## prh 形式の対応範囲

トップレベル `version: 1` と `rules`。規則は `expected`・`pattern` / `patterns`（文字列はリテラル、`/本体/フラグ` は正規表現で
フラグは g i m u。`\d` `\w` `\b` は JS と同じく ASCII、`\s` は JS と同じ Unicode の空白に広げる。
`i` は ASCII の大小だけを無視する。文字クラス内の `\S` は失格）・`regexpMustEmpty`（`$N`）・`options.wordBoundary`・`specs`。`imports`・`patterns` の省略と空リスト・名前付きグループ・
他のフラグ・未知のキーは失格にする（本家の挙動を近似して黙って違う結果を出さないため）。`specs` は本家と同じく
読み込み時に全件実行し、通らなければ失格。ヒットは置換結果がマッチと異なる箇所。`wordBoundary` は ASCII の単語文字だけを境界の判定に使う
（「jqueryを使う」は当たる）。

entity の生成規則は `expected: name` と、大小文字・全角英数・半角カナ・長音の字形・空白挿入を吸収する正規表現
（英数で始まる / 終わる名前は前後に英数が続く時は当てない）と、`forbidden` の各表記（正表記の一部なら正表記の中では当てない）。

## runner との境界（ISSUE-01 で統合）

この版は単体 CLI。次は `own-design-author` の runner に載せる時に足す: 署名付き承認台帳（`ia:` キー・名前空間 `ia-approval`・
JCS ハッシュ・対象者ダイジェスト）と遅延承認（「参照された」）、署名付き免除、判定台帳（未アンカーの残り span）、
`.resolved/` への IA の解決、成果物マニフェスト、`ship_command` と pre-push の予測マージ木、doctor、5 区分報告。
outputs の宣言を `design.yaml` とチャネル定義のどちらで持つかも統合時に runner の契約に合わせる。
