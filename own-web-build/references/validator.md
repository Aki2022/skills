---
title: Validator 仕様
---

# Validator 仕様

生成物を点検する。**客観的に機械検証できる項目はLLMの自己申告に任せず、実行可能なツールで検証する**
（再現性の担保）。各項目に検証方式を明示する。

L0とL1は確認せず毎回実行する。L2だけ人間へ実施可否を確認する。
L0失格で生成を止め、L1/L2失格で非決定論レビューを止める。
対象0件は失格、検査できない規則は未カバーとして記録する。静的だけで準拠合格を名乗らない。

詳細な実行方法・マニフェストとTailwindの契約は `own-design-author/references/web-validation.md`。

## 検証方式

- `[静的]`（軽量・ほぼ無コスト）: ブラウザを起動せず、ソースを解析して機械判定する。
  `python3 scripts/audit_html.py <file> --manifest ux/design/.resolved/web.json`（リンクされたローカルCSSも追跡）。
  - lang / viewport / landmark（main/header/footer）/ SVG・imgの代替テキスト / focus-visible の有無
  - コントラストは判定しない（描画後の実測だけが持つ。下の「描画後の実測」）
- `[実測]`（コストあり・人間承認後）: `playwright-cli` Skill で実ブラウザを起動し、**computedスタイル**で点検。
  静的では取れない「CSS変数・継承・メディアクエリ適用後の実際の色」「フォーカスリングの実表示」
  「テーマ切替後の各テーマ」を実測する。手順は `references/site-validation.md`。
  ブラウザ操作は `own-browser-route` Skill のポリシーに従う。
- `[自己点検]`: 生成した AI 自身が DESIGN.md / Media Rule と照合して確認する。客観的数値基準がない、
  または実行環境がない項目。

静的で確実に取れるものは静的で済ませ、**変数/継承/テーマ切替が絡む色や表示状態だけ実測に回す**のが
コスト効率がよい。PPTX / その他は実行可能な DOM を持たないため機械validateの対象外（参考Tier）。

## 検証ループの上限

- 不合格項目があれば修正する。**修正は最大2回まで**。
- 2回修正しても不合格なら**打ち切り**、未解決の Validator 項目を明示して成果物とともにユーザーへ報告する
  （無限ループとコスト膨張の回避）。

## Common Validator（媒体横断）

| 項目                               | 方式                |
| ---------------------------------- | ------------------- |
| DESIGN.mdに反していないか          | [自己点検]          |
| 色・フォント・余白が一貫しているか | [自己点検]          |
| コントラストが十分か               | [静的]→[実測]で確定 |
| 日本語の可読性が担保されているか   | [自己点検]          |
| ブランド固有アセットの誤用がないか | [自己点検]          |
| 参照元ブランドをコピーしていないか | [自己点検]          |

詳細チェックリストは `validators/common/*.md` を参照。

## HTML Validator（web完全対応Tier）

| 項目                       | 方式                       |
| -------------------------- | -------------------------- |
| lang / viewport がある     | [静的]                     |
| セマンティック/ landmark   | [静的]                     |
| 画像・SVGに代替テキスト    | [静的]                     |
| focus-visible がある       | [静的]→[実測]で確定        |
| コントラスト（実効値）     | [静的]→[実測]で確定        |
| ARIAが必要な箇所にある     | [実測]（axe等）→[自己点検] |
| キーボード操作が可能か     | [実測]→[自己点検]          |
| レスポンシブ対応があるか   | [実測]→[自己点検]          |
| テーマ切替が全テーマ健全か | [実測]（該当時）           |

`[静的]` は `audit_html.py` で常に取れる。CSS変数・継承・メディアクエリ・テーマ切替が絡む項目は
静的で PASS でも `[実測]` で最終確定する（人間がL2実施を認めた場合）。

## 描画後の実測

描画後の実測（溢れ・重なり・切れ・コントラスト・余白）の道具は `scripts/measure_render.mjs` の 1 つ。当てる検査は
`--profile`（`own-design-author/references/measurement-profiles.json`。当てる検査・当てない検査と理由・幅と配色）が決め、
閾値もそこにある。コントラストの計算は `own-design-author/scripts/contrast_kit.mjs` の 1 か所で、`audit_html.py` は
判定しない（以前は CSS の文字列から静的にも測り、同じ式が 2 か所にあった）。`audit_html.py` の出力は
「コントラストは未カバー」と必ず書き、L1 の合格を文字の読みやすさの保証に読み替えない。`:root[data-theme=...]` による
切替の各テーマは、そのテーマに切り替えて測る（`site-validation.md`）。

変種（ブランド）を持つ design（`ux/design/` のトークンに `variant` がある）は、HTML の `<html>` に
`data-design-variant="b"` のように使う変種を宣言する（両方のブランドを載せるページだけ空白区切りで複数）。
L2 の色・書体・角丸の照合は、宣言した変種の palette（共有トークン＋その変種）で行う。変種が 2 つ以上あるのに
宣言が無い・無い変種名を書いた時は `design-palette` で止まり、全トークンの和集合では照らさない
（取り違えが黙って通るため）。変種が 1 つ以下なら宣言は要らない。トークンへの付け方は `own-design-author/references/contract.md`。

### 文字の列の詰まりすぎ（`text-crowding`）

どの組み方にも当てる床。HTML は自由に組むので、計画に書いた狭い画面の形とは照らさない。文字を載せる最も内側の箱
（p・li・dd・dt・見出し・blockquote・figcaption・td・th・summary）が、内側の幅で全角 8 字（`crowdedMinChars`）に満たず、
2 行以上に折れていれば失格にする。測るのは測った全部の幅（375 も 1280 も）。1 行で収まる短いラベルや数字（01・価格の札など）は
対象外で、縦書きは数えない。報告は `text-crowding`（件数と、タグ・先頭 20 字・1 行の字数・行数の見本）。
直し方は、列を減らす（2 列を 1 列に戻す）・絵や余白を動かして文字に幅を返す・字を小さくしすぎない、のどれか。
閾値は `own-design-author/references/measurement-profiles.json` の `crowdedMinChars` が持ち、根拠と緩めた履歴は
`own-design-author/references/measurement-thresholds.md`。赤と緑の fixture は `own-design-author/fixtures/text-crowding/`。

独立レビュアーは、この実測が緑になってから `scripts/reviewer_gate.py` を通した時だけ起動できる（`own-deliverable-check`・
`own-website-audit`）。測った件数 0・実行の失敗・プロファイルなしでは起動しない。

## PPTX Validator（参考Tier — 機械validateしない）

PPTXは実行可能なDOMを持たないため、本Skillでは機械validateの対象外。以下は生成時の`[自己点検]`
チェックリストとして提示するに留める（実生成・整形は `pptx` Skill 等に委譲）。

- 1スライド1メッセージか / タイトルが結論を示すか / 文字サイズが過小でないか /
  余白と整列 / 図表の可読性 / テンプレートとブランドの一致

## Dashboard Validator（webはHTMLに準拠）

Dashboardをwebで出す場合はHTML Validatorに準拠（静的/実測が使える）。指標系の妥当性は`[自己点検]`。

| 項目                               | 方式       |
| ---------------------------------- | ---------- |
| KPIとチャートの関係が明確か        | [自己点検] |
| 色が意味を持って使われているか     | [自己点検] |
| チャート種別の選択が妥当か         | [自己点検] |
| 単位・期間・出典が明示されているか | [静的]     |
| 誤読を招く軸・スケールでないか     | [自己点検] |

## 出力への記録

手順12で、使用した Plugin・主要制約・**未解決 Validator 項目**（あれば）を成果物に記録する。
