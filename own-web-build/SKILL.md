---
name: own-web-build
description: "web 成果物（HTML 画面・LP・コンポーネント・ダッシュボード）を作る工場。「UI を作って」「HTML を作って」「LP を作って」「ダッシュボードを作って」「web の見た目を整えて」「ブランドに沿ったサイトを作って」で必ず使う。意匠は repo の ux/design/ と解決済みマニフェストから読み、生成前に L0 を必ず実行する。ux/design/ が無ければ own-design-author へ渡す。画像は own-image-gen に任せる。スライドは own-pptx-build。"
---

# own-web-build — web 成果物の工場（build）

作る skill（build）の名前は own-pptx-build とそろえ、`design` は layer の名前だけに使う。

repo の git toplevel の `ux/design/` を正典として HTML / dashboard を生成する。
値の定義・参照の写しは `own-design-author`、画像生成は `own-image-gen`、スライドは
`own-pptx-build` が担当する。情報設計・メッセージ設計はこの工場の範囲外。

## 処理手順

1. 媒体を判定する。web以外は `own-pptx-build` など該当工場へ渡す。
2. **生成前に毎回、`python3 <このskill>/scripts/preflight.py --repo <repo>` を実行する。**
   この入口は同梱 `own-design-author` のL0を実行する。PATHの別配布を使う場合だけ
   `DESIGN_AUTHOR_COMMAND` に単一の実行ファイルを指定できる。文字列をshell展開しない。
   PythonはL0の既存依存環境を使い、必要なら `DESIGN_AUTHOR_PYTHON` で指定する。
3. 終了コード2・起動失敗・結果不明なら生成しない。`ux/design/` が無ければ
   `own-design-author` を案内し、そこで定義を充填してL0を再実行する。1なら警告を記録して続行できる。
4. `ux/design/DESIGN.md` と L0 が今回生成した `ux/design/.resolved/web.json` / CSS を読む。
   DTCGを工場で再解決しない。倉庫・別のDESIGN.md・既定値から補完しない。
5. 値の発案や参照が必要なら `own-design-author` へ戻す。指示された参照の権利確認は
   `own-brand-resolve` の方針に従い、`proposed` と出所を `ux/design/` に記録する。
   ブランド値をAIが承認しない。
5b. **台本がある成果物（LP・チラシなど）は、台本 → 組み方の計画 → カンプ → HTML の順に作る**（詳細と検査の code は `references/plan-and-comp.md`）。`ux/infoarch/scripts/<台本 id>.yaml` があり、この成果物をoutputsに含む時に使う。台本のない依頼は上の手順を継続し、情報設計を新しく作る依頼は`own-infoarch-author`へ渡す。
   1. 台本を読む：`infoarch-author skeleton --script <id>`で骨組みのHTMLを受け取る（まとまりの印`data-ia-section`・文案の印`data-ia-copy`・固定値と正式名称`data-dict`・宣言した文言`data-ia-text`）。
   2. **台本の門（止まる）**：オーナーに台本を見せ、「進めて」を待つ。台帳には残さない。
   3. 組み方の計画を書く：`ux/build/<成果物 id>/plan.yaml`。まとまりごとに design の組み方の型（解決済みマニフェストの`layout_patterns`）の名前、型の既定と違うときだけ狭い画面の型（`narrow`）を書く。台本の文言は書かない。初回は`templates/build/.gitignore`を`ux/build/.gitignore`に写す（カンプを git に入れない）。
   4. `python3 <このskill>/scripts/plan_check.py --repo <repo> --deliverable <id>`。2なら直す。
   5. 制作の指示`ux/build/<成果物 id>/brief.yaml`（使うトークンの絞り込み・見本の画像・まとまりごとの見た目の要望）を書き、`python3 <このskill>/scripts/comp_request.py --repo <repo> --deliverable <id>`で own-image-gen の器（`comp/fields.json`）と出力先を書く。指示が無いか壊れている、または計画の検査が失格なら書かない。画像は生成しない。`--parts N`でページを N 枚に分ける。`--composition free`は組み方をモデルに提案させる試しで、流れの順がオーナーの決定待ちのため既定は`plan`。
   6. `own-image-gen`を`comp/request.json`の`commands`で流し、広い画面の絵を描く（従量課金。枚数は`parts`、既定 1）。
   7. **カンプの門（止まる）**：オーナーにカンプを見せ、「進めて」を待つ。台帳には残さない。
   8. 計画を直す：カンプが足した細部（重なり・はみ出し・接地・区切り方）をすべて`details`に写し、変えた型を直し、カンプで大きく描いた文字（値段・数字など）は`details`の文でなく`emphasis`（対象の文案か辞書の値の id と、大きくする文字）に書き、`comp.reviewed: true`にして`plan_check.py`を流し直す。カンプの文言と、カンプが新しく描いた絵は使わない。
   9. HTML を自由に組み、検査で角を削る。組む人（AI）を縛るのは 3 つだけ：骨組みの印と文言をそのまま使う、まとまりと文案の順を変えない、事実（数字・主張・声）を作らない。カンプ・計画の型と `details`・design は見本と材料として渡し、部品の中の組み方（カード・札・番号・枠・字の大きさ・強調）は組む人が決める。「計画に従う」「カンプの寸法を写さない」と縛ると、同じ design と絵でも組み方の劣る HTML になった（2026-11-27、伏せた比較 2 回でどちらも自由に組んだ版が選ばれた）。組んだあと 10 の検査を流し、失格した所だけを直す。台本に無いラベル（FAQ の Q・A、番号、「出典」など主張でない文言）が IA_DOM_UNANCHORED_TEXT で失格したら、消さずに台本の `texts` に `kind: label`（`block` 付き）で宣言し、`data-ia-text` を付ける。台本が変わるので、台本の門でオーナーに見せ、L0 を流し直す（流し直すまで L1 と計画の検査はマニフェストが古いと止まる）。主張になる文言は宣言でなくメッセージと文案にする。以下は検査が見る書き方：狭い画面の崩し方は列の規則と同じ強さ（同じ詳細度のセレクタ・同じ層）で書く。骨組みの印を残し、文言・まとまりの順・まとまりの中の文案の順を変えずにスタイリングする（文案を別のまとまりへ動かさない。足すのは属性だけ）。順は DOM の順（読む順）で判定するので、DOM が台本の順のままなら、カードやまとまりを横に並べてよい（縦に積み直さない）。見た目の順を CSS で入れ替える（`order`・`row-reverse`・`column-reverse`・`wrap-reverse`・`grid-auto-flow: dense`・DOM の前の項目より先に置くグリッドの明示の配置）と、`dom` が IA_DOM_CSS_REORDER で失格にする。`position: absolute` などで動かした印は確かめられず、未カバー（IA_DOM_ORDER_UNVERIFIED）として報告に残る。印はCSSで選ばない（L1の`ia-marks-inert`が失格にする）。ロゴやフッターの正式名称は`data-dict="name-…"`で名前だけを包む。どの印にも属さない文言を足すと`dom`がIA_DOM_UNANCHORED_TEXTで失格にするので、足したい文言は台本へ戻す（主張ならメッセージと文案、ナビやラベルなら`texts`）。CSS の`content`で文字や数字を描いても同じく失格（記号と`counter()`の連番は通る）。文案の中の数字などを、文の外にもう一度大きく置く（「大きな数字＋説明の文」）ときは、同じまとまりの中に置き`data-ia-figure="<文案の id>"`を付ける。文字は文案の一部に限り、文案そのものは 1 文のまま残す（own-infoarch-author の contract「文案の抜き出し」）。繰り返しの行動喚起も、台本の別の文案にする。
   10. 下の 7・7b の検査を流す。計画に`emphasis`があれば、`plan_check.py --repo <repo> --deliverable <id> --html <HTML>`も流す（2 つの幅に描いて、強調の文字が残りより大きいかを測る。描けなければ合格にしない）。固定値の一部を大きくするときは、`data-dict`の中を子の要素で区切ってよい（文字は変えない）。
5c. **プロダクト UI（ログイン後の画面・ダッシュボード）の blueprint を使う場合**: `ux/infoarch/blueprint/screens/<画面 id>.yaml` があり、その画面を作るなら、`python3 <このskill>/scripts/build_screen.py --repo <repo> --screen <画面 id> --out <HTML>` で作る。これは `infoarch-author skeleton --screen` の印つき HTML（`data-ia-region`・`data-ia-group`・`data-ia-action`・`data-ia-state`）に、`own-design-author/references/allocation-rules.json`（割り付け規則）から導いた幅の指定と主操作の塗りを足す。blueprint が失格なら作らない。領域の順は優先度のまま動かさず、位置と幅は規則が決める（人が規則に反して置くときは `ux/design/layout-overrides.yaml` に 1 行（id・画面・領域の組・幅・理由。幅は省けない）ずつ書き、オーナーが 1 行ずつ承認する。理由だけでは通らず、承認の無い行は失格。承認依頼の表示と記録は own-design-author の `design-author override-request`・`approve`）。作った後は必ず `python3 <このskill>/scripts/audit_screen.py --screen <画面の yaml> --html <HTML> --repo <repo> [--overrides ux/design/layout-overrides.yaml]` を流す（実ブラウザで PC とスマホの 2 つの幅に描いて、優先度と位置・主操作が 1 つだけ塗られている・宣言した状態が描ける・印がある、を確かめる。印が 0 件なら「確かめていない」で失格）。`--overrides` の各行は、台帳（`--repo` の `ux/approvals.yaml`）にオーナーの現在の承認がある行だけが、その画面・領域の組・幅の逆転を許す（報告に info で残る）。承認の無い行・幅の無い行・写した行・画面や領域が消えた行は失敗で、何も許さない。`--overrides` を渡さなければ、例外は 1 つも効かない。`audit_screen.py` は PyYAML が要る。PATH の `python3` に無いときは、`DESIGN_AUTHOR_PYTHON` の interpreter で自分を 1 回だけ起動し直す（変数が無ければ exit 2 で止まる）。
6. `media/<media>/media.yaml` / `rules.md` と必要なassetを読み、
   `references/context-composer.md` に従って生成する。媒体最低値はマニフェストを読む。
7. 生成後の L1（`scripts/audit_html.py <html> --manifest ux/design/.resolved/web.json`）は確認せず毎回実行する。L1 はコントラストを判定しない（出力に「未カバー」と出る）。描画後の実測は `scripts/measure_render.mjs <url> --profile <id> --manifest ...`（`references/validator.md`）。
   描画の L2 は `references/validator.md` / `site-validation.md` に従い人間へ実施可否を確認する。
   ブラウザ利用時は `own-browser-route` と `playwright-cli` を使う。
7b. 台本を使った場合、出力後に`infoarch-author evaluate --script <id>`を必ず実行する（check → dom → judge を順に走らせ、終了コードを 1 つにまとめる）。終了コード2なら、文言の書き換え（IA_DOM_COPY_DRIFT）・印の欠落・まとまりと文案の順序・文案のまとまり違い・印の外の文言・固定値の表示を直す。終了コード1の警告（未カバーの判定・要確認）は報告に残す。直せない時や未実施は未解決として報告する。
8. 決定論の失格は修正して再検査する。修正は最大2回。失格や未実施が残れば準拠を名乗らない。
9. 出力にL0/L1/L2の結果・提案値数・主要制約・未カバー規則・通さなかった工程と理由を残す。
   フォールバック由来の値を発見したら欠陥として報告し、`ux/design/` へ戻して修正する。

## 参照ファイル

- `references/architecture.md`: 正典と工場の境界。
- `references/context-composer.md`: ux/design/ と媒体制約の合成時に読む。
- `references/validator.md`: L1/L2の実行と報告。
- `references/site-validation.md`: L2実施時のブラウザ手順。
- `references/media-spec.md`: 媒体を追加する場合だけ読む。
- `references/plan-and-comp.md`: 組み方の計画・カンプの依頼・検査の code。
- `references/schema/plan.v1.schema.json`: 組み方の計画の形。
- `references/schema/brief.v1.schema.json`: 制作の指示（カンプの依頼が読む）の形。
- `media/<media>/` / `validators/common/`: 媒体の構造規約と点検項目。

## HTML出力前 必須チェックリスト（MUST — 毎回・例外なし）

HTMLを生成するときは、`media/html/media.yaml` の `type: constraint`（`html-*`）を**必ず**満たす。
ルールの正は media.yaml 一つで、ここには項目を重ねて書かない。省略・後回しにしない（過去に lang / viewport /
landmark / SVGラベル / focus の欠落が頻発したため必須化）。主な対応は次のとおり:

- `html-lang`（`<html lang="ja">`）／ `html-viewport`（viewport meta）
- `html-semantic`（見出し階層・`main`/`header`/`nav`/`footer`）／ `html-heading-order`（見出しを飛ばさない）
- `html-svg-label`（意味を持つ図は `role="img"` + `aria-label`、装飾は `aria-hidden="true"`）
- `html-keyboard`（`:focus-visible` を消さない）／ `html-native-first`（ネイティブ要素優先）
- `html-contrast`（本文 4.5:1 以上）／ `html-alt`（装飾は `alt=""`）

コピペ可能な骨格:

```html
<!doctype html>
<html lang="ja">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>…</title>
    <style>
      :focus-visible {
        outline: var(--focus-width) solid var(--color-primary);
        outline-offset: var(--focus-offset);
      }
    </style>
  </head>
  <body>
    <header>…</header>
    <nav aria-label="主要ナビゲーション">…</nav>
    <main>
      <!-- 意味を持つ図は必ずラベルを付ける -->
      <svg role="img" aria-label="2020–2025年の総人口推移（減少傾向）">…</svg>
    </main>
    <footer>…</footer>
  </body>
</html>
```

この MUST は**生成時に守る規律**であって、実測validate（後述）とは別。validateを実施しない場合でも
MUSTは常に満たす。

## 検査の実施境界

L0とL1は確認せず必ず実行する。L2の描画検査だけ人間に実施可否を確認する。
L2を省略した場合は「未実施」と報告し、静的検査だけで準拠を名乗らない。
コントラスト・ターゲットサイズ等の値は解決済みマニフェストの媒体最低値を読む。
このskillはweb専用。pptxの検査・生成は担当しない。

## Media Resolver（手順1の判断基準）

| 入力の手がかり                                  | 選ぶ media         |
| ----------------------------------------------- | ------------------ |
| Webページ、画面、LP、コンポーネント、React/HTML | `html`                       |
| ダッシュボード、KPI、可視化、指標画面           | `dashboard`                  |
| スライド、プレゼン、提案資料、PPTX、パワポ      | **本 skill の対象外** → `own-pptx-build` へ渡す |
| 明示がなく判断が割れる                          | ユーザーに確認する           |

`media/*/media.yaml` をスキャンして利用可能な媒体一覧を得る（ディレクトリ追加だけで増える）。
現在あるのは `html` と `dashboard` の2つだけ。

## 起動条件

UIを作る / HTMLを作る / LPを作る / ダッシュボードを作る / web の見た目を整える /
ブランドに沿ったサイトを作る — これらの依頼で起動する。

**起動しない**: スライド・PPTX・提案資料（`own-pptx-build`）。意匠の値そのものの問い合わせ
（`own-design-author`。指示された参照元の問い合わせは `own-brand-resolve`）。

## 前提検査が失敗した時

`ux/design/` の欠落・L0失格・マニフェスト不在は `own-design-author` へ戻す。
倉庫やRuntime既定値による代替生成はしない。生成後に見つけた既定値へのフォールバックも欠陥として報告する。
