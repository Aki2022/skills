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
5b. **情報設計のアウトラインを使う場合**: `ux/infoarch/outlines/<成果物 id>.yaml`があり、この成果物をoutputsに含むなら、`infoarch-author skeleton --outline <id>`で骨組みのHTMLを受け取る。セクションの印（`data-ia-section`）と文案の印（`data-ia-copy`）、固定値・正式名称（`data-dict`）・宣言した文言（`data-ia-text`）の印を残し、文言・セクションの順・セクションの中の文案の順を変えずにスタイリングする（文案を別のセクションへ動かさない。足すのは属性だけ）。印はCSSで選ばない（L1の`ia-marks-inert`が失格にする。印の付け外しで見た目が変わらないようにするため）。ロゴやフッターの正式名称は`data-dict="name-…"`（辞書の正式名称の id）で名前だけを包む。どの印にも属さない文言を足すと`dom`がIA_DOM_UNANCHORED_TEXTで失格にするので、足したい文言はアウトラインへ戻す（主張ならメッセージと文案、ナビやラベルなら`texts`）。繰り返しの行動喚起も、アウトラインの別の文案にする。アウトラインのない依頼は上の手順を継続し、情報設計を新しく作る依頼は`own-infoarch-author`へ渡す。
5c. **プロダクト UI（ログイン後の画面・ダッシュボード）の blueprint を使う場合**: `ux/infoarch/blueprint/screens/<画面 id>.yaml` があり、その画面を作るなら、`python3 <このskill>/scripts/build_screen.py --repo <repo> --screen <画面 id> --out <HTML>` で作る。これは `infoarch-author skeleton --screen` の印つき HTML（`data-ia-region`・`data-ia-group`・`data-ia-action`・`data-ia-state`）に、`own-design-author/references/allocation-rules.json`（割り付け規則）から導いた幅の指定と主操作の塗りを足す。blueprint が失格なら作らない。領域の順は優先度のまま動かさず、位置と幅は規則が決める（人が規則に反して置くときは `ux/design/layout-overrides.yaml` に理由を書く）。作った後は必ず `python3 <このskill>/scripts/audit_screen.py --screen <画面の yaml> --html <HTML> --repo <repo> [--overrides ux/design/layout-overrides.yaml]` を流す（実ブラウザで PC とスマホの 2 つの幅に描いて、優先度と位置・主操作が 1 つだけ塗られている・宣言した状態が描ける・印がある、を確かめる。印が 0 件なら「確かめていない」で失格）。`audit_screen.py` は PyYAML が要る。PATH の `python3` に無いときは、`DESIGN_AUTHOR_PYTHON` の interpreter で自分を 1 回だけ起動し直す（変数が無ければ exit 2 で止まる）。
6. `media/<media>/media.yaml` / `rules.md` と必要なassetを読み、
   `references/context-composer.md` に従って生成する。媒体最低値はマニフェストを読む。
7. 生成後の L1（`scripts/audit_html.py <html> --manifest ux/design/.resolved/web.json`）は確認せず毎回実行する。L1 はコントラストを判定しない（出力に「未カバー」と出る）。描画後の実測は `scripts/measure_render.mjs <url> --profile <id> --manifest ...`（`references/validator.md`）。
   描画の L2 は `references/validator.md` / `site-validation.md` に従い人間へ実施可否を確認する。
   ブラウザ利用時は `own-browser-route` と `playwright-cli` を使う。
7b. アウトラインを使った場合、出力後に`infoarch-author evaluate --outline <id>`を必ず実行する（check → dom → judge を順に走らせ、終了コードを 1 つにまとめる）。終了コード2なら、文言の書き換え（IA_DOM_COPY_DRIFT）・印の欠落・セクションと文案の順序・文案のセクション違い・印の外の文言・固定値の表示を直す。終了コード1の警告（未カバーの判定・要確認）は報告に残す。直せない時や未実施は未解決として報告する。
8. 決定論の失格は修正して再検査する。修正は最大2回。失格や未実施が残れば準拠を名乗らない。
9. 出力にL0/L1/L2の結果・提案値数・主要制約・未カバー規則・通さなかった工程と理由を残す。
   フォールバック由来の値を発見したら欠陥として報告し、`ux/design/` へ戻して修正する。

## 参照ファイル

- `references/architecture.md`: 正典と工場の境界。
- `references/context-composer.md`: ux/design/ と媒体制約の合成時に読む。
- `references/validator.md`: L1/L2の実行と報告。
- `references/site-validation.md`: L2実施時のブラウザ手順。
- `references/media-spec.md`: 媒体を追加する場合だけ読む。
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
