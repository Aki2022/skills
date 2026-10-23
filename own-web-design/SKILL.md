---
name: own-web-design
description: "web 成果物（HTML 画面・LP・コンポーネント・ダッシュボード）を作る工場。「UI を作って」「HTML を作って」「LP を作って」「ダッシュボードを作って」「web の見た目を整えて」「ブランドに沿ったサイトを作って」で必ず使う。意匠は repo の design/ と解決済みマニフェストから読み、生成前に L0 を必ず実行する。design/ が無ければ own-design-author へ渡す。画像は own-image-gen に任せる。スライドは own-pptx-build。"
---

# own-web-design — web 成果物の工場

repo の git toplevel の `design/` を正典として HTML / dashboard を生成する。
値の定義・参照の写しは `own-design-author`、画像生成は `own-image-gen`、スライドは
`own-pptx-build` が担当する。情報設計・メッセージ設計はこの工場の範囲外。

## 処理手順

1. 媒体を判定する。web以外は `own-pptx-build` など該当工場へ渡す。
2. **生成前に毎回、`python3 <このskill>/scripts/preflight.py --repo <repo>` を実行する。**
   この入口は同梱 `own-design-author` のL0を実行する。PATHの別配布を使う場合だけ
   `DESIGN_AUTHOR_COMMAND` に単一の実行ファイルを指定できる。文字列をshell展開しない。
   PythonはL0の既存依存環境を使い、必要なら `DESIGN_AUTHOR_PYTHON` で指定する。
3. 終了コード2・起動失敗・結果不明なら生成しない。`design/` が無ければ
   `own-design-author` を案内し、そこで定義を充填してL0を再実行する。1なら警告を記録して続行できる。
4. `design/DESIGN.md` と L0 が今回生成した `design/.resolved/web.json` / CSS を読む。
   DTCGを工場で再解決しない。倉庫・別のDESIGN.md・既定値から補完しない。
5. 値の発案や参照が必要なら `own-design-author` へ戻す。指示された参照の権利確認は
   `own-brand-resolve` の方針に従い、`proposed` と出所を `design/` に記録する。
   ブランド値をAIが承認しない。
5b. **情報設計のアウトラインを使う場合**: `design/info_arch/outlines/<成果物 id>.yaml`があり、この成果物をoutputsに含むなら、`infoarch-author skeleton --outline <id>`で骨組みのHTMLを受け取る。文言・セクション順・`data-ia-*`属性を変えずにスタイリングする。アウトラインのない依頼は上の手順を継続し、情報設計を新しく作る依頼は`own-infoarch-author`へ渡す。
6. `media/<media>/media.yaml` / `rules.md` と必要なassetを読み、
   `references/context-composer.md` に従って生成する。媒体最低値はマニフェストを読む。
7. 生成後の L1（`scripts/audit_html.py <html> --manifest design/.resolved/web.json`）は確認せず毎回実行する。
   描画の L2 は `references/validator.md` / `site-validation.md` に従い人間へ実施可否を確認する。
   ブラウザ利用時は `own-browser-route` と `playwright-cli` を使う。
7b. アウトラインを使った場合、出力後に`infoarch-author dom --outline <id>`を必ず実行する。終了コード2ならアンカーの欠落・順序・固定値の表示を直す。直せない時や未実施は未解決として報告する。
8. 決定論の失格は修正して再検査する。修正は最大2回。失格や未実施が残れば準拠を名乗らない。
9. 出力にL0/L1/L2の結果・提案値数・主要制約・未カバー規則・通さなかった工程と理由を残す。
   フォールバック由来の値を発見したら欠陥として報告し、`design/` へ戻して修正する。

## 参照ファイル

- `references/architecture.md`: 正典と工場の境界。
- `references/context-composer.md`: design/ と媒体制約の合成時に読む。
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

`design/` の欠落・L0失格・マニフェスト不在は `own-design-author` へ戻す。
倉庫やRuntime既定値による代替生成はしない。生成後に見つけた既定値へのフォールバックも欠陥として報告する。
