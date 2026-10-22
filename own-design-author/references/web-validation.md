# 解決済みマニフェストによるweb検査

実装元: biz_ops `WS-20261021-design-system-standardization` ISSUE-02 / ISSUE-03。

## 実行

L0を実行してから、工場は次を使う。L0/L1は確認不要、L2は人間の実施確認後に走らせる。

```sh
design-author check --repo .
python3 <own-web-design>/scripts/audit_html.py index.html --manifest design/.resolved/web.json --tailwind tailwind.config.cjs
node <own-website-audit>/scripts/measure_lp.mjs <url> --manifest design/.resolved/web.json --viewports 375,1280 --schemes dark,light --json <report.json>
```

consumerはマニフェストの入力ハッシュを現行のdesign/と照合し、中央の判定policyと宣言したIA resolverのdigestも照合する。変更・追加・削除後の古い生成物を使わない。
DTCGの再解決はしない。L1はブラウザを使わず、L2は既存のPlaywright環境を使う。
axe-coreは実行環境のモジュール解決か `AXE_CORE_PATH` で既存ファイルを指定する。
依存を自動導入しない。axeが無い・実行不能なら準拠を名乗らず失格を残す。

## L1とTailwind

既定ではL0が `.resolved/tailwind.config.cjs` を生成する。プロジェクトごとの設定は
`design/tailwind.json` に `schema_version: 1` と `configs` を置く。各行は repo 相対の
`target`（tailwind.config.js/cjs）・`content`・`theme`・`plugins` を持つ。
`theme` の末端は `{"$token":"color.primary"}`。構造上のCSS文字列は
`{"$format":[{"$token":"duration.fast"}," ease"]}` のように組み立てる。
数値・色の直書きは禁止。fontFamily/number/fontWeightは値、他はCSS表現を出力する。

pluginsは `{"module":"@tailwindcss/forms","options":{"strategy":"class"}}` のような宣言だけ。
L0はpluginを実行せずCommonJSのrequireを生成する。任意の関数や手書きコードは扱わない。
利用側に既存pluginの依存が必要で、自動導入しない。既存content・theme項目・pluginオプションを
比較し、実TailwindでのCSS生成を確認してから切り替える。

```sh
design-author sync --repo .
```

syncは実L0の失格時に配置先を変更しない。整合が通れば宣言したconfigとassetコピーを配置する。
repo外・design内・隠しpath・symlink先への配置を拒否し、途中の書込み失敗は既存バイトへ復元する。
生成contentとSHA256はwebマニフェストに入り、L1は解決値から期待contentを再構成して照合する。
自己申告のhashだけを信用しない。宣言したconfigは全て存在必須。
`--tailwind` を省略すると追加のconfigも探索する。Tailwindを使わないrepoはHTML構造検査のみ。
空のHTMLは失格。`asset-exports.json` の全配布先も正典と同じSHA256であることを確認する。

## 一括実行

`design-check --repo .` は実L0→L1→L2を順に呼ぶ。L0失格なら後段へ進まない。
`design/validation.json` は `schema_version: 1`、非空の `html`（repo相対path）、
`urls`（`{"env":"DESIGN_WEB_URL"}`等）、正の整数の `viewports`、`schemes`（light/dark）を持つ。
URLは実行時の環境変数で渡す。既存Nodeは `DESIGN_AUDIT_NODE`、Pythonは
`DESIGN_AUTHOR_PYTHON` で選ぶ。L2を含むため、人間の実施確認を得てから実行する。
各L2のJSONは `.resolved/l2-<index>.json` に保存し、未カバーを集約出力する。
reportはexecuted / skipped / missing / uncovered / screenedの5配列。実行したL1/L2はskippedから除き、L0のscreenedとholdout指標を引き継ぐ。
対象欠落・実行不能を成功にしない。未承認nice値だけのL0警告は集約後もexit 1に残す。

## L2と範囲

色・フォント・角の許可集合をマニフェストからブラウザのcomputed表現へ正規化する。
bodyと子孫の各要素のcolor/background/border/outline/text-decoration、fontFamily、4隅のradiusを照合する。
透明色と角ゼロも必要なら明示的なtokenにする。検査対象0件は失格。
コントラストとターゲットサイズは媒体minimumsを読む。axeのcolor-contrastと従来の実効コントラスト検査を併用し、
axeで測れない対象は未検証として残す。失敗したaxe実行を合格にしない。

擬似要素・shadow DOM・iframe内・任意の宣言的追加規則は未カバー。CSSの使用宣言と実際のfallback fontの区別は
既存のfont-silent-fallback警告で報告する。画像上のコントラスト近似等、従来検査の限界を準拠合格に読み替えない。
L1/L2が通っても、未カバー規則が消えたとは扱わない。出荷時のmust未カバーはL0 --shipで停止する。

## 検証

`own-design-author/scripts/tests/test_core.py` にL0から実L1/L2までの人工フィクスチャがある。
ブラウザ試験は `DESIGN_BROWSER_TEST=1` と `DESIGN_TEST_NODE` を指定して実行する。
`DESIGN_BASELINE_MEASURE` は同じフィクスチャを改修前スクリプトにも渡す回帰用。
通常実行のskipを実ブラウザ検証済みと報告しない。
