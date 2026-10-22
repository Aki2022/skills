# 解決済みマニフェストによるweb検査

実装元: biz_ops `WS-20261021-design-system-standardization` ISSUE-02 / ISSUE-03。

## 実行

L0を実行してから、工場は次を使う。L0/L1は確認不要、L2は人間の実施確認後に走らせる。

```sh
design-author check --repo .
python3 <own-web-design>/scripts/audit_html.py index.html --manifest design/.resolved/web.json --tailwind tailwind.config.cjs
node <own-website-audit>/scripts/measure_lp.mjs <url> --manifest design/.resolved/web.json --viewports 375,1280 --schemes dark,light --json <report.json>
```

consumerはマニフェストの入力ハッシュを現行のdesign/と照合する。変更・追加・削除後の古い生成物を使わない。
DTCGの再解決はしない。L1はブラウザを使わず、L2は既存のPlaywright環境を使う。
axe-coreは実行環境のモジュール解決か `AXE_CORE_PATH` で既存ファイルを指定する。
依存を自動導入しない。axeが無い・実行不能なら準拠を名乗らず失格を残す。

## L1とTailwind

L0は `.resolved/tailwind.config.cjs` と、そのcontent / SHA256をマニフェストに生成する。
利用するtailwind.configはこの生成物のバイト一致コピーとする。色は解決済みtokenパスのdotをhyphenに変えたキー。
L1は解決済み値から期待contentを再構成し、ハッシュと実ファイルを照合する。自己申告のハッシュだけを信用しない。
`--tailwind` を省略するとrepo内のtailwind.config群を探索し、node_modules等を除外して全て照合する。
Tailwindを使わないrepoはHTML構造検査のみ。空のHTMLは失格。

生成形式はデータだけのCommonJS（content / theme.colors / plugins空配列）に限定する。
手書きの拡張・実行関数・プラグインはまだ扱わない。既存プロジェクトへ適用する時はこの制約と既存機能を確認し、
対応を実装・検証してから切り替える。既存configを消して準拠を装わない。

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
