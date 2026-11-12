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
各URLにprofile: lp|product-uiを任意指定でき、省略時はlp。product-uiでは広告LPの最初の画面のCTAだけを適用外としてskippedへ理由付きで残し、色・a11y・字体・固定描画の制約を維持する。関連付いたcheckbox/radioのlabelは、入力を包含する連続した矩形と実hit mapが有効な時だけ操作領域に含む。
URLは実行時の環境変数で渡す。既存Nodeは `DESIGN_AUDIT_NODE`、Pythonは
`DESIGN_AUTHOR_PYTHON` で選ぶ。L2を含むため、人間の実施確認を得てから実行する。
各L2はdesign/render-environment.jsonを必須とし、固定環境で2回実行する。JSONは `.resolved/l2-<index>.json` に保存し、未カバーを集約出力する。実行前に前回の同名レポートを削除し、環境不一致や新規レポート欠落を古い成功で埋めない。
reportはexecuted / skipped / missing / uncovered / screenedの5配列。実行したL1/L2はskippedから除き、L0のscreened・補助指摘・評価用ラベル有無を引き継ぐ。精度未測定をrecall 0や合格へ読み替えない。
対象欠落・実行不能を成功にしない。未承認nice値だけのL0警告は集約後もexit 1に残す。

## 画面の棚卸し（漏れ検出）

`html` と `urls` に書いた画面だけが検査される。後から増えた画面が宣言から漏れても、
それだけではゲートは緑のまま通る。これを防ぐため、`validation.json` に任意の `pages` を置く。

```json
"pages": {
  "sources": ["apps/web/src/app/**/page.tsx", "apps/landing-page/**/*.html"],
  "covered": {"apps/web/src/app/diagnosis/page.tsx": ["DESIGN_WEB_URL"],
              "apps/landing-page/index.html": ["html", "DESIGN_LP_URL"]},
  "excluded": {"apps/web/src/app/login/page.tsx": "認証画面。外部の認証基盤が描画する"}
}
```

- `sources` は repo 相対の git pathspec glob。追跡済みと未追跡のファイルを `.gitignore` に従って列挙するため、依存やビルド出力は入らない。フレームワークごとの経路規則は持たず、画面のソースを見つける glob を repo が宣言する。
- `covered` は画面のソースから検査の証拠への対応。`"html"` はそのパスが L1 の `html` にあること、それ以外は `urls` の環境変数名。1つの URL 環境変数は1画面だけを担う。全ての URL 環境変数をどれかの画面に対応させる。
- `excluded` は対象外にする画面と、空でない理由。対象外の判断はオーナーが持つ。

`design-check` は L1 の前に棚卸しを照合し、次を exit 2 で止める（ブラウザは起動しない）。
宣言されていない画面、存在しない画面の宣言、何にも一致しない `sources`、理由のない対象外、
covered と excluded の重複、宣言していない検査対象の参照、複数画面を担う URL、画面に対応しない URL。
エラーは該当パスを名指しする。`pages` が無い repo は止めずに、`uncovered` へ `pages` を残して exit を 1 以上にし、
同じ文を `warnings` に出す。`design-ship` の記録（push 前検査・統合検査）も `warnings` にこの文を引き継ぐ。

限界: 対応が正しいか（その URL が本当にその画面を描くか）は宣言を信じる。
動的な経路（`[id]` 等）や認証の後ろの画面は、描画用の状態を用意できるまで理由付きで対象外にする。

## 採用した項目の確かめ方（コードが先、AFMは残りだけ）

L0は`defined`の全項目に確かめ方を割り当てる。mustかどうかは関係ない。順番は次の通り。

1. 証拠に宣言的rule、またはcolor・fontFamilyのトークンがある → `machine`。
2. skillが持つ対応表`references/item-checks.json`に、項目の文書と見出しが載っていて、必要な検査がすべて有効 → `machine`。
   対応表はrepoごとに書かない。全repoが同じカタログを使うので、skillで一度だけ見直す。
   paletteの検査は、比べる種類のトークン（dimension・fontWeight・shadow等）がある時だけ有効。
3. それ以外 → `afm`。

割り当ては`.resolved/web.json`の`coverage.items`と`coverage.item_routes`に入る。L2に渡す検査と部品のselectorは`item_checks`に入る。

### 検査の種類

- L2で止まる検査: palette:color / fontFamily / radius、contrast、tap-target、heading-order、layout-overflow。
  これらが失格ならL2で止まるので、項目の検証まで来た時点で合格している。
- L2が測り、採用した項目にだけ効かせる検査（L2自体は止めない）:
  - palette:fontSize / fontWeight / shadow / spacing（padding・gap。marginは自動の中央寄せと区別できないので測らない）
  - dom:placeholder / readonly-input / disabled-input / maxlength / ol / positive-tabindex / date-input / table-span / vague-link-text / text-justify
  - axe:button-name / list

### 部品の有無

対応表の`applicability`は、部品が画面にあることを示すselector。L2が描画した画面で数を数える。
数が0の画面では、その部品の項目は対象外になり、コードの検査もAFMもしない。`applicability`に無い文書は常に対象。

### 段の順番

`design-check`はL2の後に`ITEMS`段、必要なら`AFM`段を実行する。

- ITEMS: `machine`の各項目を、対象の各画面で確かめる。L2の全viewport・全schemeの結果を合わせ、1件でも外れたら項目idと画面と見本を名指しして停止する。
  測るはずの検査が測られていなければ、それも停止。結果は`.resolved/items.json`。
- AFM: 対象の画面に残った`afm`の項目だけを判定する。端末内のApple Foundation Models（`fm` CLI、systemモデル、従量課金・外部送信なし）へ、
  その画面の最も広いviewport・最初のschemeの安定キャプチャ（L2が測った同じ画像）を渡す。HTMLソースは渡さない。他のviewport・schemeはskippedに理由付きで残す。
  項目は文字数上限でまとめ、`--greedy`と構造化出力で`applies`・`violates`・`reason`を受け取る。
  結果はモデル名・OS build・プロンプト版・項目・証拠・画像のhashをキーに、repo外（`DESIGN_AFM_CACHE`、既定は利用者のcache）へ保存し、同じ入力は再判定しない。
  全判定は`.resolved/afm.json`。

AFMの結果の扱い:

- 違反（applies かつ violates）: 警告（exit 1）。精度を測るまで出荷は止めない。止めるかはbiz_ops `WS-20261112-afm-judgment-accuracy`が決める。
- 判定の欠落・壊れた出力（1回だけ単独で再試行）: 未カバー。exit 2。
- `fm`がない、systemモデルが使えない: 停止。`fm models`はPCC不可の時に終了コード1を返すので、✓の行で判定する。

`design-ship`は、ITEMS段（失格0件・ページの一致・AFM対象の組の数）と、AFM対象があればAFM段（組と判定の数の一致）を必須にする。
`DESIGN_AFM_COMMAND`で`fm`の代わりの単一実行ファイルを指定できる（テスト用）。

限界:

- 対応表は「その検査がその項目の測れる部分を確かめる」と判断したもの。項目の文章のうち、測れない言い回しの部分は確かめない。
- dimensionトークンに役割の区別がないので、文字の大きさと余白は、どちらもdimensionトークン全体と比べる。
- AFMの精度は未測定。長いページは1枚の画像として渡すので、細部は読めない。macOSで`fm`がある端末でしか動かない。

## L2と範囲

色・フォント・角の許可集合をマニフェストからブラウザのcomputed表現へ正規化する。
bodyと子孫のうち描画される色・背景・境界・outline・装飾線・文字のfontFamilyと4隅のradiusを照合する。非描画の境界・透明な背景・非表示DOMは適合の証拠にせず、描画対象0件を拒否する。寸法tokenは負の変換値も持てるが、CSSが半径として許可する値だけを半径集合に使う。正規化用の見本はページのanimation/transitionから独立させる。
透明色と角ゼロも必要なら明示的なtokenにする。検査対象0件は失格。
コントラストとターゲットサイズは媒体minimumsを読む。axeのcolor-contrastと従来の実効コントラスト検査を併用し、
axeで測れない対象は未検証として残す。失敗したaxe実行を合格にしない。

擬似要素・shadow DOM・iframe内・任意の宣言的追加規則は未カバー。
固定検査はCDPの実描画glyph情報でfont fallbackを失格にする。
環境指定のないmeasure_lp.mjs単独計測では従来のfont-silent-fallback警告を維持し、fixed-render-environmentを未カバーへ残す。この単独計測は固定L2の代用にしない。画像上のコントラスト近似等、従来検査の限界を準拠合格に読み替えない。
L1/L2が通っても、未カバー規則が消えたとは扱わない。出荷時のmust未カバーはL0 --shipで停止する。

## 検証

`own-design-author/scripts/tests/test_core.py` にL0から実L1/L2までの人工フィクスチャがある。
ブラウザ試験は `DESIGN_BROWSER_TEST=1` と `DESIGN_TEST_NODE` を指定して実行する。
`DESIGN_BASELINE_MEASURE` は同じフィクスチャを改修前スクリプトにも渡す回帰用。
通常実行のskipを実ブラウザ検証済みと報告しない。


## 固定環境

`design/render-environment.json`をrepoに置き、既存のPlaywright/Chromiumと同梱フォントをpinする。
新規インストールは行わない。環境・フォント変更時はこのpinを更新し、L0から再生成する。
環境に合わせて黙ってpinを書き換えない。

```json
{
  "schema_version": 1,
  "runtime": {"browser_version":"<exact-version>","browser_sha256":"<sha256>",
    "playwright_version":"<exact-version>","node_version":"<exact-version>",
    "platform":"<platform>","arch":"<arch>","os_release":"<exact-release>",
    "launch_args":"--disable-gpu-rasterization"},
  "clock":"2020-01-01T00:00:00.000Z","locale":"ja-JP","timezone":"UTC",
  "device_scale_factor":1,"reduced_motion":"reduce",
  "viewports":[{"width":375,"height":812},{"width":1280,"height":900}],
  "schemes":["light","dark"],
  "fonts":[{"path":"public/fonts/example.woff2","url_path":"/fonts/example.woff2",
    "family":"<actual-font-family>","sha256":"<sha256>"}]
}
```

runtimeはfixed_render.mjsのprobe(chromium, playwrightVersion)で実測する。runtimeのbrowser hashは、
chromium.executablePath()から解決し、実際のlaunchにも明示指定する同じ実行ファイルのSHA256。
これはブラウザ同梱の全資源やOS全体の完全性証明ではない。OS版・arch・Node版・Playwright版も照合する。
個別端末の実行ファイルの絶対パスは保存しない。

fontsのpathはrepo相対の通常ファイル。SHA256を実測し、同一originのurl_pathへの要求へそのバイトを配信する。
未登録のfontリクエストを拒否し、可視要素の実glyphにOSフォントや未登録familyが使われたら失格。
フォントを1つも読み込まない、glyphを1つも検査しない場合も失格。ページを別の字体へ強制置換して合格させない。
shadow DOM・iframe内のglyphは未カバーのまま。読み込み前からCDPのfontsUpdatedも記録し、実familyのsrcが登録URLに結び付かない場合は失格。同名のdata URI・local()・インライン生成FontFaceでpinを迂回しない。
フォントのライセンス・配布可否はプロジェクトが確認し、必要なNOTICEを同梱する。

固定環境のChromiumは `fixedLaunchOptions`（`--disable-gpu-rasterization`）で起動し、probeと測定で同じ描画経路を使う。起動引数はruntimeの`launch_args`としてpinし、別の描画方式で測ったpinは一致しない（作り直す）。
GPUラスタでは同じ全画面撮影が1段階ずつ揺れることがある（2026-11-11 macOS 27.2で20枚中2種類、ソフトウェアラスタで1種類）。
描画が止まった時は、測定が段階（stage）と固定コード（例 `render-unstable`・`asset-decode`）入りの診断レポートを残し、ゲートのerrorsに理由を出す。詳細は許可済みの固定文だけで、生のエラー文（URL・パスを含みうる）は保存しない。
時計はページ遷移前にPlaywrightのsetFixedTimeを適用し、Dateを固定する。タイマーは動作を続ける。
DPR・viewport・reduced-motion・locale・timezoneはcontextで指定し、ページ側の実値も照合する。
service workerはブロックし、各組合せを新しいcontextで実行する。
同じURL・viewport・schemeを2回測り、全測定結果とPNG画素のhashが同じ場合だけrepeat_equalをtrueにする。
不一致は失格かつrender-repeat-mismatchの未カバー。乱数や動的データを検査側で偽装せず、プロジェクトの再現可能な入力を用意する。
環境指紋・実字体・glyph件数・時計・画素hashを結果JSONへ残し、終了時に環境設定とフォント・runtimeを再照合する。

```sh
node <own-website-audit>/scripts/measure_lp.mjs <url> --manifest design/.resolved/web.json --environment design/render-environment.json --viewports 375,1280 --schemes light,dark --json design/.resolved/l2.json
```

viewports/schemesの順序と集合もpinに一致必須。ブラウザ版・フォントhash・環境欠落はexit 2、
測定の不一致や実fallbackはexit 1。API一次資料は[時計](https://playwright.dev/docs/clock)と
[CDP字体情報](https://chromedevtools.github.io/devtools-protocol/tot/CSS/#method-getPlatformFontsForNode)。
実ブラウザfixtureの追加条件はDESIGN_TEST_FONTに既存Inter woff2のパスを渡すこと。通常lintのskipを実測と報告しない。

各独立描画内では、スクロール等の計測で生じた再描画待ちを除くため、最大3枚の固定条件キャプチャから連続2枚の完全一致を要求する。安定しなければ失格。安定した画素のhashをさらに独立した2回の描画間で完全比較し、別runから都合のよい画像を選ばない。

描画準備はDOMContentLoaded後に本文字体と全画像のdecodeを待ち、10秒以内に揃わなければ失格にする。prefetch/analytics等の背景ストリームのnetworkidleは準備判定に使わない。測定終了後の将来の非同期応答の内容は保証しない。
