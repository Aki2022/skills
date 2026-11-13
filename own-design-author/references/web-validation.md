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
各URLにprofile: lp|product-uiを任意指定でき、省略時はlp。操作後の状態は任意の`states`で宣言する（「操作の後の状態」）。product-uiでは広告LPの最初の画面のCTAだけを適用外としてskippedへ理由付きで残し、色・a11y・字体・固定描画の制約を維持する。関連付いたcheckbox/radioのlabelは、入力を包含する連続した矩形と実hit mapが有効な時だけ操作領域に含む。
URLは実行時の環境変数で渡す。既存Nodeは `DESIGN_AUDIT_NODE`、Pythonは
`DESIGN_AUTHOR_PYTHON` で選ぶ。L2を含むため、人間の実施確認を得てから実行する。
各L2はdesign/render-environment.jsonを必須とし、固定環境で2回実行する。JSONは `.resolved/l2-<index>.json` に保存し、未カバーを集約出力する。実行前に前回の同名レポートを削除し、環境不一致や新規レポート欠落を古い成功で埋めない。
reportはexecuted / skipped / missing / uncovered / screenedの5配列。実行したL1/L2はskippedから除き、L0のscreened・補助指摘・評価用ラベル有無を引き継ぐ。精度未測定をrecall 0や合格へ読み替えない。
対象欠落・実行不能を成功にしない。未承認nice値だけのL0警告は集約後もexit 1に残す。

## 画面の棚卸し（漏れ検出）

`html` と `urls` に書いた画面だけが検査される。後から増えた画面が宣言から漏れても、
それだけではゲートは緑のまま通る。これを防ぐため、`validation.json` に `pages` を置く。出荷（`--ship`）には必須。

```json
"pages": {
  "sources": ["apps/web/src/app/**/page.tsx", "apps/landing-page/**/*.html"],
  "covered": {"apps/web/src/app/diagnosis/page.tsx": ["DESIGN_WEB_URL"],
              "apps/landing-page/index.html": ["html", "DESIGN_LP_URL"]},
  "excluded": {"apps/web/src/app/login/page.tsx": "認証画面。外部の認証基盤が描画する"},
  "routes": {"apps/landing-page/index.html": "/"}
}
```

- `sources` は repo 相対の git pathspec glob。追跡済みと未追跡のファイルを `.gitignore` に従って列挙するため、依存やビルド出力は入らない。フレームワークごとの経路規則は持たず、画面のソースを見つける glob を repo が宣言する。
- `covered` は画面のソースから検査の証拠への対応。`"html"` はそのパスが L1 の `html` にあること、それ以外は `urls` の環境変数名。1つの URL 環境変数は1画面だけを担う。全ての URL 環境変数をどれかの画面に対応させる。
- `excluded` は対象外にする画面と、空でない理由。対象外はオーナーの決定で、承認台帳に `ds:web:page:<画面のパス>` として記録する。ハッシュは画面のパスと理由に結びつくので、理由を変えると承認し直し。承認のない対象外は `--ship` で止まる。
- `routes`（任意）は、ファイルの場所から経路を決められない画面（素の HTML など）の経路。Next.js の `app/`（route group の `(x)` と並列 route の `@x` は経路に出ない）と `pages/` の画面は、ファイルの場所から経路を決める。
- URL で検査する画面は、その URL のパス（`render_servers` の `path`）が画面の経路に合うことを L0 で照らし合わせる。`[id]` は 1 区切り、`[...x]` は 1 区切り以上、`[[...x]]` は 0 区切り以上に合う。`/users/[id]` なら `/users/sample` のような見本の値で描く。経路に合わない URL・経路を決められない画面・パスのない URL は止まる。
- 画面を増やしても起動するサーバーを増やさないよう、`render_servers` の行は `{"same_as": "<別の行の環境変数>", "path": "/x"}` で、別の行が起動したサーバーの別のパスを描ける。
- ログインの後の画面は、repo の準備コマンドや描画サーバーのコマンドが、その repo の開発環境のテスト用ログイン状態を用意して描く。本番のアカウントは使わない。用意できない画面だけを、承認つきの対象外にする。

L0 と `design-check`（L1 の前）は棚卸しを照合し、次を exit 2 で止める（ブラウザは起動しない）。
宣言されていない画面、存在しない画面の宣言、何にも一致しない `sources`、理由のない対象外、
covered と excluded の重複、宣言していない検査対象の参照、複数画面を担う URL、画面に対応しない URL。
エラーは該当パスを名指しする。`pages` が無い repo は、`design-check` では止めずに警告し、`--ship` の L0 で止める。
L0 の `counts.pages` に、見つけた画面・検査する画面・対象外の数が入る。

限界: 経路の照合は、URL のパスが画面の経路の形に合うかまで。その URL が描いた中身が本当にその画面のコンポーネントかは見ない。
Next.js 以外の経路規則は持たないので、`routes` の宣言を信じる。

### 操作の後の状態（states）

静的な描画には、クリックや送信の後にだけ現れる部品がない。その部品へのruleはどの画面にも合わず、何も確かめない。
`urls`の行に`states`を置くと、そのURLを描いたあと、名前つきの状態ごとに手順を実行してから描き、画面と同じに測る。

```json
"urls": [{"env": "DESIGN_WEB_URL", "profile": "product-ui", "states": {
  "dialog-open": [{"click": "#open-settings"}],
  "email-error": [{"fill": "#email", "value": "not-an-email"}, {"submit": "#signup"}]
}}]
```

- 手順は宣言だけで、任意のコードは書けない。`click`・`focus`・`submit`はselector、`fill`はselectorと`value`。1状態に1〜20手順。
  `submit`はフォームか、フォームに属する要素（送信ボタンならそのボタンで送る）を指し、ブラウザの`requestSubmit`で送る。
- 状態の名前は英小文字・数字・ハイフン。報告では`<env>#<state>`（例 `DESIGN_WEB_URL#dialog-open`）。
- 状態ごとに、画面と同じviewport・scheme・固定環境で2回描き、L2・ITEMS・AFMを画面と同じに行う。`.resolved/l2-<n>.json`は画面、続けてその状態の順。
  手順は描画の準備（字体・画像）が揃ってから実行し、最後の手順の後にも同じ準備を待つ。
- 手順のselectorがちょうど1つの要素に合わない（0件・複数件・ブラウザが受け付けない）、または操作できない（隠れている・5秒で操作できない・`submit`の先にフォームがない）時は、
  「状態に到達できない」として失格（exit 2）。エラーは`could not reach state <env>#<state>: step <n> <手順>`。何も確かめずに通ることはない。
- 状態は同じURLの別の描画なので、画面の棚卸し（`pages.covered`）にはURLの環境変数だけを書く。状態のために描画サーバーを増やさない。
- 出荷（`design-ship`）は、宣言した全ての状態の描画があり、その報告に記録された手順が宣言と同じことを確かめる。

ruleがある状態でだけ合うことは、画面では対象外（n/a）・状態では合格か失格、として`items.json`に残る。
どの画面・状態でも一度も合わないruleは、これまで通り警告する。その決定を指す項目を未カバーにしておくかはrepoが決める。
未カバーを解くのは、状態を宣言して実際に描いた結果だけで、描かずに緑にする経路はない。

限界: 手順が意図した状態を作ったか（例: 本当にそのダイアログが開いたか）は見ない。それはruleのselectorが合うことで確かめる。
ホバー・ドラッグ・キーボード操作・時間経過の後の状態は扱わない。

### 画面を並べて描く

`validation.json` の任意の `parallel`（正の整数、既定 1）で、同時に描く画面の数を決める。
判定は宣言の順に行い、ある画面の L2 か ITEMS で外れたら、先に描き始めた画面も止める。

## 採用した項目の確かめ方（skillのコード・repoが宣言する検査・AFM）

カタログの全項目は、skillが持つ対応表`references/item-checks.json`で、次のどれか1つに割り当て済み。
割り当ては見出しの完全な経路（小見出しまで）で指す。repoごとには書かない。全repoが同じカタログを使うので、skillで一度だけ見直す。
カタログを更新して割り当てのない項目が出たら、その項目は`uncovered`になる。

| 経路 | 対象 | 確かめ方 |
|---|---|---|
| `machine` | 値・構造・属性で言える規則（色・文字・余白・影・コントラスト・押せる大きさ・見出し・ラベル・属性の禁止など） | L2が測った結果で、画面ごとにコードで確かめる |
| `decision` | repoごとに決めること（どの型を使うか、いつ部品を使うか、文章の方針、階層の体系、ロゴやライティングの方針など） | skillは検査を持てないので、repoが`checks/*.yaml`に検査を宣言し、項目から`rule:<id>`で指す。宣言した検査は画面ごとにコードで確かめる。DESIGN.mdに書いて承認しただけでは合格しない |
| `judgment` | 画面を見ないと判断できないもの（色だけで区別していないか、余白で関係性を示しているか、カードの見た目など） | 部品がある画面だけ、AFMで判定する |

L0は採用した項目（`defined`）ごとに経路を決める。mustかどうかは関係ない。

1. 証拠に画面の検査（count・text・style）を指す`rule:`がある → `machine`（その検査で確かめる）。
2. 証拠にそれ以外の宣言的ruleがある → `machine`（L0・L1で測れないruleは未カバーとして出荷で止まる）。
3. 証拠に`judgment`のruleがある → `afm`。AFMにはカタログの文ではなく、repoが書いた判定文を渡す。
4. それ以外は対応表の経路に従う。`decision`は`uncovered`（検査の宣言がない）。`machine`でも、比べる種類のトークン（dimension・fontWeight・shadow等）がなければ`uncovered`（理由つき）。
5. 対応表にない項目 → color・fontFamilyのトークンを証拠に持てば`machine`（palette）、なければ`uncovered`。
   対応表にある項目では、トークンを挙げただけでは確かめたことにならない（値があることしか言えないため）。

`decision`の項目は、宣言した検査に加えてオーナーの承認も要る。承認のハッシュは指したruleの中身を含むので、ruleを変えると承認し直しになる。
`--ship`では、`uncovered`の項目と、承認されていない`decision`の項目があれば停止する。
割り当ては`.resolved/web.json`の`coverage.items`（machine・afm・uncovered）と`coverage.item_routes`に入る。`coverage.decisions`はそのうちrepoの決定の項目。

### repoが決定ごとに宣言する検査

`checks/*.yaml`の`rules`に書く。任意のコードは書けない。ひな形は`references/decision-rule-examples.json`（どのカタログ項目向けかも書いてある）。

| kind | 書くこと | 失格になる条件 |
|---|---|---|
| `count` | `targets`（数える要素のselector）、`minimum`か`maximum`、任意で`scope`（この中ごとに数える） | `scope`の要素ごとに、中の`targets`の数が範囲外 |
| `text` | `targets`（文を読む要素）、`pattern`（正規表現）、任意で`minimum`・`maximum`（既定は0件まで） | 要素ごとの一致数が範囲外 |
| `style` | `targets`、`property`（CSSのプロパティ名）、`allowlist`（値か`token:<path>`）、任意で`state: focus`、`scope` | 要素の計算済みの値が許可した値と違う。`state: focus`ならフォーカスしてから読む |
| `judgment` | `targets`（あれば判定する要素）、`statement`（判定させる決定の文） | AFMが沿っていないと判定する（精度を測るまで警告） |

- 検査は、そのselectorに合う要素がある画面だけに効く。合う要素がなければその画面では対象外（n/a）。
- どの画面でも一度も合わなかった検査は警告する。selectorの書き間違いで何も確かめていない状態を、合格と見分けるため。
- ブラウザが受け付けないselector・正規表現は、対象外ではなく「測れなかった」として失格にする。
- `style`の`token:<path>`がトークンにないとL0で失格。
- ruleの`state`はfocusだけ。開閉・入力・送信の後にだけ現れる部品（ダイアログ・メニュー・ドロワー・送信後のエラー文・通知）は、
  `validation.json`のURLに操作後の状態（上の「操作の後の状態」）を宣言すると、その状態の描画で確かめる。ホバーは扱わない。
L2に渡す検査と部品のselectorは`item_checks`に入る。

### 検査の種類

- L2で止まる検査: palette:color / fontFamily / radius、contrast、tap-target、heading-order、layout-overflow、body-line-height、line-length。
  これらが失格ならL2で止まるので、項目の検証まで来た時点で合格している。
- L2が測り、採用した項目にだけ効かせる検査（L2自体は止めない）:
  - トークンの値: palette:fontSize / fontWeight / shadow / spacing（padding・gap。marginは自動の中央寄せと区別できないので測らない）
  - コントラスト: contrast-nontext（入力欄・ボタン等の境界と背景が3:1以上）、contrast-icon（アイコンと背景が4.5:1以上）
  - 属性・構造: placeholder・readonly・disabled・maxlength・pattern・ol・正のtabindex・`input type=date`・disabledのボタン・結合セル・thのない表・セルの配置（上寄せ、数値の右寄せ）・aria-expandedを持つ非ボタン・「ここ」リンク・リンクのカーソル・ボタンの文字のはみ出し・入力欄の高さ（36px以上）・ラベルが入力欄の上にあるか・貼り付け／コピーの禁止・フォーム内のaria-live・文字で作った箇条書き・h1が1つでtitleに含まれるか・アイコンだけのリンク／ボタン・ラベル付きの装飾アイコンの隠し忘れ・新規タブの予告・文中リンクの区別・全角空白での位置合わせ・長いイタリック・ページ背景の指定・ダイアログの枠線・見えないファイル選択・7枚以上のスライド・フォーカス順が上へ戻る
  - axe: button-name、list、label、link-name

各検査には、壊したページで必ず検出し、きれいなページでは何も出さないテストがある（`tests/test_item_collector.py`）。

### 部品の有無

対応表の`applicability`は、部品が画面にあることを示すselector。L2が描画した画面で数を数える。
数が0の画面では、その部品の項目は対象外になり、コードの検査もAFMもしない。`applicability`に無い文書は常に対象。

### 段の順番

`design-check`は画面を1つずつL2で描き、その直後にその画面の`ITEMS`を判定する。全画面が通ったら、必要なら`AFM`段を実行する。
コードの検査はすべてAFMより先に終わる。

- ITEMS: `machine`の各項目を、対象の各画面で確かめる。その画面のL2の全viewport・全schemeの結果を合わせ、1件でも外れたら項目idと画面と見本を名指しして停止する。
  次の画面は描画しないので、1回で見えるのは最初に外れた画面の分だけ。
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
