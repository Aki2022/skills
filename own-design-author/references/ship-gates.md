# 出荷ゲートの接続

`design-ship`は、cleanなcommitに対するL0 `--ship`、ソースのL1、固定環境のL2を実行する。
L0だけの成功を出荷合格としない。各段階の対象0・欠落・終了コード不一致・古い描画結果は停止する。
niceの提案値だけなら警告を保持するが、mustの未承認・未カバーは失格。

## チャネルの宣言

design.yamlの任意の`channels`は、名前をキーにした辞書。`ship: true`のチャネルだけが出荷対象。
各チャネルに次を宣言する。現在実装する媒体はweb。

- `ship`: 真偽値。`false`のチャネルは出荷しない。
- `media`: `web`。
- `ship_command`: 既存のdeploy/export入口を指す、空でないargv配列。shell文字列ではない。
- `artifacts`: 出荷するファイルまたはディレクトリのrepo相対パス配列。ディレクトリは全ファイルへ展開する。空・symlink・repo外は失格。
- `merge_base`: 予測マージの基準ref。例は`origin/main`。欠落・競合を成功扱いにしない。
- `build_command`: 必要な場合だけ、対象checkoutで成果物を構築するargv配列。
- `render_servers`: validation.jsonの全URL環境変数をキーにした辞書。各値は`command`（argv配列）と`path`（同一origin内のパス）。commandは割り当てた`{port}`を使う。起動場所は検査対象のcheckout。

既存のbuild/deploy経路を調べ、実際に出荷するファイル集合を指定する。検査中に変わるcacheを出荷物へ混ぜない。
必要な成果物を省いて検査を通すこともしない。コマンドを自動生成したり、新しい依存・認証・CIを導入したりしない。

## 実行の違い

```sh
design-ship --repo .
design-ship --repo . --doctor
design-ship --repo . --channel web
```

引数なしは全出荷チャネルの検査だけ。`--channel`を明示した場合だけ、合格後にその`ship_command`を実行する。
deploy/export自体の認可を別途守る。実装テストから本番コマンドは実行しない。
`--doctor`はチャネル宣言・有効化marker・scannerを保持するhookと現在の拡張を検査する。

ゲートが未使用のloopback portを選び、各serverを対象checkoutで起動する。既存のURL環境変数を
別checkoutの画面の証拠として流用しない。全serverを終了し、各段階の対象数・固定環境・2回分の測定と画像hashの一致を確認する。
commit/tree・design入力・出荷ファイル集合のhashが検査中に変われば停止する。
出荷コマンドにはDESIGN_SHIP_COMMIT / DESIGN_SHIP_TREE / DESIGN_SHIP_ARTIFACTS_JSONを渡し、その同じ成果物を扱わせる。
結果は`design/.resolved/ship-<channel>.json`。この生成記録を過去の成功からコピーして代用しない。

## hookとmiseへの導入

本番repoへの導入は記録済みの承認範囲で行い、既存hook・scanner・他の拡張を保持する。
core.hooksPathを変更せず、対応しないdispatcherを自動で書き換えない。

- PATHへskillのscriptsを供給し、既存依存を持つPythonをDESIGN_AUTHOR_PYTHONで選ぶ。端末固有のパスをrepoへ保存しない。
- templates/mise.design.tomlのtaskを既存設定へ統合する。design:shipは検査だけを呼ぶ。
- hooks/pre-push.d/60-design-shipを既存vibe-guardの拡張ディレクトリへ追加する。
- 有効化するrepoだけ.vibe-guard/design-ship.enabledを文字列1にする。欠落コマンド・壊れたmarkerは停止する。

vibe-guardがscanner後に保持したpush refsファイルを拡張へ渡す。拡張は
`design-ship --pre-push --refs-file <file>`を呼び、pushするcommitとチャネルの基準refを
`git merge-tree --write-tree`で統合した木を、隔離checkoutで同じ3段階検査に通す。
元checkoutや別branchの承認・起動済み画面は代用しない。削除refだけのpushは出荷なしとして報告する。
DESIGN_SHIP_COMMAND / DESIGN_AUTHOR_COMMAND / DESIGN_CHECK_COMMANDは、既存runtimeの単一実行ファイルを指定するための環境変数。
任意のshell文字列を評価しない。

## 実経路の検証

自己テストは実Chromiumと既存fontを使い、正常系と、L1違反・字体fallbackの拒否を検証する。
一時Git設定へ元設定をincludeし、実scannerと既存拡張を保持したhookコピーからbare remoteへpushする。
未承認mustの拒否と、人工fixtureの承認後の到達、同じmise入口を確認する。
ユーザーの設定・global hook・実プロジェクトの承認台帳は、このテストで変更しない。
実ブラウザ・hook用環境変数を指定しない時のskipは、実経路の合格に数えない。
これは人工repoでのゲート検証。実パイロットの全画面合格や本番デプロイ完了とは区別する。
