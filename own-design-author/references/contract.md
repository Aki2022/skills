# design/ の契約

実装元: biz_ops `WS-20261021-design-system-standardization` ISSUE-01。

## Runtime

Python 3.10以上と、既存環境に導入済みの PyYAML / jsonschema（referencingを含む）を使う。無ければ停止する。自動インストールしない。`DESIGN_AUTHOR_PYTHON` を指定すると `scripts/design-author` がその interpreter を使う。検査時はネットワークに接続せず、同梱の DTCG 2025.10 公式 draft-07 Schema を読む。

## ファイル

- `design.yaml`: JSON互換YAML。`schema_version: 1`、`skill_version: 1.0.0`、`catalog_version`、`preset`、`media`、`items`。`items` は不変idをキーにした差分だけ。`spec`は方針を持つrepo相対のdocs/specs Markdownパス、または重複しない非空のパス配列。相互参照はown-doc-updateのvalidatorが検査する。
- `design.yaml`の任意の`maintenance`は期限点検の設定。`within_days`と`review_due_dates`のみを持ち、詳細はmaintenance.md。維持検査の成功stampは`.resolved/`に置き、実台帳は変更しない。
- `design.yaml`の任意の`channels`は出荷チャネルの宣言。`ship: true`に正式のship_command・成果物・検査用serverを必須とする。ファイル契約と起動方法はship-gates.md。
- `tokens/*.tokens.json`: 値を持つ正典。名前はlower-case kebab-case、参照パスはdot区切り、aliasは `{group.name}`。`$type` はグループから継承できる。CSS名は `--group-name`、衝突は失格。同じパスの多重定義も失格。
- `DESIGN.md`: 散文とトークン参照。doc証拠は `doc:DESIGN.md#dads-000001` のような実見出しを指す。空の節・コードブロック内の偽見出しは証拠にならない。
- `profiles/web.json`: `media: web` と `minimums`。skillの `minimums.json` より厳しい数値だけ上書き可。media不明・下限の緩和・未知キーは失格。
- `checks/*.yaml`: `rules` 配列。id / kind / targetsまたはpaths / allowlist / minimum / pattern の宣言だけ。任意コマンドやコードを実行しない。実際のL1/L2実行を持たない規則は未カバー。
- `assets/`: ファイルのsha256から `assets.json` を生成する。別の手書き台帳を正典にしない。
- `approvals.yaml`: 人間確認を根拠にした行の配列。
- `judgments.json`: 任意。doc証拠とn-a理由への保存済み確率判定、pinされたpolicyと人間ラベル付きholdout。未登録入力は未カバー。
- `NOTICE`: JSON互換YAML、`sources` 配列。各参照のid / url / commit / license / copyright / editedを必須とする。`edited`は真偽値。
- `tailwind.json`: 任意の生成設定。対象別content・宣言的plugin・tokenを参照するthemeを持つ。詳細は `web-validation.md`。
- `asset-exports.json`: `schema_version: 1` と `exports` 配列。各行にassets/以下のsourceとrepo相対targetsを置き、syncで配置、L1で実バイトを照合する。
- `validation.json`: design-checkのL1/L2対象とviewport/scheme。詳細は `web-validation.md`。
- `.resolved/tailwind.config.cjs`: 解決済みの色から生成するCommonJS設定。contentとSHA256はwebマニフェストのartifactsにも保存。
- `.resolved/web.json` / `.resolved/web.css`: L0の生成物。入力ハッシュ・解決値・CSS・承認状態・媒体下限・項目・未カバー・アセットを持つ。これらを手で直しても正典は変わらない。

項目の例:

```json
{
  "dads-000001": {"state": "defined", "evidence": ["token:color.primary"]},
  "dads-000002": {"state": "n-a", "reason": "このプロジェクトには対象の操作がない"}
}
```

全項目の初期状態はプリセットから `deferred`。mustのdeferredは失格。`n-a` に空理由は不可、mustのn-aには理由のハッシュを持つ承認行も必要。`defined` は一つ以上の実体を指す。

トークンの出所例（値そのものは各プロジェクトで決める）:

```json
{"$extensions": {"org.own.design": {"origin": "proposal"}}}
```

色はDTCGのcolorSpace/componentsに加え6桁hexフォールバックを必須とする。現時点のweb変換はsrgb、componentsとhexが食い違えば失格。dimensionはpx/rem、durationはms/sの構造と単位を保ってCSSへ変換する。JSON pointerは同じトークンファイル内の値を参照する。aliasはファイルを跨いで解決する。

group `$extends`、カスタムdash配列のstrokeStyle、srgb以外のCSS変換は未対応として明示的に失格にする。対応しない構文を黙って落とさない。印刷拡張は保持し、現段階で印刷変換を保証しない。

移行で上流から消えた項目も、旧カタログの版と実効状態を `retired_items` に保持して検査する。未解決mustを削除で緑にすることはできない。同じidが戻った場合は以前の状態を復元する。

## ハッシュと承認

承認hashは媒体・実効profile・解決値/証拠の辞書をJCS（RFC 8785）で正規化したSHA256。
分類カタログの既存入力hashはこの移行では変更しない。承認schemeはhash-jcs-v1。

- tokenは解決後のtype/valueを束縛する。alias先の変更も失効する。
- defined itemは証拠参照と解決値の辞書を束縛する。
- n-a itemはstate/reasonを束縛する。mustのn-aは承認行が必須。

行はkey / value_hash / approver / ISO dateを持つ。
keyはds:web:item:<id> またはds:web:token:<path>。IAのキーは固定の信頼されたIA resolverが現在hashを返した場合だけ登録対象に含める。実装がないIA入力は失格。
旧token_path/item_id行や署名行を自動移行しない。現在のJCS hashを確認して新しい記録を作る。
署名・鍵・allowed_signersは要求しない。承認者の真正性は技術的に証明せず、人間の明示承認を運用上の根拠とする。
AIは提案までを自律で行い、実台帳は書かない。オーナー自身が確認後に記録する。詳細はapproval-upgrade.md。
consumerはapproval_schemeと全入力hashを検査し、承認台帳の削除・変更後の古いmanifestも拒否する。

一致するdefined item行は直接参照するtokenの承認にもなる。散文内のtoken文字列だけでは伝播しない。
source-hash規則で全targetsがdesign/assets/以下なら実SHA256とallowlistをL0で検証し、
item証拠hashに規則と実測を含める。他の宣言規則を自己申告で対応済みにしない。

判定台帳・5区分報告とIA接続は `judgments-and-modules.md`。design.yamlのmodules.iaに実装版をpinする。

## 終了コードと境界

- 0: 自己整合に問題なし。
- 1: 提案値、未カバー規則、skill版の不一致などの警告。生成は可能だが警告を報告する。
- 2: 欠落・未知版/id・空のmust・証拠不在・台帳不一致・下限緩和・対象0件など。生成しない。

`--ship` は提案状態のmustと未カバーmustも2にする。未カバーに対応済みという自己申告を書いて合格にする経路はない。L1/L2の実行方法と制約は `web-validation.md`、hook/miseは `ship-gates.md`。散文が意味的に適切かはL0の決定論では判定しない。

`presentation` presetはwebプレビューを選ぶ。pptxの媒体規約とconsumerは従属WSで実装する。
