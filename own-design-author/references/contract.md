# design/ の契約

実装元: biz_ops `WS-20261021-design-system-standardization` ISSUE-01。

## Runtime

Python 3.10以上と、既存環境に導入済みの PyYAML / jsonschema（referencingを含む）を使う。無ければ停止する。自動インストールしない。`DESIGN_AUTHOR_PYTHON` を指定すると `scripts/design-author` がその interpreter を使う。検査時はネットワークに接続せず、同梱の DTCG 2025.10 公式 draft-07 Schema を読む。

## ファイル

- `design.yaml`: JSON互換YAML。`schema_version: 1`、`skill_version: 1.0.0`、`catalog_version`、`preset`、`media`、`items`。`items` は不変idをキーにした差分だけ。
- `tokens/*.tokens.json`: 値を持つ正典。名前はlower-case kebab-case、参照パスはdot区切り、aliasは `{group.name}`。`$type` はグループから継承できる。CSS名は `--group-name`、衝突は失格。同じパスの多重定義も失格。
- `DESIGN.md`: 散文とトークン参照。doc証拠は `doc:DESIGN.md#dads-000001` のような実見出しを指す。空の節・コードブロック内の偽見出しは証拠にならない。
- `profiles/web.json`: `media: web` と `minimums`。skillの `minimums.json` より厳しい数値だけ上書き可。media不明・下限の緩和・未知キーは失格。
- `checks/*.yaml`: `rules` 配列。id / kind / targetsまたはpaths / allowlist / minimum / pattern の宣言だけ。任意コマンドやコードを実行しない。実際のL1/L2実行を持たない規則は未カバー。
- `assets/`: ファイルのsha256から `assets.json` を生成する。別の手書き台帳を正典にしない。
- `approvals.yaml`: 人間確認を根拠にした行の配列。
- `NOTICE`: JSON互換YAML、`sources` 配列。各参照のid / url / commit / license / copyright / editedを必須とする。`edited`は真偽値。
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

`value_hash` はUnicodeを保ったsort-keys・空白なしのJSONをUTF-8化したSHA256。

- token行は `{type, value}` の解決後値を束縛する。aliasの参照先が変わればハッシュも変わる。
- defined item行は `{evidence-reference: resolved-evidence-value}` の辞書を束縛する。
- n-a item行は `{state: "n-a", reason: <理由>}` を束縛する。

行は `token_path` または `item_id` のどちらか一つ、`value_hash`、`approver`、ISO `date` を持つ。重複行・存在しない対象・ハッシュ不一致は失格。台帳の作成コマンドは提供しない。自己テスト内の承認行は人工フィクスチャで、実プロジェクトの承認ではない。

## 終了コードと境界

- 0: 自己整合に問題なし。
- 1: 提案値、未カバー規則、skill版の不一致などの警告。生成は可能だが警告を報告する。
- 2: 欠落・未知版/id・空のmust・証拠不在・台帳不一致・下限緩和・対象0件など。生成しない。

`--ship` は提案状態のmustと未カバーmustも2にする。未カバーに対応済みという自己申告を書いて合格にする経路はない。L1/L2との接続は後続sliceが実装する。散文が意味的に適切かはL0の決定論では判定しない。

`presentation` presetはwebプレビューを選ぶ。pptxの媒体規約とconsumerは従属WSで実装する。
