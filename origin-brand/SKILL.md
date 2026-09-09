---
name: origin-brand
description: >-
  ブランド意匠の倉庫。文脈（corporate-presentation / public-sector / dashboard 等）と媒体
  （html / dashboard / pptx 等）を渡すと、解決済みの意匠（値＋出所＋ライセンス可否）を返す窓口。
  必ず使うこと: 成果物にブランドの配色・タイポ・余白を反映する時、「自社ブランドで」
  「デジタル庁風に」「このサービスの意匠で」と言われた時、複数ブランドを切り替える時、
  各工場 skill（origin-web-design / origin-pptx / 将来の印刷）が意匠を必要とする工程。
  使わない場面: 成果物そのものの生成（各工場 skill が持つ）／媒体固有の制約や
  レイアウト規約（各工場の media/ が持つ）／画像アセットの生成（origin-image-gen）／
  品質の採点と再評価ループ（origin-output-eval）／情報設計・コンテンツ構造・メッセージ設計
  （スタイリング層の外）。
---

# Origin Brand

**この skill はブランド意匠のデータと、それを選ぶロジックだけを持つ。** 成果物は作らない。
「文脈と媒体を渡すと、解決済みの意匠（値＋出所＋ライセンス可否）を返す」窓口として振る舞う。

なぜ倉庫を分けたか（`SPEC-web-design-skill-scope` 決定7）: 選択ロジックを各工場に置くと、
同じ規則が複数実装になり `extends` の解釈がズレる。**ズレは仮定ではなく実測済み** —
同一ブランドの値が2箇所で完全に別物だった（一方は色相を持つ配色、他方はグレー階調）。
さらに**ライセンスゲートは守る対象と同じ場所に置く**必要がある。門が工場側にあると、
工場は門を通らずに第三者資産の値を取得できてしまう。

## 同梱物

| パス | 役割 |
| --- | --- |
| plugins/my_company | `owned` の箱（自社ブランド） |
| plugins/digital-agency | `observed` の箱（上流が公式 token を公開） |
| plugins/notion | `observed` の箱（散文のみ） |
| plugins/awesome-design-md-jp-import | 第三者 DESIGN.md を内部形式へ写像する入口 |
| references/plugin-spec.md | `plugin.yaml` の正式スキーマ（等級・取得メタ・`usage_policy`） |
| references/resolver.md | Plugin Resolver・`extends` 1階層・タイブレークの規則と受け入れテスト |

## 等級は2つだけ（決定8）

`owned`（自社ブランド）と `observed`（外部観察）。**当初案の `licensed` は廃止した** —
デジタル庁も `awesome-design-md-jp` も観察対象であり、違いは等級ではなく
**上流が公開している情報の豊かさ**にすぎない。

**ライセンスゲート（値を出力に使う許可）は等級ではなく直交する属性**である。
`observed` でもゲートを通れば値を使えるし、通っていなければ使えない。

| 等級 | 構造 | 例 |
| --- | --- | --- |
| `owned` | DESIGN.md + `tokens/` + **媒体プロファイル**（pptx=pt 実値 / web=px・rem / 将来 印刷=mm） | `my_company` |
| `observed`（上流が公式 token を公開） | DESIGN.md + `tokens/` + 媒体別資産。値の出力利用はゲート通過が条件 | `digital-agency` |
| `observed`（散文のみ） | DESIGN.md のみ | `notion` |

**DESIGN.md が全等級共通のハブ**で、そこから必要な下位構造へ分岐する。
`owned` が媒体プロファイルとして**実値**を持つのは、単位変換を各工場に書かせると
Resolver を1本化した理由と同じズレが単位変換で再発するため。

## 取得メタは機械が読める位置に置く

各 `plugin.yaml` の `id:` の直後に `tier` と `source` を持つ。**上流が版を公開していない場合や
取得時の版を記録していなかった場合は、それを正直に書く**（`unversioned` / `unrecorded`）。
でっち上げた版は鮮度検査を黙って通してしまう。

```yaml
id: <box>
tier: owned | observed
source:
  upstream: <URL または none>
  upstream_version: <commit / 日付 / unversioned / unrecorded>
  fetched_at: <YYYY-MM-DD または n/a>
```

上流は流動的で定期的に変わる。**ズレは検出される側に回す** — 値を手で追い続けるのではなく、
鮮度と同値を検査で見る（`origin-skill-commonize` の `check_mirrors` と同じ姿勢）。

## ライセンスゲート（この skill が守る）

第三者アセット（`observed` の箱）の**資産値**（色トークン実値・フォント・ロゴ・テンプレート）を
成果物へ反映する前に、その箱の `references.md` に**ライセンス確認と人間承認の記録**があり、
`plugin.yaml` の `usage_policy.license_gate_approved` が `true` であることを確認する。
未承認なら値を出力に使わない（**方針や雰囲気の参照は可、実値の反映は不可**）。

**第三者 DESIGN.md の参照元は `awesome-design-md-jp` に限定する。** 個別サイトはユーザー指示制で、
この skill が自律的に外部サイトを探索・スクレイピングしない。

## Plugin Resolver

1. `plugins/*/plugin.yaml` をスキャンし、`supported_media` に対象媒体を含む箱を候補にする。
2. `priority.contexts` のうちタスク文脈に一致するスコアが高い箱を選ぶ。
   文脈一致がなければ `priority.default` を使う。
3. **同点のタイブレーク**: (a) プロジェクトローカルに近い箱、(b) それも同点なら id の辞書順で先。
   決めきれない・候補が拮抗する場合はユーザーに確認する。
4. `extends` があれば親を**1階層だけ**辿って合成対象に含める。**多段継承は不可。**

規則の詳細と受け入れテストは references/resolver.md。

## 返すもの・返さないもの

- 返す: 解決した箱の id（継承元を含む）／DESIGN.md と tokens の値／**出所**（`source`）／
  **ライセンス可否**（ゲートの結果）／媒体プロファイルの値（`owned` のみ）
- 返さない: 成果物／媒体固有の制約（工場の `media/` が持つ）／情報設計・メッセージ設計

## やってはいけないこと

- **ゲートを通らずに `observed` の資産値を出力へ流す。** 方針の参照とは別物。
- **多段継承を辿る。** 1階層で止める（`extends` の解釈がズレた実測がある）。
- **取得メタの版をでっち上げる。** 分からないなら `unrecorded` と書く。
- **この skill で成果物を作る。** 倉庫は値と選択規則だけを持つ。
- **同じ選択ロジックを工場側にも書く。** 2実装になった時点でズレが始まる。
