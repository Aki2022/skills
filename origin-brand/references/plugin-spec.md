---
title: plugin.yaml の正式スキーマ（等級・取得メタ・ライセンスゲート）
---

# plugin.yaml スキーマ仕様

各箱（`plugins/<box>/`）の `plugin.yaml` の正式スキーマ。倉庫はこれをスキャンして
候補一覧を作る。**`media.yaml` の仕様は本書には無い** — 媒体は工場の所掌なので
`origin-web-design/references/media-spec.md` にある。

## 等級と取得メタ（決定8・2026-09-09 追加）

`id:` の直後に置く。**機械が読む位置を固定する**ため、順序を変えない。

```yaml
id: <box>                 # 必須。ディレクトリ名と一致
tier: owned | observed    # 必須。等級は2つだけ（licensed は廃止）
source:                   # 必須
  upstream: <URL または none>
  upstream_version: <commit / 日付 / unversioned / unrecorded>
  fetched_at: <YYYY-MM-DD または n/a>
```

| 値 | 意味 |
| --- | --- |
| `upstream: none` | 自社資産。外部上流を持たない（`owned` はこれ） |
| `upstream_version: unversioned` | **上流が版を公開していない** |
| `upstream_version: unrecorded` | 上流に版はあるが**取得時に記録しなかった** |
| `fetched_at: n/a` | 取得という行為が無い（`owned`） |

**`unversioned` と `unrecorded` を書き分ける。** 前者は上流の性質、後者はこちらの記録漏れで、
直し方が違う。**でっち上げた版は鮮度検査を黙って通す**ので、分からないなら `unrecorded` と書く。

## plugin.yaml

```yaml
id: digital-agency # 必須。ディレクトリ名と一致させる
name: Digital Agency Design System Plugin # 必須。人間可読名
version: 0.1.0 # 必須。semver
status: experimental # experimental | stable
license: "See references.md" # 必須。詳細は references.md に集約
design_md_version: draft # 参照した DESIGN.md 仕様バージョン（Google仕様の変化に備える）
language:
  primary: ja
  supported: [ja, en]
extends: [] # 継承する親 plugin id。MVPは最大1個・1階層。空配列=継承なし

supported_media: [html, dashboard, pptx] # 対応媒体。Media Resolver がこれで候補を絞る

priority:
  default: 50 # 文脈一致なし時のスコア
  contexts: # タスク文脈ごとのスコア（高いほど優先）
    public-sector: 90
    dashboard: 95
    corporate-presentation: 40

entrypoints:
  design_md: DESIGN.md # 必須。デザイン方針の入口
  references: references.md # 必須。出典・ライセンス・承認記録

assets:
  tokens: [tokens/color.json, tokens/typography.json, tokens/spacing.json]
  icons: [assets/icons/]
  templates:
    {
      html: templates/html/,
      dashboard: templates/powerbi/,
      pptx: templates/pptx/,
    }
  examples: [examples/dashboard/, examples/html/]

validators: [validators/accessibility.md, validators/dashboard.md]

usage_policy:
  allow_reference: true # 参照利用の可否
  allow_asset_copy: depends_on_license # true | false | depends_on_license
  require_attribution: true # クレジット表記の要否
  prohibit_font_redistribution: true # フォント再配布の禁止
  license_gate_approved: false # 【重要】人間承認が済むまで false。true になるまで資産値を出力に使わない
```

### 必須フィールド

`id`, `name`, `version`, `status`, `license`, `supported_media`, `entrypoints.design_md`,
`entrypoints.references`, `usage_policy`。他は任意（無い場合は該当機能を使わない）。

### license_gate_approved の扱い

- `false`（デフォルト）の間は、DESIGN.md の**原則プローズ**は参照してよいが、tokens の実値・
  フォント・ロゴ・テンプレート等の**資産**を成果物へ反映してはいけない。
- 人間レビューで承認されたら `references.md` に承認記録を書き、`true` に更新する。
