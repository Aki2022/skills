---
title: Plugin / Media スキーマ仕様
---

# Plugin / Media スキーマ仕様

`plugin.yaml`（Design/Asset Plugin）と `media.yaml`（Media Plugin）の正式スキーマ。
Runtime はこれらをスキャンして候補一覧を作る（`architecture.md` の発見機構参照）。

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

## media.yaml

```yaml
id: html # 必須。ディレクトリ名と一致
name: HTML / Web Media # 必須
version: 0.1.0
description: レスポンシブなHTML/Web画面の媒体ルール
rules:
  - id: html-contrast
    type: constraint # constraint（制約）| recommendation（推奨）
    summary: 本文テキストのコントラスト比は 4.5:1 以上
    validator: contrast # 対応する validator 名（[自動]検証に使う）
  - id: html-grid-cols
    type: recommendation
    summary: 標準グリッドは 12 カラムを目安にする
```

### type の意味（context-composer.md と連動）

- `constraint`: 壊れると使い物にならない / アクセシビリティ上問題になる制約。Context Composer の
  優先度 **2**。Design Plugin の意匠より優先される（例: コントラスト比、ARIA、スライド比率）。
- `recommendation`: 見た目の推奨・目安。優先度 **5**。Design Plugin のレイアウト方針が優先される
  （例: グリッド列数の目安、サンプルレイアウト）。

Runtime は `type: constraint` を「必ず守る」、`type: recommendation` を「Plugin指定がなければ従う」
として扱う。
