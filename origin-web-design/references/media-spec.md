---
title: media.yaml の正式スキーマ
---

# media.yaml スキーマ仕様

`media/<media>/media.yaml` の正式スキーマ。工場はこれをスキャンして利用可能な媒体一覧を得る
（ディレクトリを足すだけで媒体が増える）。

**`plugin.yaml` の仕様は本書には無い** — 意匠は倉庫の所掌なので
`origin-brand/references/plugin-spec.md` にある。

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
