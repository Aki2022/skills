---
title: "UC-02: My Companyブランドの新規事業提案PPTX"
media: pptx
plugin: my_company
context: [corporate-presentation]
---

# UC-02: 新規事業提案資料（PPTX）

## 入力

- 目的: 新規事業提案資料
- 媒体: PPTX
- デザイン: My Company（明示指定）

## 期待される Runtime 挙動

1. Media Resolver が `pptx` を選ぶ。
2. Plugin Resolver が `my_company` を選ぶ。`my_company` は `extends: [digital-agency]`（1階層）なので、
   親 digital-agency のアクセシビリティ/公共性の原則を合成対象に含める。
3. Context Composer:
   - `media/pptx` の `type: constraint`（スライド比率・最小フォント・1スライド1メッセージ）を最優先（優先度2）。
   - my_company（子, 優先度3）の意匠 → digital-agency（親, 優先度4）の順で合成。子が親を上書き。
4. `[自動]`（最小フォントサイズ）→ `[自己点検]`（1スライド1メッセージ、タイトルが結論、余白/整列、
   図表の可読性、テンプレート/ブランド一致）の順に点検。
5. 出力に使用Plugin（my_company←digital-agency 継承）・主要制約・未解決項目を記録。

## 期待される出力

- スライド構成
- スライドデザイン仕様
- PPTX生成指示（テンプレート/マスター/レイアウト重視。編集可能構造に寄せる）
- チェックリスト

## 検証ポイント（受け入れテスト）

my_company 単体の指定で、digital-agency の原則が1階層だけ継承合成されること。多段継承が起きないこと。
