---
title: "UC-03: awesome-design-md-jp を参考にした日本語LP"
media: html
plugin: awesome-design-md-jp-import
context: [japanese-ui, landing-page]
---

# UC-03: 日本語SaaS風LP（参照利用）

## 入力

- 目的: 日本語SaaS風のLPを作成
- 媒体: HTML
- 参考: ユーザーが `awesome-design-md-jp` から指定した個別サイトの DESIGN.md

## 期待される Runtime 挙動

1. **ユーザーが参考にする個別サイトを指示する**（Skillは自律探索しない）。
2. Media Resolver が `html` を選ぶ。
3. `awesome-design-md-jp-import` プラグインが、指示された DESIGN.md を取得し内部形式に正規化する。
4. Context Composer:
   - 参照 DESIGN.md は**優先度6**。抽象化した方針・密度・CJKトーンのみ取り込み、具体値はコピーしない。
   - 実際の色・フォント・レイアウトは自社 Plugin（例: my_company, 優先度3）から採る。
5. `[自動]`（コントラスト・セマンティック・alt）→ `[自己点検]`（CJK可読性、参照元コピーの有無）で点検。
   Brand Consistency Validator で「参照元ブランドをコピーしていないか」を必ず確認。
6. 出力に参照元と自社ブランドの差分説明を含める。

## 期待される出力

- 日本語UIに適したLP設計
- CJKタイポグラフィ配慮
- 参照元と自社ブランドの差分説明

## 禁止事項（import-policy と連動）

既存サービスのUIコピー / ロゴ・商標・固有ビジュアルの再利用 / 競合と誤認される表現。
