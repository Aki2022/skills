---
title: HTML Media Rules
---

# HTML Media Rules

`media.yaml` の各ルールの補足。`type: constraint` は必ず守り、`recommendation` は Plugin 指定が
なければ従う（優先度は `references/context-composer.md`）。

## 責務

- レスポンシブ設計 / セマンティックHTML / ARIA / キーボード操作 / カラーモード
- Tailwind / CSS 設計方針 / コンポーネント利用ルール

## 制約（constraint）

- **コントラスト**（html-contrast）: WCAG AA。`validators/common/contrast.md` で `[自動]` 検証。
- **セマンティック / ARIA / alt / キーボード**（html-semantic/aria/alt/keyboard）:
  `validators/common/accessibility.md`。alt・セマンティック・ARIA は `[自動]`、キーボードは `[自己点検]`。
- **レスポンシブ**（html-responsive）: 主要ブレークポイントで破綻しないこと。`[自己点検]`。

## 推奨（recommendation）

- **グリッド**（html-grid-cols）: 12カラム目安。Plugin にレイアウト指定があればそちらを優先。
- **カラーモード**（html-color-scheme）: ライト/ダーク対応が望ましい。

## 実装方針

- Tailwind か素の CSS かは Plugin / プロジェクト方針に従う。指定がなければユーティリティ優先で提案。
- コンポーネントは Plugin の examples / templates を優先利用する（ライセンス承認済みの場合）。
