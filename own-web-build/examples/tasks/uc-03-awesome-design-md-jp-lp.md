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

1. `scripts/preflight.py` でL0を実行する。design/不在なら生成せずown-design-authorへ渡す。
2. 指示された個別サイトだけを参照し、import-policyの権利確認を行う。
3. 方針と値をdesign/へproposedとして写し、reference出所とNOTICEを付ける。
4. L0を再実行してから、解決済みマニフェストで生成する。倉庫の値を継承しない。
5. L1は必ず実行し、L2は人間へ実施可否を確認する。
6. 提案状態・出所・未カバー・実施しなかった工程と理由を報告する。

## 期待される出力

- 日本語UIに適したLP設計
- CJKタイポグラフィ配慮
- 参照元と自社ブランドの差分説明

## 禁止事項（import-policy と連動）

既存サービスのUIコピー / ロゴ・商標・固有ビジュアルの再利用 / 競合と誤認される表現。
