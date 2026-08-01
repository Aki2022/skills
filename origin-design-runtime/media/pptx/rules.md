---
title: PPTX Media Rules
---

# PPTX Media Rules

## 責務

- スライド比率（16:9 / 4:3）/ スライドマスター / 余白 / グリッド
- タイトル・本文・脚注の文字サイズ / 図表・表・チャートの配置
- 1スライド1メッセージ原則 / テンプレートPPTX・theme（.crtx等）利用方針

## 制約（constraint）

- **比率**（pptx-aspect-ratio）: 既定 16:9。指定があれば従う。
- **最小フォント**（pptx-min-font）: 本文18pt・注釈12pt相当を下回らない。`[自動]`（pptx-font-size）。
- **1スライド1メッセージ**（pptx-one-message）: タイトルは要点/結論を述べる。`[自己点検]`。
- **マスター利用**（pptx-master）: レイアウトを使い、要素の直接配置による乱れを避ける。

## 推奨（recommendation）

- **余白/グリッド**（pptx-margin-grid）: 整列を統一。Plugin 指定があれば優先。
- **編集可能構造**（pptx-editable）: 再現性のため画像化を避け、テキスト/表/ネイティブ図形に寄せる。

## 連携

PPTX の実生成は `pptx` Skill（テンプレート + style.yaml）と連携できる。design-runtime は
デザイン方針・構成・チェックリストを与え、実ファイル生成はそちらに委譲してよい。
