---
title: My Company 提案書PPTX Example
---

# My Company 提案書PPTX Example（UC-02 対応）

新規事業提案資料のスライド構成の参考。

## スライド構成（例）

1. 表紙（タイトル + 一言サマリ）
2. エグゼクティブサマリ（結論先出し）
3. 課題認識
4. 提案の全体像
5. 詳細（複数枚、1スライド1メッセージ）
6. 実行計画・体制
7. 期待効果（KPI）
8. まとめ・ネクストアクション

## デザイン適用

- 余白を広めに取り、1スライド1メッセージ（media/pptx pptx-one-message）。
- タイトルは結論を述べる（例:「〇〇により△△が可能」）。
- 本文18pt以上（media/pptx pptx-min-font, `[自動]`）。
- 色は my_company tokens の primary/accent を最小限に。データ系列は dataViz.categorical。
- アクセシビリティ制約（コントラスト等）は親 digital-agency から継承し順守。

## 継承の確認ポイント

my_company 単体指定で、digital-agency の原則が1階層だけ合成されること（多段継承なし）。
PPTX 実生成は `pptx` Skill に委譲してよい。
