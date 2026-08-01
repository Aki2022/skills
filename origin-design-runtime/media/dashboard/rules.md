---
title: Dashboard Media Rules
---

# Dashboard Media Rules

## 責務

- KPIカード / 時系列チャート / 棒グラフ / 地図 / フィルター / 凡例
- 色の意味付け / データ粒度 / 誤読防止 / アクセシビリティ

## 制約（constraint）

- **メタ明示**（dash-meta）: 単位・期間・出典。`[自動]`（dashboard-meta）で必須フィールド存在を確認。
- **色の意味**（dash-color-meaning）: 意味を持って一貫使用。色のみに依存した情報伝達をしない。
- **軸・スケール**（dash-axis）: 誤読を招かない。棒グラフの縦軸は原則0起点。

## 推奨（recommendation）

- **チャート選択**（dash-chart-choice）: 時系列=折れ線 / カテゴリ比較=棒 / 地理分布=地図 / 構成比=積み上げ。
- **KPI階層**（dash-kpi-hierarchy）: 上部にKPIカード、下部に根拠チャート。関係を明確に。
- **フィルター**（dash-filter）: 凡例・粒度切替をわかりやすく。

## 参考

デジタル庁「ダッシュボードデザインの実践ガイドブック」を参照方針とする（digital-agency plugin の
references.md でライセンス・出典を管理。承認前は原則のみ参照）。
