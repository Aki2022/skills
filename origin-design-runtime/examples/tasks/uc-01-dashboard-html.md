---
title: "UC-01: Digital Agency風 行政向けダッシュボードHTML"
media: html
plugin: digital-agency
context: [public-sector, dashboard]
---

# UC-01: 自治体向け人口・財政ダッシュボード（HTML）

## 入力

- 目的: 自治体向けの人口・財政ダッシュボード
- 媒体: HTML
- デザイン: Digital Agency（明示指定）

## 期待される Runtime 挙動

1. Media Resolver が `html` を選ぶ。
2. Plugin Resolver が `digital-agency` を選ぶ（`priority.contexts.dashboard=95` / `public-sector=90`）。
3. Context Composer が優先度順に合成する:
   - `media/html` の `type: constraint`（コントラスト・ARIA・レスポンシブ）を最優先で確定（優先度2）。
   - digital-agency の DESIGN.md / tokens を意匠として適用（優先度3）。
   - グリッド等の recommendation は Plugin 指定を優先（優先度5 < 3）。
4. `[自動]` Validator（コントラスト・セマンティック・alt・ARIA）→ `[自己点検]`（KPIとチャートの整合、
   色の意味づけ、単位/期間/出典）の順に点検。
5. 出力に「使用Plugin: digital-agency / 主要制約: WCAG AA, レスポンシブ / 未解決項目」を記録。

## 期待される出力

- HTML/CSS(/JS または React/Tailwind)の実装方針
- 画面構成（KPIカード / 時系列 / 棒グラフ / 地図 / フィルター）
- 利用アセット一覧
- 品質チェック結果

## ライセンス注意

digital-agency の `usage_policy.license_gate_approved` が `true` の場合のみ tokens 実値・テンプレートを反映。
未承認なら原則プローズ＋Runtimeデフォルトで生成し、その旨を明記する。
