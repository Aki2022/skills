---
title: Common Validator — Brand Consistency
method: self-check
---

# Brand Consistency Validator [自己点検]

ブランド一貫性と、第三者ブランドの誤用・コピーがないかを点検する。

## チェック項目

- 使用した色・フォント・余白が選択 Plugin の DESIGN.md / tokens の範囲内か。
- 複数 Plugin の意匠が混在して一貫性を欠いていないか。
- **参照元ブランド（awesome-design-md-jp 等）の UI・ロゴ・商標・固有表現をコピーしていないか**。
  参照は抽象化した方針・密度・トーンのみ。具体ビジュアルの再利用は不可。
- ライセンス未承認（`usage_policy.license_gate_approved: false`）の Plugin の資産値を使っていないか。
- 公的機関を想起させる表現で、公式資料と誤認される作りになっていないか（digital-agency 系で特に注意）。

## 不合格時

コピー・誤用箇所を除去し、参照は抽象特性の反映に留める。ライセンス未承認資産は Runtime デフォルトへ差し替える。
