---
title: Context Composer — コンフリクト解決アルゴリズム
---

# Context Composer

DESIGN.md（プロジェクト / Plugin / 参照）と Media Rule、tokens の間で値が競合したとき、
どの値を採用するかを決める。SKILL.md 手順8 で必ずこの順序を適用する。

## 優先度（先勝ち。数字が小さいほど強い）

1. **プロジェクトローカル DESIGN.md**（リポジトリ直下等にある場合）
2. **Media Rule のうち `type: constraint`**（アクセシビリティ基準、ARIA、コントラスト比、
   レスポンシブ・ブレークポイント、スライド比率、1スライド1メッセージ等の構造制約）
3. **明示指定された Design Plugin（子）** の DESIGN.md / tokens
4. **子が継承する親 Plugin**（`extends`、1階層）
5. **Media Rule のうち `type: recommendation`**（グリッド列数の目安、サンプルレイアウト等）
6. **参照 DESIGN.md**（awesome-design-md-jp 等。雰囲気・密度・CJK の参考に限定）
7. **Runtime デフォルト方針**（`architecture.md`）

## 判断の要点

- **Media Rule は constraint と recommendation で強さが変わる**。コントラスト比 4.5:1 は
  constraint（優先度2）なので Plugin の配色より優先される。一方グリッド列数は recommendation
  （優先度5）なので Plugin のレイアウト方針が勝つ。
- constraint / recommendation の分類は `media/<media>/media.yaml` の `rules[].type` で判定する。
  Composer は自分で分類を推測しない。
- 参照 DESIGN.md（優先度6）は**具体値をコピーしない**。抽象化した方針・密度・トーンだけを
  取り込み、実際の色・フォント・レイアウトは優先度3〜5から採る。

## 合成手順

```text
1. 空のコンテキストを用意する。
2. 優先度7→1の順に各ソースの値を書き込む（後に書いた強いソースが上書きする）。
   ※「先勝ち」を上書き実装で表現: 弱い順に入れ、強い順で上書きする。
3. type: constraint（優先度2）は最後まで保持されるよう、Plugin意匠(3,4)より後に上書きしない。
   → 実装上は「constraintは書き込み後ロックし、以降のソースで上書き禁止」とする。
4. 各値に出所（どのソース由来か）を付記し、生成指示に含める。
5. 競合で捨てた値があれば記録し、必要なら成果物の注記に残す。
```

## ライセンスゲートとの関係

Design Plugin（優先度3・4）の**資産値**（tokens 実値・フォント・ロゴ）は、その plugin の
`usage_policy.license_gate_approved: true` かつ `references.md` に承認記録がある場合のみ合成に含める。
未承認なら DESIGN.md の原則プローズのみを使い、資産値は Runtime デフォルト（優先度7）で埋める。
