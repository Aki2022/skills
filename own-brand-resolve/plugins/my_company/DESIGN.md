# DESIGN.md — My Company

My Company の提案書・営業資料・事業開発資料・AI生成UIに共通するブランド基盤。
`digital-agency` を1階層継承し、アクセシビリティ・公共性・ダッシュボード原則を引き継いだ上で、
コンサルティング資料としての知性・高級感・余白を重ねる。

## Visual Theme

- 知的・高級感・落ち着き。余白を活かした「静かな説得力」。
- コンサルティング資料として読みやすく、公共・大企業向けにも通用する品格。
- 日本語資料に最適化。

## Color Roles（継承 + 上書き）

- 親（digital-agency）の高コントラスト・非色依存の原則を継承。
- **原則グレー階調。色相を持つのはトーン色だけ**（positive / negative）。
  「落ち着いた濃色の Primary + 金のアクセント」は 2026-09-09 に**破棄した下書き値**で、採らない。
- 背景は白基調、面はごく薄いグレー（`tokens/color.json` の `shared`）。
- **本文の濃度は媒体で違う** — pptx は `#7F7F7F`（会社標準・3.0:1 運用）、web は `#666666`（5.74:1）。
  `media/html` の 4.5:1 は意匠より強い制約なので、pptx の値を web に持ち込めない。

## Typography

- **フォントは媒体別に持つ**（`tokens/typography.json` の `mediaProfiles`）。
  pptx は `BIZ UDPGothic`（全デッキ一律）、web は `Inter` + `Noto Sans JP`（会社サイトの実値）。
  どちらもフォント同梱はしない。
- 見出しは抑制の効いたウェイト運用。本文は行間広めで読みやすさ優先（CJK は 1.7 以上）。

## Spacing / Layout

- 親の 8px スケールを継承。**余白を親より広めに取り**、情報密度を落として高級感を出す。
- 1スライド1メッセージ、1画面1論点。

## Components

- 図表は装飾を削ぎ、データと結論を主役に。罫線は最小限。

## Content Tone

- 論理的で簡潔。結論先出し（PREP / ピラミッド構造）。過度な修飾を避ける。

## Do / Don't

- Do: 余白・抑制・一貫・結論先出し・根拠の明示。
- Don't: 詰め込み、装飾過多、トーンの不統一、親のアクセシビリティ制約を破ること。

## 継承の扱い

- `extends: [digital-agency]`（1階層）。競合時は Context Composer の優先度で
  my_company（子, 優先度3）が digital-agency（親, 優先度4）を上書きする。
- ただし Media Rule の `type: constraint`（アクセシビリティ等）は双方より優先される。
