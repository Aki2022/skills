---
title: Digital Agency HTML Example
---

# Digital Agency HTML Example

行政・公共向け HTML 画面の参考構成。tokens（color/typography/spacing）を CSS 変数に写像する例。

## CSS 変数への写像（例）

```css
:root {
  --color-primary: #0017c1; /* tokens/color.json roles.primary.default */
  --color-text: #1a1a1c;
  --color-bg: #ffffff;
  --color-border: #d8d8db;
  --font-sans:
    "Noto Sans JP", sans-serif; /* フォントは同梱せずWebフォント/環境依存 */
  --space-4: 16px; /* tokens/spacing.json */
  --space-5: 24px;
}
body {
  color: var(--color-text);
  background: var(--color-bg);
  font-family: var(--font-sans);
  line-height: 1.7;
}
```

## 構成の考え方

- `header` / `nav` / `main` / `footer` のランドマークを使う。
- 本文 16px・行間 1.7 以上で CJK 可読性を確保。
- コントラストは WCAG AA（本文 4.5:1 以上）。contrast Validator で自動検証。
- Material Symbols 由来アイコン（Apache 2.0）は意味の明確なものを最小限に。

## 注意

- デジタル庁ロゴ・紋章は使わない。公式資料と誤認される表現を避ける。
- コードスニペット（MIT）を参考にする場合、改変前提。未改変公開時は出典表記が必要（references.md）。
