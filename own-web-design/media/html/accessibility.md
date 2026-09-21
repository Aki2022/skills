---
title: HTML Accessibility
---

# HTML Accessibility

`validators/common/accessibility.md` を HTML 向けに具体化したもの。

## 必須（constraint）

- 見出しは `h1`→`h2`→`h3` の順に飛ばさず入れ子にする。
- ページ構造に `header` / `nav` / `main` / `footer` のランドマークを使う。
- インタラクティブ要素はネイティブ要素（`button`/`a`/`input`）を優先。div にロールを足すのは最後の手段。
- フォーム項目に `label` を関連づける。
- 画像は意味に応じて `alt` を付ける（装飾画像は `alt=""`）。
- フォーカス可視化（`:focus-visible`）を消さない。
- コントラストは WCAG AA。

## 自己点検

- キーボードのみで全機能に到達・操作できるか。
- 色だけで状態（エラー等）を伝えていないか（アイコン/テキストを併用）。
- `prefers-reduced-motion` に配慮しているか。
