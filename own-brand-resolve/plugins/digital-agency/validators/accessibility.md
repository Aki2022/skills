---
title: Digital Agency — Accessibility Validator
method: mixed
---

# Digital Agency Accessibility Validator

`validators/common/accessibility.md` に加え、行政・公共向けに厳格化した項目。

## 追加チェック

- コントラストは WCAG **AA を必須**、重要導線は AAA を目指す（`[自動]` contrast）。
- フォーム: ラベル関連付け・エラーメッセージのテキスト明示・必須項目の非色依存表示（`[自己点検]`）。
- キーボード操作・フォーカス順序が論理的か（`[自己点検]`）。
- 公的サービス想定で、支援技術（スクリーンリーダー）での読み上げ順が破綻しないか（`[自己点検]`）。
- 色のみで状態（エラー/成功）を伝えていないか（`[自己点検]`）。

## 不合格時

constraint 該当（コントラスト・ARIA・alt）は Plugin 意匠より優先して修正する。
