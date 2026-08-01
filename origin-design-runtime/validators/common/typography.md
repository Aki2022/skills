---
title: Common Validator — Typography
method: self-check
---

# Typography Validator [自己点検]

タイポグラフィが DESIGN.md / Media Rule と一貫しているか点検する。

## チェック項目

- フォントファミリが DESIGN.md 指定（または Runtime デフォルト）と一致するか。
- 見出し/本文/注釈の階層（サイズ・ウェイト）が一貫しているか。
- 本文サイズが媒体の最小サイズ制約を下回っていないか（PPTX の最小フォントは `[自動]` 側で判定）。
- 日本語（CJK）の行間・字間が読みやすいか。和欧混植でベースラインが破綻していないか。
- 1画面/1スライド内でフォントファミリを増やしすぎていないか（原則2種まで）。

## 不合格時

指定のタイプスケールに揃える。CJK 可読性は `validators/common/*` と DESIGN.md の Content Tone を参照。
