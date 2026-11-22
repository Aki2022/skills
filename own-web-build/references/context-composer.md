---
title: Context Composer — design/ と媒体構造規約
---

# Context Composer

生成する意匠は repo の `design/` だけから読む。L0が今回生成した解決済みマニフェストの値を使い、
DTCGのaliasや継承をここで解決しない。

1. `design/DESIGN.md` の方針・トークン参照と `.resolved/web.json` の値を読む。
2. 媒体の `type: constraint`（HTML構造・ARIA・レスポンシブ等）を守る。
   数値のアクセシビリティ下限はマニフェストの `minimums` を使い、緩めない。
3. `type: recommendation` は意匠に矛盾しない範囲だけ適用する。
4. 競合・値不足は `own-design-author` へ戻し、L0を再実行する。

倉庫のPlugin優先度・extends・参照DESIGN.md・Runtime既定値は合成しない。
参照が指示された場合、許可元と権利を `own-brand-resolve` で確認したうえで、方針と値を
`proposed` として `design/` に写す。`reference:<id>` と `design/NOTICE` を伴い、
ブランド値の承認とは分ける。本家の英語UIのtypography値は写さない。

各値の出所・提案状態・競合・フォールバック値を報告する。フォールバックは準拠値として扱わない。
