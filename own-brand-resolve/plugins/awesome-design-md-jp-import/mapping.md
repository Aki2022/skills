---
title: 参照 DESIGN.md の正規化マッピング
---

# 正規化マッピング

指示された DESIGN.md を repo の `design/` に写す。`import-policy.md` の許可元・権利確認に従う。

| カテゴリ | 写せる内容 | 写さないもの |
| --- | --- | --- |
| Visual Theme | 一般化した方針・トーンを DESIGN.md の散文に | ブランド名を想起させる固有表現 |
| Color Roles | 色と役割を proposed トークンに | ロゴ・固有ビジュアル |
| Typography | -jp またはオーナー指定の和文フォントを proposed に | 本家からの typography 値、フォントファイル |
| Spacing / Layout | 余白・寸法を proposed トークンに | 画面構成・実装のコピー |
| Components | 角・影を proposed トークンに | 固有コンポーネントの意匠 |
| Iconography | 一般化したテイストを散文に | アイコンの実ファイル |
| Motion / Accessibility | 一般化した方針を散文に | 媒体最低値を緩める変更 |
| Content Tone / Do-Don't | 一般化した文体・原則を散文に | 固有のコピー文言 |
| Examples | 方向性の言語化 | スクリーンショット・実装 |

数値は `tokens/*.tokens.json` だけに置き、散文は `token:<path>` で参照する。
全ての参照値は `reference:<id>` と対応する `NOTICE` 行を伴う。参照値を採っただけでは承認されない。
本家の typography は空のまま未決定と報告し、別の既定値で埋めない。
