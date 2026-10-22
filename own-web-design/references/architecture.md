---
title: Architecture — design/ と工場の境界
---

# Architecture

| 層 | 正典 | 責務 |
| --- | --- | --- |
| ブランド定義 | repo の `design/` | 方針・トークン・媒体プロファイル・アセット・出所・承認台帳 |
| 定義とL0 | `own-design-author` | 整合検査・DTCG解決・マニフェスト生成 |
| 参照ライブラリ | `own-brand-resolve` | 指示された参照だけを権利情報とともに渡す |
| 工場 | `own-web-design` | 前提検査・媒体判定・生成・L1/L2 |
| 媒体構造 | `media/` | HTML・dashboardの構造上の制約と推奨 |

経路は User Task → Media判定 → `scripts/preflight.py`（L0）→ design/の方針と解決済み値 →
生成 → L1 → 人間の実施確認後L2 → `own-design-route`。
L0の失格では生成せず、警告は明示して続行できる。L0/L1は確認せず実行する。

値が無ければ `own-design-author` へ戻す。Runtimeの中立値・倉庫の値を自動採用しない。
工場はDTCGを再解決しない。repoの設定に端末固有パスを書かない。
媒体追加は `media/<media>/media.yaml` で行い、pptxはこの工場へ追加しない。
