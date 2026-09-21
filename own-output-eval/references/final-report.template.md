---
title: 最終報告（ループ終了時に必ず）
---

# 評価ループ最終報告 — <対象>

| 項目 | 値 |
| --- | --- |
| 終了理由 | pass（run_round exit 0）/ not-converging（exit 3）/ max-rounds 到達（exit 4）。**exit 2 で終わった場合は報告を書かない** — 入力の不備なので直して再実行する |
| 巡数 | k / N |
| スコア推移 | r1: min/avg/must_fix/unclear/blocking_count/findings_count → r2: … （各 `round_K/verdict.json` から。どちらも機械が数えた値。**収束判定に使われたのは `blocking_count`**） |
| 実行した方式と理由 | A: …／B: …／C: … |
| 実行しなかった方式と理由 | 例: A — 対象がデザインでテストが書けない |
| 前提条件（通した決定論ゲート） | 呼び出し側の宣言をそのまま |
| C 保留リスト | 正典変更に及ぶ指摘。人間が次に判断する材料 |
| 却下した提案 | `rejected` の一覧と理由 |
| **評価者の起動条件** | 各巡の `round_K/provenance.json` をそのまま転記（`verdict.json` の `provenance`）。手書きしない — 判定時に機械が検証済みの申告である |
| 未収束・上限フラグ | あり／なし（未収束なら比較した巡の `blocking_count` と方式集合を併記） |
| 状態ファイル | `state.json` と `round_k/` のパス |

この報告の後に人間ゲートが1つある。**本 skill の pass は人間確定の代わりにならない。**
