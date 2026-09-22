---
id: ADR-20260921-separate-historical-analysis-from-response-status
status: accepted
scope: development
created_at: 2026-09-21
updated_at: 2026-09-21
source_workstreams: []
source_issues: [ISSUE-20260921-historical-trouble-mining]
related_specs: []
related_guides: []
supersedes: ""
superseded_by: ""
---

# Decision: 履歴分析をresponse statusから分離する

## Context

過去レポートに列挙されたが当時のresponse evidence欄が無い226件は、現役の評価backlogへ
戻すと存在しない証拠の補作と大量の未評価負債を生む。一方で履歴枠へ置くだけでは、同型比較や
形式化候補の発見に再利用できない。候補対応の類似度と、status変更に必要な一次証拠を同じ列で
扱うと、テーマの近さだけで実装済み・再発・有効性を誤認する。

## Decision

履歴分析を独立した多対多索引として保持し、response statusから分離する。各entryを
`direct` / `mechanism` / `thematic` / `unclassified` の証拠強度でpatternへ紐付ける。
`mechanism` と `thematic` は候補検索だけに使い、statusを変更しない。status変更は対応する
`direct` 行と状態固有の一次証拠がある場合に限る。entry本文・既存レポートはappend-onlyとし、
分析snapshotと判断は日付付きの不変レポートに残す。

## Alternatives Considered

- 226件を現役backlogへ戻す: 証拠の無い評価を強制し、現役の評価件数を歪めるため採らない。
- 履歴を削除する: 再発比較と形式化候補の根拠を失うため採らない。
- entry本文へタグを追記する: append-onlyと並行記録の設計を壊すため採らない。
- 候補responseを直ちに正式紐付けする: テーマ類似と直接証拠を混同するため採らない。

## Consequences

### Positive

- 全件の分析有無と未分類理由を機械検証できる。
- 横断検索に使いつつ、現役の未評価件数を増やさない。
- 後日の新しい証拠で多対多リンクを追加できる。

### Negative or Follow-up

- パターン正典とリンク索引の保守が増える。
- 形式化候補は人間承認後に別工程で実装する必要がある。
- `direct` 判定は本文または旧レポートの一次証拠を読み直す必要がある。

## Links

- [ISSUE-20260921-historical-trouble-mining](../issues/archive/ISSUE-20260921-historical-trouble-mining.md)
- `own-trouble-log/SKILL.md`
- `own-trouble-log/references/status.md`
