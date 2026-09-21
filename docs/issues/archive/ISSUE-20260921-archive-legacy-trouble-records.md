---
schema_version: 2
id: ISSUE-20260921-archive-legacy-trouble-records
status: archived
created_at: 2026-09-21
updated_at: 2026-09-21
branch: main
pr: ""
related_specs: []
related_guides: []
guide_impact: none
guide_impact_reason: "own-trouble-log の現行仕様は SKILL.md と references/status.md が正典であり、別ガイドは増やさない"
---

# 履歴トラブル記録を現役評価負債から分離する

## Goal

台帳導入前の `legacy_unrecorded` 226件を監査用履歴として保持しつつ、現在評価すべき
`implemented_unverified` とは別の件数として表示し、解消不能な backlog を作らない。

## Acceptance

- verify: machine — python3 -m pytest -q own-trouble-log/scripts/tests && status_ledger.py summary が historical_unlinked と evaluation_pending を分離する
  <!-- `machine — <command and expected result>`, or `human-review — <who reviews what>` -->

## Current Status

as of 2026-09-21 — 履歴226件を現役評価待ちから分離し、検証後にissueをarchive済み。

## Next Actions

- 今回の作業に追加対応はなく、再発または新しい一次証拠が出た履歴だけをresponseへ紐付ける。

## Guide Impact

- Decision: none
- Target or reason: 現行挙動の正典である `own-trouble-log/SKILL.md` と
  `own-trouble-log/references/status.md` を直接更新したため、別ガイドは増やさない。

## Notes

- ADR: `docs/adrs/ADR-20260921-archive-legacy-trouble-records.md`。
- entry本文と過去レポートは削除・書換えせず、存在しない実装証拠も補わない。
- `historical_unlinked_entries` は監査用履歴、`evaluation_pending_entries` は
  `implemented_unverified` の現役評価対象だけを表す。

## Log

- 2026-09-21 — 集計の分離テストを先に追加し、変更前の失敗と変更後の成功を確認した。
- 2026-09-21 — 226件の可変台帳 `next_action` を、再発または新証拠がある場合だけ再開する方針へ更新した。
- 2026-09-21 — 専用テスト34件と台帳645行のvalidateが成功した。

## Completion

- [x] Implementation completed or intentionally not needed
- [x] Specs updated if direction or requirements changed
- [x] Guide impact classified before implementation
- [x] Guides updated in the same slice if implemented behavior changed
- [x] Branch merged and cleaned up (or intentionally kept — direct maintenance on `main`)
- [x] 00_index.md updated
- [x] Moved to docs/issues/archive/ when complete
