---
schema_version: 2
id: ISSUE-20260921-historical-trouble-mining
status: active
created_at: 2026-09-21
updated_at: 2026-09-22
branch: codex/historical-trouble-mining
pr: ""
related_specs: []
related_guides: []
guide_impact: none
guide_impact_reason: "現行挙動は own-trouble-log/SKILL.md と references/status.md が正典であり、別ガイドは増やさない"
---

# 履歴トラブル226件を検索可能な知識へ変換する

## Goal

履歴枠にある226件を現役backlogへ戻さず、全件を検索・比較可能なパターン索引へ変換する。
候補 response と正式 status を証拠強度で分離し、entry 本文と既存レポートは不変に保つ。

## Acceptance

- verify: machine — python3 -m pytest -q own-trouble-log/scripts/tests && historical_index.py validate が全snapshot entryを被覆し exit 0
- verify: machine — status_ledger.py validate と skill_lint.sh が exit 0
- verify: human-review — 形式化候補の hook / skill / permission / manual 推奨を実装前に人間が承認する

## Current Status

as of 2026-09-22 — 226件の索引化・実データ適用・2件の重複pattern統合・対象skillの検証が完了した。全skill lintは別作業由来のown-doc-updateテスト2件が未解消のため、branch統合を停止している。

## Next Actions

- `codex/historical-trouble-mining` の変更をreviewして統合する。
- 承認済み: `destructive-action-without-safe-check`（5件、`origin-git-cleanup` 所有、warn-only skill案）。実装は別issueで行う。
- 承認済み: `nondeterministic-or-invalid-measurement`（4件、`own-trouble-log` 所有、warn-only/manual-review案）。実装は別issueで行う。
- 承認済み: `semantic-source-of-truth-drift`（4件、`origin-doc-update` 所有、warn-only/manual-review案）。実装は別issueで行う。
- `completion-race-stale-run`（3件）は既存 `verification-target-mismatch` へ統合済み。
- 残り25件の新規形式化候補は、人間承認後に所有先ごとの別issueで実装する。
- 全skill lintのown-doc-update 2件失敗は、グローバルdocs-validator hookがfixtureの意図的な不正docsをcommit拒否する既存環境問題。hookを無効化せず、別作業として解消するまでmergeしない。

## Guide Impact

- Decision: none
- Target or reason: 現行挙動の正典である own-trouble-log/SKILL.md と references/status.md を同時更新するため、別guideは不要。

## Notes

- `legacy_unrecorded` は評価負債ではないが、分析入力として随時利用できる。
- `mechanism` / `thematic` の候補対応では status を更新しない。
- 新しいhook・skillは提案までとし、このissueでは実装しない。
- `formalization_state=candidate` は候補の分類を示すため、承認済み候補もこの索引値を維持する。承認はこのissueのログで管理する。
- 実行結果: mechanism 203、thematic 23、direct 0、unclassified 0。
- response候補付きlinkは156件だが正式紐付けは0件で、status遷移も0件。
- 初回実行時は `legacy_unrecorded` 226件と `evaluation_pending_entries` 30件の分離を維持した。後続の別セッション更新後の現在値は31件で、今回の統合によるstatus変更ではない。

## Log

- 2026-09-21 — 約1.17MB・22,869行の226件を対象に確定。旧レポートからentry単位で復元できる既知クラスタは16件のみ。
- 2026-09-21 — historical_index.py の赤テストを確認後、snapshot / validate / summary / atomic apply を実装。専用テスト7件が合格。
- 2026-09-21 — 12件pilotが12/12件でschema合格。75・75・76件を3 agentで各1回digestし、再試行0回。
- 2026-09-21 — 生pattern 79件を重複統合して71件へ整理。226/226件、欠落0件で保管ルートへ適用。
- 2026-09-21 — entry本文226件の不変、historical index validate、status ledger validate、docs validator、全skill lintを確認。
- 2026-09-22 — `runtime-context-mismatch` が既存 `verification-target-mismatch` と重複すると人間承認され、11 linkを統合。pattern 71→70、link 226、status遷移0を確認。
- 2026-09-22 — `destructive-action-without-safe-check`（5件、origin-git-cleanup、warn-only skill案）を人間承認。実装・hook変更・status遷移はこのissueでは行わない。
- 2026-09-22 — `nondeterministic-or-invalid-measurement`（4件、own-trouble-log、warn-only/manual-review案）を人間承認。実装・hook変更・status遷移はこのissueでは行わない。
- 2026-09-22 — `semantic-source-of-truth-drift`（4件、origin-doc-update、warn-only/manual-review案）を人間承認。実装・hook変更・status遷移はこのissueでは行わない。
- 2026-09-22 — `completion-race-stale-run` の3 linkを既存 `verification-target-mismatch` へ統合。pattern 70→69、link 226、status遷移0を確認。同日2件目のためsuffix付き不変history reportを作成。
- 2026-09-22 — 対象skill 44テスト、historical/status/docs validatorは合格。全skill lintはown-doc-updateのfixture commit拒否2件で停止し、hook回避なしでmerge保留。

## Completion

- [x] Implementation completed or intentionally not needed
- [x] Specs updated if direction or requirements changed
- [x] Guide impact classified before implementation
- [x] Guides updated in the same slice if implemented behavior changed
- [ ] Branch merged and cleaned up (or intentionally kept — note why)
- [x] 00_index.md updated
- [ ] Moved to docs/issues/archive/ when complete
