---
schema_version: 2
id: ISSUE-20260921-historical-trouble-mining
status: active
created_at: 2026-09-21
updated_at: 2026-09-22
branch: codex/historical-trouble-mining
pr: "https://github.com/Aki2022/skills/pull/9"
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

as of 2026-09-22 — 226件の索引化・実データ適用・3件の重複pattern統合を実施し、対象skillの検証結果を確認した。commit 2c0bfa4をpushしてPR #9を作成したが、全skill lintは別作業由来のown-doc-updateテスト2件が未解消のため、mergeは人間ゲートで停止している。

## Next Actions

- PR #9の変更をreviewし、own-doc-updateのfixture/hook境界を解消して全skill lintが緑になった後にmergeする。
- 承認済み: `destructive-action-without-safe-check`（5件、`origin-git-cleanup` 所有、warn-only skill案）。実装は別issueで行う。
- 承認済み: `nondeterministic-or-invalid-measurement`（4件、`own-trouble-log` 所有、warn-only/manual-review案）。実装は別issueで行う。
- 承認済み: `semantic-source-of-truth-drift`（4件、`origin-doc-update` 所有、warn-only/manual-review案）。実装は別issueで行う。
- 承認済み: `decision-history-not-read`（3件、`own-trouble-log` 所有、履歴確認skill案）。実装は別issueで行う。
- 承認済み: `idempotence-postcondition-gap`（3件、`origin-doc-update` 所有、warn-only案）。実装は別issueで行う。
- 承認済み: `archive-parser-shape-collision`（2件、`own-doc-update` 所有、parser回帰テスト案）。実装は別issueで行う。
- 承認済み: `artifact-state-not-recorded`（2件、`origin-doc-update` 所有、docs更新後readback案）。実装は別issueで行う。
- 承認済み: `destructive-write-overwrite`（2件、`settings-hook-owner` 所有、dirty-state警告とmanual review案）。実装は別issueで行う。
- 承認済み: `formatter-boundary-gap`（2件、`origin-doc-update` 所有、warn-only案）。実装は別issueで行う。
- 承認済み: `formatter-mutation-bypass`（1件、`origin-skill-commonize` 所有、hook/skill境界確認案）。実装は別issueで行う。
- 承認済み: `git-history-reachability-gap`（1件、`origin-git-cleanup` 所有、warn-only案）。実装は別issueで行う。
- 承認済み: `permission-preflight`（2件、`own-trouble-log` 所有、permission preflight案）。実装は別issueで行う。
- 保留: `command-shape-permission-friction`（1件、`settings-hook-owner` 所有、warn-only案）。許可境界の実測後に再判断する。
- 保留: `premature-completion-state`（1件、`origin-doc-update` 所有、完了前readback案）。追加の直接証拠後に再判断する。
- 保留: `safe-operation-alternative`（1件、`own-trouble-log` 所有、warn-only案）。追加の直接証拠後に再判断する。
- 保留: `security-guard-boundary`（1件、`repository-security-hook` 所有、warn-only案）。意図的テストと実運用の判別証拠後に再判断する。
- 承認済み: `security-placeholder-exposure`（1件、`repository-security-hook` 所有、security lint/skill review案）。実装は別issueで行う。
- 承認済み: `sensitive-data-reuse`（2件、`repository-security-hook` 所有、redaction guardとmanual review案）。実装は別issueで行う。
- 保留: `hook-signal-quality`（1件、`settings-hook-owner` 所有、warn-only再評価案）。真陽性・誤警告の測定後に再判断する。
- 保留: `patch-context-staleness`（1件、`own-trouble-log` 所有、patch直前readback案）。再現性の追加証拠後に再判断する。
- 保留: `dependency-usage-not-checked`（2件、`origin-doc-update` 所有、usage search手順案）。利用実態の追加証拠後に再判断する。
- 件数1〜2かつ高リスク領域に該当しない次の候補は、証拠件数が増えるまで一括保留する: `deploy-scope-mismatch`、`diagnostic-state-not-persisted`、`ephemeral-delivery-reference`、`guard-read-write-discrimination`、`operator-visible-completion-evidence`、`wrong-corpus-or-durable-location`。
- `completion-race-stale-run`（3件）は既存 `verification-target-mismatch` へ統合済み。`silent-write-noop`（3件）は既存 `doc-hygiene-postconditions` へ統合済み。
- 承認済み14件は所有先ごとの別issueで実装する。保留13件は証拠件数が増えるまで承認質問を再開しない。
- 全pattern・候補を再照合し、追加の明確な重複がないことを確認した。response名だけが一致する候補は失敗機構が異なるため統合していない。
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
- 2026-09-22 — `decision-history-not-read`（3件、own-trouble-log、履歴確認skill案）を人間承認。実装・hook変更・status遷移はこのissueでは行わない。
- 2026-09-22 — `idempotence-postcondition-gap`（3件、origin-doc-update、warn-only案）を人間承認。実装・hook変更・status遷移はこのissueでは行わない。
- 2026-09-22 — `silent-write-noop` の3 linkを既存 `doc-hygiene-postconditions` へ統合。pattern 69→68、link 226、status遷移0を確認。同日3件目のためsuffix付き不変history reportを作成。
- 2026-09-22 — 全68 pattern・残候補を再照合。`archive-parser-shape-collision`、`formatter-mutation-bypass`、`permission-preflight` はresponse名の一致だけで失敗機構が異なるため、重複統合の対象外と判定。
- 2026-09-22 — `archive-parser-shape-collision`（2件、own-doc-update、parser回帰テスト案）を人間承認。実装・hook変更・status遷移はこのissueでは行わない。
- 2026-09-22 — `artifact-state-not-recorded`（2件、origin-doc-update、docs更新後readback案）を人間承認。実装・hook変更・status遷移はこのissueでは行わない。
- 2026-09-22 — `destructive-write-overwrite`（2件、settings-hook-owner、dirty-state警告とmanual review案）を人間承認。実装・hook変更・status遷移はこのissueでは行わない。
- 2026-09-22 — `formatter-boundary-gap`（2件、origin-doc-update、warn-only案）を人間承認。実装・hook変更・status遷移はこのissueでは行わない。
- 2026-09-22 — `formatter-mutation-bypass`（1件、origin-skill-commonize、hook/skill境界確認案）を人間承認。実装・hook変更・status遷移はこのissueでは行わない。
- 2026-09-22 — `git-history-reachability-gap`（1件、origin-git-cleanup、warn-only案）を人間承認。実装・hook変更・status遷移はこのissueでは行わない。
- 2026-09-22 — `permission-preflight`（2件、own-trouble-log、permission preflight案）を人間承認。実装・hook変更・status遷移はこのissueでは行わない。
- 2026-09-22 — `premature-completion-state`（1件、origin-doc-update、完了前readback案）を保留。追加の直接証拠なしに形式化へ進めない。
- 2026-09-22 — `safe-operation-alternative`（1件、own-trouble-log、warn-only案）を保留。追加の直接証拠なしに形式化へ進めない。
- 2026-09-22 — `security-guard-boundary`（1件、repository-security-hook、warn-only案）を保留。意図的テストと実運用を区別する証拠なしに形式化へ進めない。
- 2026-09-22 — `security-placeholder-exposure`（1件、repository-security-hook、security lint/skill review案）を人間承認。実装・hook変更・status遷移はこのissueでは行わない。
- 2026-09-22 — `sensitive-data-reuse`（2件、repository-security-hook、redaction guardとmanual review案）を人間承認。実装・hook変更・status遷移はこのissueでは行わない。
- 2026-09-22 — 候補の最終整理: 重複統合3件、承認済み14件、保留13件。保留13件は証拠件数増加まで次の承認質問を行わない。
- 2026-09-22 — `hook-signal-quality`（1件、settings-hook-owner、warn-only再評価案）を保留。真陽性・誤警告の測定なしに形式化へ進めない。
- 2026-09-22 — `patch-context-staleness`（1件、own-trouble-log、patch直前readback案）を保留。再現性の追加証拠なしに形式化へ進めない。
- 2026-09-22 — `command-shape-permission-friction`（1件、settings-hook-owner、warn-only案）を保留。許可境界の実測なしに形式化へ進めない。
- 2026-09-22 — `dependency-usage-not-checked`（2件、origin-doc-update、usage search手順案）を保留。利用実態の追加証拠なしに形式化へ進めない。
- 2026-09-22 — 低重要度・低件数の6候補（deploy-scope-mismatch、diagnostic-state-not-persisted、ephemeral-delivery-reference、guard-read-write-discrimination、operator-visible-completion-evidence、wrong-corpus-or-durable-location）を一括保留。件数増加まで承認質問を行わない。
- 2026-09-22 — 対象skill 44テスト、historical/status/docs validatorは合格。全skill lintはown-doc-updateのfixture commit拒否2件で停止し、hook回避なしでmerge保留。
- 2026-09-22 — own-session-closeでdocs hygiene/reviewを記録し、commit 2c0bfa4をpushしてPR #9を作成。human-gatedのためmergeとbranch削除は行わない。

## Completion

- [x] Implementation completed or intentionally not needed
- [x] Specs updated if direction or requirements changed
- [x] Guide impact classified before implementation
- [x] Guides updated in the same slice if implemented behavior changed
- [ ] Branch merged and cleaned up (or intentionally kept — note why)
- [x] 00_index.md updated
- [ ] Moved to docs/issues/archive/ when complete
