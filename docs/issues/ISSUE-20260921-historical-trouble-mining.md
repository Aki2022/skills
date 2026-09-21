---
schema_version: 2
id: ISSUE-20260921-historical-trouble-mining
status: active
created_at: 2026-09-21
updated_at: 2026-09-21
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

as of 2026-09-21 — 226件の索引化・実データ適用・全検証が完了し、branch統合待ち。

## Next Actions

- `codex/historical-trouble-mining` の変更をreviewして統合する。
- 30件の新規形式化候補は、人間承認後に所有先ごとの別issueで実装する。

## Guide Impact

- Decision: none
- Target or reason: 現行挙動の正典である own-trouble-log/SKILL.md と references/status.md を同時更新するため、別guideは不要。

## Notes

- `legacy_unrecorded` は評価負債ではないが、分析入力として随時利用できる。
- `mechanism` / `thematic` の候補対応では status を更新しない。
- 新しいhook・skillは提案までとし、このissueでは実装しない。
- 実行結果: mechanism 203、thematic 23、direct 0、unclassified 0。
- response候補付きlinkは156件だが正式紐付けは0件で、status遷移も0件。
- `legacy_unrecorded` 226件と `evaluation_pending_entries` 30件の分離を維持した。

## Log

- 2026-09-21 — 約1.17MB・22,869行の226件を対象に確定。旧レポートからentry単位で復元できる既知クラスタは16件のみ。
- 2026-09-21 — historical_index.py の赤テストを確認後、snapshot / validate / summary / atomic apply を実装。専用テスト7件が合格。
- 2026-09-21 — 12件pilotが12/12件でschema合格。75・75・76件を3 agentで各1回digestし、再試行0回。
- 2026-09-21 — 生pattern 79件を重複統合して71件へ整理。226/226件、欠落0件で保管ルートへ適用。
- 2026-09-21 — entry本文226件の不変、historical index validate、status ledger validate、docs validator、全skill lintを確認。

## Completion

- [x] Implementation completed or intentionally not needed
- [x] Specs updated if direction or requirements changed
- [x] Guide impact classified before implementation
- [x] Guides updated in the same slice if implemented behavior changed
- [ ] Branch merged and cleaned up (or intentionally kept — note why)
- [x] 00_index.md updated
- [ ] Moved to docs/issues/archive/ when complete
