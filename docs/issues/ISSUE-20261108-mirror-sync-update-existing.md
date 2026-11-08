---
schema_version: 2
id: ISSUE-20261108-mirror-sync-update-existing
status: in_progress
workstream: WS-20261108-third-party-skill-mirror-sync
created_at: 2026-11-08
updated_at: 2026-11-08
branch: "WS-20261108-third-party-skill-mirror-sync"
pr: ""
related_specs: [docs/specs/third-party-skill-mirror-update.md]
related_guides: []
guide_impact: none
guide_impact_reason: "skill の挙動は own-skill-commonize/SKILL.md が正典で、同じ slice で更新する（docs/guides は使わない）"
---

# plugin の版から正典へ置き換える同期（--update-existing の移植と修正）

## Goal

plugin の導入版で、正典の現役 mirror と台帳を置き換える同期と、ずれを知らせる検査を作る。

## Acceptance

- verify: machine — mise run verify が exit 0。新しいテストで、指定しない注記・local_copy・版が台帳から消えないこと、途中失敗で正典と台帳が変わらないことを確かめ、赤から緑を一度見る
  <!-- `machine — <command and expected result>`, or `human-review — <who reviews what>` -->

## Current Status

as of 2026-11-08 — `--update-existing`（dry-run と `--apply`、完全 SHA 必須、欄の保持、失敗時の復元、`--plugin`）と `check_mirrors.py` の drift・件数表示を実装。テストは赤を見てから緑（own-skill-commonize 120 passed）、作業場所に向けた skill_lint は exit 0。

## Next Actions

1. PR を merge し、Completion の残りを埋めて status: complete と archive_issue.py で閉じる
2. その後 ISSUE-20261108-retire-wip-mirror-branch へ進む

<!-- First bullet = the very next command or step (or what unblocks a blocked issue). Required once work starts; the validator rejects an empty section. -->

## Guide Impact

- Decision: none
- Target or reason: skill の挙動は own-skill-commonize/SKILL.md が正典で、同じ slice で更新した

## Notes

## Log

- 2026-11-08 — wip の update 節を土台に main から作り直した。wip にあった欠陥（指定しない注記・local_copy の消失、版が unknown に戻る、HEAD を取る）は、欄の保持と完全 SHA 必須で塞いだ。

## Completion

- [x] Implementation completed or intentionally not needed
- [x] Specs updated if direction or requirements changed
- [x] Guide impact classified before implementation
- [x] Guides updated in the same slice if implemented behavior changed
- [ ] Branch merged and cleaned up (or intentionally kept — note why)
- [ ] 00_index.md updated
- [ ] Moved to docs/issues/archive/ when complete
