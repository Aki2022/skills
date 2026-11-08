---
schema_version: 2
id: ISSUE-20261108-mirror-sync-update-existing
status: active
workstream: WS-20261108-third-party-skill-mirror-sync
created_at: 2026-11-08
updated_at: 2026-11-08
branch: ""
pr: ""
related_specs: []
related_guides: []
guide_impact: none
guide_impact_reason: "skill の挙動は own-skill-commonize/SKILL.md が正典で、同じ slice で更新する（docs/guides は使わない）"
---

# plugin の版から正典へ置き換える同期（--update-existing の移植と修正）

## Goal

## Acceptance

- verify: machine — mise run verify が exit 0。新しいテストで、指定しない注記・local_copy・版が台帳から消えないこと、途中失敗で正典と台帳が変わらないことを確かめ、赤から緑を一度見る
  <!-- `machine — <command and expected result>`, or `human-review — <who reviews what>` -->

## Current Status

<!-- Snapshot for the next session, overwritten each time: first line `as of 2026-11-08 — <state in one sentence>`. History goes to ## Log, not here. -->

## Next Actions

1. spec が active になったら、main から枝を切り、wip の mirror_skill.py の update 節とテストを選択的に移す

<!-- First bullet = the very next command or step (or what unblocks a blocked issue). Required once work starts; the validator rejects an empty section. -->

## Guide Impact

- Decision: none
- Target or reason:

## Notes

## Log

<!-- Append-only, dated: `- 2026-11-08 — what changed / what was learned`. The only place chronology belongs. -->

## Completion

- [ ] Implementation completed or intentionally not needed
- [ ] Specs updated if direction or requirements changed
- [ ] Guide impact classified before implementation
- [ ] Guides updated in the same slice if implemented behavior changed
- [ ] Branch merged and cleaned up (or intentionally kept — note why)
- [ ] 00_index.md updated
- [ ] Moved to docs/issues/archive/ when complete
