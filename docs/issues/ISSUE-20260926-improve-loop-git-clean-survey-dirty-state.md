---
schema_version: 2
id: ISSUE-20260926-improve-loop-git-clean-survey-dirty-state
status: in_progress
workstream: none
priority: high
due: none
created_at: 2026-09-26
updated_at: 2026-09-26
branch: codex/global-skill-commonization-20260926
pr: ""
related_specs: []
related_guides: []
guide_impact: none
guide_impact_reason: "own-git-clean/SKILL.md と scripts/survey.sh が現行手順と実装の正典で、docs/guides はこの repo では使用しない"
---

# own-git-clean survey の false-clean 判定を塞ぐ

## Goal

`own-git-clean/scripts/survey.sh` が dirty な checkout を `(clean)` と誤報したり、存在しない integration ref の diff エラーを空差分として squash-merged と誤認したりする経路を塞ぎ、各状態を回帰テストで固定する。

## Acceptance

- verify: machine — python3 -m unittest discover -s own-git-clean/scripts/tests -v が4件成功し、clean/dirty 状態と local main・origin/main・比較 ref 不在時の判定を検証する
  <!-- `machine — <command and expected result>`, or `human-review — <who reviews what>` -->

## Current Status

as of 2026-09-26 — false-clean and missing-ref conditions are fixed and four focused regression tests pass; the implementation and docs are committed on the commonization branch, pending PR integration.

## Next Actions

- Push and merge `codex/global-skill-commonization-20260926` with the related commonization changes, then archive this issue.

## Guide Impact

- Decision: none
- Target or reason: `own-git-clean/SKILL.md` and `scripts/survey.sh` are the current operational source; this repository does not use `docs/guides` for that procedure.

## Notes

- The pre-fix survey emitted `(clean)` while `git status --porcelain` returned 27 lines and its own later worktree detail said `dirty: yes`.
- A fixture with 1500 untracked paths failed before the fix and now passes; a clean fixture remains reported clean.
- With no local `main`, the survey previously ran `git diff main..HEAD` against a nonexistent ref, treated its empty error output as patch-equivalent, and reported unknown ahead/behind counts. The survey now compares with `origin/main`; if both refs are unavailable it reports UNKNOWN without a merge classification.

## Log

- 2026-09-26 — Replaced early-exit `grep -q` checks with complete status snapshots and added clean/dirty repository fixtures.
- 2026-09-26 — Resolved remote-tracking integration refs when the local branch is absent, guarded classifications when no ref exists, and added missing-ref regression fixtures.
- 2026-09-26 — Reconfirmed the issue's four-test acceptance result; branch integration remains.

## Completion

- [x] Implementation completed or intentionally not needed
- [x] Specs updated if direction or requirements changed
- [x] Guide impact classified before implementation
- [x] Guides updated in the same slice if implemented behavior changed
- [ ] Branch merged and cleaned up (or intentionally kept — note why)
- [x] 00_index.md updated
- [ ] Moved to docs/issues/archive/ when complete
