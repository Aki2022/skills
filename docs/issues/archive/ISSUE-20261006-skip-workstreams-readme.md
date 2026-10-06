---
schema_version: 2
id: ISSUE-20261006-skip-workstreams-readme
status: archived
workstream: none
priority: medium
due: none
created_at: 2026-10-06
updated_at: 2026-10-06
branch: "codex/fix-workstreams-readme-validator"
pr: "#32"
related_specs: []
related_guides: []
guide_impact: none
guide_impact_reason: "Internal validator maintenance; no user-facing guide behavior changes."
---

# Exclude the workstreams README from lifecycle validation

## Goal

`docs/workstreams/README.md` explains the directory and must not be treated as a managed
workstream by validation, ownership routing, or lifecycle hygiene scans.

## Acceptance

- verify: machine — python3 -m pytest own-doc-update/scripts/tests -q — all tests pass
- verify: machine — python3 own-doc-update/scripts/validate_repo_docs.py <targetTest4> — exit 0 and no diagnostics for docs/workstreams/README.md
- verify: machine — `VIBE_GUARD_DOCS_VALIDATOR_REF=codex/fix-workstreams-readme-validator git -C <fixture> commit -m update` — exit 0 with a staged README edit
  <!-- `machine — <command and expected result>`, or `human-review — <who reviews what>` -->

## Current Status

as of 2026-10-06 — PR #32 landed as 799a33e; all acceptance checks pass, the source branch/worktree are cleaned, and this issue is archived.

## Next Actions

- None; this issue is archived.

<!-- First bullet = the very next command or step (or what unblocks a blocked issue). Required once work starts; the validator rejects an empty section. -->

## Guide Impact

- Decision: none
- Target or reason:

## Notes

## Log

<!-- Append-only, dated: `- 2026-10-06 — what changed / what was learned`. The only place chronology belongs. -->

- 2026-10-06 — Added regression tests for validation, ownership routing, and lifecycle hygiene discovery; all three failed before the fix and pass after README is skipped.
- 2026-10-06 — The complete own-doc-update test suite passes (350 tests). Shared skill lint was run; environment-level failures are confined to an unrelated Python 3.9 incompatibility in own-pptx-build and pre-existing Cloudflare references/symlink drift.
- 2026-10-06 — Ran the validator against targetTest4; it exited 0 and emitted no diagnostics for docs/workstreams/README.md.
- 2026-10-06 — The installed pre-commit was tested with a staged README edit: the committed baseline validator rejected it, while this branch's validator passed it.
- 2026-10-06 — PR #32 merged as 799a33e. The closeout hygiene report found no repair or judgment candidates.
- 2026-10-06 — Removed the merged source branch/worktree and the temporary baseline worktree; all completion checklist items are now met.
- 2026-10-06 — Archived the completed issue and removed its active index route.

## Completion

- [x] Implementation completed or intentionally not needed
- [x] Specs updated if direction or requirements changed
- [x] Guide impact classified before implementation
- [x] Guides updated in the same slice if implemented behavior changed
- [x] Branch merged and cleaned up (or intentionally kept — note why)
- [x] 00_index.md updated
- [x] Moved to docs/issues/archive/ when complete
