---
schema_version: 2
id: ISSUE-20261006-skip-workstreams-readme
status: active
workstream: none
priority: medium
due: none
created_at: 2026-10-06
updated_at: 2026-10-06
branch: "codex/fix-workstreams-readme-validator"
pr: ""
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

as of 2026-10-06 — Rebased on the latest main; all 350 own-doc-update tests pass. targetTest4 validation is clean, and the installed pre-commit rejects the README on the old validator ref but passes on this branch.

## Next Actions

1. Push this branch and open a PR; merge after checks pass, then clean the branch and archive this issue.

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

## Completion

- [ ] Implementation completed or intentionally not needed
- [ ] Specs updated if direction or requirements changed
- [ ] Guide impact classified before implementation
- [ ] Guides updated in the same slice if implemented behavior changed
- [ ] Branch merged and cleaned up (or intentionally kept — note why)
- [ ] 00_index.md updated
- [ ] Moved to docs/issues/archive/ when complete
