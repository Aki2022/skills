---
schema_version: 2
id: ISSUE-20261003-fix-origin-picker-metadata
status: archived
workstream: none
priority: low
due: none
created_at: 2026-10-03
updated_at: 2026-10-03
branch: ISSUE-20261003-fix-origin-picker-metadata
pr: ""
related_specs: []
related_guides: []
guide_impact: none
guide_impact_reason: "This changes skill picker labels and an internal metadata identifier only; skill instructions and operating procedures are unchanged, and docs/guides are not used for current skill behavior."
---

# Align own skill metadata with own names

## Goal

Remove stale `Origin` labels and identifiers from active own-skill metadata while preserving third-party mirror naming.

## Acceptance

- verify: machine — targeted metadata assertions and `git diff --check` exit 0; the required full `skill_lint.sh` run is recorded with unrelated S4/S10 findings, while its S8 suites pass.
  <!-- `machine — <command and expected result>`, or `human-review — <who reviews what>` -->

## Current Status

as of 2026-10-03 — own skill picker labels and the PPTX config identifier use the own identity; targeted assertions pass and unrelated full-lint findings are recorded.

## Next Actions

- None; acceptance is verified and the issue is archived.

<!-- First bullet = the very next command or step (or what unblocks a blocked issue). Required once work starts; the validator rejects an empty section. -->

## Guide Impact

- Decision: none
- Target or reason: picker labels and an internal config identifier changed; skill instructions and operating procedures are unchanged, and this repo records current skill behavior in each SKILL.md rather than docs/guides.

## Notes

- `quarto-authoring` remains unchanged: `mirrors.yaml` records it as the `posit-dev/skills` third-party mirror. Its directory is gitignored and has no HEAD entry, so byte comparison against this repository is not applicable; this task changed no file under that mirror.
- Full lint returned 2 with findings unrelated to this change: S4 reports a missing reference in the third-party `cloudflare` mirror, and S10 reports alias topology drift (`agents-sdk`, `cloudflare`, and dangling `own-cloudflare-route` entries across Codex/Gemini roots). Targeted metadata assertions and `git diff --check` passed; S8 pytest suites reported by lint passed.
- Branch `ISSUE-20261003-fix-origin-picker-metadata` was cut from main for this PR-based closeout, and the issue records it.

## Log

<!-- Append-only, dated: `- 2026-10-03 — what changed / what was learned`. The only place chronology belongs. -->
- 2026-10-03 — Changed the picker labels for `own-goal-run`, `own-spec-grill`, and `own-quarto-build` to match their canonical names; changed the PPTX config `_meta.name` from `origin-pptx-skill-config` to `own-pptx-skill-config`. Kept `quarto-authoring` under its upstream name. Targeted assertions passed; full lint findings are recorded in Notes.
- 2026-10-03 — Ran the mandated full skill lint (exit 2) and docs hygiene/validator; the remaining lint findings are unrelated to this change, and docs hygiene reported no fix or judgment candidates.
- 2026-10-03 — Cut the issue branch from main for PR-based closeout and recorded it here.

## Completion

- [x] Implementation completed or intentionally not needed
- [x] Specs updated if direction or requirements changed
- [x] Guide impact classified before implementation
- [x] Guides updated in the same slice if implemented behavior changed
- [x] Branch merged and cleaned up (or intentionally kept — note why)
- [x] 00_index.md updated
- [x] Moved to docs/issues/archive/ when complete
