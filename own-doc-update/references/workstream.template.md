---
schema_version: 2
id: WS-YYYYMMDD-short-slug
status: active
created_at: YYYY-MM-DD
updated_at: YYYY-MM-DD
branch: WS-YYYYMMDD-short-slug
pr: ""
human_boundary_confirmed_at: YYYY-MM-DD
next_human_gate: short-gate-name
related_specs: []
related_guides: []
---

# Title

## Goal

## Success Criteria

## Authorization Envelope

- Approved scope:
- Autonomous actions allowed:
- Confirm first:
- Merge policy: CD (default) — a PR whose recorded quality gates are green merges autonomously
- Cost or usage ceiling:
- Out of scope:

## Human Gates

- Start gate: confirmed on YYYY-MM-DD
- Next gate: short-gate-name
  <!-- a named human checkpoint, or `none-workstream-complete` when nothing gates completion -->
- Stop conditions:

## Issue Queue

| Issue               | Status  | Depends on | Outcome            |
| ------------------- | ------- | ---------- | ------------------ |
| ISSUE-01-short-slug | pending | none       | One vertical slice |

### ISSUE-01-short-slug

- status: pending
- depends_on: []
- runnability:
  <!-- `ready`, or `gated on <the human decision / missing input>` — executors treat a missing value as gated and stop -->
- guide_impact: required
- related_guides: [GUIDE-short-slug]
- guide_impact_reason: ""

#### Goal

#### Acceptance

- verify:
  <!-- `machine — <command and expected result>`, or `human-review — <who reviews what>` -->

#### Current Status

#### Next Actions

## Decisions

Record only decisions that are difficult to reverse or surprising without context.

## Completion

- [ ] Every issue meets its acceptance criteria
- [ ] Every issue records guide impact as required or none
- [ ] Required guides describe current implemented behavior
- [ ] Specs reflect any changed direction or requirements
- [ ] Next human gate reached or the workstream intentionally stopped
- [ ] 00_index.md updated
- [ ] Workstream archived when complete
