---
schema_version: 2
id: ISSUE-20260909-measure-skill-usage
status: archived
created_at: 2026-09-09
updated_at: 2026-09-09
branch: ISSUE-20260909-measure-skill-usage
pr: ""
related_specs: []
related_guides: [GUIDE-skill-usage-metrics]
guide_impact: required
guide_impact_reason: "The new collector and its privacy/coverage guarantees are operator-visible."
---

# Codex/Claude skill usage measurement

## Goal

Provide one deterministic, read-only collector that normalizes explicit skill invocations from
Claude Code and Codex seat histories into aggregate, secret-safe metrics. It must distinguish
observed use from unknown coverage and must never emit prompts, project paths, session IDs, or
account state.

## Acceptance

- verify: machine — python3 -m pytest -q origin-skill-commonize/scripts/tests/test_measure_skill_usage.py && bash origin-skill-commonize/scripts/skill_lint.sh; both exit 0
  <!-- `machine — <command and expected result>`, or `human-review — <who reviews what>` -->

## Current Status

Implemented `origin-skill-commonize/scripts/measure_skill_usage.py` and ten fixture/CLI tests.
The collector accepts explicit Claude/Codex roots, counts only identifiable user messages, keeps
ambiguous Codex messages fail-closed, derives dependency-only status from the canonical reference
parser, and writes no snapshot unless explicitly requested.

The four-account read-only run on 2026-09-09 found 14 used, 5 dependency-only, 50 unknown, and
zero deletion candidates across the 69 directories currently containing a canonical `SKILL.md`.
The extra directory is surfaced by the measurement scope; registry/topology reconciliation remains
a separate human-gated change.

## Next Actions

For recurring measurement, an operator may schedule the documented read-only command and choose a
retention policy. Scheduler installation, registry updates, and any skill retirement require a
separate human-approved change.

## Guide Impact

- Decision: required
- Target or reason: GUIDE-skill-usage-metrics

## Notes

- Human gate: no deletion, disabling, network access, or authentication changes are in scope.
- The default output is stdout; persistent snapshots are opt-in and contain aggregates only.
- `skill-doctor` remains an account-local corroborating source, not the canonical collector.
- ADR: `docs/adrs/ADR-20260909-skill-usage-metadata.md`.

## Completion

- [x] Implementation completed or intentionally not needed
- [x] Specs updated if direction or requirements changed
- [x] Guide impact classified before implementation
- [x] Guides updated in the same slice if implemented behavior changed
- [x] Branch merged and cleaned up (or intentionally kept — note why)
- [x] 00_index.md updated
- [x] Moved to docs/issues/archive/ when complete
