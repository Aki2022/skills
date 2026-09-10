---
schema_version: 2
id: ISSUE-20260909-skill-usage-schedule
status: archived
created_at: 2026-09-09
updated_at: 2026-09-09
branch: ISSUE-20260909-skill-usage-schedule
pr: ""
related_specs: []
related_guides: [GUIDE-skill-usage-metrics]
guide_impact: required
guide_impact_reason: "The schedule, storage location, retention, and disable procedure are operator-visible."
---

# Schedule aggregate Codex/Claude skill usage snapshots

## Goal

Install one per-user launchd job that runs the existing secret-safe collector once per day for
Claude/Codex seat1/seat2. Keep only aggregate JSONL snapshots for 180 days, never copy raw
transcripts, and make the job safe to disable or rerun without touching aliases or account state.

## Acceptance

- verify: machine — python3 -m pytest -q origin-skill-commonize/scripts/tests/test_skill_usage_schedule.py && python3 origin-skill-commonize/scripts/install_skill_usage_launchd.py --check; both exit 0
  <!-- `machine — <command and expected result>`, or `human-review — <who reviews what>` -->

## Current Status

The canonical runner, LaunchAgent template, installer, and schedule tests are implemented. The
user LaunchAgent was installed and bootstrapped with the authorized label at 03:30 local time;
one bounded kickstart completed with exit code 0. The resulting local state contains one aggregate
snapshot for all four explicit account roots, with owner-only permissions and all privacy flags false.
The snapshot summary is 69 canonical skills: 14 used, 5 dependency-only, 50 unknown, and 0 candidates.

## Next Actions

Continue observing the daily snapshot cadence. Revisit the 30-day window, 180-day retention, or
03:30 schedule only as a separately reviewed operational change. Disable with the documented
`launchctl bootout` procedure if collection is no longer wanted.

## Guide Impact

- Decision: required
- Target or reason: `GUIDE-skill-usage-metrics` documents the launchd cadence, state paths, retention, and disable procedure.

## Notes

- No scheduler existed with this label at the start of the change.
- Snapshot data is aggregate-only; the runner refuses malformed snapshot lines during pruning.
- The install action is limited to `~/Library/LaunchAgents` and the runner's local state directory.
- Registry, alias topology, skill content, authentication, and transcript files are out of scope.
- ADR: `docs/adrs/ADR-20260909-skill-usage-schedule.md`.

## Completion

- [x] Implementation completed or intentionally not needed
- [x] Specs updated if direction or requirements changed
- [x] Guide impact classified before implementation
- [x] Guides updated in the same slice if implemented behavior changed
- [x] Branch merged and cleaned up (or intentionally kept — note why)
- [x] 00_index.md updated
- [x] Moved to docs/issues/archive/ when complete
