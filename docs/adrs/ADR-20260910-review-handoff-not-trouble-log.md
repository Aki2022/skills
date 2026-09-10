---
id: ADR-20260910-review-handoff-not-trouble-log
status: accepted
scope: development
created_at: 2026-09-10
updated_at: 2026-09-10
source_workstreams: []
source_issues: [ISSUE-20260910-improve-loop-review-handoff]
related_specs: []
related_guides: [GUIDE-skill-usage-metrics]
supersedes: ""
superseded_by: ""
---

# Decision: Keep human review handoff outside trouble logs

## Context

The skill-usage scheduler has four human decisions that must remain visible before a commit or
push. Placing those reminders in `origin-trouble-log` would mix a standing procedure with
append-only incident evidence. That skill also intentionally has no automatic reminder or scheduler.

## Decision

Keep the current review checklist in `GUIDE-skill-usage-metrics.md`, where operator-visible
behavior and approval boundaries already live. Add a generic closeout pointer in
`origin-close-session` so the related guide's `承認前レビュー` section is read before Git
integration. Leave `origin-trouble-log` unchanged; record there only an actual omission, false
completion, or other observed friction.

## Alternatives Considered

- Add the checklist to `origin-trouble-log`: rejected because evidence and standing policy would
  be mixed, and the skill's no-automatic-trigger contract would not prevent forgetting.
- Create a second permanent review ledger: rejected because it duplicates the guide and creates a
  new synchronization surface.
- Rely only on the archived issue: rejected because archived work units are historical context,
  not the current operator procedure.

## Consequences

### Positive

The guide is the single current location for scheduler-specific review decisions, and closeout
has a predictable place to surface them. Actual review omissions remain measurable in the trouble
log without changing its schema or triage semantics.

### Negative or Follow-up

A future checklist for another feature belongs in that feature's guide and is reached through the
same closeout pointer.

## Links

- Issue: `../issues/archive/ISSUE-20260910-improve-loop-review-handoff.md`
- Guide: `../guides/GUIDE-skill-usage-metrics.md`
