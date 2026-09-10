---
id: ADR-20260909-skill-usage-metadata
status: accepted
scope: development
created_at: 2026-09-09
updated_at: 2026-09-09
source_workstreams: []
source_issues: [ISSUE-20260909-measure-skill-usage]
related_specs: []
related_guides: [GUIDE-skill-usage-metrics]
supersedes: ""
superseded_by: ""
---

# Decision: Aggregate-only Codex/Claude skill usage metrics

## Context

Usage evidence is currently assembled by ad-hoc scans of Claude and Codex histories. The two
products store different JSONL shapes, and injected instructions can contain skill names that are
not invocations. A future collector must be repeatable across seats without copying prompts or
account state into a shared report.

## Decision

Add a read-only collector under `origin-skill-commonize/scripts/` with explicit Claude-project and
Codex-session roots. It normalizes only actual user messages, counts canonical skill tokens per
account and window, derives dependency-only status from canonical cross-skill references, and
labels no-evidence rows as `unknown`. JSON/JSONL output contains counts, coverage metadata, and
stable labels only; raw prompts, paths, session IDs, and credentials are never emitted. Writing a
snapshot is opt-in, and no scheduler or account-state symlink is installed by this change.

## Alternatives Considered

- `skill-doctor` only: useful corroboration, but account-local and not available for Codex.
- Parse every JSONL string: rejected because injected instructions and tool results create false
  skill hits.
- Persist raw prompts for later analysis: rejected because it expands secret and personal-data
  exposure.
- Run every skill as a smoke test: rejected because skills may deploy, mutate, or require external
  credentials; behavior tests remain targeted and human-approved.

## Consequences

### Positive

- Seat-separated, reproducible counts can be compared over time.
- Unknown coverage is fail-closed and cannot silently become a deletion candidate.
- The same aggregate schema works for Claude and Codex while preserving product-specific parsers.

### Negative or Follow-up

- Implicit routing that leaves no transcript marker is not counted; `skill-doctor` or future native
  telemetry remains supplemental.
- A recurring schedule and retention policy are intentionally left to a later human-approved
  operational change.
- Third-party skill quality is limited to provenance/parity checks; their bodies are not modified.

## Links

- Issue: `../issues/archive/ISSUE-20260909-measure-skill-usage.md`
- Guide: `../guides/GUIDE-skill-usage-metrics.md`
