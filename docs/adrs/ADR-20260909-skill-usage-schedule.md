---
id: ADR-20260909-skill-usage-schedule
status: accepted
scope: development
created_at: 2026-09-09
updated_at: 2026-09-09
source_workstreams: []
source_issues: [ISSUE-20260909-skill-usage-schedule]
related_specs: []
related_guides: [GUIDE-skill-usage-metrics]
supersedes: ""
superseded_by: ""
---

# Decision: Run daily aggregate skill usage collection with bounded retention

## Context

Manual collection gives no comparable time series, while copying transcripts into a shared ledger
would expose prompts and account state. The collector already emits aggregate-only reports for four
explicit roots. A recurring job needs a stable cadence, a bounded retention policy, and a lock so a
wake-from-sleep retry or manual run cannot interleave writes.

## Decision

Install a per-user macOS LaunchAgent with label `com.origin.skill-usage-metrics`, scheduled daily at
03:30 local time and not run at load. It invokes the canonical runner, which scans the four explicit
Claude/Codex roots over a 30-day window, appends one aggregate JSON object to a local JSONL snapshot,
and atomically prunes entries older than 180 days. The runner uses an exclusive local lock and
refuses to rewrite malformed or non-aggregate lines. No scheduler is shared through skill aliases,
and no raw transcript, credential, prompt, path, or session ID is copied.

## Alternatives Considered

- Alternative: run on every invocation — rejected because it adds latency and repeated scans to
  interactive work.
- Alternative: keep an unbounded history — rejected because retention and personal-data exposure
  would grow without an operational bound.
- Alternative: use a shared cron/skill alias — rejected because launchd is the native per-user
  scheduler and account state must not be symlinked.

## Consequences

### Positive

- Daily snapshots support trend comparisons while preserving the fail-closed `unknown` state.
- A lock and atomic replacement make retries recoverable without partial JSONL writes.
- The state directory and files are owner-only, and malformed/non-aggregate data fails closed
  instead of being carried into a later snapshot.

### Negative or Follow-up

- The chosen cadence and 180-day retention are defaults that can be changed in a later reviewed
  operational change.
- Launchd availability is macOS-specific; manual invocation remains the fallback on other hosts.
- Skill token usage is explicit evidence only; implicit routing and token/context cost remain outside
  this collector.

## Links

- Issue: `../issues/archive/ISSUE-20260909-skill-usage-schedule.md`
- Guide: `../guides/GUIDE-skill-usage-metrics.md`
