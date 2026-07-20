---
name: origin-goal
description: Run a long-lived, workstream-driven objective inside explicit human boundaries. Use when the user asks to keep working until a goal or checkpoint is reached, resume or execute a docs workstream, orchestrate subagents over multiple slices, control retries or metered cost, or prepare a durable autonomous run with human approval gates. Require a workstream and complete the preflight interview before starting or resuming execution.
---

# Origin Goal

Treat the workstream as durable truth, the runtime goal as the persistence
engine, and the plan as the current execution view. Keep the parent agent
responsible for authorization, integration, budgets, and completion.

## 1. Establish the target workstream

1. Read `docs/00_index.md` when present, then read only the relevant active
   workstream, linked specs, guides, and necessary code.
2. Make workstream selection the first human interaction. If exactly one active
   workstream matches, name it and ask whether to resume it. If several are
   plausible, ask which one to run and recommend the best match.
3. Do not use a standalone issue as the control unit for long-running work. Ask
   whether to create a workstream or attach the issue as a bounded slice; do not
   migrate history without approval.
4. If no matching workstream exists, say so and ask first whether to create one
   through `origin-doc-update`, proposing a workstream name and one-sentence scope.
   After approval, invoke `origin-doc-update`, follow its human-boundary interview, and
   create the workstream only after the complete boundary is confirmed. Do not
   ask downstream implementation questions before this create-or-select choice.
5. Reuse the workstream's recorded branch or worktree. Do not create a parallel
   branch for resumed work.

## 2. Complete preflight before execution

Ask one question at a time, in dependency order. Inspect the repository first
and do not ask for facts that are discoverable. Give a recommended answer with
each question. Reuse still-valid answers already recorded in the workstream and
ask only for missing, ambiguous, or changed boundaries.

Confirm and record all applicable fields:

1. Objective, observable success criteria, and out-of-scope work.
2. Actions allowed autonomously.
3. Actions requiring confirmation, including external writes, sends,
   submissions, deploys, destructive changes, data migrations, authentication,
   secrets, dependencies, network access, and metered services.
4. Resource ceilings appropriate to the task: money and currency, tokens,
   elapsed time, external calls, scan volume, iterations, retries, concurrency,
   or another measurable unit. Do not hard-code a monetary default. Record
   `none` when no metered work is possible.
5. Test commands, evidence, quality thresholds, and required reviewers.
6. The next human checkpoint and early stop conditions.
7. Subagent division, write ownership, model tier, and concurrency.

Treat creation and merge of a PR that passes its recorded test and CI gates as
autonomous integration by default. Record a PR-specific human gate only when
the user or repository policy explicitly requires one. This default does not
override branch protection, failed checks, unavailable authority, or a
repository-specific safety rule.

Before any metered batch, estimate:

`unit cost × execution count × scan volume × model/retry multiplier`

State the assumptions and reserve. If monetary cost cannot be observed or
reliably converted, ask for a measurable surrogate ceiling such as requests,
tokens, runtime, or attempts. Never claim that a currency ceiling is enforced
without usable price and usage telemetry. Stop before the approved ceiling
would be exceeded.

Do not start or resume the goal until the workstream records the confirmed
Authorization Envelope and Human Gates.

## 3. Start or resume the runtime goal

1. If a matching goal exists, inspect it and resume it without replacing its
   objective or budget.
2. Treat an explicit `$origin-goal` request to run the workstream as permission
   to create a matching runtime goal after preflight when the runtime provides a
   goal primitive.
3. Express one durable objective and one verifiable stopping condition. Point
   the goal at the workstream instead of duplicating its full issue queue.
4. Use the runtime's native goal lifecycle. Do not mark a goal complete until
   every required acceptance and guide contract is satisfied. Do not mark it
   blocked unless the runtime's blocked-state requirements are also satisfied.
5. If no goal primitive is available, continue safe vertical slices within the
   current session and leave a precise workstream handoff. Do not enable Goal,
   edit runtime configuration, or launch an external loop automatically.

Maintain a compact plan for the current slice and update it as work progresses.
The plan is not a second backlog and must not replace the workstream.

## 4. Route subagents deliberately

Delegate only concrete, bounded work that can proceed independently. Prefer
parallel subagents for read-heavy exploration, test execution, log analysis,
triage, and independent review. Avoid parallel writes to shared files.

Choose the cheapest capable available model and reasoning tier:

- Use a fast, lower-cost tier for searches, inventories, mechanical checks,
  summarization, and straightforward isolated edits.
- Use the strongest appropriate tier for ambiguous design, cross-cutting
  implementation, integration, security, and difficult failure analysis.
- Inherit the parent setup when no override has a clear benefit. Select only
  models exposed by the current runtime; do not invent or pin stale model names.

Give each subagent a single deliverable, exact scope, expected evidence, and
return format. Keep the parent agent responsible for the workstream, cost
ledger, shared-file writes, conflict resolution, and final verification. Use
separate worktrees or disjoint file ownership for parallel write tasks.

Include subagent tokens, calls, time, and retries in the shared resource
ceiling. Do not multiply agents merely because concurrency is available.

## 5. Execute vertical slices

For each queued issue:

1. Confirm dependencies, acceptance criteria, and `guide_impact` before edits.
2. Prefer red-green-refactor where practical.
3. Implement the smallest complete vertical slice.
4. Run the agreed tests and inspect evidence from the environment.
5. Update required guides in the same slice.
6. Update workstream status, material decisions, budget state, and Next Actions.
7. Continue automatically while inside the Authorization Envelope.

Update documentation at slice completion, a material decision, a changed
budget, a blocker, or a human gate. Do not append a diary entry for every tool
call, retry, or minor observation.

Report progress compactly: current slice, verified result, remaining work,
resource state when applicable, and blocker or next gate.

## 6. Control failures and stopping

Classify failures by root cause using the failing action, stable error signature,
and relevant environment state. Count the initial failure as attempt one.

Stop after three consecutive failed attempts with the same root cause. Record
the evidence, attempted remedies, last error, and the smallest human decision or
external change needed. Do not count successful progress iterations or distinct
root causes against this limit. Reset the counter only when evidence shows that
the root cause changed, not merely because the prompt or wording changed.

Stop immediately when:

- the next action requires confirmation;
- a resource ceiling would be exceeded;
- the next human gate is reached;
- acceptance requires unavailable evidence or authority;
- scope must materially expand; or
- a repository or runtime safety rule requires stopping.

Do not leave an approval or input request waiting silently. Persist the exact
question and current state in the workstream, then surface it to the user.

## 7. Finish or hand off

Before declaring completion, verify every acceptance criterion, test gate,
guide impact, resource constraint, and the runtime goal's stopping condition.
Update the workstream and guides through `origin-doc-update`. Archive the workstream
only when no required work remains and the recorded human gate permits closure.

When stopping short of completion, leave only the durable facts needed to
resume: completed slice, verification evidence, current failure count and root
cause, remaining budget, blocker or gate, and the next executable action.
