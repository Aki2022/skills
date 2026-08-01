---
name: origin-ws-loop
description: >-
  Autonomously drain a repository's queue of active workstreams in one
  continuous run: pick the next workstream, execute it via origin-goal, close
  it via origin-close-session (merging green PRs autonomously under the
  default CD merge policy), park finished-but-unreviewed workstreams on a
  bounded review shelf, accumulate improvement observations as filed issues,
  and stop the entire loop the moment any human question arises. Use whenever the user
  wants accumulated workstreams processed in bulk or wants work to continue
  unattended until human input is needed, e.g. "wsを一気に消化", "溜まったws
  を処理して", "キューを回して", "自律で進められるところまで進めて", "run the
  ws queue", "drain the backlog", "process all workstreams", or /origin-ws-loop
  — typically right after mass-creating workstreams with origin-grill /
  origin-doc-update. Requires docs governance (docs/00_index.md +
  docs/workstreams/). Do NOT use for a single workstream (use origin-goal
  directly), for creating workstreams or specs (use origin-doc-update /
  origin-grill), or for recurring scheduled tasks unrelated to workstreams
  (use /loop alone).
---

# origin-ws-loop

Drain the workstream queue; stop at the first human question. This skill is an
orchestrator in the mold of `origin-close-session`: it owns no destructive
behavior of its own. Execution belongs to `origin-goal`, closeout to
`origin-close-session`, and workstream/issue creation to `origin-doc-update` —
each keeps its own safety contract, and this skill must not replicate or
shortcut their procedures. The only logic this skill owns is **selection**,
**observation**, and the **journal**.

One loop iteration = one workstream's run of consecutive ready issues. Safety
over throughput: efficiency comes from workstreams being created question-free
(runnability at issue granularity, a CD merge policy) and from shelving
finished-but-unreviewed work — never from skipping a question.

## Prerequisites

This skill consumes workstreams; it does not prepare them.

- No docs governance (`docs/00_index.md` + `docs/workstreams/` absent): refuse
  to run. Explain that the queue must be stocked first via `origin-grill`
  (specs) and `origin-doc-update` (workstreams), then stop. Do not initialize
  the scaffold or create workstreams on the user's behalf — those steps carry
  human decisions this skill must not absorb.
- Governance present but zero active workstreams: report "queue empty" as a
  normal, successful exit. Do not invent work.

## Preflight — once per session, interactive

Complete all five checks before the first iteration. The purpose is to move
every foreseeable human interaction to this single conversation so the run
itself never has to wait.

1. **Inventory.** Read `docs/00_index.md` and list active workstreams with
   their recorded runnability at **issue granularity** (`ready` / `gated`, from
   each workstream's Human Gates section; treat a missing runnability line as
   `gated`). Present the queue to the user. If the index's active list and the
   workstream directory disagree (e.g. completed workstreams still sitting in
   the active area), do not fix it inline — file one improvement issue for the
   drift; stale inventory poisons every later selection.
2. **Merge policy.** Confirm how PRs land. The default is continuous delivery:
   a PR whose recorded quality gates are green (CI when present, otherwise the
   workstream's recorded local gates) is merged autonomously by
   `origin-close-session` — this is what keeps the loop looping. Honor a
   workstream whose envelope explicitly marks merge as human-gated (it becomes
   a review gate, below), but point it out here so the user can lift it by
   editing the workstream if CD is what they actually want.
3. **Permission check.** Anticipate the commands the queue will need (tests,
   package scripts, gh, etc.) and check them against **both** permission lists
   in every settings layer that applies (user, project, and local — e.g.
   `~/.claude/settings.json`, `.claude/settings.json`,
   `.claude/settings.local.json`). The two lists fail in different ways and
   need different handling:
   - **Allowlist gaps.** An unattended run dies silently on a permission
     prompt, so fill gaps now: propose the additions, get the user's approval,
     and record them in settings before starting. Never loosen permissions
     autonomously.
   - **Deny collisions.** A command matched by a `deny` rule cannot be
     unblocked by adding an allow entry — deny is a deliberate guardrail, and
     relaxing it is a different, heavier human decision. Do not propose
     removing deny rules. Instead, downgrade the runnability of every issue
     whose core work needs that command to `gated`, record the matching deny
     rule as the reason, and raise it in this preflight conversation so the
     user can either drop the issue from the run or decide separately to
     change the guardrail. Catching this here costs one question; catching it
     mid-run costs a wasted iteration.
4. **Auto-promotion bounds.** Confirm the bounds for promoting improvement
   issues to workstreams without a question (see Improvement observations).
   Default — all conditions AND, expansion is human-only:
   1. it is a repo improvement (loop-process improvements are never promoted);
   2. it only touches docs, development scripts, tests, or in-repo tooling —
      nothing that alters production behavior, public APIs, package
      boundaries, or releases;
   3. no new dependencies, no external sends, no metered cost;
   4. its runnability is `ready`;
   5. at most 3 auto-promotions per session; overflow waits for triage.
5. **Stop conditions.** Defaults, adjustable here: a maximum of 10 iterations,
   and a review shelf cap of 3 workstreams (see Stop discipline). On the
   **first run in a repository**, recommend a pilot cap of 2 iterations
   instead — it proves workstream quality, permissions, and CD wiring cheaply
   before committing tokens to a full unattended run. Do not add loop-level
   time/token/cost ceilings — each workstream's Authorization Envelope
   already owns its resource limits, and a second guardian invites
   conflicting accounting.

## Each iteration — one workstream

1. **Select.** Runnability is judged at issue granularity: a workstream is
   startable when its **next pending issue** is `ready`; never execute a
   `gated` issue. A workstream that depends on a shelved workstream is not
   independent — skip it. If nothing startable remains outside the shelf,
   surface the nearest gate as the stopping question. Record a one-line
   selection reason in the journal. No human ordering approval is needed.
2. **Execute** via `origin-goal`, through the workstream's consecutive ready
   issues. The workstream's Authorization Envelope and Human Gates are already
   recorded, so its preflight must reuse them without re-interviewing. Its
   3-strike failure rule and stop discipline apply unchanged. Stop the
   workstream's run when its next issue is gated.
3. **Close** via `origin-close-session`. Under the CD merge policy (the
   default), obtain an independent review before the autonomous merge: a
   fresh-context reviewer (the `/code-review` skill or a reviewer subagent)
   that has not seen this iteration's reasoning — a reviewer inside the same
   context inherits the same blind spots, and this review is what justifies
   merging without a human. Fix confirmed findings within the iteration; a
   finding that needs a human decision is a question gate. With gates green
   and the review passed, the PR merges autonomously, so every finished
   workstream lands on a clean, merged main before the next begins. Under a
   human-gated merge policy, stop after the PR exists — that workstream has
   reached a review gate and goes to the shelf.
4. **Observe.** File improvement observations (below), bounded per iteration.
   Also close the lesson loop on failures: when a failure was diagnosed and
   fixed during this iteration, ask whether the lesson generalizes. If it
   does and encoding it is within the envelope (a guide line, a repo
   convention note, a config default), encode it in the same iteration;
   otherwise file it as an improvement issue. A failure fixed only in place
   will be re-fixed from scratch by a later iteration.
5. **Journal.** Append one entry: workstream id, outcome, PR link, selection
   reason, observation count, and turns/tokens when the runtime exposes them
   (`/goal`, `/usage`) — that is what lets the human tune iteration caps and
   model routing later.

## Stop discipline — two kinds of gate

Gates are not all alike, and the distinction decides whether the loop
survives:

- **Question gate** — a decision, missing input, or new permission is needed
  before work can proceed (an escalation, a 3-strike failure, a `gated` issue
  reached, an unforeseen permission prompt). Persist the exact question in the
  workstream first (`origin-goal` already does this), then stop the **entire
  loop**. Parking an unanswered decision and continuing would accumulate
  choices the human never sequenced.
- **Review gate** — the workstream's work is finished and pushed as a PR;
  only human review/merge blocks further progress in _that_ workstream (a
  human-gated merge policy, a recorded slice-review gate). Move the
  workstream to the **review shelf** (record the pending PR in the journal)
  and continue with the next independent ready workstream. This is safe
  precisely because nothing unanswered accumulates — only finished,
  independent work awaiting review.

The shelf is bounded: at most 3 shelved workstreams (preflight-adjustable).
Reaching the cap stops the loop — the human's review budget, not the queue,
is the scarce resource, and one batched review beats ten interruptions.

Also stop the entire loop when the queue is exhausted (no startable
workstream outside the shelf) or the iteration limit is reached. On every
stop, produce the final digest — including the shelf, so pending reviews are
presented in one batch.

## Improvement observations

While executing, notice friction: stuck points, repeated manual steps, the
same helper being rebuilt across workstreams. Record these as standalone
issues in `docs/issues/` (via `origin-doc-update`'s improvement-issue
convention) instead of interrupting the run. **File every observation as an
issue before the digest is produced** — filing is a docs-only autonomous
action; an observation that exists only in the digest text is lost the moment
the chat scrolls away, which is exactly the silent loss this mechanism
exists to prevent. Only _promotion_ to a workstream is bounded, never the
filing itself.

- **Two scopes, named apart.** `ISSUE-YYYYMMDD-improve-<slug>` targets the
  repository itself; `ISSUE-YYYYMMDD-improve-loop-<slug>` targets the loop
  machinery (this skill and the skills it composes, canonical in
  `~/.agents/skills/`). The scopes have different owners and approval paths,
  so the name must reveal the scope at a glance.
- **Bounded.** Deduplicate mechanically against existing issue slugs/titles
  (no LLM judgment), and file at most 2 observations per iteration. An
  unbounded observer drowns the triage queue and its own signal.
- **Auto-promotion.** A repo improvement meeting all preflight bounds may be
  promoted to a workstream via `origin-doc-update` without a question — cite
  the preflight approval as its confirmed boundary — and joins the tail of
  the queue. Everything else (including every loop-process improvement) stays
  an issue for human triage; loop-process changes are applied later by a
  human through `origin-skill-commonize`.

## Scheduled runs

Pairing this skill with `/schedule` (e.g. a nightly run) turns it into a
proactive loop in the official loop taxonomy. No design change is needed —
the question-gate full stop already makes unattended runs fail safe, and the
review shelf becomes the morning inbox: the human wakes to a batch of
finished PRs and filed observations instead of a stalled session. Match the
schedule to how fast the queue actually refills — a loop that wakes hourly
against a queue that refills weekly burns tokens on empty preflights.

## Journal and final digest

Keep the journal as an append-only file in the session scratchpad — one entry
per iteration, written as it happens. Never reconstruct the run from the
transcript afterwards; incremental appending is what keeps reporting cheap.
The repository keeps no run log: durable facts live in the workstreams, PRs,
and issues the run already produced.

On any stop, report in chat with this fixed shape — pointers plus one-line
outcomes, never content copies:

```
TLDR: <n> ws done, <s> on review shelf, <stopped why>, <k> improvements filed (<m> auto-promoted)
| ws | outcome | PR |
レビュー待ち: | ws | PR | 何を承認すると何が進むか |   ← shelf; omit when empty
停止理由: <question + where it is persisted, or "queue empty" / "shelf cap" / "iteration limit">
improvement: <filed issue ids, triage-pending marked>
次のアクション: <single next step, e.g. "review the shelf PRs, then rerun /origin-ws-loop">
```

## Token efficiency

- Per iteration read only `docs/00_index.md` and the selected workstream (plus
  what `origin-goal` itself requires). Never rescan all of `docs/`.
- The digest points; it does not duplicate. Workstream/PR/issue bodies are the
  record.
- Route read-heavy inventory or search subtasks to cheaper model tiers, as
  `origin-goal` already prescribes for subagents.

## Never during this skill

- initialize a docs scaffold, create a spec, or interview for a new
  workstream boundary (except citing the preflight approval for an in-bounds
  auto-promotion);
- execute a `gated` issue or answer a gate question by guessing;
- continue any iteration after a question gate has arisen (a review gate
  shelves the workstream; it never excuses continuing _inside_ it);
- put anything with an unanswered decision on the review shelf — the shelf is
  for finished work awaiting review only;
- leave an improvement observation unfiled (digest text is not a record);
- edit skills or anything under `~/.agents` (loop-process improvements are
  recorded, not applied);
- loosen permissions, bypass hooks, or take any action the composed skills'
  own safety contracts would forbid.
