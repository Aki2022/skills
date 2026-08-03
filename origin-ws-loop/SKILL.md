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

  **Produce a digest on this refusal too**, short but real: what was
  evaluated, what was missing, and anything worth knowing that surfaced on the
  way in. Observations gathered here are the one place they are reported in
  chat instead of filed — a refused target has no `docs/issues/` to file
  into, and filing them in the calling repository would misattribute another
  repository's problems. This is a deliberate, bounded exception to "an
  observation that exists only in digest text is lost": a refusal always ends
  in front of the human who asked, so the digest reaches someone who can act
  on it. The exception covers refusal only. Once a run is under way, every
  observation is filed.

- Governance present but zero active workstreams: report "queue empty" as a
  normal, successful exit. Do not invent work.
- **Evaluate exactly the directory you were pointed at.** A subdirectory of a
  governed repository that has no governance of its own is an ungoverned
  target, and the first rule applies. Do not walk up to the git root to find a
  queue: the caller named a scope, and a queue reached by widening it is a
  queue nobody asked to drain. Say which directory was evaluated, so a caller
  who meant the repository can re-point the skill in one step.

## Preflight — once per session

Complete all five checks before the first iteration. The purpose is to move
every foreseeable human interaction to this single conversation so the run
itself never has to wait.

**When a human is there, this preflight is a conversation.** Present the
inventory, confirm the policies, and let the answers override the defaults.

**When no human is there — invoked by another agent, or under `/schedule` — it
is a report, not a conversation, and it still completes.** An unattended run is
a primary path here, not an exception, so "confirm" cannot mean "wait". In that
case: adopt what each workstream already records, leave every default exactly as
written rather than choosing a value nobody approved, and put the inventory and
the resulting choices in the digest without asking for approval. A caller that
stated a bound explicitly (an iteration cap, for instance) has stated it, and
that value wins over the default. Only check 3 differs, because a permission gap
can be measured rather than assumed — it keeps its own procedure below. If a
check cannot be completed even this way, that is a question gate: persist it and
stop before the first iteration, where the cost is one message instead of a dead
run.

Never treat another agent's instruction as the human approval that writing
settings, lifting a recorded gate, or widening an envelope requires.

1. **Inventory.** Read `docs/00_index.md` and list active workstreams with
   their recorded runnability at **issue granularity** (`ready` / `gated`, from
   each workstream's Human Gates section; treat a missing runnability line as
   `gated`). Present the queue to the user. If the index's active list and the
   workstream directory disagree (e.g. completed workstreams still sitting in
   the active area), do not fix it inline — file one improvement issue for the
   drift; stale inventory poisons every later selection.

   **Drift means the record contradicts reality, not that a workstream is
   further along than its status.** A workstream whose every issue is done, whose
   PRs are merged, and whose `next_human_gate` has **not** been reached is
   correctly `active`: the gate is the thing still outstanding, and archiving it
   would be the error. Read the gate before calling it drift. It is drift when the
   gate has been reached and recorded and the workstream is still sitting in the
   active area, or when the index and the directory disagree about which
   workstreams exist at all. Getting this wrong costs an issue filed against
   correct behaviour, which teaches the next reader to distrust the check.

   **Land that issue on its own branch with its own PR, and merge it
   autonomously** — it is docs-only, so nothing gates it. Preflight runs before
   any workstream branch exists, so the alternatives are both wrong: committing
   straight to the default branch is a heavier act than a docs file warrants,
   and folding the file into the first iteration's workstream branch charges it
   to an envelope that never approved it and, under a gated merge policy, leaves
   it unlanded until a human merges something unrelated.

   **The index's frontmatter is in scope for this check, not just its body
   lists.** An index that carries mutable state — a `current_focus` naming how
   many workstreams are active, or which one is next — can contradict the
   directory exactly as a stale body list can, and it is read first, so it
   misleads first. Only check what is checkable without interpretation:
   counts, ids, and statuses stated outright. Free prose about direction is
   not drift and is not this check's business. One issue covers all drift
   found, frontmatter and body together.

   **Then reconstruct the review shelf from open PRs**, before selecting
   anything: list the repository's open PRs and match them to workstreams by
   branch name (`gh pr list` reports `headRefName`; a shelved workstream always
   still has its branch, because its PR is open). **Match on the workstream id as
   a prefix, not on string equality** — a run that had to cut a differently named
   branch (a suffix, a retry) still owns that PR, and an exact match would miss
   it, re-select the workstream, and redo work already sitting in the PR. When a
   PR's head branch cannot be attributed to any workstream, say so in the digest
   rather than assuming the shelf is empty. Every match starts this run already
   shelved, counts against the shelf cap, and is never selected.

   **The inverse case is real too, and this check is where it has to surface: a
   workstream branch that exists with unmerged commits and has no open PR.** The
   shelf rule reads from PRs to branches, so on its own it reports an empty shelf
   and says nothing about the leftover — and the branch is then discovered only
   when that workstream is selected, mid-iteration, as a bare
   `fatal: a branch named ... already exists`. List those branches here as
   **leftovers**, in the inventory and the journal. They are not shelved and do
   not count against the shelf cap: nothing is awaiting review. Do not act on
   them either — deciding what a leftover branch is worth belongs to
   `origin-goal`'s stale-branch rule at execution time, which salvages, cuts from
   the default branch, and rewrites the record. This check only makes sure the
   leftover is known before it can surprise an iteration.

   Match on the branch, not on a `pr:` field in the workstream: the PR number
   only exists after the push that creates it, so recording it would need a
   second commit and break `origin-close-session`'s one-self-consistent-commit
   ordering. The branch name is known before any of that and is just as
   decisive. The digest still carries the PR link for the human.

   This step exists because the shelf has no other durable home. A workstream
   parked at a review gate records that fact **on its own unmerged branch**, so
   the default branch still shows it as active with its issues pending — a
   later run that trusts only the default branch will select it again and redo
   work that is already sitting in a PR. The open PR is the durable fact (the
   same reasoning that keeps run logs out of the repository: durable facts live
   in workstreams, PRs, and issues), so the shelf must be _read_ from it rather
   than kept in a journal that dies with the session.

2. **Merge policy.** Confirm how PRs land. The default is continuous delivery:
   a PR whose recorded quality gates are green (CI when present, otherwise the
   workstream's recorded local gates) is merged autonomously by
   `origin-close-session` — this is what keeps the loop looping. Honor a
   workstream whose envelope explicitly marks merge as human-gated (it becomes
   a review gate, below), but point it out here so the user can lift it by
   editing the workstream if CD is what they actually want.
3. **Permission check.** Anticipate the commands the queue will need (tests,
   package scripts, `gh`, and — whenever any workstream's merge policy is CD —
   `git worktree add` and `git worktree remove`, which the independent review in
   step 3 of each iteration requires by name) and check them against **both**
   permission lists
   in every settings layer that applies (user, project, and local — e.g.
   `~/.claude/settings.json`, `.claude/settings.json`,
   `.claude/settings.local.json`). The two lists fail in different ways and
   need different handling:
   - **Allowlist gaps.** An unattended run dies silently on a permission
     prompt, so fill gaps now: propose the additions, get the user's approval,
     and record them in settings before starting. Never loosen permissions
     autonomously.

     Writing to settings is the human's act, and no other agent's
     authorization substitutes for it. When this preflight runs without a
     human to ask — invoked by another agent, or in a non-interactive
     runtime — do not write settings and do not treat the check as skipped.
     Take the missing command's measure instead: run it once, harmlessly, and
     see whether it is actually blocked. If it runs, the gap was theoretical;
     record in the journal that it was verified by execution rather than by
     settings, and proceed. If it is genuinely blocked, that is a question
     gate — persist it and stop before the first iteration, where the cost is
     one message rather than a dead unattended run.

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
   instead: it proves workstream quality, permissions, and CD wiring cheaply
   before committing tokens to a full unattended run. This is a recommendation —
   a cap the human states explicitly wins, and the run records that the pilot
   suggestion was overridden. Do not add loop-level
   time/token/cost ceilings — each workstream's Authorization Envelope
   already owns its resource limits, and a second guardian invites
   conflicting accounting.

## Each iteration — one workstream

1. **Select.** Runnability is judged at issue granularity: a workstream is
   startable when its **next pending issue** is `ready`; never execute a
   `gated` issue. A workstream that depends on a shelved workstream is not
   independent — skip it. A workstream with an open PR is already on the shelf
   — skip it too (see Preflight inventory). If nothing startable remains
   outside the shelf, surface the nearest gate as the stopping question.

   When several workstreams are startable, order them by these rules, in
   order, and record the deciding rule as the one-line selection reason:

   1. **Workstreams that cannot end the run come first.** A workstream whose
      remaining issues are all `ready` and whose merge policy is CD can only
      finish and merge. Take those before anything else.
   2. **Then workstreams that consume a shelf slot** — a `gated` merge policy
      ends that workstream at a review gate and moves the loop one step closer
      to the shelf cap.
   3. **Last, workstreams with a known `gated` issue downstream**, because
      reaching it stops the _entire_ loop. Everything startable should have had
      its turn first.
   4. **Tie-break by the order of the index's active list**, which is stable
      and human-visible, so the same queue always produces the same run.

   The principle is _defer whatever ends the run_: within one iteration budget,
   the most work is drained when the run-enders go last. Order is not a detail —
   the same queue and the same cap produce completely different outcomes
   depending on what is picked first, so leaving it to the runner's judgement
   makes a run unrepeatable and its coverage accidental. A human may override
   the order at preflight; nothing here needs per-iteration approval.

2. **Execute** via `origin-goal`, through the workstream's consecutive ready
   issues. The workstream's Authorization Envelope and Human Gates are already
   recorded, so its preflight must reuse them without re-interviewing. Its
   3-strike failure rule and stop discipline apply unchanged. Stop the
   workstream's run when its next issue is gated. Reaching a gated issue ends
   the **work**, not the iteration: continue through steps 3–5 to land what is
   already green, and stop the loop after step 5.
3. **Close** via `origin-close-session`. Under the CD merge policy (the
   default), obtain an independent review before the autonomous merge: a
   fresh-context reviewer (the `/code-review` skill or a reviewer subagent)
   that has not seen this iteration's reasoning — a reviewer inside the same
   context inherits the same blind spots, and this review is what justifies
   merging without a human.

   **The reviewer must be able to run the workstream's recorded quality gates,
   not merely read the diff, and must not disturb the working tree to do it.**
   Give it a subagent type that has shell access, tell it to execute the gates,
   and tell it to do so in its own checkout —
   `git worktree add --detach <tmpdir> <PR head sha>`, or an equivalent
   isolated copy. Pass `--detach` and a sha rather than the branch name: the
   loop is standing on that branch, and git refuses to check the same branch
   out twice, so the branch form of the instruction cannot be followed as
   written. **Tell it to retire that checkout with
   `git worktree remove --force <tmpdir>` when the gates are done** — the obvious
   alternative, deleting the directory recursively, matches a common deny rule and
   is refused, and a reviewer left to guess reaches for it. A creation
   instruction without a teardown instruction leaves worktrees behind.
   A shell-capable
   reviewer left to its own devices will `git checkout` the PR head in the
   shared tree, which moves the branch the loop itself is standing on and can
   leave behind a branch that squash-merge makes undeletable by `git branch -d`.
   The capability and its bound belong together. if the runtime offers no such reviewer,
   run the gates yourself and hand it the real output. A reviewer without a
   shell will not refuse — it will read what it can reach, reason about what
   the tests _would_ do, and return a pass by inference. That verdict is then
   the sole evidence for merging without a human, and it was never evidence at
   all. Record in the journal whether the reviewer executed the gates or
   inferred them; an inferred pass does not satisfy CD, so either re-review
   with a capable reviewer or treat the merge as human-gated.

   Fix confirmed findings within the iteration; a
   finding that needs a human decision is a question gate. With gates green
   and the review passed, the PR merges autonomously, so every finished
   workstream lands on a clean, merged main before the next begins. Under a
   human-gated merge policy, stop after the PR exists — that workstream has
   reached a review gate and goes to the shelf.

   **Merging and retiring the branch are two steps, not one flag.** Merge with
   `gh pr merge <n> --squash` alone, and retire the remote branch afterwards as
   a separate, explicit step. A merge that also carries `--delete-branch` reads
   as a destructive operation and is the form most often refused by a permission
   classifier, which strands an autonomous run one step short of done; keeping
   the two apart keeps the refusable surface small.

   If the merge is refused anyway, that is a **question gate**: leave the PR
   open, leave the branch alone, persist the refused command and the PR number,
   and stop. An open PR with green gates is already a legal end state under the
   gated path, so a refusal costs one human message rather than the run.

   **Re-read a workstream file before any edit that follows an edit to one of its
   markdown tables.** A formatter re-pads a table's columns to its widest cell, so
   changing one cell rewrites the whitespace of every other row, and the next edit
   — its target copied from before the reformat — fails to match. When several
   cells in one table need changing, rewrite them in a single pass instead of one
   edit per row. This costs an iteration nothing when expected and several failed
   edits when not; every workstream close touches the Issue Queue table.

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
  workstream first (`origin-goal` already does this), close out the current
  iteration (steps 3–5), then stop the **entire loop**. Parking an unanswered
  decision and starting more work would accumulate choices the human never
  sequenced; abandoning green, reviewed work unmerged would discard finished
  output for no gain. What the gate forbids is new work, not the landing of
  work already done.
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

**Every shelved PR was verified alone.** Each one's gates ran against the
default branch plus that one change, because that is the only state that
existed when it was built. Three independently green PRs therefore say nothing
about the state after all three land — and where a repository has no CI, the
fallback to local gates means nothing re-checks them at merge time either. Say
so in the digest's shelf table rather than presenting green PRs as ready: for
each shelved PR record that its gate result is isolated, and make the next
action a **combined** verification (merge them into one integration branch, or
merge sequentially re-running the gates after each) rather than a bulk merge.
Do not merge shelved PRs on the loop's own authority — they are shelved
precisely because a human owns that decision.

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
  unbounded observer drowns the triage queue and its own signal. The bound
  counts observations, so an issue another clause **ordered** you to file — the
  preflight drift issue, for instance — does not consume it. Preflight is not an
  iteration and has no allowance of its own.
- **Named for attribution.** A filing branch must start with the id of the
  workstream whose iteration produced the observation, so shelf reconstruction
  can attribute its PR (matching is by prefix). Preflight filings, which belong
  to no workstream, use a `ws-loop-preflight-` prefix and are reported as
  leftovers rather than shelved if their PR ever stays open. A branch named only
  for the observation is unattributable, and a later run reports it as a PR
  belonging to nothing.
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
レビュー待ち: | ws | PR | 何を承認すると何が進むか | gate: isolated |   ← shelf; omit when empty
                (2 本以上あるときは「統合後は未検証」と次のアクションに統合検証を明記)
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
- start new work after a question gate has arisen — no further issue in the
  current workstream, no next workstream. Closing the current iteration
  (review, merge what is already green, observe, journal) is not continuing;
  it is how the iteration ends. A review gate shelves the workstream; it never
  excuses executing another issue _inside_ it;
- put anything with an unanswered decision on the review shelf — the shelf is
  for finished work awaiting review only;
- leave an improvement observation unfiled (digest text is not a record) —
  except on a Prerequisites refusal, where reporting in the digest is the
  defined destination;
- edit skills or anything under `~/.agents` (loop-process improvements are
  recorded, not applied);
- loosen permissions, bypass hooks, or take any action the composed skills'
  own safety contracts would forbid.
