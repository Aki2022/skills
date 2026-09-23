---
name: own-doc-update
description: Keep repository documentation aligned with implementation during development, including recording architectural and implementation decisions as ADRs. Use when starting or resuming features, bugs, refactors, API/schema/config/command changes, docs work, workstream or issue creation, unfinished work, docs reorganization, or a hard-to-reverse or surprising design/development decision. Trigger especially when the user asks to create a workstream, when implemented behavior may require docs/guides updates, when an ADR should be created or superseded, or when docs/00_index.md exists. Do not use for pure operational checks, log inspection, status reporting, or read-only explanations with no behavior or documentation change.
---

# own-doc-update

Treat `docs/` as persistent AI context. Keep each fact in one layer only.

| Path                | Owns                                                                           |
| ------------------- | ------------------------------------------------------------------------------ |
| `docs/00_index.md`  | Small routing index; read first                                                |
| `docs/specs/`       | Intent, requirements, and design policy                                        |
| `docs/workstreams/` | Multi-issue autonomous work between human gates                                |
| `docs/issues/`      | Issue files; each declares `workstream: WS-…` (owned) or `workstream: none`     |
| `docs/guides/`      | Current implemented behavior; source of truth                                  |
| `docs/adrs/`        | Decision rationale, alternatives, and consequences; historical decision record |
| `*/archive/`        | Historical work context, not current truth                                     |
| `docs/log/`         | Narrative moved out of the index and hygiene reports; history, never current truth |

Do not copy implementation history into guides or current behavior into workstreams. Keep decision rationale in an ADR and link to it from the relevant spec, workstream/issue, or guide.

## Start every task

1. Read `docs/00_index.md` if present.
2. Read the active workstream or issue.
3. Read only its related specs, guides, and ADRs when present.
4. Choose one work unit:
   - Use a **workstream** when multiple vertical-slice issues can run under the same authorization envelope until the same next human gate.
   - Use a **standalone issue** for one bounded change, an unrelated blocker, or work with an independent lifecycle.
   - Use **docs only** when no implementation changes.
5. Reuse the recorded branch or worktree. Do not create a second branch for resumed work.

Default to one workstream file with embedded issue blocks. Split it only when independent branches, parallel ownership, or file size makes one file materially harder to resume.

## Issue ownership and routing

Ownership is decided by whose authorization envelope and next human gate an issue runs under — not by whether it has its own file, branch, or PR (SPEC-doc-governance, biz_ops).

- Every file in `docs/issues/` declares `workstream: WS-…` or `workstream: none`; that field is the only source of truth. An owned issue is routed from its workstream's `## Split Issues`; only `none` (standalone) issues are routed from the index's Active Issues.
- Standalone issues and workstreams carry `priority: high|medium|low` and `due: YYYY-MM-DD` (or `due: none` when there is truly no deadline — never an invented date), set by the session that drafts them.
- Lists between `own-doc-update:generated` markers are derived from front matter. Do not edit them; `create_issue.py`, `create_workstream.py`, `archive_issue.py` and `docs_hygiene.py --fix` rewrite them. The Active Workstreams row shows the earliest `due` of the workstream and its owned issues, with its source.
- Orphans are prevented at the three operations that can create them: creation requires the ownership choice and writes the route row, issue archive drops the row from both the index and the owning workstream, and workstream archive refuses while an active issue still declares it. `docs_hygiene.py --fix` is the backstop: it repairs what front matter determines and lists the rest under Unassigned Issues without rewriting the declared owner.
- Files created before `ownership.ROLLOUT_DATE` only warn when a field is missing and stay routed as standalone; adding `workstream:` to them is optional, and is what moves them out of Active Issues.
- Validator errors name the file whose change caused them (the pre-commit hook blocks only errors on staged files). Overdue dates are not checked; humans read them in the index.

## Create a workstream

Do not create the file immediately when the user asks for a workstream. First establish the human boundary.

1. Inspect the code and existing docs. Answer anything discoverable without asking the user.
2. Identify decisions that change authorization or the path of work.
3. Ask one question at a time, in dependency order, and include a recommended answer.
4. Confirm these items before writing the workstream:
   - goal, success criteria, and out-of-scope work;
   - actions the agent may take autonomously;
   - actions requiring confirmation, especially external writes, sends, submissions, deploys, destructive changes, data migrations, auth/secrets, dependencies, network access, and metered services;
   - cost or usage ceiling when metered work is possible;
   - test and quality gates;
   - merge policy: continuous delivery is the default — a PR whose recorded quality gates pass (CI when present, otherwise the recorded local gates) is merged autonomously. Record a human merge gate only as a named exception with its reason (e.g. live/production impact, spend, new dependencies); an unexplained merge gate silently kills autonomous runs downstream. The template's `- Merge policy:` line in the Authorization Envelope carries the choice;
   - the next human checkpoint and early stop conditions.
5. Classify runnability for every planned issue and fill its block's `runnability:` field: `ready` (completable with only current permissions and currently available information) or `gated on <the human decision, missing input, or new permission>`. Surface every gated point as a question now, during creation — a question asked here costs one interview turn, while the same question discovered mid-execution stops an entire autonomous run. As part of the same pass, judge whether each issue's acceptance is machine-verifiable (counts, thresholds, passing tests) or needs human review, and fill the `- verify:` line under its `#### Acceptance`: `machine — <command and expected result>` or `human-review — <who reviews what>`. Push subjective acceptance toward a quantifiable restatement; where human judgment is genuinely required, `human-review` records that the issue ends at a review gate — an agent cannot self-verify a subjective goal and will either stall or overclaim. These fields are not optional prose: `validate_repo_docs.py` rejects a missing or malformed value, because executors (`own-goal-run`, `own-ws-drain`) treat an unrecorded runnability as `gated` and stop the whole run at a gate nobody set.
6. State that `own-doc-update` is pausing creation until these boundaries are confirmed.
7. After confirmation, create from `references/workstream.template.md` or run `create_workstream.py` with the confirmed boundary fields.

Record only decisions that are hard to reverse or surprising without context. Do not create ADRs for routine implementation choices, temporary investigation notes, or ordinary history already captured by the workstream/issue.

Do not audit or revoke IAM roles, API scopes, bucket bindings, or other session-granted access here. `own-permission-audit` owns that lifecycle; `own-session-close` invokes it before this skill when applicable and passes the verified outcome into the documentation slice.

## Record an ADR

Use the repository-level `docs/adrs/` directory for both kinds of decision:

- `scope: spec` — intent, requirements, architecture, design policy, security posture, or other decisions that change what the system is meant to be. Update the related `docs/specs/` document with the current policy and link back to the ADR.
- `scope: development` — implementation, integration, migration, operational, dependency, or tooling tradeoffs made while delivering a workstream or issue. Link the ADR from that work unit and update a guide when the resulting behavior is user-, operator-, integrator-, or agent-visible.

Create the record when the decision is made, not only at session close. Use `references/adr.template.md` or `scripts/create_adr.py <slug> --scope <spec|development> --status <proposed|accepted|rejected>`. The generated filename is `ADR-YYYYMMDD-<slug>.md`. Keep accepted and rejected records in `docs/adrs/`; when a decision changes, create a new ADR and mark the old one `superseded` with a link instead of rewriting its decision history.

An ADR must state the context/problem, the decision, alternatives considered, consequences, and links to its source workstream/issue and affected specs/guides. Use `status: proposed` while a human gate is pending and `status: accepted` or `rejected` after the decision is settled. Record only the rationale here; current policy belongs in specs and current behavior belongs in guides.

## Improvement issues

An improvement issue is a standalone issue in `docs/issues/` that records a friction observation from development work — a stuck point, a repeated manual step, an inefficiency worth fixing later — rather than a requested change. Name it `ISSUE-YYYYMMDD-improve-<slug>` when the improvement targets the repository itself, and `ISSUE-YYYYMMDD-improve-loop-<slug>` when it targets the development-loop machinery (skills canonical under `~/.agents/skills/`). The two scopes have different owners and approval paths, so the name must reveal the scope at a glance. Autonomous runs (e.g. `own-ws-drain`) file observations here instead of interrupting their work; humans triage them later. Create with `create_issue.py` as usual.

Boundary with `own-trouble-log`: an improvement issue is an actionable change request against this repository or the loop machinery. An observation about how the agent itself worked wrong — a silent no-op, a false completion report, a vacuous check, guidance friction — must ALSO be recorded as one `own-trouble-log` entry, and when it is only an observation (no concrete change to implement yet) it goes ONLY there; a repo issue filed instead of a trouble entry is invisible to the cross-repo triage and was measured to get lost (2026-08-10 observation-filed-to-wrong-corpus).

## Implement a workstream or issue

For each issue, complete one vertical slice:

1. Set `guide_impact` before implementation:
   - `required`: name every guide that must describe the resulting behavior.
   - `none`: write a concrete reason, such as internal refactor with unchanged behavior.
2. Define acceptance criteria and dependencies. The `- verify:` line under Acceptance is the executable form: run the recorded machine check, or route to the recorded reviewer.
3. Capture each qualifying decision in `docs/adrs/` during the slice and link it from the issue/workstream. Mark a human-gated decision as `proposed` until the gate is resolved.
4. Implement and test, preferring red-green-refactor where practical.
5. Update the target guide in the same slice, before marking the issue complete.
6. Update current status and next actions. An issue's `status` must be exactly
   one of `pending`, `in_progress`, `blocked`, `complete` — `validate_repo_docs.py`
   rejects anything else, and plausible words like `done` are the usual way to
   find that out the hard way.

7. Continue automatically while inside the authorization envelope.
8. Stop at the next human gate or any recorded stop condition.

Never defer all guide work to workstream close. A guide is part of the definition of done for the issue that changed behavior.

## Guide contract

Update or create a guide when behavior observable by a user, operator, integrator, or future agent changes, including:

- commands, configuration, schemas, APIs, supported workflows, and defaults;
- operational procedures, safety constraints, failure handling, and known limitations;
- behavior needed to use, maintain, debug, or extend the implementation correctly.

Do not update a guide for a pure internal refactor with identical behavior. Record `guide_impact: none` and why.

Write guides as current truth, not as a changelog. Include what the system does, how to use it, guarantees or constraints, maintenance/verification notes, and known limitations. Add the source workstream or issue ID in front matter.

## Update specs and history

- Update a spec only when intent, requirements, architecture, or design policy changes.
- When that change follows a qualifying decision, update the related spec with the current policy and link the ADR; do not duplicate the full rationale in the spec.
- Keep chronological investigation and abandoned approaches in the active work unit, then archive it.
- Keep `docs/00_index.md` as links plus one-line routing descriptions. Do not add a second progress dashboard unless ordering across many workstreams cannot fit in the index.
- The index must stay **under 32 KB with no line over 500 characters**; `validate_repo_docs.py` rejects both. The SessionStart hook injects a routing digest of it (see Hooks), so the file itself is what a session opens next. Progress narrative, postmortems, and metrics belong in the work unit or in `docs/log/`, never in the index. `docs_hygiene.py --fix` moves them there and shortens over-long rows; when the ceiling still cannot be met the row count is the problem and only closing work fixes it (see Cleanup cadence).

## Write for the next session

Every document is read by a session that remembers nothing. Measured 2026-09-18 across
four repositories, the failures that cost the most were all writing habits, not missing
information: a 291 KB index of progress prose that the hook could not deliver, 81 of 166
"active" rows that said 完了, and guides that had become dated decision logs. The rules
below are the shape that survives handoff; the parenthesis names what enforces each.

- **Index row = one link + one line of routing, no status.** Under 200 characters after
  the link; the frontmatter `status` is the only status. Current Focus holds at most five
  entries. What happened goes to the work unit; how it happened goes to `docs/log/`.
  (validator: line length, size; hygiene A2 moves and shortens.)
- **Work unit = snapshot + next step + log.** `## Current Status` is overwritten, first
  line `as of YYYY-MM-DD — <one sentence>`. `## Next Actions` names the very next command
  or step, or what unblocks a blocked issue; it is never empty once work has started.
  Chronology goes under `## Log` as dated append-only bullets. (validator: empty Next
  Actions on an open issue is an error once `updated_at` differs from `created_at`.)
- **Say it with status, not prose.** The moment acceptance's `verify:` passes, set
  `status: complete` in the same commit; archive follows at close-session. Writing
  完了 / ✅ / done in an open issue's status line or its index row is the defect that
  produced the 81 rows. (hygiene R7 reports it.)
- **Guide = current truth, spec = current policy.** No dated headings, no "2026-09-05
  revision" sections: a decision becomes an ADR, a superseded paragraph is deleted, the
  narrative of getting there goes to the work unit or `docs/log/`. A guide over 60 KB or
  with more than five dated headings is a split or a cleanup waiting to happen.
  (hygiene R4 reports it.)
- **Dates are maintained by git, not by memory.** `updated_at` older than the file's last
  commit is corrected from git. (hygiene A5.)
- **Paths named in a guide must exist.** Write the path a reader can open; a generated
  artifact that is absent by design is gitignored so the check can tell. (hygiene R3.)

## Cleanup cadence

Cleanup has three layers, and none of them is a calendar: work that is not closed in the
session that finished it is the only source of accumulation, so the cadence is tied to
sessions.

1. **Every session start** — the hook delivers a routing digest bounded to what arrives
   inline. Nothing to do; if the digest says it dropped entries, the index is oversized.
2. **Every close-session** — `docs_hygiene.py <repo> --fix --report`. Mechanical repairs
   land in the same commit as the session's docs; the report names what needs a decision.
   Decide the items that belong to the session's own work right there.
3. **Sweep when the report says so** — when R1 + R2 + R7 exceeds ten, or the index is over
   its ceiling after `--fix`, write `docs_hygiene.py <repo> --sweep`. It produces
   `docs/log/sweep-YYYYMMDD.md`: one checkbox per closure candidate with its evidence.
   Tick `[x]` to archive; leave `[ ]` and append `— keep: <reason>` to keep. The agent may
   tick rows whose evidence is mechanical (branch merged and deleted, Current Status says
   done); everything else waits for the human. `--apply-sweep <file>` archives exactly the
   ticked rows and stamps the file, so the decision is recorded next to the work it closed
   and the next session does not re-ask it. The ticking is done by `own-docs-maintain`'s
   judge subagent on primary evidence (Completion boxes, artifacts that exist, commits on
   main); a human reads the file afterwards instead of being asked twenty questions.

4. **Review at every close-session; 14 days is the backstop.** Closing work keeps
   the execution layer honest, but nothing above re-weights the current-truth layer: a
   guide every issue once linked becomes one nobody reads, a spec stays `active` after the
   work that needed it is archived, and a 170 KB guide is still "one guide". `--report`
   therefore also writes `docs/log/review-YYYYMMDD.md` (deterministic, seconds): every
   guide/spec/ADR tiered by who links to it (hot = active work or another guide/spec, warm
   = index only, cold = nothing living), demotion candidates (cold and older than 90 days)
   as `[x]`-able archive rows, and split candidates (over 32 KB or more than five dated
   headings) with their H2 sections sized so the cut is mechanical. Decide demotions in the
   review file and apply the archive rows with `--apply-sweep`; splits go one guide per task,
   keeping the old path as a pointer with its original front matter. **Who decides:**
   `own-docs-maintain` — a judge subagent ticks the sweep and review rows on primary
   evidence and split/condense subagents restructure at most two documents per
   close-session; humans read the recorded decisions, they are not asked for them
   (ADR-20260921-autonomous-docs-curation, biz_ops). The validator
   warns and the SessionStart hook says one line when the newest review is older than 14
   days — that fires for a repository no session has closed in two weeks.

**Context budget.** docs/ is context, so its cost is measured, not guessed. Hygiene's R9
reports three numbers every close-session: index entries the digest had to drop (must be
0 — each dropped entry is work the next session cannot find), guides/specs over 32 KB
(a guide is read whole, so one such file costs more than the whole routing budget; the
validator warns on each), and the total size of guides/specs linked from active work (what
a working session actually opens). When dropped > 0 the fix is closing or archiving
rows; when a guide is over 32 KB the fix is the split candidate in the review; when the
hot set is large the fix is fewer, smaller guides per active issue.

Never run `--fix` or `--apply-sweep` in a repository with uncommitted docs/ changes that
are not yours, and never against a deliberately shaped fixture repository.

## Complete work

Before archive:

1. Verify every issue acceptance criterion.
2. Verify every issue has `guide_impact: required` or `none`.
3. Verify required guides describe the implemented behavior and reference the source work.
4. Verify qualifying decisions have an ADR in `docs/adrs/`, with a settled status or an explicit proposed human gate, and that related specs/guides/work units link to it.
5. Update specs if direction changed.
6. Reach the recorded human gate or record why the workstream stopped.
7. Run `docs_hygiene.py <repo path> --fix --report`, then read the report it
   names. The fixes are mechanical (archive what says `complete`, move index
   narrative to `docs/log/`, normalize status aliases, add missing front matter
   from git dates); the report lists what needs a decision — issues whose
   branch is gone, untouched work, dead references, history mixed into guides —
   and every count is printed, zero included. Decide the reported items that
   belong to this session's work; leave the rest in the report. When the report's
   R1 + R2 + R7 exceeds ten or the index is still over its ceiling, write a sweep
   (`--sweep`) and hand the checklist to the human as this session's one question.
8. Run `validate_repo_docs.py <repo path>`. Name the repository rather than
   relying on the current directory: reached through an orchestrator, the current
   directory is a different repository, whose docs would validate clean and be
   reported as this one's result. Check the `validated:` line it prints.
9. Archive the work unit and update `docs/00_index.md`. The archive scripts accept an issue/
   workstream id, `.md` filename, or path, stage the document/index changes before applying them,
   and roll back both files if a later replacement fails. They print the removal count plus the
   exact index lines they changed. Treat a zero or unexpected count as a stop condition and
   inspect the diff before continuing. An entry the row matcher does not recognize (prose, a
   nested bullet) is repointed at the archive path instead of being left pointing at the file
   that just moved; pass `--keep-row` when the index's own policy keeps completed rows, and the
   row is repointed in place rather than removed.
10. Hand merged branch cleanup to `own-git-clean`.

## Resume and onboard

When resuming, follow `docs/00_index.md` to the active work unit, reuse its branch, then continue from `Next Actions` without rescanning the repository.

When onboarding scattered docs, initialize the scaffold, classify each file by the ownership table, preserve history with `git mv`, confirm ambiguous removals, update cross-references, and validate.

## Hooks

Keep hooks best-effort and non-blocking:

- `session_start.sh` injects a routing digest of `docs/00_index.md` built by
  `index_digest.py`, never the file itself. A hook's output reaches the agent inline
  only up to about 10,000 characters (measured 2026-09-19: largest delivered 8,957,
  smallest spilled 10,019); past that it is written to a file and the agent receives
  a 2 KB preview, so a whole-file injection is read by nobody and nothing reports it.
  The digest keeps every active entry with one line of routing, counts what it
  dropped, and names the full file.
- `stop_nudge.sh` emits one short message at turn end: a reminder when non-doc
  changes lack docs changes, and — when docs/ changed this session — the validator's
  error count with the first five errors. It must be registered under `Stop` in
  `~/.claude/settings.json` (measured 2026-09-21: it had existed for weeks and was
  registered nowhere, so "the hook checks" was nominal). Warn-only, capped output,
  20-second timeout; never blocks.
- Do not force-load this full skill from a hook and do not block commits or task completion from a semantic guess.

Use template fields and `validate_repo_docs.py` for deterministic enforcement. Hooks cannot reliably infer whether behavior changed and hard enforcement creates false positives and repeated token cost.

## Resources

Templates in `references/`:

- `00_index.template.md`
- `workstream.template.md`
- `issue.template.md`
- `guide.template.md`
- `adr.template.md`
- `spec.template.md`

Scripts in `scripts/`:

- `init_repo_docs.py [repo]`
- `create_workstream.py <slug> --issue <slug> --scope <text> --confirmed-at YYYY-MM-DD --next-human-gate <name> --autonomous <text> --confirm-first <text> (--verify-machine <text> | --verify-human <text>) (--guide <GUIDE-id> | --no-guide-reason <text>) --priority <high|medium|low> --due <YYYY-MM-DD|none> [--merge-policy <text>] [--gated-on <text>] [--repo <repo>]`
  Writes the generated Active Workstreams row in the same run.
  The interview's confirmed boundaries are required arguments: envelope
  (`--autonomous`, `--confirm-first`), acceptance (`--verify-*`), and — when the
  initial issue is not immediately runnable — `--gated-on <reason>`. A file created
  without them validates red, and executors treat the missing record as a gate.
- `create_issue.py <slug> (--workstream <WS-id> | --standalone --priority <high|medium|low> --due <YYYY-MM-DD|none>) (--verify-machine <text> | --verify-human <text>) (--guide <GUIDE-id> | --no-guide-reason <text>) --next-action <text> [--title <title>] [--repo <repo>]`
  Pass a slug, not a full issue id — the `ISSUE-<date>-` prefix is added for you.
  The guide decision, the acceptance decision and `--next-action` are required, as
  the first two are for `create_workstream.py`: name the guide this issue must
  update (or why not), state how acceptance is verified, and give the very next
  step. Without them the generated file cannot pass `validate_repo_docs.py`.
  `--next-action` replaced an exemption: the `## Next Actions` check used to skip a
  file whose `updated_at` still equalled `created_at`, but that condition is a
  hand-maintained value, so an issue whose `updated_at` was never touched stayed
  exempt forever while being worked on (measured 2026-09-20: 56 of 477 open work
  units across 13 repositories, unreported for over two weeks). The generator now
  produces a real first step, so the check needs no exemption and depends on
  neither the clock nor git. The ownership choice writes the route row in the same
  run (the owning workstream's Split Issues, or the index's Active Issues); an
  inactive workstream is refused and nothing is created.
- `create_adr.py <slug> --scope <spec|development> [--status <proposed|accepted|rejected>] [--title <title>] [--repo <repo>]`
- `archive_workstream.py <workstream> [--repo <repo>]`
- `archive_issue.py <issue> [--repo <repo>]`
- `archive_transaction.py`: stage and atomically roll back archive/index file updates
- `docs_hygiene.py <repo> [--fix] [--report] [--json]`: `--fix` applies the
  mechanical repairs above; `--report` writes `docs/log/hygiene-YYYYMMDD.md` with
  the judgment candidates (R1 stale + branch gone, R2 untouched 60 days, R3 dead
  `npm run`/workflow/path references, R4 history in guides or specs, R5
  non-canonical directories and duplicate basenames, R6 baseline debt, R7 open
  issues whose Current Status or index row already says done). `--fix` also syncs
  `updated_at` from git (A5). `--sweep` writes `docs/log/sweep-YYYYMMDD.md`, the
  checklist of closure candidates; `--apply-sweep <file>` archives its `[x]` rows
  and stamps the file (a review file's archive rows work the same way). `--review`
  writes `docs/log/review-YYYYMMDD.md`, the re-weighting of guides/specs/ADRs (due every 14 days)
  (R8 reports when it is overdue); `--report` writes it too, so every close-session
  reviews. R9 reports the context budget (digest drops, oversized docs, hot-set size).
  A6 regenerates the ownership lists from front matter and drops the hand-written rows
  they replace; R10 counts where every active issue is routed from (orphans and
  unassigned issues, zero included).
  Without `--fix` it is a dry run that counts. It uses git dates and never an LLM, so it
  is safe to run across every governed repository. Exit 2 when the repository has
  no `docs/00_index.md`. Do not run `--fix` against a deliberately shaped fixture
  repository (e.g. `ws-loop-fixture`).
- `index_digest.py <path to docs/00_index.md>`: the routing digest the SessionStart
  hook injects, bounded in characters rather than bytes
- `validate_repo_docs.py <repo>` (prints the repository it validated, and a
  `coverage:` line saying how many open work units the `Next Actions` check
  actually examined. Being green and having looked at everything are different
  facts; without the count, a check that silently skips most of its population
  is indistinguishable from one that passes.)

  It also resolves every relative link under `docs/**/*.md` and fails on any
  that does not exist — archiving moves a file one level deeper and leaves its
  referrers behind, and that breakage is otherwise invisible until someone
  follows a link. Only inline code spans are excluded — write an unresolvable
  path as `` `[x](../placeholder.md)` `` and it is ignored. **A link inside a
  fenced block or an HTML comment is reported**: block-level parsing was tried
  and removed, because across 26 repositories using this convention it cost 5
  loud false positives in 2 of them and never found a link the simpler rule
  misses, while repeatedly opening regions it never closed and silently deleting
  every link to the end of a document. A repository adopting the check
  with existing rot can record it as debt in `docs/validator-link-baseline.txt`
  (one `path<TAB>target` per line, preferring the target-scoped form); the list
  only shrinks, so an entry that is now resolvable is itself an error.

Read the matching template before creating a file manually. Keep legacy archives in place; promote useful current knowledge into a guide or spec instead of renaming history.
