---
name: own-session-close
description: >-
  Close out a finished repository task by optionally running own-permission-audit,
  then own-trouble-log, then origin-doc-update, then own-git-clean. Commit/merge/sync the finished
  work and remove only approved stale branches and worktrees. Use for "店じまい",
  "後片付け", "片付けて", "クリーンに", "整理", "wrap up", "close out",
  "get back to a clean main", or after a PR is merged. Prefer this over
  own-git-clean alone when docs/ governance exists. Do not use for mid-task
  rebase, cherry-pick, status-only checks, or docs edits without cleanup intent.
---

# own-session-close

The single entry point for "I'm done with this chunk of work — tidy everything."
Bundling exists for one reason: **documentation and git must be closed out
together, in that order.** Doing only own-git-clean (the common habit) leaves the
docs stale, and doing docs after the commit means the doc edits miss the commit
that own-git-clean makes. own-session-close guarantees the order so neither happens.

This skill owns no new destructive behavior of its own. It orchestrates an optional
permission-audit preflight and two closeout phases, each of which keeps its own
safety contract — most importantly own-git-clean's
**survey → propose plan → execute integration → verify** flow.
Commits, pushes, and green PR merges are autonomous; only destructive cleanup
(branch/worktree/ref deletion or a forced operation) requires confirmation.

## The order, and why it matters

0. **own-permission-audit first, when applicable — decide, act, verify, _then_
   let Step 2 document it.** If this session granted elevated access, invoke
   `own-permission-audit` before writing docs about it. That skill owns the
   discovery, classification, revocation, and verification contract.
1. **own-trouble-log next — capture while the context is still here.** Sweep the
   session for troubles that were not already recorded at the moment they were
   pointed out, and write one entry per trouble. It writes outside git, so it does
   not interact with the commit ordering below. It runs before the git phase on
   purpose: that phase can stop for a confirmation, and anything placed after it
   would silently never run.
2. **origin-doc-update next — edit, don't commit.** Update the active
   workstream/issue, classify guide impact (required vs none), and make any
   `docs/` edits in the working tree — including the outcome of Step 0, if it
   ran. `origin-doc-update` owns ADR recording. Leave the edits uncommitted;
   own-git-clean will pick them up as part of
   the same commit in the next phase.
3. **own-git-clean last — commit everything, then clean.** It surveys the repo
   (now seeing the doc edits as uncommitted changes), proposes a single combined
   plan (commit incl. docs → merge → sync main → delete merged branches/
   worktrees), executes the integration steps, and verifies the tree ends clean
   on main. Any deletion steps remain separately subject to confirmation.

If these ran in the opposite order, own-git-clean would commit and merge the code,
and the doc edits would land in a separate afterthought commit — or get
forgotten. First-permissions-then-docs-then-git keeps one reviewable,
self-consistent commit whose docs match what actually happened to any access
this session was granted.

### Human review handoff

Before Step 3, if a related current guide contains a `## 承認前レビュー` section, read it and
surface each listed decision in the handoff. Honor any item that the guide marks as human-gated;
tests, lint, and an archived issue are evidence, not a substitute for that approval. Do not copy
the checklist into `own-trouble-log`: that skill stores observed trouble evidence, while the
guide remains the current operational procedure.

## How to run it

### Step 0 — own-permission-audit (only if this session granted elevated access)

Skip this step, and say so, when nothing in this session (or a resumed session
whose history you can see) required elevated access beyond the repository's
committed baseline. When a grant exists, invoke `own-permission-audit` and
follow its discovery → classify → revoke → verify → report workflow. Pass its
per-grant report to Step 2; do not duplicate its permission logic in this skill.

### Step 1 — own-trouble-log

Invoke the **own-trouble-log** skill and sweep this session for troubles that
were not already recorded when they happened: silent no-ops, completion reports
that turned out to be wrong, questions asked about something already automated or
already approved, checks that passed without checking anything.

Record one entry per trouble. If there were none, say so — do not invent one. This
step writes outside git and never blocks the git phase.

Its main trigger is a user pointing out how the work was done, which fires during
the session; this sweep is the backstop for the ones nobody pointed out.

### Step 2 — origin-doc-update

Invoke the **origin-doc-update** skill and follow its SKILL.md against the current work.
Concretely: read `docs/00_index.md` if present, pass the Step 0 report through,
and invoke `origin-doc-update` for the active workstream/issue. It owns ADR
recording, guide impact, and current-document updates. **Stop before committing**
— own-session-close commits via own-git-clean in Step 3.

As part of this step, run the repository's hygiene pass and name the repository:
`python3 ~/.agents/skills/origin-doc-update/scripts/docs_hygiene.py <repo> --fix --report`.
It archives what is marked complete, keeps `docs/00_index.md` under its size
ceiling, and writes `docs/log/hygiene-YYYYMMDD.md` with the items that need a
decision. Read the counts it prints; act on the reported items that belong to
this session's work and leave the rest recorded. This is the only scheduled
maintenance docs/ gets (decision 2026-09-18: no weekly job), so skipping it here
means the repository is never cleaned. When the report's R1 + R2 + R7 exceeds ten
or the index is still over its ceiling, run `--sweep` and present the resulting
`docs/log/sweep-YYYYMMDD.md` as this closeout's one batched question; apply the
ticked rows with `--apply-sweep` before Step 3 so the closure lands in the same
commit.

Skip this step only when the repo has no `docs/` governance (`docs/00_index.md`
absent). Say so, then go straight to Step 3.

### Step 3 — own-git-clean

Invoke the **own-git-clean** skill and follow its SKILL.md, **passing it both
of its inputs**: the repository being closed out, and the merge policy.

The repository, because when this closeout was reached through an orchestrator the
current directory is that orchestrator's own repository, not the target, and a
survey that defaults to the current directory reports the wrong repository into a
deletion plan.

The merge policy, because `own-git-clean` cannot infer it and its defaults all
point at merging. State `human-gated` whenever the workstream's envelope reserves
the merge or the caller said not to merge — a reservation carried only in this
skill's prose gets merged by the skill that performs the merge, and the report
reads as an ordinary landing.

Its Stage 1 survey now includes the doc edits from Step 2 as uncommitted changes,
so its Stage 3 plan should propose committing code **and** docs together. Execute
that integration plan autonomously, and finish on main, clean, and synced.

**Under a gated merge policy, the finish line is different and must not be
chased past.** When the workstream's envelope marks merge as human-gated, or
the caller says the PR is not to be merged, closeout is complete once the work
is committed, pushed, and the PR exists — with the PR still open and the
branch still present. That is the finished state, not a partial one: do not
merge to reach "clean on main", and do not delete the branch the open PR needs.
Report the PR and stop. The tree being clean matters; the branch it is standing
on does not have to be main.

If cleanup would
delete a branch, worktree, or ref **that this session did not create**, obtain
confirmation for those deletion steps. Deleting the branch this session made and
just merged is part of the merge (see own-git-clean's deletion authority).

The current branch you are standing on is the usual cleanup target — own-git-clean
now handles switching to main before deleting it (local and remote). Don't
short-circuit that; let own-git-clean's stages run. If the current session
worktree itself is an approved clean deletion target, own-git-clean must
run `cd <safe-worktree> && git worktree remove <target-worktree>` in one shell
invocation so the removal process is no longer inside the directory it removes.

## Destructive-cleanup approval principle

Do not request approval for ordinary integration. origin-doc-update's edits are
ordinary, reviewable file changes; surface them and let own-git-clean commit,
push, and merge them after its checks pass. If destructive cleanup is needed,
present only those deletion targets and obtain one confirmation before acting.

## When part of the flow doesn't apply

- **No `docs/` governance** → skip Step 2, run Step 1 and Step 3 only.
- **Nothing to document** (pure refactor with no behavior/API change, and
  origin-doc-update concludes "guide impact: none") → record that conclusion, then
  Step 3.
- **Already on a clean main with nothing to merge** → origin-doc-update may still have
  index/workstream updates; otherwise own-session-close is a no-op beyond confirming the
  clean state. Say so rather than inventing work.

## What "done" looks like

Same end state own-git-clean verifies, plus docs current:

- active workstream/issue reflects the finished work; guides updated where
  behavior changed; qualifying decisions are recorded and linked in `docs/adrs/`,
- working tree clean, root on `main`, `main == origin/main`,
- only intentional branches/worktrees remain, each with a stated disposition,
- the session's troubles are recorded, or their absence is stated.
