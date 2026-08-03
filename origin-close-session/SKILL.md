---
name: origin-close-session
description: >-
  Close out a finished repository task by optionally running origin-permission-audit,
  then origin-doc-update, then origin-git-cleanup. Commit/merge/sync the finished
  work and remove only approved stale branches and worktrees. Use for "店じまい",
  "後片付け", "片付けて", "クリーンに", "整理", "wrap up", "close out",
  "get back to a clean main", or after a PR is merged. Prefer this over
  origin-git-cleanup alone when docs/ governance exists. Do not use for mid-task
  rebase, cherry-pick, status-only checks, or docs edits without cleanup intent.
---

# origin-close-session

The single entry point for "I'm done with this chunk of work — tidy everything."
Bundling exists for one reason: **documentation and git must be closed out
together, in that order.** Doing only origin-git-cleanup (the common habit) leaves the
docs stale, and doing docs after the commit means the doc edits miss the commit
that origin-git-cleanup makes. origin-close-session guarantees the order so neither happens.

This skill owns no new destructive behavior of its own. It orchestrates an optional
permission-audit preflight and two closeout phases, each of which keeps its own
safety contract — most importantly origin-git-cleanup's
**survey → propose plan → execute integration → verify** flow.
Commits, pushes, and green PR merges are autonomous; only destructive cleanup
(branch/worktree/ref deletion or a forced operation) requires confirmation.

## The order, and why it matters

0. **origin-permission-audit first, when applicable — decide, act, verify, _then_
   let Step 1 document it.** If this session granted elevated access, invoke
   `origin-permission-audit` before writing docs about it. That skill owns the
   discovery, classification, revocation, and verification contract.
1. **origin-doc-update next — edit, don't commit.** Update the active
   workstream/issue, classify guide impact (required vs none), and make any
   `docs/` edits in the working tree — including the outcome of Step 0, if it
   ran. `origin-doc-update` owns ADR recording. Leave the edits uncommitted;
   origin-git-cleanup will pick them up as part of
   the same commit in the next phase.
2. **origin-git-cleanup last — commit everything, then clean.** It surveys the repo
   (now seeing the doc edits as uncommitted changes), proposes a single combined
   plan (commit incl. docs → merge → sync main → delete merged branches/
   worktrees), executes the integration steps, and verifies the tree ends clean
   on main. Any deletion steps remain separately subject to confirmation.

If these ran in the opposite order, origin-git-cleanup would commit and merge the code,
and the doc edits would land in a separate afterthought commit — or get
forgotten. First-permissions-then-docs-then-git keeps one reviewable,
self-consistent commit whose docs match what actually happened to any access
this session was granted.

## How to run it

### Step 0 — origin-permission-audit (only if this session granted elevated access)

Skip this step, and say so, when nothing in this session (or a resumed session
whose history you can see) required elevated access beyond the repository's
committed baseline. When a grant exists, invoke `origin-permission-audit` and
follow its discovery → classify → revoke → verify → report workflow. Pass its
per-grant report to Step 1; do not duplicate its permission logic in this skill.

### Step 1 — origin-doc-update

Invoke the **origin-doc-update** skill and follow its SKILL.md against the current work.
Concretely: read `docs/00_index.md` if present, pass the Step 0 report through,
and invoke `origin-doc-update` for the active workstream/issue. It owns ADR
recording, guide impact, and current-document updates. **Stop before committing**
— origin-close-session commits via origin-git-cleanup in Step 2.

Skip this step only when the repo has no `docs/` governance (`docs/00_index.md`
absent). Say so, then go straight to Step 2.

### Step 2 — origin-git-cleanup

Invoke the **origin-git-cleanup** skill and follow its SKILL.md, **passing it the
repository being closed out**. When this closeout was reached through an
orchestrator, the current directory is that orchestrator's own repository, not
the target, and a survey that defaults to the current directory reports the wrong
repository into a deletion plan.

Its Stage 1 survey now includes the doc edits from Step 1 as uncommitted changes,
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
just merged is part of the merge (see origin-git-cleanup's deletion authority).

The current branch you are standing on is the usual cleanup target — origin-git-cleanup
now handles switching to main before deleting it (local and remote). Don't
short-circuit that; let origin-git-cleanup's stages run. If the current session
worktree itself is an approved clean deletion target, origin-git-cleanup must
run `cd <safe-worktree> && git worktree remove <target-worktree>` in one shell
invocation so the removal process is no longer inside the directory it removes.

## Destructive-cleanup approval principle

Do not request approval for ordinary integration. origin-doc-update's edits are
ordinary, reviewable file changes; surface them and let origin-git-cleanup commit,
push, and merge them after its checks pass. If destructive cleanup is needed,
present only those deletion targets and obtain one confirmation before acting.

## When part of the flow doesn't apply

- **No `docs/` governance** → skip Step 1, run Step 2 only.
- **Nothing to document** (pure refactor with no behavior/API change, and
  origin-doc-update concludes "guide impact: none") → record that conclusion, then
  Step 2.
- **Already on a clean main with nothing to merge** → origin-doc-update may still have
  index/workstream updates; otherwise origin-close-session is a no-op beyond confirming the
  clean state. Say so rather than inventing work.

## What "done" looks like

Same end state origin-git-cleanup verifies, plus docs current:

- active workstream/issue reflects the finished work; guides updated where
  behavior changed; qualifying decisions are recorded and linked in `docs/adrs/`,
- working tree clean, root on `main`, `main == origin/main`,
- only intentional branches/worktrees remain, each with a stated disposition.
