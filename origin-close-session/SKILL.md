---
name: origin-close-session
description: >-
  End-of-work closeout for a repository: bring BOTH documentation and version
  control to a clean, next-task-ready state in one pass. Runs origin-doc-update first
  (update the active workstream/issue, classify guide impact, land any docs/
  changes in the working tree) and then origin-git-cleanup (commit — including those
  doc edits — merge the finished branch into main, sync, and remove stale
  branches/worktrees). Use this whenever the user signals a chunk of work is
  finished and wants to tidy up, e.g. "店じまい", "後片付け", "片付けて",
  "クリーンに", "整理して", "実装完了したので整理", "wrap up", "close out",
  "get back to a clean main", or right after a PR is merged. Key signals:
  "店じまい", "片付けて", "後片付け", "クリーンに", "掃除", "整理", "wrap up",
  "closeout", "ブランチを消して", "全部コミットした？", "docsも更新した？",
  "きれいになった？". Prefer this over invoking origin-git-cleanup alone at end of task,
  because origin-git-cleanup by itself skips the documentation layer and the doc edits
  miss the commit. Do NOT use for mid-task git operations (rebase, cherry-pick,
  status-only checks) or for docs edits with no intent to clean up git.
---

# origin-close-session

The single entry point for "I'm done with this chunk of work — tidy everything."
Bundling exists for one reason: **documentation and git must be closed out
together, in that order.** Doing only origin-git-cleanup (the common habit) leaves the
docs stale, and doing docs after the commit means the doc edits miss the commit
that origin-git-cleanup makes. origin-close-session guarantees the order so neither happens.

This skill owns no new destructive behavior of its own. It orchestrates two
existing skills, each of which keeps its own safety contract — most importantly
origin-git-cleanup's **survey → propose plan → execute integration → verify** flow.
Commits, pushes, and green PR merges are autonomous; only destructive cleanup
(branch/worktree/ref deletion or a forced operation) requires confirmation.

## The order, and why it matters

1. **origin-doc-update first — edit, don't commit.** Update the active
   workstream/issue, classify guide impact (required vs none), and make any
   `docs/` edits in the working tree. Leave them uncommitted. origin-git-cleanup will
   pick them up as part of the same commit in the next phase.
2. **origin-git-cleanup second — commit everything, then clean.** It surveys the repo
   (now seeing the doc edits as uncommitted changes), proposes a single combined
   plan (commit incl. docs → merge → sync main → delete merged branches/
   worktrees), executes the integration steps, and verifies the tree ends clean
   on main. Any deletion steps remain separately subject to confirmation.

If these ran in the opposite order, origin-git-cleanup would commit and merge the code,
and the doc edits would land in a separate afterthought commit — or get
forgotten. First-docs-then-git keeps one reviewable, self-consistent commit.

## How to run it

### Step 1 — origin-doc-update

Invoke the **origin-doc-update** skill and follow its SKILL.md against the current work.
Concretely: read `docs/00_index.md` if present, update the active
workstream/issue with what changed, and update `docs/guides/` in the same slice
when implemented behavior changed. **Stop before committing** — origin-close-session commits
via origin-git-cleanup in Step 2.

Skip this step only when the repo has no `docs/` governance (`docs/00_index.md`
absent). Say so, then go straight to Step 2.

### Step 2 — origin-git-cleanup

Invoke the **origin-git-cleanup** skill and follow its SKILL.md. Its Stage 1 survey now
includes the doc edits from Step 1 as uncommitted changes, so its Stage 3 plan
should propose committing code **and** docs together. Execute that integration
plan autonomously, and finish on main, clean, and synced.

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
short-circuit that; let origin-git-cleanup's stages run.

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
  behavior changed,
- working tree clean, root on `main`, `main == origin/main`,
- only intentional branches/worktrees remain, each with a stated disposition.
