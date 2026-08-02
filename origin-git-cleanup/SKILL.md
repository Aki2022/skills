---
name: origin-git-cleanup
description: >-
  Wrap up and clean a git/GitHub repository after a chunk of work is done.
  Commits/pushes the appropriate changes, merges the finished branch into main,
  syncs main with origin/main, and removes stale merged branches and worktrees —
  leaving a clean tree on main, ready for the next task. Use this whenever the
  user signals work is finished and wants to tidy up version control, e.g.
  "git をクリーンに", "後片付け", "実装完了したので整理して", "main に戻して整理",
  "ブランチを掃除", "wrap up", "clean up branches", "get back to a clean main",
  or right after a PR is merged. Key signals: "後片付け", "片付けて", "クリーンに",
  "掃除", "整理", "merge to main and delete branch", "finished with this branch",
  "PR マージ後の整理", "ブランチを消して", "cleanupもした？", "ブランチきれい？",
  "きれいになった？", "全部コミットした？", "push した？". Use for end-of-task git
  housekeeping. Do NOT use for mid-task git operations like rebasing,
  cherry-picking, checking git status only, or committing individual files
  without cleanup intent.
---

# origin-git-cleanup

> **Wrapping up a whole chunk of work?** If the repo uses `docs/` governance,
> prefer the **origin-close-session** skill instead of calling origin-git-cleanup
> directly — it runs
> origin-doc-update first so documentation and git are closed out together and the doc
> edits land in the same commit. Use origin-git-cleanup directly for git-only
> housekeeping (no docs to update).

Bring a repository to a clean, ready-for-next-task state after a unit of work is
done. Target end state:

- current branch is **main** (or the repo's integration branch),
- working tree is **clean**,
- **main == origin/main**,
- branches and worktrees that were only needed for the finished work are gone,
- everything still present is there on purpose.

## Why this needs judgment, not a script

The dangerous failure mode is acting mechanically. A repo often has **parallel
work in flight** — other branches, other worktrees, someone else's open PR. If
you blindly delete branches or sweep stray edits into a commit, you destroy work
that wasn't yours to touch. So the whole skill is built around one shape:

**Survey (read-only) → Classify every branch/worktree → Propose a plan →
Execute integration → Verify root/main cleanliness.**

When anything is ambiguous, stop and ask rather than guess. Leaving something
alone is always safe; deleting the wrong thing is not.

## Cardinal rules

- **Read before you write.** Never mutate state before you've surveyed it.
- **Propose before you destroy.** Present the full plan before execution.
  Commit, push, and a green PR's merge are autonomous; get approval before any
  forced operation or other destructive cleanup.
- **Deletion authority follows provenance.** Removing a branch **this run itself
  created and has just merged** is part of that merge, not a separate
  destruction — `gh pr merge --squash --delete-branch` stays the ordinary merge
  command and needs no extra approval. Deleting anything this run did not create
  — a pre-existing branch, someone else's worktree, a remote ref you did not
  push — still requires approval. What the rule protects against is losing work
  you never saw, which cannot describe work you just wrote, merged, and
  verified.
- **Don't guess about parallel work.** If a branch or worktree might belong to
  other active work, leave it and say so.
- **Prefer the reversible.** Use `git branch -d` (never `-D`); never
  `git push --force`, `git reset --hard`, or `git clean -fd` without an explicit
  instruction from the user.
- **Follow repo conventions.** Use the `gh` CLI for all GitHub operations. Match
  the repository's existing commit-message style.
- **Classify every leftover explicitly.** Nothing should remain as "probably
  active." Each branch/worktree must end as `integrate`, `delete`, `keep`, or
  `investigate`, with a short reason.
- **Root worktree must finish clean on main.** Cleanup is not complete if the
  repository root is still on a feature branch, detached HEAD, dirty, or behind
  origin.
- **Patch-equivalence beats branch ancestry.** Squash merges and cherry-picks
  often make `git branch --merged` lie by omission. Use patch/stat inspection
  before deciding a branch is unmerged work.
- **Write context-rich commit messages.** A good commit message explains not just
  _what_ changed but _why_ — the problem that prompted the work, the decision made,
  and any non-obvious constraints. Future readers (including AI agents doing code
  archaeology) should be able to reconstruct the reasoning without needing to ask.
  A one-line summary is rarely enough. When proposing the commit message in
  Stage 3, draft a multi-line body that captures: what was broken or missing, what
  was investigated, what approach was chosen, and what was confirmed to work.

---

## Stage 1 — Survey (read-only)

Run the bundled snapshot script from the repo root. It only reads, never writes:

```bash
bash ~/.agents/skills/origin-git-cleanup/scripts/survey.sh
```

It reports: current branch, remote, root cleanliness, every worktree's branch /
HEAD / dirty status, upstream ahead/behind, stashes, local and remote branches,
branches merged vs not merged into main, recent history, and open PRs.

If the script isn't available, gather the same picture manually with
`git status`, `git branch -vv`, `git branch -a`, `git worktree list --porcelain`,
`git stash list`, `git log --graph --oneline -15`,
`git branch --merged main` / `--no-merged main`, and `gh pr status`.

If the repo has `docs/issues/` (origin-doc-update in use), also run this — it
cross-checks origin-doc-update's issue `branch` fields against real git state so
Stage 2 doesn't have to guess by name alone:

```bash
python3 ~/.agents/skills/origin-git-cleanup/scripts/check_active_issue_branches.py
```

It reports which active issues' recorded branch is missing (a signal that
resuming the issue needs a fresh branch, or that the branch was cleaned up
without archiving the issue), and which `ISSUE-*`-shaped branches have no
active issue referencing them (orphan candidates). It no-ops cleanly if
`docs/issues/` doesn't exist or the repo isn't using origin-doc-update.

## Stage 2 — Classify and decide

Work through each branch/worktree and assign exactly one disposition:

| Disposition   | Meaning                                                            |
| ------------- | ------------------------------------------------------------------ |
| `integrate`   | Contains useful work that should be committed/merged/cherry-picked |
| `delete`      | Confirmed obsolete, duplicate, generated residue, or superseded    |
| `keep`        | Intentional parallel work; leave it and record why                 |
| `investigate` | Not enough evidence; inspect diffs/PRs before any mutation         |

Do not leave a branch/worktree unclassified. Note the conclusion and evidence so
it can be presented in Stage 3.

### Uncommitted / untracked changes

Read the actual diff (`git diff`, `git diff --staged`, and look at untracked
files). Also check the **modification dates** of untracked files (`ls -la` or
`stat`). Files with old dates (weeks or months ago) may be forgotten artifacts —
flag them separately so the user can decide whether to commit, archive, or delete
them rather than silently folding them into the current task's commit.

Then decide:

- **Part of the finished task** → propose committing them. Draft a message that
  includes the subject line (repo style) **and** a multi-line body explaining
  the background: what was broken or missing, what was investigated, what approach
  was chosen, and what was verified. This context helps future readers — human or
  AI — reconstruct the reasoning without having to dig through Slack or PR
  comments.
- **Unrelated or half-done** → propose `git stash` (with a label) or simply
  leaving them in place. Do **not** fold unrelated edits into the task's commit.
- **Generated junk / secrets / large artifacts** → flag it; don't commit.

When intent is unclear, ask which bucket the changes fall into.

### Current branch — the usual cleanup target, and the easiest to miss

You are almost always running this skill **while checked out on the very branch
you just finished**. That branch is normally the main thing to remove, but two
traps make cleanup silently leave it behind — locally and on the remote:

- **A squash-merged branch looks unmerged locally.** After a squash (or
  cherry-pick), the branch's commits are not ancestors of your local main, so
  `git branch --merged main` omits it and the survey may list it under "not
  merged / likely active work." Don't trust that label for the branch you're on:
  confirm via its PR state (`gh pr view`) or patch-equivalence
  (`git diff main..HEAD` is empty → already in main). If it's merged, it's a
  `delete`, not `keep`.
- **You cannot delete the branch you're standing on.** `git branch -d <current>`
  fails while it's checked out. The branch must be deleted **after** the root
  worktree is switched back to main (Stage 4 enforces this order).

Classify it explicitly:

- Merged (by ancestry, squash, or cherry-pick)? → `delete`. Plan to switch to
  main first, then delete the local branch **and** its remote/tracking ref.
- Has unpushed commits / no PR yet? → `integrate` (merge it first, then delete).
- Is it main itself? → there's no feature branch to merge; focus on syncing and
  pruning.

### Patch-equivalence and obsolete branch checks

Use these when a branch is not ancestry-merged but looks already handled:

- Compare patch identity for a suspected squash/duplicate commit:

  ```bash
  git show <branch-commit> -- <paths> | git patch-id --stable
  git show <main-commit> -- <paths> | git patch-id --stable
  ```

- Compare the actual payload before deleting:

  ```bash
  git show --stat --name-status <branch>
  git diff --stat main..<branch>
  git log --oneline --left-right main...<branch>
  ```

- If the useful work was cherry-picked or manually reimplemented on main, mark
  the old branch `delete (superseded by <main-sha>)`. This is one of the few
  valid cases for `git branch -D`, but only after explicit user approval.

### Merge strategy (auto-detect)

- **GitHub remote present** → `gh` PR + **squash** merge. This gives a numbered,
  reviewable record, lets CI gate the merge, keeps main as a clean one-commit-
  per-feature history, and makes reverts trivial.
- **No remote (local-only repo)** → local **fast-forward** merge into main.
- Cross-check against recent history: if the repo clearly uses merge commits
  rather than squash, follow what the repo already does.

### Branch deletion candidates

A local branch is a candidate **only if all** hold:

- it appears under `git branch --merged main`, **and**
- it is not a protected branch (`main`/`master`/`develop`/`release/*`), **and**
- it doesn't look like active parallel work (see below).

Delete with `git branch -d` (safe; refuses if not merged). Never `-D`.

Note the common case: after a **squash** merge, the feature branch is merged on
the remote but git's local `--merged` check (against your local HEAD) may still
warn it's "not fully merged." Once main is updated from origin, deleting it is
safe — `gh pr merge --delete-branch` handles the remote side, and `git branch -d`
works locally after main is synced. If `-d` still refuses for a branch you've
confirmed is squash-merged via its PR, surface that to the user and let them
decide; don't reach for `-D` on your own.

### Worktrees

Remove a worktree only if it was for a branch you just merged/deleted **and** it
is clean. **Never remove a worktree with uncommitted changes** — report it
instead.

Special cases:

- A repository root worktree cannot be removed. If it is on a doomed branch, first
  move another worktree off `main` if needed, then checkout `main` in the root.
- A detached clean worktree may be removed when it points to an obsolete
  temporary deployment or inspection commit and has no uncommitted files.
- Dirty worktrees require a decision: commit, stash, keep, or discard. Discard
  (`git checkout -f`, `rm`, `git worktree remove --force`) only with explicit
  user approval and after listing the files.

### Protecting parallel work

A branch matching an active issue's `branch` field (per
`check_active_issue_branches.py`'s OK list) is `keep` unconditionally — this
is a decisive signal, not a guess, and overrides any name-based heuristic
below. Reason: "active issue `<id>`".

Otherwise, leave a branch/worktree alone if any of these are true:

- its name is unrelated to the finished task,
- it has recent commits you didn't make as part of this work,
- it has an open PR,
- it's checked out in another worktree,
- you're simply not sure.

An orphan-candidate branch from the check script (looks issue-shaped, no
active issue references it) is not an automatic `delete` — treat it as
`investigate` input alongside the usual patch-equivalence checks.

Surface what you're leaving and why, so the user can override if they want.

### Deletion protocol for destructive cleanup

`git branch -D`, `git worktree remove --force`, `rm` of untracked files, and
remote branch deletion are allowed only when all are true:

1. The target has been inspected (`status`, `diff`/`show`, and PR/remote state
   where relevant).
2. The target is classified as `delete`.
3. The user has approved deletion or explicitly instructed deletion.
4. The final report names what was deleted and why.

Never use `git reset --hard` or `git clean -fd` as a shortcut for this protocol.

## Stage 3 — Propose the plan and isolate destructive cleanup

Present a single structured summary. Execute its commit, push, PR creation, and
green PR merge steps autonomously. If it includes branch/worktree/ref deletion,
separate those steps and request one confirmation before performing only those
actions:

```
## origin-git-cleanup plan

Will do autonomously:
- Commit <files> on <branch>  — "<message>"  (reason)
- Push <branch> and merge via <PR+squash | local ff> into main
- Switch to main and sync with origin/main

Requires confirmation before deletion:
- Delete merged branch: <name>  (merged via PR #N)
- Remove worktree: <path>  (clean, branch merged)
- Prune stale remote-tracking refs

Will leave alone:
- Branch <name>  — keep: unrelated active work, has open PR #M
- Worktree <path>  — investigate: has uncommitted changes
- Stash <ref>  — unrelated
```

If the user wants changes, adjust and re-present. Don't execute anything
destructive until they approve.

## Stage 4 — Execute (safe order, stop on first error)

1. Commit or stash uncommitted changes per the decision.
2. Push the current branch if it has unpushed commits.
3. Merge into main:
   - GitHub: `gh pr create …` then `gh pr merge <N> --squash --delete-branch`.
   - Local: `git checkout main && git merge --ff-only <branch>`.
4. Sync main: `git checkout main && git pull --ff-only`.
5. Return the repository root worktree to main if it is not already there. **Do
   this before any branch deletion** — you cannot delete the branch you're on, so
   the just-finished current branch stays undeletable until you've switched away.
6. Delete merged local branches: `git branch -d <name>`. This now includes the
   branch you started on, once you've switched to main. If `-d` refuses a branch
   you've confirmed squash-merged via its PR, surface it (don't reach for `-D`).
   `gh pr merge --delete-branch` already removed its remote counterpart; if the
   branch was merged without that flag, delete the remote ref in step 10.
7. Delete superseded branches with `git branch -D <name>` only when approved.
8. Remove approved clean worktrees: `git worktree remove <path>`.
9. Remove approved dirty/superseded worktrees with `git worktree remove --force`
   only after file-level inspection and approval.
10. Delete approved remote branches with `git push origin --delete <name>`.
11. Prune stale remote refs: `git remote prune origin`.

If a step errors, stop and report — don't push past failures.

If a deleted/merged branch matched an active issue's `branch` field, add a
line to the final report suggesting the issue be archived — don't archive it
yourself; whether specs/guides/index updates are complete is origin-doc-update's
call, not this skill's:

```
- Branch <name> deleted (issue <id>) — if the work is complete, archive it:
  python3 ~/.agents/skills/origin-doc-update/scripts/archive_issue.py <id>
```

## Stage 5 — Verify

Confirm and report the end state:

- `git status` is clean,
- repository root is on `main` (or the integration branch),
- `git rev-parse main` == `git rev-parse origin/main` (or report ahead/behind),
- `git worktree list` contains only intentional worktrees,
- `git branch -vv` contains only intentional branches,
- every remaining branch/worktree is listed with its disposition and reason.

You should finish **on main, clean, and synced** — ready to branch off for the
next task.

---

## Quick reference: command safety

| Safe (reversible)                                                                                | Needs care (in plan)                                                                                | Forbidden without explicit ask                                           |
| ------------------------------------------------------------------------------------------------ | --------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------ |
| `git status`, `git branch`, `git log`, `git diff`, `survey.sh`, `check_active_issue_branches.py` | `git commit`, `git push`, `gh pr merge`, `git branch -d`, `git worktree remove`, `git remote prune` | `git branch -D`, `git push --force`, `git reset --hard`, `git clean -fd` |
