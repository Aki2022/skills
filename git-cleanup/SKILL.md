---
name: git-cleanup
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
  "PR マージ後の整理", "ブランチを消して". Use for end-of-task git housekeeping.
  Do NOT use for mid-task git operations like rebasing, cherry-picking, checking
  git status only, or committing individual files without cleanup intent.
---

# git-cleanup

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

**Survey (read-only) → Classify → Propose a plan → Get one approval → Execute → Verify.**

When anything is ambiguous, stop and ask rather than guess. Leaving something
alone is always safe; deleting the wrong thing is not.

## Cardinal rules

- **Read before you write.** Never mutate state before you've surveyed it.
- **Propose before you destroy.** Present the full plan and get approval once
  before any commit/push/merge/delete.
- **Don't guess about parallel work.** If a branch or worktree might belong to
  other active work, leave it and say so.
- **Prefer the reversible.** Use `git branch -d` (never `-D`); never
  `git push --force`, `git reset --hard`, or `git clean -fd` without an explicit
  instruction from the user.
- **Follow repo conventions.** Use the `gh` CLI for all GitHub operations. Match
  the repository's existing commit-message style.
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
bash ~/.claude/skills/git-cleanup/scripts/survey.sh
```

It reports: current branch, remote (and whether it's GitHub), uncommitted/
untracked files, upstream ahead/behind, stashes, local branches with tracking,
all branches, branches **merged** vs **not merged** into main, worktrees, recent
history (a hint at squash-vs-merge convention), and open PRs.

If the script isn't available, gather the same picture manually with
`git status`, `git branch -vv`, `git branch -a`, `git worktree list`,
`git stash list`, `git log --graph --oneline -15`,
`git branch --merged main` / `--no-merged main`, and `gh pr status`.

## Stage 2 — Classify and decide

Work through each decision. Note your conclusion and the reason — you'll present
these in Stage 3.

### Uncommitted / untracked changes

Read the actual diff (`git diff`, `git diff --staged`, and look at untracked
files). Then decide:

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

### Current branch

- Already merged into main? → it just needs deleting (Stage 4).
- Has unpushed commits / no PR yet? → it needs to be merged first.
- Is it main itself? → there's no feature branch to merge; focus on syncing and
  pruning.

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

### Protecting parallel work

Leave a branch/worktree alone if any of these are true:

- its name is unrelated to the finished task,
- it has recent commits you didn't make as part of this work,
- it has an open PR,
- it's checked out in another worktree,
- you're simply not sure.

Surface what you're leaving and why, so the user can override if they want.

## Stage 3 — Propose the plan, get one approval

Present a single structured summary and wait for one confirmation:

```
## git-cleanup plan

Will do:
- Commit <files> on <branch>  — "<message>"  (reason)
- Push <branch> and merge via <PR+squash | local ff> into main
- Switch to main and sync with origin/main
- Delete merged branch: <name>  (merged via PR #N)
- Remove worktree: <path>  (clean, branch merged)
- Prune stale remote-tracking refs

Will leave alone:
- Branch <name>  — unrelated active work, has open PR #M
- Worktree <path>  — has uncommitted changes
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
5. Delete merged local branches: `git branch -d <name>`.
6. Remove clean, merged worktrees: `git worktree remove <path>`.
7. Prune stale remote refs: `git remote prune origin`.

If a step errors, stop and report — don't push past failures.

## Stage 5 — Verify

Confirm and report the end state:

- `git status` is clean,
- `git rev-parse main` == `git rev-parse origin/main` (or report ahead/behind),
- list remaining branches/worktrees and confirm each is intentional.

You should finish **on main, clean, and synced** — ready to branch off for the
next task.

---

## Quick reference: command safety

| Safe (reversible)                                              | Needs care (in plan)                                                                                | Forbidden without explicit ask                                           |
| -------------------------------------------------------------- | --------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------ |
| `git status`, `git branch`, `git log`, `git diff`, `survey.sh` | `git commit`, `git push`, `gh pr merge`, `git branch -d`, `git worktree remove`, `git remote prune` | `git branch -D`, `git push --force`, `git reset --hard`, `git clean -fd` |
