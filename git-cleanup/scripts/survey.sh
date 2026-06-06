#!/usr/bin/env bash
#
# git-cleanup survey: print a read-only snapshot of repository state.
#
# This script NEVER writes. It only reads, so it is always safe to run first
# to understand a repo before proposing any cleanup actions. If you find
# yourself wanting to add a command that mutates state (commit, push, branch
# delete, worktree remove, fetch, pull, prune), it does NOT belong here.
#
# Usage: survey.sh [main_branch]
#   main_branch defaults to "main", falling back to "master" if main is absent.

set -uo pipefail

section() { printf '\n=== %s ===\n' "$1"; }

# Resolve the integration branch name.
main_branch="${1:-}"
if [[ -z "$main_branch" ]]; then
  if git show-ref --verify --quiet refs/heads/main; then
    main_branch="main"
  elif git show-ref --verify --quiet refs/heads/master; then
    main_branch="master"
  else
    main_branch="main"
  fi
fi

if ! git rev-parse --is-inside-work-tree >/dev/null 2>&1; then
  echo "Not inside a git work tree. Aborting survey." >&2
  exit 1
fi

section "REPO"
echo "toplevel: $(git rev-parse --show-toplevel)"
echo "current branch: $(git rev-parse --abbrev-ref HEAD)"
echo "integration branch (assumed): $main_branch"

section "REMOTE"
if git remote -v | grep -q .; then
  git remote -v
  # Detect GitHub specifically (affects merge strategy: PR+squash vs local ff).
  if git remote -v | grep -qiE 'github\.com'; then
    echo "github_remote: yes"
  else
    echo "github_remote: no (non-GitHub remote)"
  fi
else
  echo "github_remote: no (no remote configured)"
fi

section "WORKING TREE (uncommitted / untracked)"
if git status --porcelain | grep -q .; then
  git status --short
else
  echo "(clean)"
fi

section "UPSTREAM AHEAD/BEHIND"
git status --short --branch | head -1

section "STASHES"
git stash list || true
[[ -z "$(git stash list 2>/dev/null)" ]] && echo "(none)"

section "LOCAL BRANCHES (with upstream tracking)"
git branch -vv

section "ALL BRANCHES (incl. remotes)"
git branch -a

section "BRANCHES MERGED INTO $main_branch (deletion candidates, excl. protected)"
# git marks the current branch with '*' and worktree-checked-out branches with
# '+'. Drop '+' lines entirely — a branch checked out in another worktree is
# active work and git won't let you delete it anyway. Strip '* ' / '  ' markers
# from the rest, then exclude protected branches.
git branch --merged "$main_branch" 2>/dev/null \
  | grep -v '^+ ' \
  | sed 's/^[* ] //' \
  | grep -vE "^(${main_branch}|master|develop|release/.*)$" \
  || true

section "BRANCHES CHECKED OUT IN A WORKTREE (KEEP — git refuses to delete these)"
git branch --merged "$main_branch" 2>/dev/null | grep '^+ ' | sed 's/^+ //' || true
git branch --no-merged "$main_branch" 2>/dev/null | grep '^+ ' | sed 's/^+ //' || true

section "BRANCHES NOT MERGED INTO $main_branch (KEEP — likely active work)"
git branch --no-merged "$main_branch" 2>/dev/null | grep -v '^+ ' | sed 's/^[* ] //' || true

section "WORKTREES"
git worktree list

section "RECENT HISTORY (merge-style hint: squash vs merge commits)"
git log --graph --oneline -15

section "OPEN PULL REQUESTS (gh)"
if command -v gh >/dev/null 2>&1 && git remote -v | grep -qiE 'github\.com'; then
  gh pr status 2>/dev/null || echo "(gh available but pr status failed — check auth)"
else
  echo "(skipped: gh not available or no GitHub remote)"
fi

printf '\n=== END SURVEY ===\n'
