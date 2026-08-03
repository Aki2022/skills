#!/usr/bin/env bash
#
# origin-git-cleanup survey: print a read-only snapshot of repository state.
#
# This script NEVER writes. It only reads, so it is always safe to run first
# to understand a repo before proposing any cleanup actions. If you find
# yourself wanting to add a command that mutates state (commit, push, branch
# delete, worktree remove, fetch, pull, prune), it does NOT belong here.
#
# Usage: survey.sh [--repo <path>] [main_branch]
#   --repo      the repository to survey. Defaults to the current directory,
#               which is only right when the caller is standing in the target.
#               An orchestrator invoked *against* a repository (origin-ws-loop,
#               origin-close-session) is not, so it must pass this.
#   main_branch defaults to "main", falling back to "master" if main is absent.

set -uo pipefail

repo=""
args=()
while [[ $# -gt 0 ]]; do
  case "$1" in
  --repo)
    [[ $# -ge 2 ]] || {
      echo "--repo needs a path" >&2
      exit 2
    }
    repo="$2"
    shift 2
    ;;
  --repo=*)
    repo="${1#--repo=}"
    shift
    ;;
  *)
    args+=("$1")
    shift
    ;;
  esac
done
set -- ${args[@]+"${args[@]}"}

if [[ -n "$repo" ]]; then
  [[ -d "$repo" ]] || {
    echo "--repo path does not exist: $repo" >&2
    exit 2
  }
  cd "$repo" || exit 2
  git rev-parse --is-inside-work-tree >/dev/null 2>&1 || {
    echo "--repo path is not a git work tree: $repo" >&2
    exit 2
  }
fi

section() { printf '\n=== %s ===\n' "$1"; }
subsection() { printf '\n--- %s ---\n' "$1"; }

print_worktree_status() {
  local path="$1"
  local branch="$2"
  local head="$3"
  local marker="$branch"
  [[ -z "$marker" ]] && marker="detached HEAD"

  subsection "$path [$marker]"
  echo "head: ${head:-unknown}"
  (
    cd "$path" || exit
    git status --short --branch
    if git status --porcelain | grep -q .; then
      echo "dirty: yes"
      echo "untracked_count: $(git status --porcelain | grep -c '^??' || true)"
    else
      echo "dirty: no"
      echo "untracked_count: 0"
    fi
  )
}

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
if [[ -z "$repo" ]]; then
  # Surveying the wrong repository is not a harmless mistake: this output is the
  # input to the deletion plan. Say plainly which repository was read and how it
  # was chosen, so a caller that meant a different one notices here.
  echo "target_source: current directory (no --repo given)"
  echo "WARNING: verify this is the repository you meant to survey. An orchestrator invoked against a repository must pass --repo <path>; its own cwd is a different repository."
else
  echo "target_source: --repo $repo"
fi
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

section "ROOT WORKTREE COMPLETION CHECK"
root_path="$(git rev-parse --show-toplevel)"
root_branch="$(git -C "$root_path" rev-parse --abbrev-ref HEAD)"
echo "root_path: $root_path"
echo "root_branch: $root_branch"
if [[ "$root_branch" == "$main_branch" ]] && ! git -C "$root_path" status --porcelain | grep -q .; then
  echo "root_ready: yes"
else
  echo "root_ready: no (cleanup must finish with root on $main_branch and clean)"
fi

section "CURRENT BRANCH (the usual cleanup target — you are standing ON it)"
# The branch you're checked out on is almost always the thing you just finished
# and want gone. Two traps make cleanup silently skip it:
#   1. A squash-merged branch is NOT an ancestor of local main (different SHA),
#      so `git branch --merged main` omits it and it lands in the "NOT MERGED /
#      KEEP" list below — mislabeled as active work. Patch-equivalence
#      (`git diff main..HEAD` empty) reveals it's actually done.
#   2. You cannot `git branch -d` the branch you're on. It must be deleted AFTER
#      switching the root worktree back to main.
cur_branch="$(git rev-parse --abbrev-ref HEAD)"
echo "current_branch: $cur_branch"
if [[ "$cur_branch" == "$main_branch" ]]; then
  echo "status: on integration branch — no feature branch to delete from here"
else
  ahead="$(git rev-list --count "$main_branch..HEAD" 2>/dev/null || echo '?')"
  behind="$(git rev-list --count "HEAD..$main_branch" 2>/dev/null || echo '?')"
  echo "commits ahead of $main_branch: $ahead"
  echo "commits behind $main_branch: $behind"
  # A branch with no commits of its own is an ancestor of main by definition,
  # which made this report "MERGED — delete it" for a freshly cut branch whose
  # work was still uncommitted in the working tree. That is the most dangerous
  # thing this survey can say, and it is the NORMAL state for any workflow that
  # edits files before surveying (origin-close-session prescribes exactly that
  # order). Check for it before the ancestry test.
  # Narrow: ONLY the dirty case is special. A clean branch with no commits of
  # its own is genuinely indistinguishable from a fast-forward-merged one, and
  # reporting it as unmerged would break the survey's most useful signal — so it
  # falls through to the ancestry test below, exactly as before.
  if [[ "$ahead" == "0" && -n "$(git status --porcelain 2>/dev/null)" ]]; then
    echo "merge_state: NO COMMITS YET, working tree dirty — this branch holds uncommitted work. Do NOT delete it; commit first."
  elif git merge-base --is-ancestor HEAD "$main_branch" 2>/dev/null; then
    echo "merge_state: MERGED into $main_branch (ancestry) — delete after switching to $main_branch"
  elif [[ -z "$(git diff "$main_branch"..HEAD 2>/dev/null)" ]]; then
    echo "merge_state: PATCH-EQUIVALENT to $main_branch (likely squash/cherry-pick merged) — verify via PR, then delete after switching to $main_branch"
  else
    echo "merge_state: has unmerged changes vs $main_branch — merge/integrate before deleting"
  fi
  echo "reminder: switch root worktree to $main_branch FIRST, then 'git branch -d $cur_branch' and delete its remote/tracking ref"
fi

section "UPSTREAM AHEAD/BEHIND"
git status --short --branch | head -1

section "STASHES"
git stash list || true
[[ -z "$(git stash list 2>/dev/null)" ]] && echo "(none)"

section "LOCAL BRANCHES (with upstream tracking)"
git branch -vv

section "LOCAL BRANCHES WITH GONE UPSTREAM"
git branch -vv | grep ': gone]' || echo "(none)"

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
# Exclude the current branch (marked '*'): a squash-merged branch you're standing
# on lands here too, but it's the cleanup target, not parallel work. It's already
# analyzed in the CURRENT BRANCH section above with a proper merge_state check.
git branch --no-merged "$main_branch" 2>/dev/null | grep -v '^+ ' | grep -v '^\* ' | sed 's/^  //' || true
if git branch --no-merged "$main_branch" 2>/dev/null | grep -q '^\* '; then
  echo "(current branch '$cur_branch' also shows as not-merged — see CURRENT BRANCH section; may be squash-merged, not active work)"
fi

section "WORKTREES"
git worktree list

section "WORKTREE STATUS DETAILS"
current_path=""
current_head=""
current_branch=""
while IFS= read -r line || [[ -n "$line" ]]; do
  if [[ "$line" == worktree\ * ]]; then
    if [[ -n "$current_path" ]]; then
      print_worktree_status "$current_path" "$current_branch" "$current_head"
    fi
    current_path="${line#worktree }"
    current_head=""
    current_branch=""
  elif [[ "$line" == HEAD\ * ]]; then
    current_head="${line#HEAD }"
  elif [[ "$line" == branch\ refs/heads/* ]]; then
    current_branch="${line#branch refs/heads/}"
  elif [[ -z "$line" ]]; then
    if [[ -n "$current_path" ]]; then
      print_worktree_status "$current_path" "$current_branch" "$current_head"
    fi
    current_path=""
    current_head=""
    current_branch=""
  fi
done < <(git worktree list --porcelain)
if [[ -n "$current_path" ]]; then
  print_worktree_status "$current_path" "$current_branch" "$current_head"
fi

section "RECENT HISTORY (merge-style hint: squash vs merge commits)"
git log --graph --oneline -15

section "OPEN PULL REQUESTS (gh)"
if ! git remote -v | grep -qiE 'github\.com'; then
  echo "(skipped: no GitHub remote)"
elif command -v gh >/dev/null 2>&1; then
  gh pr status 2>/dev/null || echo "(gh available but pr status failed — check auth)"
else
  echo "(skipped: gh not available on PATH)"
fi

section "CLASSIFICATION REMINDER"
cat <<'EOF'
Classify every remaining branch/worktree before mutating:
- integrate: useful work to merge/cherry-pick/commit
- delete: inspected obsolete/duplicate/generated/superseded residue
- keep: intentional parallel work with a reason
- investigate: not enough evidence yet

For squash/cherry-pick duplicates, compare patch IDs or file stats before delete.
Finish with the repository root on main, clean, and main == origin/main.
EOF

printf '\n=== END SURVEY ===\n'
