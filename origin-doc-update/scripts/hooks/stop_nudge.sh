#!/usr/bin/env bash
# origin-doc-update :: Stop hook (Claude Code / Codex shared shim)
#
# Best-effort, NON-BLOCKING reminder: if files changed this session but nothing
# under docs/ was touched, nudge the agent to classify guide impact.
# Read-only `git status` only — this is NOT a git hook and never blocks a commit.
# Contract: must never block (no continue:false), must never error, silent when
# there is nothing to say.
set -eu

input="$(cat 2>/dev/null || true)"
cwd=""
if [ -n "${input}" ] && command -v python3 >/dev/null 2>&1; then
  cwd="$(printf '%s' "${input}" \
    | python3 -c 'import sys,json;print(json.load(sys.stdin).get("cwd","") or "")' \
    2>/dev/null || true)"
fi
[ -n "${cwd}" ] || cwd="${PWD}"

# Only repos that use the doc-governance system.
[ -f "${cwd%/}/docs/00_index.md" ] || exit 0
command -v git >/dev/null 2>&1 || exit 0
git -C "${cwd}" rev-parse --is-inside-work-tree >/dev/null 2>&1 || exit 0

status="$(git -C "${cwd}" status --porcelain 2>/dev/null || true)"
[ -n "${status}" ] || exit 0

# Crude path extraction is acceptable for a nudge: strip the 3-char status
# prefix, and for renames ("old -> new") keep the new path.
paths="$(printf '%s\n' "${status}" | sed 's/^...//' | sed 's/.* -> //')"
nondocs="$(printf '%s\n' "${paths}" | grep -v '^docs/' || true)"
docschanged="$(printf '%s\n' "${paths}" | grep '^docs/' || true)"

if [ -n "${nondocs}" ] && [ -z "${docschanged}" ]; then
  # Emit JSON systemMessage: surfaced as a non-blocking UI warning (Codex).
  # No "continue":false, so the turn still ends normally — never blocks.
  printf '%s\n' '{"systemMessage":"origin-doc-update nudge: non-doc files changed. Update the active workstream/issue, classify guide impact as required or none, and update docs/guides/ in the same slice when behavior changed."}'
fi
exit 0
