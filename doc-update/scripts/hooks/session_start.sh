#!/usr/bin/env bash
# doc-update :: SessionStart hook (Claude Code / Codex shared shim)
#
# Injects docs/00_index.md into the agent context when the current repository
# uses the doc-governance system. Graceful no-op for every other repository.
# Contract: must be fast (<1s), must never error, must stay silent when there
# is nothing to inject (global hook fires in every session in every repo).
set -eu

# Drain stdin (hook JSON) and try to read cwd from it; fall back to $PWD.
input="$(cat 2>/dev/null || true)"
cwd=""
if [ -n "${input}" ] && command -v python3 >/dev/null 2>&1; then
  cwd="$(printf '%s' "${input}" \
    | python3 -c 'import sys,json;print(json.load(sys.stdin).get("cwd","") or "")' \
    2>/dev/null || true)"
fi
[ -n "${cwd}" ] || cwd="${PWD}"

index="${cwd%/}/docs/00_index.md"
[ -f "${index}" ] || exit 0

printf '%s\n' "Repository documentation index (read this first; do not scan all of docs/):"
printf '%s\n' "----- docs/00_index.md -----"
cat "${index}"
exit 0
