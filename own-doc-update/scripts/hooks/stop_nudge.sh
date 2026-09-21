#!/usr/bin/env bash
# own-doc-update :: Stop hook (Claude Code / Codex shared shim)
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

msg=""
if [ -n "${nondocs}" ] && [ -z "${docschanged}" ]; then
  msg="own-doc-update nudge: non-doc files changed. Update the active workstream/issue, classify guide impact as required or none, and update docs/guides/ in the same slice when behavior changed."
fi

# 2026-09-21: when docs/ changed this session, run the validator and say what is
# red. The validator was only a close-session step, so a session that never
# reached close-session left red docs that nobody saw. Syntactic checks only, so
# warn-only is the right strength (ADR-20260816); never blocks, capped output.
if [ -n "${docschanged}" ] && command -v python3 >/dev/null 2>&1; then
  validator=""
  for candidate in \
    "$(cd "$(dirname "$0")" 2>/dev/null && pwd)/../validate_repo_docs.py" \
    "${HOME}/.agents/skills/own-doc-update/scripts/validate_repo_docs.py"; do
    [ -f "${candidate}" ] && { validator="${candidate}"; break; }
  done
  if [ -n "${validator}" ]; then
    out="$(timeout 20 python3 "${validator}" "${cwd}" 2>/dev/null || true)"
    errs="$(printf '%s\n' "${out}" | grep -c '✗' || true)"
    if [ "${errs:-0}" -gt 0 ]; then
      first="$(printf '%s\n' "${out}" | grep '✗' | head -5 | sed 's/^[[:space:]]*✗ //' | tr '\n' ' ' | cut -c1-600)"
      msg="${msg}${msg:+ | }docs validator: ${errs} error(s) after this session's docs changes — ${first}"
    fi
  fi
fi

if [ -n "${msg}" ]; then
  # Emit JSON systemMessage: surfaced as a non-blocking UI warning.
  # No "continue":false, so the turn still ends normally — never blocks.
  python3 -c 'import json,sys; print(json.dumps({"systemMessage": sys.argv[1]}, ensure_ascii=False))' "${msg}"
fi
exit 0
