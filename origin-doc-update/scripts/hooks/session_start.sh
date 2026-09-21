#!/usr/bin/env bash
# origin-doc-update :: SessionStart hook (Claude Code / Codex shared shim)
#
# Injects a routing digest of docs/00_index.md into the agent context when the
# current repository uses the doc-governance system. Graceful no-op for every
# other repository. Contract: must be fast (<1s), must never error, must stay
# silent when there is nothing to inject (global hook fires in every session in
# every repo).
#
# Why a digest and not the file (2026-09-19, measured over every session on one
# machine): Claude Code delivers a hook's output to the agent inline only up to
# about 10,000 CHARACTERS -- largest ever delivered inline 8,957, smallest ever
# spilled 10,019, and that spilled one was this very injection. Past the limit
# the output is written to a file and the agent is handed a 2 KB preview, so
# injecting the file whole made every "read this first" rule downstream run
# against nothing, with no failure reported anywhere. index_digest.py compresses
# the index to fit and states what it left out; the agent opens the file itself
# when it needs the full text.
set -eu

# Used only when the digest cannot be built (no python3, script missing).
# Deliberately a byte bound at the character limit: bytes >= characters, so it
# is conservative and can never over-deliver.
FALLBACK_MAX_BYTES=8000

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

hook_dir="$(cd "$(dirname "$0")" 2>/dev/null && pwd || true)"
digest_script=""
for candidate in \
  "${hook_dir:-.}/../index_digest.py" \
  "${HOME}/.agents/skills/origin-doc-update/scripts/index_digest.py"; do
  if [ -f "${candidate}" ]; then
    digest_script="${candidate}"
    break
  fi
done

digest=""
if [ -n "${digest_script}" ] && command -v python3 >/dev/null 2>&1; then
  digest="$(python3 "${digest_script}" "${index}" 2>/dev/null || true)"
fi

# Review reminder (2026-09-20): closing work keeps issues honest, but nothing
# re-weights guides/specs as the repository grows. docs_hygiene.py --review does
# that and leaves docs/log/review-YYYYMMDD.md; past 14 days, say so in one line.
review_note=""
newest="$(ls "${cwd%/}/docs/log/" 2>/dev/null | sed -n 's/^review-\([0-9]\{8\}\)\.md$/\1/p' | sort | tail -1)"
if [ -z "${newest}" ]; then
  review_note="docs review: never run (python3 ~/.agents/skills/origin-doc-update/scripts/docs_hygiene.py <repo> --review)"
else
  now_days=$(( $(date +%s) / 86400 ))
  then_days=$(( $(date -j -f %Y%m%d "${newest}" +%s 2>/dev/null || date -d "${newest}" +%s 2>/dev/null || echo 0) / 86400 ))
  if [ "${then_days}" -gt 0 ] && [ $(( now_days - then_days )) -gt 14 ]; then
    review_note="docs review: last ${newest}, over 14 days — run docs_hygiene.py --review"
  fi
fi

printf '%s\n' "Repository documentation index (read this first; do not scan all of docs/):"
[ -n "${review_note}" ] && printf '%s\n' "NOTE: ${review_note}"

# index_digest.py prints the marker line too: it decides whether the index fits
# as written or has to be compressed, and says which one it sent.
if [ -n "${digest}" ]; then
  printf '%s\n' "${digest}"
  exit 0
fi

# No digest available: send the head of the file rather than the whole of it,
# because the whole of it would arrive as a 2 KB preview of itself.
size="$(wc -c < "${index}" | tr -d ' ')"
if [ "${size}" -le "${FALLBACK_MAX_BYTES}" ]; then
  printf '%s\n' "----- docs/00_index.md -----"
  cat "${index}"
  exit 0
fi
printf '%s\n' "WARNING: index_digest.py did not run, so only the head of docs/00_index.md follows. Read the file for the rest."
printf '%s\n' "----- docs/00_index.md (truncated) -----"
head -c "${FALLBACK_MAX_BYTES}" "${index}"
printf '\n'
exit 0
