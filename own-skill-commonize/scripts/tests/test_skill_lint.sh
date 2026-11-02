#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR=$(cd "$(dirname "$0")" && pwd)
LINTER="$SCRIPT_DIR/../skill_lint.sh"
tmp_root=$(mktemp -d)
trap 'rm -rf "$tmp_root"' EXIT

write_skill() {
  local name="$1"
  mkdir -p "$tmp_root/$name"
  printf '%s\n' '---' "name: $name" 'description: fixture' '---' "$2" > "$tmp_root/$name/SKILL.md"
}

# Same-skill references remain relative to the current skill.
write_skill provider 'See `references/guide.md`.'
mkdir -p "$tmp_root/provider/references"
printf '%s\n' 'provider guide' > "$tmp_root/provider/references/guide.md"

# A named canonical skill followed by its relative resource is also cross-skill.
write_skill agents-sdk 'Agents SDK reference.'
mkdir -p "$tmp_root/agents-sdk/references"
printf '%s\n' 'MCP reference' > "$tmp_root/agents-sdk/references/mcp.md"
write_skill named-consumer 'Use `agents-sdk` skill, its `references/mcp.md` before continuing.'

# A canonical skill name followed by a relative resource is a cross-skill reference.
write_skill consumer 'Load `/provider` → `references/guide.md` before continuing.'
write_skill full-consumer 'Load `/provider/references/guide.md` before continuing.'

if ! output=$(bash "$LINTER" "$tmp_root" 2>&1); then
  printf '%s\n' "$output" >&2
  echo 'cross-skill reference should pass' >&2
  exit 1
fi

# Missing resources in the current skill must remain a hard failure.
write_skill same-missing 'See `references/missing.md` before continuing.'
set +e
output=$(bash "$LINTER" "$tmp_root" 2>&1)
status=$?
set -e
if [ "$status" -eq 0 ] || ! grep -q 'FAIL S4 same-missing' <<<"$output"; then
  printf '%s\n' "$output" >&2
  echo 'missing same-skill reference should fail' >&2
  exit 1
fi

# Missing resources in the referenced canonical skill must still fail closed.
write_skill broken 'Load `/provider` → `references/missing.md` before continuing.'
set +e
output=$(bash "$LINTER" "$tmp_root" 2>&1)
status=$?
set -e
if [ "$status" -eq 0 ] || ! grep -q 'FAIL S4 broken' <<<"$output"; then
  printf '%s\n' "$output" >&2
  echo 'missing cross-skill reference should fail' >&2
  exit 1
fi

# Traversal must never escape the referenced skill root.
write_skill unsafe 'Load `/provider` → `references/../secret.md` before continuing.'
set +e
output=$(bash "$LINTER" "$tmp_root" 2>&1)
status=$?
set -e
if [ "$status" -eq 0 ] || ! grep -q 'unsafe reference' <<<"$output"; then
  printf '%s\n' "$output" >&2
  echo 'path traversal should fail' >&2
  exit 1
fi

echo 'OK: skill_lint cross-skill references, missing resources, and traversal'
