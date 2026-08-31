#!/usr/bin/env bash
# check_entry_format.sh — deterministic checks for new origin-trouble-log entries.

set -u

FAIL=0
error() {
  FAIL=1
  printf 'FAIL %s: %s\n' "$1" "$2"
}

section_body() {
  local heading=$1
  local file=$2
  awk -v wanted="$heading" '
    $0 == wanted { on=1; next }
    on && /^## / { exit }
    on { print }
  ' "$file"
}

check_file() {
  local file=$1
  [ -f "$file" ] || { error "$file" "file not found"; return; }

  local fm
  fm=$(awk 'NR == 1 && $0 != "---" { exit } NR > 1 && $0 == "---" { exit } NR > 1 { print }' "$file")
  [ "$(sed -n '1p' "$file")" = "---" ] || error "$file" "frontmatter must start with ---"

  local key
  for key in date summary skills canon paths; do
    printf '%s\n' "$fm" | grep -Eq "^${key}:" \
      || error "$file" "frontmatter missing ${key}:"
  done

  if printf '%s\n' "$fm" | grep -Eq '/(Users|home)/[^/[:space:]]+/'; then
    error "$file" "frontmatter contains a user-named absolute path"
  elif grep -Eq '/(Users|home)/[^/[:space:]]+/' "$file"; then
    error "$file" "entry contains a user-named absolute path"
  fi

  local heading count body
  while IFS= read -r heading; do
    count=$(grep -Fxc "$heading" "$file" || true)
    [ "$count" -eq 1 ] || error "$file" "required heading appears ${count} times: $heading"
  done <<'HEADINGS'
## 意図 — 何をしようとしていたか
## 実際にやったこと — 実行したコマンド・書いたコードを実文で
## 観測された壊れ方 — ログ / エラーを実文で。実害を含む
## 気づいた経緯 — 誰が / 何がどうやって気づいたか
## 既に存在していた正解 — 具体パス。「無かった」「不明」も可
## 本来どうすべきだったか（仮説）
## なぜそうしなかったか（仮説）
HEADINGS

  for heading in \
    '## 実際にやったこと — 実行したコマンド・書いたコードを実文で' \
    '## 観測された壊れ方 — ログ / エラーを実文で。実害を含む'; do
    body=$(section_body "$heading" "$file")
    printf '%s\n' "$body" | grep -Fq '```' \
      || error "$file" "section has no verbatim code block: $heading"
  done
}

if [ "$#" -eq 0 ]; then
  printf 'usage: %s ENTRY.md [...]\n' "$0" >&2
  exit 2
fi

for file in "$@"; do
  check_file "$file"
done

if [ "$FAIL" -eq 0 ]; then
  printf 'OK: entry format checks passed\n'
fi
exit "$FAIL"
