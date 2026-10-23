#!/usr/bin/env bash
# Tests for sops-set-secret.sh. Uses a throwaway age key, so decrypting here
# exposes nothing real. Run: bash scripts/sops-set-secret.test.sh
set -uo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SET="$HERE/sops-set-secret.sh"
work="$(mktemp -d)"; trap 'rm -rf "$work"' EXIT
age-keygen -o "$work/key.txt" 2>/dev/null
pub="$(age-keygen -y "$work/key.txt")"
export SOPS_AGE_KEY_FILE="$work/key.txt"
mkdir -p "$work/repo"
printf '{"EXISTING":"keep"}' | sops --encrypt --age "$pub" --input-type json --output-type json /dev/stdin > "$work/repo/secrets.enc.json"
fail=0
check() { # check <label> <want-rc> <got-rc>
  if [[ "$2" == "$3" ]]; then printf '  ok    %s\n' "$1"; else printf '  FAIL  %s (want rc=%s got rc=%s)\n' "$1" "$2" "$3"; fail=1; fi
}
value_of() { sops --decrypt --extract "[\"$1\"]" "$2" 2>/dev/null; }
digest() { shasum "$work/repo/secrets.enc.json" | cut -d' ' -f1; }

(cd "$work/repo" && printf 'sk-test123' | bash "$SET" NEW_KEY >/dev/null 2>&1); check "default file is cwd secrets.enc.json" 0 $?
[[ "$(value_of NEW_KEY "$work/repo/secrets.enc.json")" == "sk-test123" ]]; check "value written" 0 $?
[[ "$(value_of EXISTING "$work/repo/secrets.enc.json")" == "keep" ]]; check "other keys kept" 0 $?
grep -q 'sk-test123' "$work/repo/secrets.enc.json"; check "value stored encrypted" 1 $?

cp "$work/repo/secrets.enc.json" "$work/other.enc.json"
printf 'v2' | bash "$SET" --file "$work/other.enc.json" OTHER >/dev/null 2>&1; check "--file target" 0 $?
[[ "$(value_of OTHER "$work/other.enc.json")" == "v2" ]]; check "--file value written" 0 $?

before="$(digest)"
(cd "$work/repo" && printf '' | bash "$SET" EMPTY >/dev/null 2>&1); check "empty value refused" 1 $?
(cd "$work/repo" && printf 'two words' | bash "$SET" WS >/dev/null 2>&1); check "whitespace refused" 1 $?
(cd "$work/repo" && printf 'xx-abc' | EXPECT_PREFIX=sk- bash "$SET" PFX >/dev/null 2>&1); check "prefix mismatch refused" 1 $?
(cd "$work/repo" && printf 'v' | bash "$SET" bad-name >/dev/null 2>&1); check "bad key name refused" 2 $?
[[ "$(digest)" == "$before" ]]; check "refusals leave the file unchanged" 0 $?
(cd "$work" && printf 'v' | bash "$SET" NOFILE >/dev/null 2>&1); check "missing file refused" 1 $?
out="$(cd "$work/repo" && printf 'sk-secretvalue' | bash "$SET" SHOWN 2>&1)"
[[ "$out" != *sk-secretvalue* ]]; check "value never printed" 0 $?

echo; [[ $fail -eq 0 ]] && echo "RESULT: all expectations met" || echo "RESULT: SOME EXPECTATIONS FAILED"
exit $fail
