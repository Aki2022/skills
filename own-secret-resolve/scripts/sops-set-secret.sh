#!/usr/bin/env bash
# sops で暗号化した JSON（既定は cwd の secrets.enc.json）に、秘密を 1 つ足す・置き換える。
# 値は標準入力から読む。引数・履歴・ps・画面には出さない。
#
#   bw get password '<item>' | bash sops-set-secret.sh KEY_NAME
#   EXPECT_PREFIX=sk-ant- bash sops-set-secret.sh ANTHROPIC_API_KEY   # 端末なら実行後に貼り付ける（表示しない）
#   bash sops-set-secret.sh --file path/to/secrets.enc.json KEY_NAME
#
# 元は seedion の scripts/sops-set-secret.sh。空の値を拒むのは 2026-10-15 の事故による。bw が項目を
# 見つけられず何も出さなかったのにパイプラインが止まらず、sops が空の文字列を暗号化されない形で書き、
# mise が復号に失敗した。
set -euo pipefail

usage() { echo "usage: <value on stdin> | $0 [--file FILE] KEY_NAME  (KEY_NAME: upper-case letters, digits, _)" >&2; exit 2; }

file="secrets.enc.json"
if [[ "${1:-}" == "--file" ]]; then
  [[ -n "${2:-}" ]] || usage
  file="$2"; shift 2
fi
key="${1:-}"
[[ "$key" =~ ^[A-Z][A-Z0-9_]*$ ]] || usage
if [[ ! -f "$file" ]]; then
  echo "ERROR: $file not found; run from the repository root or pass --file" >&2
  exit 1
fi

if [[ -t 0 ]]; then
  # 端末から直接: 実行してから貼り付ける（表示しない）。クリップボードの取り違えを避けられる。
  read -rs -p "paste the value for $key, then press Enter (input is hidden): " value
  echo >&2
else
  value="$(cat)"
fi
value="${value%$'\n'}"
if [[ -z "$value" ]]; then
  echo "ERROR: empty value on stdin; nothing written (did the upstream command fail?)" >&2
  exit 1
fi
# 鍵・トークンに空白は入らない。クリップボードに鍵でない文章が入ったまま書かれたことがある。
# 空白を含む秘密を入れるときだけ ALLOW_WHITESPACE=1。
if [[ "$value" =~ [[:space:]] && "${ALLOW_WHITESPACE:-}" != "1" ]]; then
  echo "ERROR: value contains whitespace (${#value} chars); nothing written. Is the clipboard holding the key?" >&2
  exit 1
fi
if [[ -n "${EXPECT_PREFIX:-}" && "$value" != "$EXPECT_PREFIX"* ]]; then
  echo "ERROR: value does not start with the expected prefix; nothing written" >&2
  exit 1
fi

if command -v sops >/dev/null 2>&1; then sops=(sops); else sops=(nix shell nixpkgs#sops -c sops); fi
export SOPS_AGE_KEY_FILE="${SOPS_AGE_KEY_FILE:-${MISE_SOPS_AGE_KEY_FILE:-$HOME/.config/sops/age/keys.txt}}"

printf '%s' "$value" | jq -Rs . | tr -d '\n' | "${sops[@]}" set --value-stdin "$file" "[\"$key\"]"
echo "set $key in $file (value not shown)"
