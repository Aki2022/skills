#!/usr/bin/env bash
#
# refresh_adc.sh — 複数リポジトリの per-repo ADC を一括再生成する横断ツール
#
# 背景:
#   gcloud/bq CLI は .mise.toml の CLOUDSDK_AUTH_IMPERSONATE_SERVICE_ACCOUNT を読むが、
#   Node/Python の Google SDK は読まない。SDK は ADC
#   (~/.config/gcloud/application_default_credentials.json) を見る。
#   グローバル ADC は単一ファイルなので複数リポジトリで衝突する。
#   そこでリポジトリごとに ADC ファイルを分離し、.mise.toml の
#   GOOGLE_APPLICATION_CREDENTIALS で cd 時に切り替える。
#
# 仕組み (ブラウザ再ログイン不要・テンプレート生成方式):
#   グローバル ADC に埋め込まれた source_credentials (authorized_user + refresh_token) を
#   共有し、各リポジトリの SA を impersonate する per-repo ADC を生成する。
#
# 使い方:
#   bash ~/.claude/skills/gcp-project-setup/refresh_adc.sh
#
# 再認証が必要になったとき (refresh_token が失効した場合のみ):
#   gcloud auth application-default login   # 素ログイン (impersonate指定なし)
#   bash ~/.claude/skills/gcp-project-setup/refresh_adc.sh
#
set -euo pipefail

GCLOUD_DIR="${CLOUDSDK_CONFIG:-$HOME/.config/gcloud}"
GLOBAL_ADC="${GCLOUD_DIR}/application_default_credentials.json"
CODE_DIR="${CODE_DIR:-$HOME/code}"

PYTHON="$(command -v python3 || true)"
[ -z "$PYTHON" ] && { echo "❌ python3 が見つからない"; exit 1; }
[ -f "$GLOBAL_ADC" ] || { echo "❌ グローバル ADC が無い: $GLOBAL_ADC"; echo "   先に: gcloud auth application-default login"; exit 1; }

echo "=== source_credentials を抽出中 ==="
# グローバル ADC から authorized_user の source を取り出す。
# 既に impersonated_service_account なら .source_credentials を、
# 素の authorized_user ならトップレベルを source として使う。
SOURCE_JSON="$("$PYTHON" - "$GLOBAL_ADC" <<'PY'
import json, sys
with open(sys.argv[1]) as f:
    d = json.load(f)
t = d.get("type")
if t == "impersonated_service_account":
    src = d.get("source_credentials")
    if not src:
        sys.exit("グローバル ADC に source_credentials が無い")
elif t == "authorized_user":
    src = {k: d[k] for k in d}
else:
    sys.exit(f"未対応の ADC type: {t}")
if src.get("type") != "authorized_user" or "refresh_token" not in src:
    sys.exit("source は authorized_user + refresh_token である必要がある")
print(json.dumps(src))
PY
)"
echo "  ✅ authorized_user source を取得"

# .mise.toml を走査して repo -> SA を収集
echo ""
echo "=== impersonation リポジトリを走査 (${CODE_DIR}/*/.mise.toml) ==="
shopt -s nullglob
GENERATED=""   # 改行区切りの "repo|sa|adc" レコード (bash 3.2 互換のため文字列で保持)
for MISE in "${CODE_DIR}"/*/.mise.toml; do
  REPO_PATH="$(dirname "$MISE")"
  REPO_NAME="$(basename "$REPO_PATH")"
  SA_LINE="$(grep -E '^[[:space:]]*CLOUDSDK_AUTH_IMPERSONATE_SERVICE_ACCOUNT' "$MISE" 2>/dev/null | head -1 || true)"
  SA_EMAIL="$(printf '%s' "$SA_LINE" | sed -E 's/.*=[[:space:]]*"?([^"]+)"?.*/\1/' | tr -d '[:space:]')"
  [ -z "$SA_EMAIL" ] && continue

  ADC_PATH="${GCLOUD_DIR}/${REPO_NAME}_adc.json"
  IMP_URL="https://iamcredentials.googleapis.com/v1/projects/-/serviceAccounts/${SA_EMAIL}:generateAccessToken"

  echo "  • ${REPO_NAME}  ->  ${SA_EMAIL}"

  # per-repo ADC を生成
  SOURCE_JSON="$SOURCE_JSON" IMP_URL="$IMP_URL" "$PYTHON" - "$ADC_PATH" <<'PY'
import json, os, sys
adc = {
    "delegates": [],
    "service_account_impersonation_url": os.environ["IMP_URL"],
    "source_credentials": json.loads(os.environ["SOURCE_JSON"]),
    "type": "impersonated_service_account",
}
with open(sys.argv[1], "w") as f:
    json.dump(adc, f, indent=2)
PY
  chmod 600 "$ADC_PATH"

  # .mise.toml に GOOGLE_APPLICATION_CREDENTIALS が無ければ追記
  if ! grep -qE '^[[:space:]]*GOOGLE_APPLICATION_CREDENTIALS' "$MISE"; then
    # [env] セクション直下に挿入。無ければ末尾に追記。
    if grep -qE '^\[env\]' "$MISE"; then
      "$PYTHON" - "$MISE" "$ADC_PATH" <<'PY'
import sys
mise, adc = sys.argv[1], sys.argv[2]
lines = open(mise).read().splitlines()
out, inserted = [], False
for ln in lines:
    out.append(ln)
    if not inserted and ln.strip() == "[env]":
        out.append(f'# SDK(Node/Python)用 ADC。CLIのIMPERSONATE変数はSDKに効かないため必須')
        out.append(f'GOOGLE_APPLICATION_CREDENTIALS = "{adc}"')
        inserted = True
if not inserted:
    out.append(f'GOOGLE_APPLICATION_CREDENTIALS = "{adc}"')
open(mise, "w").write("\n".join(out) + "\n")
PY
    else
      printf '\nGOOGLE_APPLICATION_CREDENTIALS = "%s"\n' "$ADC_PATH" >> "$MISE"
    fi
    echo "    ↳ .mise.toml に GOOGLE_APPLICATION_CREDENTIALS を追記"
  fi

  GENERATED="${GENERATED}${REPO_NAME}|${SA_EMAIL}|${ADC_PATH}"$'\n'
done
shopt -u nullglob

[ -z "$GENERATED" ] && echo "  (impersonation リポジトリなし)"

# グローバル ADC を素の authorized_user に戻す
echo ""
echo "=== グローバル ADC を素のユーザー認証へ戻す ==="
SOURCE_JSON="$SOURCE_JSON" "$PYTHON" - "$GLOBAL_ADC" <<'PY'
import json, os, sys
src = json.loads(os.environ["SOURCE_JSON"])
with open(sys.argv[1], "w") as f:
    json.dump(src, f, indent=2)
PY
chmod 600 "$GLOBAL_ADC"
echo "  ✅ グローバル ADC = authorized_user (impersonate なし)"

# 疎通確認
echo ""
echo "=== 疎通確認 (各 per-repo ADC で access token 取得) ==="
printf '%s' "$GENERATED" | while IFS='|' read -r REPO_NAME SA_EMAIL ADC_PATH; do
  [ -z "$REPO_NAME" ] && continue
  if GOOGLE_APPLICATION_CREDENTIALS="$ADC_PATH" \
     gcloud auth application-default print-access-token >/dev/null 2>&1; then
    echo "  ✅ ${REPO_NAME}: OK"
  else
    echo "  ⚠️  ${REPO_NAME}: トークン取得失敗 (IAM 反映待ち or TokenCreator 未付与)"
  fi
done

echo ""
echo "完了。各リポジトリで 'cd' すると mise が ADC を自動切替する。"
