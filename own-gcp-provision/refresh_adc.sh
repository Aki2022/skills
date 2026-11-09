#!/usr/bin/env bash
#
# refresh_adc.sh — 複数リポジトリの per-repo ADC を一括再生成する横断ツール
#
# 背景:
#   gcloud/bq CLI は mise 設定の CLOUDSDK_AUTH_IMPERSONATE_SERVICE_ACCOUNT を読むが、
#   Node/Python の Google SDK は読まない。SDK は ADC
#   (~/.config/gcloud/application_default_credentials.json) を見る。
#   グローバル ADC は単一ファイルなので複数リポジトリで衝突する。
#   そこでリポジトリごとに ADC ファイルを分離し、mise 設定の
#   GOOGLE_APPLICATION_CREDENTIALS で cd 時に切り替える。
#
# 走査対象: ${CODE_DIR}/*/.mise.toml と ${CODE_DIR}/*/mise.toml（ドットなし）の両方。
#   1 つの repo に両方あり、別々の SA を指すときは何も書かずに exit 1 で止まる。
#   同じ SA なら 1 件として扱い、GOOGLE_APPLICATION_CREDENTIALS がどちらにも無ければ
#   SA を書いている方のファイルに追記する。
#
# 仕組み (ブラウザ再ログイン不要・テンプレート生成方式):
#   グローバル ADC に埋め込まれた source_credentials (authorized_user + refresh_token) を
#   共有し、各リポジトリの SA を impersonate する per-repo ADC を生成する。
#
# 使い方:
#   通常は各 repo から `mise run reauth`
#   共通フローを直接呼ぶ場合のみ:
#   bash ~/.agents/skills/own-gcp-provision/refresh_adc.sh
#   対象を一覧するだけ（ADC も設定ファイルも書かない。gcloud も source ADC も要らない）:
#   bash ~/.agents/skills/own-gcp-provision/refresh_adc.sh --list
#   出力は 1 行 1 repo、タブ区切りで「repo名<TAB>SA<TAB>SA を書いている設定ファイル」
#
# 再認証が必要になったとき (refresh_token が失効した場合のみ):
#   各 repo から `mise run reauth`
#   もしくは `bash ~/.agents/skills/own-gcp-provision/re-auth.sh`
#
# active gcloud config に auth/impersonate_service_account が永続設定されている場合、
# 上の素ログインでも impersonated_service_account ADC が作られることがある。
# その場合は re-auth.sh を使うか、一時 CLOUDSDK_CONFIG を分離して authorized_user ADC を作る。
#
set -euo pipefail

LIST_ONLY=0
for arg in "$@"; do
  case "$arg" in
    --list) LIST_ONLY=1 ;;
    *) echo "使い方: refresh_adc.sh [--list]" >&2; exit 2 ;;
  esac
done

MISE_FILES=(.mise.toml mise.toml)

sa_of() {
  grep -E '^[[:space:]]*CLOUDSDK_AUTH_IMPERSONATE_SERVICE_ACCOUNT' "$1" 2>/dev/null | head -1 \
    | sed -E 's/.*=[[:space:]]*"?([^"]+)"?.*/\1/' | tr -d '[:space:]' || true
}

resolve_gcloud() {
  if command -v gcloud >/dev/null 2>&1; then
    command -v gcloud
    return 0
  fi
  for candidate in \
    /opt/homebrew/share/google-cloud-sdk/bin/gcloud \
    /usr/local/share/google-cloud-sdk/bin/gcloud \
    "$HOME/google-cloud-sdk/bin/gcloud"; do
    if [ -x "$candidate" ]; then
      printf '%s\n' "$candidate"
      return 0
    fi
  done
  return 1
}

GCLOUD_DIR="${CLOUDSDK_CONFIG:-$HOME/.config/gcloud}"
GLOBAL_ADC="${GCLOUD_DIR}/application_default_credentials.json"
SOURCE_ADC="${SOURCE_ADC_PATH:-$GLOBAL_ADC}"
CODE_DIR="${CODE_DIR:-$HOME/code}"
# 走査: repo -> SA を収集（書き込みはまだしない）
echo "=== impersonation リポジトリを走査 (${CODE_DIR}/*/{$(IFS=,; echo "${MISE_FILES[*]}")}) ===" >&2
TARGETS=""     # 改行区切りの "repo|sa|sa_file|has_gac" レコード (bash 3.2 互換のため文字列で保持)
CONFLICTS=""
for REPO_PATH in "${CODE_DIR}"/*/; do
  REPO_PATH="${REPO_PATH%/}"
  REPO_NAME="$(basename "$REPO_PATH")"
  SA_EMAIL=""; SA_FILE=""; HAS_GAC=0; SEEN=""
  for NAME in "${MISE_FILES[@]}"; do
    F="${REPO_PATH}/${NAME}"
    [ -f "$F" ] || continue
    if grep -qE '^[[:space:]]*GOOGLE_APPLICATION_CREDENTIALS' "$F"; then HAS_GAC=1; fi
    SA="$(sa_of "$F")"
    [ -z "$SA" ] && continue
    SEEN="${SEEN} ${NAME}=${SA}"
    if [ -z "$SA_EMAIL" ]; then
      SA_EMAIL="$SA"; SA_FILE="$F"
    elif [ "$SA" != "$SA_EMAIL" ]; then
      CONFLICTS="${CONFLICTS}  ❌ ${REPO_NAME}:${SEEN}"$'\n'
    fi
  done
  [ -z "$SA_EMAIL" ] && continue
  TARGETS="${TARGETS}${REPO_NAME}|${SA_EMAIL}|${SA_FILE}|${HAS_GAC}"$'\n'
done

if [ -n "$CONFLICTS" ]; then
  echo "❌ 1 つの repo の .mise.toml と mise.toml が別々の SA を指している。どちらかに揃えてから再実行する:" >&2
  printf '%s' "$CONFLICTS" >&2
  exit 1
fi

if [ "$LIST_ONLY" = 1 ]; then
  printf '%s' "$TARGETS" | while IFS='|' read -r REPO_NAME SA_EMAIL SA_FILE HAS_GAC; do
    [ -z "$REPO_NAME" ] && continue
    printf '%s\t%s\t%s\n' "$REPO_NAME" "$SA_EMAIL" "$SA_FILE"
  done
  [ -z "$TARGETS" ] && echo "  (impersonation リポジトリなし)" >&2
  exit 0
fi

GCLOUD_BIN="$(resolve_gcloud || true)"

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PYTHON="$(command -v python3 || true)"
[ -z "$PYTHON" ] && { echo "❌ python3 が見つからない"; exit 1; }
[ -n "$GCLOUD_BIN" ] || {
  echo "❌ gcloud が見つかりません"
  echo "   PATH を確認するか、Google Cloud SDK の bin ディレクトリを追加してください。"
  exit 1
}
export PATH="$(dirname "$GCLOUD_BIN"):$PATH"
[ -f "$SOURCE_ADC" ] || { echo "❌ source ADC が無い: $SOURCE_ADC"; echo "   先に: gcloud auth application-default login"; exit 1; }

echo "=== source_credentials を抽出中 ==="
# グローバル ADC から authorized_user の source を取り出す。
# 既に impersonated_service_account なら .source_credentials を、
# 素の authorized_user ならトップレベルを source として使う。
SOURCE_JSON="$("$PYTHON" "${SCRIPT_DIR}/scripts/adc_source.py" "$SOURCE_ADC")"
echo "  ✅ authorized_user source を取得"

echo ""
echo "=== per-repo ADC を生成 ==="
GENERATED=""   # 改行区切りの "repo|sa|adc" レコード
while IFS='|' read -r REPO_NAME SA_EMAIL SA_FILE HAS_GAC; do
  [ -z "$REPO_NAME" ] && continue
  ADC_PATH="${GCLOUD_DIR}/${REPO_NAME}_adc.json"
  IMP_URL="https://iamcredentials.googleapis.com/v1/projects/-/serviceAccounts/${SA_EMAIL}:generateAccessToken"

  echo "  • ${REPO_NAME}  ->  ${SA_EMAIL}"

  # per-repo ADC を生成
  SOURCE_JSON="$SOURCE_JSON" IMP_URL="$IMP_URL" "$PYTHON" "${SCRIPT_DIR}/scripts/write_repo_adc.py" "$ADC_PATH"
  chmod 600 "$ADC_PATH"

  # どの設定ファイルにも GOOGLE_APPLICATION_CREDENTIALS が無ければ、SA を書いている方へ追記
  if [ "$HAS_GAC" = 0 ]; then
    "$PYTHON" "${SCRIPT_DIR}/scripts/mise_insert_adc.py" "$SA_FILE" "$ADC_PATH"
    echo "    ↳ $(basename "$SA_FILE") に GOOGLE_APPLICATION_CREDENTIALS を追記"
  fi

  GENERATED="${GENERATED}${REPO_NAME}|${SA_EMAIL}|${ADC_PATH}"$'\n'
done < <(printf '%s' "$TARGETS")   # here-string は使わない（scripts/*.py の docstring を参照）

[ -z "$GENERATED" ] && echo "  (impersonation リポジトリなし)"

# グローバル ADC を素の authorized_user に戻す
echo ""
echo "=== グローバル ADC を素のユーザー認証へ戻す ==="
SOURCE_JSON="$SOURCE_JSON" "$PYTHON" "${SCRIPT_DIR}/scripts/restore_global_adc.py" "$GLOBAL_ADC"
chmod 600 "$GLOBAL_ADC"
echo "  ✅ グローバル ADC = authorized_user (impersonate なし)"

# 疎通確認
echo ""
echo "=== 疎通確認 (各 per-repo ADC で access token 取得) ==="
printf '%s' "$GENERATED" | while IFS='|' read -r REPO_NAME SA_EMAIL ADC_PATH; do
  [ -z "$REPO_NAME" ] && continue
  if GOOGLE_APPLICATION_CREDENTIALS="$ADC_PATH" \
     "$GCLOUD_BIN" auth application-default print-access-token >/dev/null 2>&1; then
    echo "  ✅ ${REPO_NAME}: OK"
  else
    echo "  ⚠️  ${REPO_NAME}: トークン取得失敗 (IAM 反映待ち or TokenCreator 未付与)"
  fi
done

echo ""
echo "完了。各リポジトリで 'cd' すると mise が ADC を自動切替する。"
