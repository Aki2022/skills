#!/usr/bin/env bash
#
# re-auth.sh — 統合再認証スクリプト
#
# GCP（全リポジトリ）と Google Ads OAuth を1回のコマンドで更新する。
# Workspace RAPT (24h) ポリシーによる失効後に実行すること。
#
# 使い方:
#   bash ~/.agents/skills/origin-gcp-project-setup/re-auth.sh
#
# 何が起きるか:
#   1. gcloud ユーザー認証 + ADC 更新 (browser 1回を想定)
#   2. 全リポジトリの per-repo ADC を一括再生成 (refresh_adc.sh)
#   3. Google Ads refresh_token 更新 (browser, --ads オプション時のみ)
#
# オプション:
#   --ads     Google Ads の OAuth も更新する（通常は不要）
#   --gcp     GCP のみ更新する（デフォルト動作と同じ）
#
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
UPDATE_ADS=false
TMP_GCLOUD_CONFIG="$(mktemp -d)"
trap 'rm -rf "$TMP_GCLOUD_CONFIG"' EXIT

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

GCLOUD_BIN="$(resolve_gcloud || true)"
[ -n "$GCLOUD_BIN" ] || {
  echo "❌ gcloud が見つかりません"
  echo "   PATH を確認するか、Google Cloud SDK の bin ディレクトリを追加してください。"
  exit 1
}
export PATH="$(dirname "$GCLOUD_BIN"):$PATH"

for arg in "$@"; do
  case "$arg" in
    --ads) UPDATE_ADS=true ;;
    --gcp) ;;  # デフォルト
    *) echo "不明なオプション: $arg" >&2; exit 1 ;;
  esac
done

echo "==================================================="
echo " 統合再認証スクリプト"
echo "==================================================="
echo ""

# --- Step 1: gcloud ユーザー認証 + ADC 更新 ---
echo "【Step 1/3】 gcloud ユーザー認証 + ADC 更新..."
echo "  次にブラウザ承認待ちへ入ります。CLI はここで待機するので、そのまま承認を完了してください。"
echo "  承認画面が見えない場合は、gcloud が表示する URL を確認してください。"
env -u GOOGLE_APPLICATION_CREDENTIALS -u CLOUDSDK_AUTH_IMPERSONATE_SERVICE_ACCOUNT \
  CLOUDSDK_CONFIG="$TMP_GCLOUD_CONFIG" \
  "$GCLOUD_BIN" auth login --update-adc

echo ""
echo "  browser 承認が完了したので、一時 ADC の内容を確認します..."
SOURCE_ADC_PATH="${TMP_GCLOUD_CONFIG}/application_default_credentials.json"
if [ ! -f "$SOURCE_ADC_PATH" ]; then
  echo "❌ 一時 ADC が生成されませんでした: $SOURCE_ADC_PATH"
  echo "   active gcloud config に impersonation が混入していないか確認し、再実行してください。"
  exit 1
fi

ADC_TYPE="$(python3 - "$SOURCE_ADC_PATH" <<'PY'
import json, sys
with open(sys.argv[1]) as f:
    print(json.load(f).get("type", ""))
PY
)"
if [ "$ADC_TYPE" != "authorized_user" ]; then
  echo "❌ 一時 ADC は authorized_user ではありませんでした: $ADC_TYPE"
  echo "   active gcloud config の impersonation 混入を避けるため、一時 CLOUDSDK_CONFIG を使っています。"
  echo "   別シェルの gcloud 設定を確認してから再実行してください。"
  exit 1
fi
echo "  ✅ 一時 ADC = authorized_user"

# --- Step 3: per-repo ADC を一括再生成 ---
echo ""
echo "【Step 2/3】 全リポジトリの per-repo ADC を再生成..."
SOURCE_ADC_PATH="$SOURCE_ADC_PATH" \
  bash "${SCRIPT_DIR}/refresh_adc.sh"

# --- Step 4 (optional): Google Ads OAuth ---
if [ "$UPDATE_ADS" = true ]; then
  echo ""
  echo "【Step 3/4】 Google Ads OAuth refresh_token を更新..."
  MOCKUP_DIR="${CODE_DIR:-$HOME/code}/mockup"
  if [ ! -d "$MOCKUP_DIR" ]; then
    echo "❌ mockup リポジトリが見つかりません: $MOCKUP_DIR"
    echo "   CODE_DIR 環境変数で場所を指定してください"
    exit 1
  fi
  # mise 経由で実行することで OAUTH_CLIENT_ID, OAUTH_CLIENT_SECRET がロードされる
  (cd "$MOCKUP_DIR" && mise exec -- node scripts/utils/refresh_google_ads_token.cjs)
fi

echo ""
echo "==================================================="
echo " 再認証完了"
echo "==================================================="
echo ""
echo "  GCP: 全リポジトリの ADC を更新済み"
echo "  通常の入口は各 repo の 'mise run reauth' です"
if [ "$UPDATE_ADS" = true ]; then
  echo "  Google Ads: ~/.config/gads/mockup_token.env を更新済み"
else
  echo "  Google Ads: スキップ (更新が必要な場合は --ads オプションを追加)"
fi
echo ""
echo "次回 RAPT 失効後の手順:"
echo "  bash ~/.agents/skills/origin-gcp-project-setup/re-auth.sh"
echo "  (Google Ads も更新: bash ... re-auth.sh --ads)"
