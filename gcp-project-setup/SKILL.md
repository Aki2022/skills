---
name: gcp-project-setup
description: |
  新しいリポジトリをGCP複数プロジェクト環境にimpersonation方式（鍵レス）で追加するセットアップスキル。
  サービスアカウント作成・最小権限付与・mise.toml生成・end-to-end検証まで一括実行する。

  以下のようなリクエストで必ず使うこと:
  - 「新しいリポジトリをGCPに繋げたい」
  - 「effectuation_score を innovation-score プロジェクトでセットアップして」
  - 「GCPの認証を設定して」「mise.toml を作って」
  - 「impersonation でセットアップ」「鍵レスで設定」
  - 「別プロジェクトのリポジトリに認証を追加」
  - GCPプロジェクトIDとリポジトリ名が両方出てくるセットアップ文脈
---

# GCP Project Setup（impersonation方式・鍵レス）

## 概要

このスキルは、新しいリポジトリを GCP 複数プロジェクト環境に安全に追加するための
**impersonation（鍵レス）方式**セットアップを自動化する。

秘密鍵をディスクに保存せず、ユーザー認証を土台にサービスアカウント（SA）になりすます方式。
`cd` するだけでそのリポジトリ専用の GCP 認証が有効になる。

> **CLI と SDK は認証経路が違う**（重要）
>
> - gcloud / bq CLI は `.mise.toml` の `CLOUDSDK_AUTH_IMPERSONATE_SERVICE_ACCOUNT` を読む。
> - Node / Python の Google SDK は**この変数を読まない**。ADC
>   (`~/.config/gcloud/application_default_credentials.json`) を見る。
> - ADC はグローバル単一ファイルなので複数リポジトリで衝突する。
>   そこで **per-repo ADC ファイルを分離**し、`.mise.toml` の `GOOGLE_APPLICATION_CREDENTIALS`
>   で切り替える（Step 6.5）。SDK を使うリポジトリでは必須。
>
> **「鍵レス」の正確な意味**：SA キー（JSON 鍵）はディスクに置かない。ただし per-repo ADC には
> ユーザー本人の OAuth refresh_token が埋め込まれる（SA キーより低リスクだが長命資格情報ではある）。

## 前提確認

作業開始前に以下を確認すること:

```bash
# ユーザー認証が生きているか
gcloud auth print-access-token 2>&1 | head -1
# → トークンが返れば OK。エラーなら `gcloud auth login` を先に実行
```

認証が切れていたらユーザーに `gcloud auth login` の実行を依頼して止まること。

## 引数

| 引数              | 必須 | 例                                       | 説明                           |
| ----------------- | ---- | ---------------------------------------- | ------------------------------ |
| リポジトリ名      | ✅   | `effectuation_score`                     | `~/code/` 以下のディレクトリ名 |
| GCPプロジェクトID | ✅   | `innovation-score`                       | GCP コンソールのプロジェクトID |
| 権限セット        | ❌   | `bigquery+gcs`（デフォルト）/ `vertexai` | 付与する権限の種類             |

## 実行ステップ

### Step 1: 変数セット

```bash
REPO_NAME="<リポジトリ名>"
PROJECT_ID="<GCPプロジェクトID>"
USER_ACCOUNT="toshiaki.okimoto@office-ohana.net"
SA_NAME="${REPO_NAME//_/-}-local-dev"           # アンダースコアをハイフンに変換
SA_EMAIL="${SA_NAME}@${PROJECT_ID}.iam.gserviceaccount.com"
REPO_PATH="$HOME/code/${REPO_NAME}"
```

> SA 名はアンダースコアが使えないため `_` → `-` に変換する。

### Step 2: SA 作成

```bash
gcloud iam service-accounts create "${SA_NAME}" \
  --display-name="Local Dev (impersonation)" \
  --project="${PROJECT_ID}"
```

既存の場合は `already exists` エラーが出るが続行して問題ない。

### Step 3: 権限付与

**デフォルト（bigquery+gcs）**:

```bash
for ROLE in \
  "roles/bigquery.dataEditor" \
  "roles/bigquery.jobUser" \
  "roles/serviceusage.serviceUsageConsumer" \
  "roles/storage.objectAdmin"; do
  gcloud projects add-iam-policy-binding "${PROJECT_ID}" \
    --member="serviceAccount:${SA_EMAIL}" \
    --role="${ROLE}" \
    --condition=None \
    --format="value(etag)"
done
```

> `bigquery.dataEditor` だけではクエリ実行ができない。`bigquery.jobUser` が必ず必要。
> `serviceusage.serviceUsageConsumer` は SDK/quota project 経由で API を呼ぶ際に必須
> （無いと `Caller does not have required permission to use project ...` で失敗する）。

**vertexai オプション追加時**（上記に加えて）:

```bash
gcloud projects add-iam-policy-binding "${PROJECT_ID}" \
  --member="serviceAccount:${SA_EMAIL}" \
  --role="roles/aiplatform.user" \
  --condition=None
```

### Step 4: TokenCreator 権限（impersonation の核心）

```bash
# ⚠️ --project を必ず明示すること。省略すると active config のプロジェクトを見に行って失敗する
CLOUDSDK_CORE_PROJECT="${PROJECT_ID}" \
gcloud iam service-accounts add-iam-policy-binding "${SA_EMAIL}" \
  --member="user:${USER_ACCOUNT}" \
  --role="roles/iam.serviceAccountTokenCreator" \
  --project="${PROJECT_ID}"
```

### Step 5: IAM 反映待ち（最大 60 秒）

IAM 変更は即時反映されない。トークンが取れるまでリトライする:

```bash
echo "IAM 反映待ち..."
for i in $(seq 1 6); do
  TOKEN=$(gcloud auth print-access-token \
    --impersonate-service-account="${SA_EMAIL}" 2>/dev/null)
  if [ -n "$TOKEN" ]; then
    echo "✅ ${i}回目で反映完了"
    break
  fi
  echo "  ${i}/6 まだ反映中... (10秒待機)"
  sleep 10
done
[ -z "$TOKEN" ] && echo "❌ 60秒経過しても反映されず。数分後に再試行を" && exit 1
```

### Step 6: .mise.toml 作成

```bash
cat > "${REPO_PATH}/.mise.toml" << EOF
# ${REPO_NAME} (${PROJECT_ID}) ローカル開発環境
# impersonation方式: 秘密鍵を持たずSAになりすます
[env]
# SDK（Python/R/クライアントライブラリ）が読む既定プロジェクト
GOOGLE_CLOUD_PROJECT = "${PROJECT_ID}"
# gcloud / bq CLI が読む既定プロジェクト（CLIはGOOGLE_CLOUD_PROJECTを見ない）
CLOUDSDK_CORE_PROJECT = "${PROJECT_ID}"
# quota/課金プロジェクトを固定（ADC使用時の誤プロジェクト参照を防止）
GOOGLE_CLOUD_QUOTA_PROJECT = "${PROJECT_ID}"
# SDK・gcloud がこのSAになりすます（鍵レス）※CLIのみ有効。SDKはStep 6.5のADCを使う
CLOUDSDK_AUTH_IMPERSONATE_SERVICE_ACCOUNT = "${SA_EMAIL}"
# SDK(Node/Python)用 ADC。CLIのIMPERSONATE変数はSDKに効かないため必須（Step 6.5で生成）
GOOGLE_APPLICATION_CREDENTIALS = "${HOME}/.config/gcloud/${REPO_NAME}_adc.json"
EOF
```

> `GOOGLE_CLOUD_PROJECT` と `CLOUDSDK_CORE_PROJECT` は別物。両方必要。
> SDK（Python/R）は前者、gcloud/bq CLI は後者を参照する。
> `GOOGLE_APPLICATION_CREDENTIALS` は SDK 用 ADC のパス。実体は Step 6.5 で生成する。

### Step 6.5: SDK 用 per-repo ADC を生成（SDK を使うリポジトリは必須）

`CLOUDSDK_AUTH_IMPERSONATE_SERVICE_ACCOUNT` は gcloud/bq CLI 専用で、Node/Python の
Google SDK には効かない。SDK は ADC を見るが、グローバル ADC は単一ファイルのため複数
リポジトリで衝突する。横断ヘルパーで per-repo ADC（`~/.config/gcloud/<repo>_adc.json`）を
生成する。ブラウザ再ログインは不要（グローバル ADC の refresh_token を共有して生成する）。

```bash
bash ~/.claude/skills/gcp-project-setup/refresh_adc.sh
```

このヘルパーは `~/code/*/.mise.toml` を走査し、`CLOUDSDK_AUTH_IMPERSONATE_SERVICE_ACCOUNT`
を持つ全リポジトリの per-repo ADC を生成・更新し、`GOOGLE_APPLICATION_CREDENTIALS` 行が
無ければ `.mise.toml` に追記する。最後にグローバル ADC を素のユーザー認証へ戻す
（未設定リポジトリが誤った SA で動くのを防ぐ）。

> 前提：グローバル ADC が必要。無ければ先に `gcloud auth application-default login` を実行。

### Step 7: mise trust

```bash
cd "${REPO_PATH}" && /opt/homebrew/bin/mise trust .mise.toml
```

### Step 8: end-to-end 検証

```bash
cd "${REPO_PATH}"
eval "$(/opt/homebrew/bin/mise env -s bash)"

echo "=== プロジェクト確認 ==="
echo "GOOGLE_CLOUD_PROJECT: $GOOGLE_CLOUD_PROJECT"
echo "CLOUDSDK_CORE_PROJECT: $CLOUDSDK_CORE_PROJECT"
echo "IMPERSONATE: $CLOUDSDK_AUTH_IMPERSONATE_SERVICE_ACCOUNT"

echo "GAC(SDK用ADC): $GOOGLE_APPLICATION_CREDENTIALS"

echo ""
echo "=== CLI 実行確認（bq が誰として動くか）==="
bq query --use_legacy_sql=false --format=pretty \
  'SELECT SESSION_USER() AS running_as' 2>&1 | grep -vE "^$" | tail -6

echo ""
echo "=== SDK 実行確認（Node が誰として動くか）==="
node -e "const {BigQuery}=require('@google-cloud/bigquery'); new BigQuery({projectId:process.env.GOOGLE_CLOUD_PROJECT}).query('SELECT SESSION_USER() AS u').then(([r])=>console.log('SDK running_as:', r[0].u)).catch(e=>console.error('ERR:', e.message))"
```

期待する出力（CLI・SDK 両方）:

```
| running_as                                                   |
| <sa-name>@<project-id>.iam.gserviceaccount.com              |
SDK running_as: <sa-name>@<project-id>.iam.gserviceaccount.com
```

ユーザーアカウント（`@office-ohana.net`）が表示されたら impersonation が効いていない。
CLI は SA だが SDK だけユーザーになる場合は Step 6.5 の per-repo ADC 未生成
（`refresh_adc.sh` を実行）。

## 完了後の報告

以下をユーザーに伝える:

1. **作成した SA**: `${SA_EMAIL}`
2. **付与した権限**: 付与したロール一覧
3. **作成ファイル**: `${REPO_PATH}/.mise.toml`、`~/.config/gcloud/${REPO_NAME}_adc.json`（SDK用ADC）
4. **毎日の使い方**: `cd ~/code/${REPO_NAME}` するだけで CLI も SDK も自動切替
5. **再認証が必要な場合**（refresh_token 失効時のみ・数週間〜数ヶ月に1回程度）:
   ```bash
   gcloud auth login                          # gcloud/bq CLI 用
   gcloud auth application-default login       # ADC（SDK）用の素ログイン
   bash ~/.claude/skills/gcp-project-setup/refresh_adc.sh   # per-repo ADC を一括再生成
   ```
   > per-repo ADC は生成時の refresh_token を焼き込むため、再認証では自動更新されない。
   > 必ず `refresh_adc.sh` を再実行して全リポジトリ分を作り直すこと。

## トラブルシューティング

| エラー                                                                                     | 原因                                                         | 対処                                                                                                 |
| ------------------------------------------------------------------------------------------ | ------------------------------------------------------------ | ---------------------------------------------------------------------------------------------------- |
| `NOT_FOUND: Service account ... does not exist`                                            | --project 省略で active config のプロジェクトを参照          | Step 4 の `CLOUDSDK_CORE_PROJECT` 指定を確認                                                         |
| `PERMISSION_DENIED: Failed to impersonate`                                                 | IAM 未反映 or TokenCreator 未付与                            | Step 5 のリトライを待つ。それでも失敗なら Step 4 を再実行                                            |
| `SESSION_USER()` がユーザーアカウントを返す                                                | impersonation が効いていない                                 | `CLOUDSDK_AUTH_IMPERSONATE_SERVICE_ACCOUNT` 環境変数を確認                                           |
| `bq: Reauthentication failed`                                                              | 土台のユーザー認証が期限切れ                                 | `gcloud auth login` を実行                                                                           |
| SDK だけ `Caller does not have required permission to use project` / `serviceusage` エラー | per-repo ADC 未生成。SDK がグローバル ADC（別 SA）を見ている | Step 6.5 の `refresh_adc.sh` を実行し、`.mise.toml` に `GOOGLE_APPLICATION_CREDENTIALS` があるか確認 |
| CLI は SA だが SDK だけユーザー/別 SA で動く                                               | `GOOGLE_APPLICATION_CREDENTIALS` 未設定 or 指す ADC が古い   | `.mise.toml` を確認し `refresh_adc.sh` を再実行                                                      |
