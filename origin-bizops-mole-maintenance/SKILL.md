---
name: origin-bizops-mole-maintenance
description: >-
  Macの週次メンテナンスを Mole CLI (mo) で実行するスキル。
  使うこと: 「Mac の掃除」「週次メンテ」「ディスク整理」「mole を使って」「キャッシュを掃除」
  「node_modules を削除」「ストレージ空き」「mac maintenance」「clean up mac」「mole weekly」
  などのキーワードが含まれる場合に必ずこのスキルを使う。
  破壊的・no-return な操作（mo uninstall, mo uninstall --permanent, mo remove）は実行しない。
  常に dry-run または読み取り専用プレビューを提示してからユーザー確認を取り、承認後に実行する。
---

# Mole Weekly Maintenance

Mole CLI (`mo`) を使って Mac を週次メンテナンスするワークフロー。
破壊的・no-return な操作は原則除外する。必要な整理でも対象を限定し、明示確認後に実行する。ファイルは可能な限りTrash経由とする。

## 前提

- Mole がインストール済みであること: `which mo` で確認
- インストール方法: `brew install mole` (tw93/tap)
- 現端末バージョン: v1.44.1 (arm64)

## 実行フロー

### Step 1: ヘルスチェック（読み取り専用）

```bash
mo status --json
```

表示する情報:

- `health_score` (100点満点)
- CPU使用率・メモリ使用率
- ディスク空き容量
- アップタイム

### Step 2: 全体容量と見落としやすい領域の確認（読み取り専用）

キャッシュだけでなく、仮想ディスク・同期データ・ゴミ箱・ローカルスナップショットを確認する。

```bash
df -h /
du -x -d 1 -k ~/Library/Containers 2>/dev/null | sort -nr | head -20
find ~/Library/Containers -xdev -type f -size +1G -exec ls -lh {} \; 2>/dev/null | head -30
tmutil listlocalsnapshots /
```

`du` はAPFSのスパースファイルの論理サイズと実使用量が異なる場合があるため、必ず `df` と併記する。

特に `~/Library/Containers/com.docker.docker/Data/vms/0/data/Docker.raw` は直接削除しない。Dockerの整理は後述のDocker手順を使う。

### Step 3: キャッシュ肥大の確認（読み取り専用）

```bash
mo analyze ~/Library/Caches --json | python3 -c "
import json, sys
data = json.load(sys.stdin)
entries = sorted(data.get('entries', []), key=lambda x: -x.get('size', 0))
for e in entries[:10]:
    print(f\"{e['size']//1024//1024:>6}MB  {e['name']}\")
print(f\"Total: {data.get('total_size', 0)//1024//1024}MB\")
"
```

### Step 4: キャッシュ掃除（dry-run → 確認 → 実行）

**まず dry-run でプレビュー:**

```bash
mo clean --dry-run
```

プレビュー結果をユーザーに提示し、確認を取る。承認後:

```bash
mo clean
```

**重要: Mole のデフォルト whitelist で以下は保護済み（消えない）:**

- `~/Library/Caches/ms-playwright*` — Playwright ブラウザバイナリ（528MB）
- `~/.cache/huggingface*` — HuggingFace モデル（vector/embedding含む）
- `~/.ollama/models/*` — Ollama ローカルLLMモデル
- `~/Library/Caches/pypoetry/virtualenvs*` — Python仮想環境
- `~/Library/Caches/JetBrains*` — JetBrains IDE インデックス
- `~/Library/Caches/com.nssurge.surge-mac/*` — Surge プロキシ設定
- `~/.m2/repository/*` — Maven依存ライブラリ
- `~/.gradle/caches/*` — Gradle依存ライブラリ

**パフォーマンスへの影響（軽微、許容範囲）:**

- Go build cache が再構築される（初回ビルドが若干遅くなる）
- ブラウザキャッシュが消えるため、初回ページ読み込みが若干遅い
- `com.openai.codex` (1.3GB) は Sparkle 更新キャッシュ（AIモデルではない）→ 次回起動時に再構築

**MoleがTTY・権限・ログ書き込みで失敗した場合:**

- エラーを明示し、成功したものとして扱わない
- `df`、`du`、`find` で読み取り専用の代替集計を行う
- `mo clean --dry-run` の代わりに手動で対象を列挙し、確認後にゴミ箱へ移動する

### Step 5: インストーラーファイル掃除（dry-run → 確認 → 実行）

Downloads・Desktop・Homebrew cache 等の `.dmg` / `.pkg` / `.zip` を検出する。

```bash
mo installer --dry-run
```

プレビュー結果をユーザーに提示し、確認を取る。承認後:

```bash
mo installer
```

### Step 6: Dockerの整理（dry-run相当の確認 → 確認 → 実行）

Docker Desktopの使用量はMoleだけでは把握しにくいため、Docker CLIで確認する。

```bash
docker system df -v
docker ps -a --size
docker volume ls
docker image ls --format '{{.Repository}}\t{{.Tag}}\t{{.ID}}\t{{.CreatedSince}}\t{{.Size}}'
```

安全な既定順序:

1. `docker system df -v` でビルドキャッシュ・イメージ・コンテナ・ボリュームを分けて提示する
2. 実行中コンテナとボリュームを保護する
3. ビルドキャッシュは再生成可能だが、`docker builder prune -a` は完全削除なので別途確認を取る
4. 古いイメージは作成日時とコンテナ参照を確認し、対象IDを明示して `docker image rm <id>` を実行する

`docker system prune -a --volumes` は一括で広範囲を削除するため、自動実行しない。`Docker.raw` を直接削除・移動しない。Docker CLIが使えない場合はDocker Desktopの起動を案内し、仮想ディスク内部を手動操作しない。

Dockerのpruneとimage rmはゴミ箱を経由しないため、実行前に対象・再取得の影響・見込み容量を明示して確認を取る。実行後は `docker system df` と `df -h /` を再確認する。

### Step 7: プロジェクトビルド成果物の掃除（端末で手動実行）

`mo purge` は対話的TUIのため、Claude からは直接実行できない。端末で実行してもらう。

```bash
mo purge
```

**注意点:**

- 7日以内に更新されたプロジェクトはデフォルトで選択外（誤削除防止）
- TUIでプロジェクト名・最終利用日・サイズを確認しながら個別に選択する
- 実行時は Trash に移動（`rm -rf` ではない）→ 誤って消しても Trash から復元可能
- `node_modules` を削除したプロジェクトは次回 `npm install` / `pnpm install` が必要

**⚠️ Rust / Go など「コンパイル済みバイナリを本番で使っているプロジェクト」の扱い:**

`mo purge` は `target/`（Rust）や `bin/`（Go）を丸ごと削除する。
`target/release/<バイナリ名>` が稼働中のサービスで使われている場合、削除すると障害になる。

対処方法は2つ:

1. **TUIで該当プロジェクトを毎回除外する**（確実）
2. **ビルドゴミだけ手動で削除する**（`target/release/` は残したい場合）

```bash
# Rust: debug ビルドのみ削除（release バイナリは保持）
cargo clean --manifest-path ~/code/<project>/Cargo.toml --profile dev

# または debug ディレクトリを直接削除
rm -rf ~/code/<project>/target/debug
```

### Step 8: システム最適化（dry-run → 確認 → 実行）

キャッシュ再構築・Finder/Dock更新・ネットワークサービスリセット等。

```bash
mo optimize --dry-run
```

プレビュー結果をユーザーに提示し、確認を取る。承認後:

```bash
mo optimize
```

### Step 9: 完了サマリー

実行したアクション・ゴミ箱へ移動した容量・実際に `df` で増えた空き容量を区別して報告する。APFSやゴミ箱の影響で `du` の合計と `df` の変化が一致しない場合は、その旨を明記する。

---

## 除外コマンド（絶対に実行しない）

| コマンド                   | 理由                                       |
| -------------------------- | ------------------------------------------ |
| `mo uninstall`             | アプリ削除は週次メンテに不適切、重大すぎる |
| `mo uninstall --permanent` | rm -rf、NO RETURN（Trash 非経由）          |
| `mo remove`                | Mole 自体のアンインストール                |

## whitelist の追加・確認

ユーザーが特定のキャッシュを保護したい場合:

```bash
mo clean --whitelist   # TUI で保護対象を選択
```

whitelist 設定ファイル: `~/.config/mole/whitelist`

## トラブルシューティング

| 問題                       | 対処                                             |
| -------------------------- | ------------------------------------------------ |
| `mo: command not found`    | `brew install mole` または `which mo` でパス確認 |
| purge のスキャンが遅い     | `brew install fd` でスキャンが高速化             |
| 特定キャッシュを保護したい | `mo clean --whitelist` で whitelist に追加       |
| 誤って削除した             | Trash を確認（`~/.Trash/`）→ 復元可能            |
