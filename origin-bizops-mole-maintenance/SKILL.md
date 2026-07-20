---
name: origin-bizops-mole-maintenance
description: >-
  Macの週次メンテナンスを Mole CLI (mo) で実行するスキル。
  使うこと: 「Mac の掃除」「週次メンテ」「ディスク整理」「mole を使って」「キャッシュを掃除」
  「node_modules を削除」「ストレージ空き」「mac maintenance」「clean up mac」「mole weekly」
  などのキーワードが含まれる場合に必ずこのスキルを使う。
  破壊的・no-return な操作（mo uninstall, mo uninstall --permanent, mo remove）は実行しない。
  常に dry-run でプレビューを提示してからユーザー確認を取り、承認後に実行する。
---

# Mole Weekly Maintenance

Mole CLI (`mo`) を使って Mac を週次メンテナンスするワークフロー。
破壊的・no-return な操作は除外し、すべて Trash 経由（復元可能）または可逆操作のみ使用する。

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

### Step 2: キャッシュ肥大の確認（読み取り専用）

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

### Step 3: キャッシュ掃除（dry-run → 確認 → 実行）

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

### Step 4: インストーラーファイル掃除（dry-run → 確認 → 実行）

Downloads・Desktop・Homebrew cache 等の `.dmg` / `.pkg` / `.zip` を検出する。

```bash
mo installer --dry-run
```

プレビュー結果をユーザーに提示し、確認を取る。承認後:

```bash
mo installer
```

### Step 5: プロジェクトビルド成果物の掃除（端末で手動実行）

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

### Step 6: システム最適化（dry-run → 確認 → 実行）

キャッシュ再構築・Finder/Dock更新・ネットワークサービスリセット等。

```bash
mo optimize --dry-run
```

プレビュー結果をユーザーに提示し、確認を取る。承認後:

```bash
mo optimize
```

### Step 7: 完了サマリー

実行したアクション・解放した容量をまとめて報告する。

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
