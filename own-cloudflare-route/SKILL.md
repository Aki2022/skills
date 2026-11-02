---
name: own-cloudflare-route
description: >-
  Cloudflare の作業で cf と Wrangler のどちらを使うか選び、適切な CLI と skill へ振り分ける。
  Cloudflare CLI の導入・選定、Workers の開発・build・deploy、Cloudflare API / account resource の操作を
  行うときに使う。既存の Wrangler project を cf へ移行するか判断するときも使う。
---

# Cloudflare CLI の選択

## 最初に project と作業対象を確認する

- `cloudflare.config.ts`、`wrangler.json` / `wrangler.jsonc` / `wrangler.toml`、`package.json`、lockfile を確認する。
- 対象が Worker の開発・build・deploy なのか、Cloudflare account/API resource の管理なのかを分ける。
- Cloudflare の仕様と CLI は更新されるため、実行前に該当する公式ドキュメントを読む。

## CLI を選ぶ

- 新しい Cloudflare project と `cloudflare.config.ts` がある project では `cf` を使う。必要な場合は `cloudflare` skill で製品を選び、`cf` の最新公式ドキュメントを参照する。
- `wrangler.json` / `wrangler.jsonc` / `wrangler.toml` があり、`cloudflare.config.ts` がない既存 project では、その Worker の dev・build・deploy・設定に Wrangler を使う。`cf dev`、`cf build`、`cf deploy` を先に実行しない。
- 既存 Wrangler project でも、Cloudflare の account/API resource を操作する依頼には `cf` を使える。Worker の実行設定や配備経路と混同しない。
- Wrangler project の移行は、ユーザーが移行を求めた場合に限る。公式の移行手順を読み、必要な設定変更を確認してから進める。CLI 名だけを置き換えて移行済みと判断しない。
- どちらの config も無い場合は、目的に合わせて新規なら `cf` を選ぶ。既存構成の手掛かりが足りず影響が決められないときは、変更前に確認する。

## `cf` の使い方

- `cf` は beta。記憶した flag や Wrangler の対応コマンドを流用せず、[Cloudflare CLI docs](https://developers.cloudflare.com/cf/) と [coding agents guide](https://developers.cloudflare.com/cf/agents/) を読む。
- コマンドを探すときは `cf cli search "<task>"` を使い、実行前に該当 command の `--help` または `cf schema` で現在の引数を確認する。
- `cf` は Node.js 22.18 以上が必要。Codex の非対話 shell では `mise exec -- cf ...` を使い、対話 shell でも Node の実効 version を確認する。要件未満の Node で再試行を繰り返さない。
- `cf` と Wrangler の認証は別々に扱う。依頼が無い限り login や認証状態の変更を始めない。
- 書き込み・deploy・削除などのリモート操作は、対象 account・resource・環境を特定し、依頼された範囲だけを実行する。利用可能なら dry-run を先に使う。

## Wrangler の使い方

- Wrangler が必要な既存 project では `wrangler` skill を読む。project の package manager と lockfile に従い、project が pin した CLI を使う。
- 現行の command、config、権限、remote/local の違いは [Wrangler docs](https://developers.cloudflare.com/workers/wrangler/) で確認する。

