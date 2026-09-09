---
title: 実行機構 — 安全モード・コマンド解決・保存先の罠・完了判定
---

# 実行機構

Codex 組み込み `image_gen`（gpt-image-2）を安全・決定的に呼ぶための機構。
**罠はすべて実測で確認されたもの**で、禁止形を文書に書くだけでは破られた実績がある。
だから窓口スクリプトの内側に閉じ込めている。

## 認証と課金

ChatGPT サブスクの OAuth を使う。`OPENAI_API_KEY` は不要で、**従量課金は発生しない**。
消費するのは**シートの画像利用枠**（有限）。枠切れは
`Your workspace is out of credits` で表れる。

## コマンド解決

```bash
CODEX_BIN=$(command -v codex2 || command -v codex)
```

- `codex2` は `CODEX_HOME=~/.codex-seat2` で同じバイナリを起動する**別シートのラッパー**。
  シートごとに認証と利用枠が分かれる。メインシートが枠切れを返したらサブシートで実行する。
- ログイン確認はシート側で行う（`codex2 login status` → `Logged in using ChatGPT`）。
- **生成物の保存先もシートに従う**（`<CODEX_HOME>/generated_images/<session-id>/`）。
  回収は session id から全シートを自動探索する。**呼び出し側で `CODEX_HOME` を渡さない** —
  ラッパーは子プロセス内だけで切り替えるので、呼び出し側シェルの `CODEX_HOME` に頼る実装は
  別シートの生成物を見失う。

事前確認: `"$CODEX_BIN" features list | grep image_generation` が `stable true` を返す。

## 安全モード（固定）

```bash
"$CODEX_BIN" exec --sandbox workspace-write \
  -c sandbox_workspace_write.network_access=true --cd "$PWD" "<prompt>" < /dev/null
```

- 既定サンドボックス（`read-only`）はネットワークを遮断するので image_gen がブロックされる。
  **フルバイパスは不要** — `workspace-write` を維持したままネットワークだけ許可すれば正常動作する。
- 作業ディレクトリが git リポジトリの外（一時ディレクトリ等）なら `--skip-git-repo-check` を足す。
  無いと「Not inside a trusted directory」で即終了する。
- **`< /dev/null` は必須。** バックグラウンド実行で stdin を閉じないとハングする。
- **`--dangerously-bypass-approvals-and-sandbox` は使わない。** 安全モードで動くうえ、
  Claude Code の分類器がこのフラグをハードブロックするため AI からは実行できない。
  窓口が組み立てるコマンドにこの文字列が現れないことを検査で守っている。

## 保存先の罠（2度誤判定した）

image_gen は生成物を**既定で**次へ保存する。**プロンプト内で「Xとして保存して」と書いても
そのパスには保存されない。**

```
<CODEX_HOME>/generated_images/<session-id>/ig_*.png
```

この後処理が欠けると**画像自体は生成されているのに期待した場所に無い**状態になり、
過去に「生成失敗」と誤判定した原因はこれだった。生成 → 探索 → コピーを
**ワンセットの手順**として扱う。窓口がこれを内側でやる。

**もう一つの誤判定源**: `ERROR: Reconnecting... N/5` →
`Falling back from WebSockets to HTTPS transport`。これはストリーミング接続の
フォールバックで**失敗ではない**（表示されたが全画像が正常生成された実測がある）。
この行で中断しない。

## 回収は「自セッションのディレクトリ」から

並列に codex セッションを走らせると、共有ルートを glob で拾う実装は**別セッションの画像を
掴む**（11枚のうち6枚が別の用途の画像になった実測がある）。回収は**必ず session id 起点**で行う。
共有ルートの glob は禁止。

## 完了判定3原則

1. **旧成果物を退避する。** 実行前に既存の出力先を `.stale` へ移す。退避しないと、
   生成が失敗しても旧ファイルが残り**成功に見える**（stale pass-through）。
2. **回収失敗を exit 非0 に反映する。** 生成数 < 出力名数なら停止し、**不足を埋めない**。
   埋めると「成功に見えて中身が違う」になる。
3. **1枚目をスモークとして直列実行する。** 様式が合っていることを確認してから残りを並列にする。
   枠切れもここで検出できる（バッチ途中で切れると回収が半端になる）。

`.stale` に退避したファイルは**成果物パスとして返さない**。返すと退避の意味が消える。
