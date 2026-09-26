---
name: own-seat2-build
description: >-
  デスクトップアプリの二席目（Claude-Seat2 / Codex-Seat2 等、同じ製品を別アカウントで
  同時に動かすための複製 bundle）を、本体 app から作り直して版を追従させる。
  必ず使うこと: 「Seat2 がアップデートできない」「二席目のアプリが古い」
  「Claude-Seat2 / Codex-Seat2 を更新して」「seat2 を作り直して」「二席目を作りたい」
  「アプリ内アップデートが失敗する」と言われた時、本体 app を更新した後、
  macOS の二重起動・複数アカウント運用の構成を点検する時。
  「アプリ内の更新機能を直す」方向で調べ始める前に必ずこれを読む — 識別子を書き換えた
  複製では更新機能は原理的に成功しないので、その方向は必ず空振りする。
  使わない場面: 本体 app 自体の更新（各アプリの更新機能に任せる）／エージェント設定や
  skill の symlink 統一（own-skill-commonize が持つ）／アカウントの認証・ログイン操作。
---

# デスクトップ二席目の再生成

同じ製品を2アカウントで同時に動かすために、本体 app を複製して
`CFBundleIdentifier` を書き換えた bundle（以下 seat2）を `/Applications` に置く運用がある。
このスキルは **seat2 を本体から作り直す**手順と、その再生成スクリプトを持つ。

## まず知っておくこと: seat2 のアプリ内更新は直せない

両社の更新機構（Claude = Squirrel.Mac、Codex = Sparkle）は、ダウンロードした更新物の中から
**実行中アプリと同じ bundle identifier を持つ bundle を探して差し替える**。
seat2 は識別子を書き換えてあるので一致せず、必ずこのログで終わる:

```
[updater] Found an update, downloading
[updater] Auto-update error: Could not locate update bundle for
  local.launchers.claude-seat2 within .../local.launchers.claude-seat2.ShipIt/update.XXXX/
  domain: SQRLUpdaterErrorDomain
```

仮に一致させても解決しない。更新は bundle ごと置き換えるので、seat 分離の要である
shim（`MacOS/<exe>` = プロファイルを渡すシェルスクリプト、`MacOS/<exe>Seat2Payload` = 本体バイナリ）が
消え、ただの1席目に戻る。**seat2 の更新経路は「本体から作り直す」以外にない。**

「更新機能を直す」「識別子を戻す」「更新サーバを差し替える」方向は全部ここで止まる。
ユーザーが「アップデートできない」と言ったときに調べるべきは更新機能ではなく、**ドリフトの量**。

## 手順

### 1. ドリフトを見る

```bash
bash ~/.agents/skills/own-seat2-build/scripts/seat2_clone.sh --check
```

`OK` なら本体と同版、`DRIFT` なら再生成が要る（exit 1）。
seat2 が未インストールなら `SKIP` を出して何もしない。

### 2. 本体 app を先に最新にする

seat2 は本体の複製なので、**本体が古ければ作り直しても古いまま**。
本体（`Claude.app` / `ChatGPT.app`）のアプリ内更新を先に済ませる。ここは更新機能が正常に働く。

### 3. 対象の seat2 を終了してから再生成

```bash
bash ~/.agents/skills/own-seat2-build/scripts/seat2_clone.sh --install all
```

`all` の代わりに `claude` / `codex` で片方だけも作れる。

**自分が動いている席は作り直せない。** エージェントのセッションが seat2 の上で動いているとき、
そのセッションから `--install` を呼ぶとスクリプトが起動中を検出して拒否する（これは正しい挙動）。
その場合は素のターミナルから実行するようユーザーに渡す。反対側の席（例: Claude 上から codex）は
そのまま作り直せる。

### 4. 確認して伝える

`--check` が両方 `OK` になることを見る。そのうえでユーザーに伝えるべきことが2つある:

- プロファイル（`~/.claude-seat2` / `~/.codex-seat2`）は触らないので、アカウント・設定・
  セッションは残る。
- bundle は別物になるので、**Automation / アクセシビリティ / 画面収録の許可は初回に再度聞かれる**。
  権限は bundle の署名に紐づくため、これは避けられない。

## スクリプトが作るもの

`scripts/seat2_clone.sh` は本体 app を `ditto` で複製し、**3点だけ**変えて ad-hoc 署名する。
手で直すときもこの3点以外に触らない — 触るほど次の本体更新との差分が増える。

1. **shim**: `Contents/MacOS/<exe>` を、プロファイルディレクトリを `--user-data-dir` として
   注入して `<exe>Seat2Payload`（改名した本体バイナリ）を `exec` するシェルスクリプトにする。
   呼び出し側が `--user-data-dir` を明示していれば尊重する。Codex は加えて `CODEX_HOME` と
   `CODEX_ELECTRON_USER_DATA_PATH` を export する。
2. **Info.plist**: `CFBundleDisplayName`・`CFBundleIdentifier`・URL scheme。
   scheme は列挙ではなく**接尾辞を付けて**書き換える（本体が scheme を増やしても seat2 が
   黙って二重登録しないため）。`http` / `https` は落とす（seat2 が既定ブラウザを争わないため）。
3. **ad-hoc の `--deep` 再署名**。

Info.plist の変換処理は同じ `scripts/` の `patch_plist.py` に置き、`seat2_clone.sh` は
そのPythonファイルを直接起動する。Pythonコードを Bash の here-document に埋め込まない。
macOS ではパイプ容量がシステム状態で縮み、Bash が子プロセス起動前に here-document を
書ききれず停止する場合がある。helper は同じ `scripts/` に保つ。

### `--deep` は必須

バイナリを改名した時点でベンダー署名は無効になる。外側だけ署名して payload をベンダー署名の
まま残すと、**起動時に SIGKILL される**（rc=137。2026-09-23 に実測）。
`--deep` は非推奨と言われるが、ここでは nested バイナリ全部を ad-hoc に揃えるために要る。

### 起動中判定に `pgrep` を使わない

`ps -Ao pid=,command=` を使う。**エージェントセッションを載せている当の app 本体プロセスが
`pgrep -f` に出てこない**（helper プロセスは出る）のを実測した（2026-09-23）。
`pgrep` で書くと「起動中なのに起動していないと判定して、動作中の bundle を差し替える」という
黙って壊れる経路になる。bundle 配下のプロセスが1つでもあれば拒否する。

## 失敗した更新の残骸を掃除する

失敗した更新は `/Applications/.<Name>.previous-<ns>.rollback` を残し、1回ごとに
app 1個ぶん（数百MB〜1.4GB）積み上がる。ドリフト調査のときに必ず見る:

```bash
du -sh /Applications/.*.rollback 2>/dev/null
```

削除はユーザーの確認を取ってから。復元用の退避なので、seat2 が正常に動いていることを
確かめてから消す。

## 席を増やすとき

`scripts/seat2_clone.sh` の `SEATS` に1行、`shim_env_for` にプロファイルの定義を足す。
`seat : 本体 app 名 : 生成名 : 表示名 : bundle identifier` の形。
プロファイルディレクトリは `$HOME` 起点で書き、絶対パスを埋めない。

## テスト

`scripts/tests/test_seat2_clone.py` が偽の app bundle を組み立てて、生成・shim の引数注入・
環境変数・scheme 書き換え・ドリフト検出・起動中の拒否を実物のスクリプトで確かめる。
`skill_lint.sh` の S8 から走る。実アプリには触らない。

```bash
python3 -m pytest -q ~/.agents/skills/own-seat2-build/scripts/tests
```

**この検査が見ていないもの**を承知しておく。`ps` を `pgrep` に戻しても偽 bundle の
テストは緑のまま通る — `pgrep` が取りこぼすのはエージェントセッションを載せている
当の app 本体プロセスだけで、テストが起動する普通のプロセスは見つかるため。
起動中判定を書き換えるときは、実機で seat2 を動かしたまま `--install` が拒否することを
手で確かめる。同様に「vendor 署名のまま起動すると SIGKILL」も偽 bundle では再現できないので、
`--deep` の有無は payload の署名識別子で代替検査している。

## 関連

- 二席目の起動方法そのもの（AppleScript ランチャー方式）は過去の設計で、
  更新を受け取れないことが分かって複製 + shim に移った。経緯は reinstall repo の
  `ISSUE-20260828-unify-desktop-second-seat-launchers-on-wrappers`。
- skill・設定の symlink 統一は `own-skill-commonize`。
