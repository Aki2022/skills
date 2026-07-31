---
name: origin-skill-commonize
description: >-
  複数のコーディングエージェント（Claude Code / Codex CLI / Antigravity・Gemini CLI 等）の
  設定ファイル、スキル、commands、非秘密のMCP設定を symlink で単一ソースに統一する規約と手順。正典は `.agents/`
  （グローバルは `~/.agents/`、リポジトリ単位はリポジトリルートの `.agents/`）。
  必ずこのスキルを使うこと: CLAUDE.md / AGENTS.md / GEMINI.md / skills.md などのエージェント
  設定ファイルや `.claude/skills` `.agents/skills` `.codex/skills`、複数account間で共有する
  commands / MCP設定を**新規作成・編集・移動・削除**
  しようとする時、「設定を共通化／統一」「symlink で揃える」「新しいリポジトリにエージェント設定を入れる」
  と言われた時、あるいは既存の symlink 構成を壊しかねない操作（symlink を実体ファイルに置換、別の場所へ
  コピー作成、正典ディレクトリの削除）をしようとする時。グローバルでもリポジトリ単位でも適用される。
  スキル品質の保守もこのスキルが所有する: スキルの静的チェック（lint）、スキル使用時の摩擦・
  不発・手戻りの記録（摩擦ログ FRICTION.md への追記）、「スキルを改善したい」「スキルの品質を
  チェックして」と言われた時もこのスキルを使う。
---

# Agent Config Symlink 統一

Claude Code・Codex CLI・Antigravity/Gemini CLI など複数のエージェントは、それぞれ別名の
設定ファイル（`CLAUDE.md` / `AGENTS.md` / `GEMINI.md`）と別ディレクトリのスキル
（`.claude/skills` / `.agents/skills` / `.codex/skills`）を読む。さらに同じ製品の複数accountは、
commandsやMCP設定を別々のconfiguration directoryに持ち得る。放置すると同じ内容が
複数箇所に分岐し、「どれが最新か分からない」状態になる。

これを防ぐため、**正典を 1 つ決め、他は正典への symlink にする**。これにより
どのパスを編集しても正典が更新され、全エージェントに即時反映される（常にフレッシュ）。

## 正典（single source of truth）の場所

| スコープ                   | 正典                                                                           | symlink で向けるもの                                                |
| -------------------------- | ------------------------------------------------------------------------------ | ------------------------------------------------------------------- |
| グローバル・ツール横断     | `~/.agents/AGENTS.md`、`~/.agents/skills/`                                     | `~/.claude/CLAUDE.md`、`~/.codex/AGENTS.md`、`~/.claude/skills`ほか |
| グローバル・ツール固有共有 | `~/.agents/<tool>/`配下。Claude例: `commands/`、非秘密の`mcp.json`             | primary/secondaryを含む各tool configuration directoryの対応path     |
| リポジトリ単位             | `<repo>/AGENTS.md`、`<repo>/.agents/skills/`、必要なら`<repo>/.agents/<tool>/` | `<repo>/CLAUDE.md`、`<repo>/.claude/skills`、tool固有aliasほか      |

`.agents/` を正典にする理由: Codex CLI がユーザースキルとして `~/.agents/skills` を公式に読み、
かつ `.agents` はツール非依存の中立な名前のため。

## 共有設定とaccount stateの境界

このSkillは**共有する静的設定とsymlink topology**を所有する。shell、Home Manager、各製品の
account切替functionはsymlinkを作成・更新せず、選択したconfiguration directoryから利用する。

| Classification           | Examples                                                     | Policy                                                                |
| ------------------------ | ------------------------------------------------------------ | --------------------------------------------------------------------- |
| Cross-agent shared       | instructions、skills                                         | `.agents/`直下を正典にして各agentからsymlinkする                      |
| Tool-specific shared     | Claude commands、非秘密のMCP定義                             | `.agents/<tool>/`を正典にして、そのtoolの全accountから直接symlinkする |
| Account-specific mutable | 認証、session、history、project state、log、cache、telemetry | accountごとのconfiguration directoryに分離し、symlinkしない           |
| Secret                   | token、Cookie、private endpoint、MCP credential              | Gitと共有正典へ置かず、Bitwardenまたは実行時環境から注入する          |

primary account directoryをsecondary accountの正典にしない。例えば
`~/.claude-seat2/commands → ~/.claude/commands`ではなく、両方を
`~/.agents/claude/commands`へ直接向ける。これによりprimary directoryの移動・削除が
secondaryへ連鎖しない。

MCP設定は値を確認せず機械的に正典化しない。secretを含まないことを確認できる構造だけを
`.agents/<tool>/`へ置き、credentialは参照名または環境変数だけにする。判定できない場合は移動を止め、
既存fileを維持したままhuman reviewを求める。

## 不変条件（これを破ると分岐が復活する）

これは常に守ること。symlink 構成を前提に動く。

1. **symlink 経由の編集は正しい。** `~/.claude/CLAUDE.md` や `.claude/skills/foo/SKILL.md` を
   編集してよい — それは正典を編集することになり、全エージェントに反映される。`#` キーでの
   書き戻しも同様に正典へ届く。
2. **symlink を実体ファイル／ディレクトリに置き換えない。** `rm` してから `Write` で作り直す、
   といった操作は分岐を復活させる。編集は in-place（symlink を保ったまま中身を書く）で行う。
   エディタによっては「保存時に symlink を置換」する設定があるため注意。
3. **正典の外に新しいコピーを作らない。** 「念のため別名でも置いておく」はやらない。
4. **正典ディレクトリ（`.agents/`）を安易に削除しない。** 全エージェントに波及する。
5. **判断に迷う・正典が見つからない場合は破壊操作の前に確認する。**
6. **account固有の可変stateを共有しない。** 認証、session、history、cache等をsymlink対象にしない。
7. **symlinkをbackupとみなさない。** 正典自体を秘密を含まないprivate Git repositoryまたは
   同等のversioned backupで復元可能にする。

迷ったら、まず対象パスが symlink かどうかを確認する:
`ls -l <path>` / `readlink <path>`。symlink なら不変条件 1〜2 に従う。

## セットアップ手順（新規にグローバル or リポジトリを統一する）

棚卸し → 正典決定 → 内容統合 → バックアップ → symlink 化 → 検証、の順で進める。

### 1. 棚卸し（inventory）

対象スコープ内の設定ファイルとスキルディレクトリを列挙し、それぞれ
**実体 / symlink / 不在** を判定する。

```bash
# 設定ファイル例（リポジトリルートで）
for f in AGENTS.md CLAUDE.md GEMINI.md; do
  if [ -L "$f" ]; then echo "$f: symlink → $(readlink "$f")";
  elif [ -e "$f" ]; then echo "$f: 実体 ($(wc -l < "$f") 行)";
  else echo "$f: 不在"; fi
done

# スキルディレクトリの symlink 状態を確認
for d in .agents/skills .claude/skills .codex/skills; do
  if [ -L "$d" ]; then echo "$d: symlink → $(readlink "$d")";
  elif [ -d "$d" ]; then echo "$d: 実体ディレクトリ ($(ls "$d" 2>/dev/null | wc -l | tr -d ' ') 個)";
  else echo "$d: 不在"; fi
done

# Claude複数accountの共有静的設定を確認（内容は読まない）
for p in \
  ~/.agents/claude/commands ~/.agents/claude/mcp.json \
  ~/.claude/commands ~/.claude/mcp.json \
  ~/.claude-seat2/commands ~/.claude-seat2/mcp.json
do
  if [ -L "$p" ]; then echo "$p: symlink → $(readlink "$p")";
  elif [ -d "$p" ]; then echo "$p: 実体ディレクトリ";
  elif [ -f "$p" ]; then echo "$p: 実体ファイル";
  else echo "$p: 不在"; fi
done
```

**棚卸し時の判定ルール（スキル）:**

- `.agents/skills/` に中身があり、`.claude/skills` が **不在または実体ディレクトリ** → symlink 化が必要
- `.claude/skills → .agents/skills` の symlink が存在する → 正常、対応不要
- `.claude/skills` が不在でも `.agents/skills/` が空なら → 対応不要

`.claude/skills` の不在は「問題なし」ではなく、`.agents/skills/` の中身と合わせて判断すること。
中身があるのに symlink がなければ、Claude Code がプロジェクトスキルを読めない可能性がある。

commandsやMCP設定も同様に、複数accountのうち一つが実体で他がそこへのsymlinkなら未統一と判定する。
`~/.agents/<tool>/`正典へ直接向いて初めて統一済みとする。ただしMCPはsecret-safe確認前に移動しない。

### 2. 正典を決める

上表の正典（`.agents/` 側）を採用する。正典がまだ無ければ、最も内容が充実した実体を
正典の場所へ移して正典にする。tool固有の共有設定は`.agents/<tool>/`配下へ置き、
別toolへ誤って公開しない。

MCP候補は内容のsecret-safe判定とcredential分離が終わるまで正典へ移さない。commandsが空でも、
複数accountで将来分岐させない必要がある場合は空の正典directoryを作ってよい。

### 3. 内容を統合する（分岐がある場合）

複数の実体が**異なる内容**を持つ場合は、機械的に上書きせず差分を確認してから統合する:

```bash
diff <(cat AGENTS.md) <(cat CLAUDE.md)
```

- ツール名のハードコード（`# Claude AI 設定` 等）は中立な見出し（`# AI 設定`）にする。
- 特定ツール向けの記述（例「Codex は補助」）も、他ツールが読んで無害なら残してよい。
- 判断が割れる差分はユーザーに提示して選んでもらう。

### 4. バックアップ

破壊操作の前に必ず退避する。symlink 構造ごと保持するため `cp -a` を使う。

```bash
D=$(date +%Y%m%d)
cp -a AGENTS.md "AGENTS.md.bak_$D" 2>/dev/null || true
```

### 5. symlink 化

`scripts/unify_config.sh` を使うと、バックアップ・分岐検出・symlink 作成・検証を安全に行える:

```bash
# ファイル: CLAUDE.md を AGENTS.md（正典）へ向ける
bash scripts/unify_config.sh AGENTS.md CLAUDE.md

# ディレクトリ: .claude/skills を .agents/skills（正典）へ向ける
bash scripts/unify_config.sh .agents/skills .claude/skills

# Claude固有共有: primary/secondaryを同じ中立な正典へ直接向ける
bash scripts/unify_config.sh ~/.agents/claude/commands \
  ~/.claude/commands ~/.claude-seat2/commands
bash scripts/unify_config.sh ~/.agents/claude/mcp.json \
  ~/.claude/mcp.json ~/.claude-seat2/mcp.json
```

MCPの例は、正典fileがsecret-safeであることを人が確認した後だけ実行する。

手で行う場合（中身を理解した上で）:

```bash
mv CLAUDE.md CLAUDE.md.old_$(date +%Y%m%d)   # 実体を退避
ln -s AGENTS.md CLAUDE.md                      # 同階層なら相対パスでよい
```

- **同一ディレクトリ内**（`AGENTS.md` ↔ `CLAUDE.md`）は相対パス（`ln -s AGENTS.md CLAUDE.md`）。
- **ディレクトリをまたぐ**（`~/.claude/skills` → `~/.agents/skills`）は絶対パスが安全。

### 6. 検証

```bash
ls -l CLAUDE.md                       # → AGENTS.md を指していること
head -1 CLAUDE.md && head -1 AGENTS.md  # 同一内容が見えること
```

フレッシュ性テスト: 正典を 1 行だけ一時編集し、別名側から同じ変更が見えることを確認して元に戻す。

複数accountでは、各aliasがprimary account経由ではなく`.agents/<tool>/`正典へ直接解決されることを
`readlink`で確認する。認証・session・history・cacheがsymlinkでないことも確認する。

## リポジトリ単位での注意（グローバルとの違い）

### AGENTS.md の2種類の役割を区別する

リポジトリに `AGENTS.md` がある場合、その内容が何かを確認すること。

- **プロジェクト固有ルール**（ファイル管理方針・ディレクトリ構造・ワークフロー等）→ この
  リポジトリ専用の内容。残す価値がある。
- **全般的な AI 行動設定**（口調・ツール選択・並列実行方針等）→ グローバルの
  `~/.agents/AGENTS.md` で管理すべき内容が誤ってリポジトリに置かれている可能性がある。

### リポジトリ CLAUDE.md は必須ではない

`~/.claude/CLAUDE.md → ~/.agents/AGENTS.md` のグローバル symlink が既に設定済みなら、
リポジトリに `CLAUDE.md` を作らなくても Claude Code はグローバル設定を読む。

リポジトリに `CLAUDE.md` を作る（`AGENTS.md` への symlink）価値があるのは:

- そのリポジトリ専用の `AGENTS.md` があり、Claude Code に自動ロードさせたい場合
- Codex CLI と Claude Code 両方でプロジェクト固有ルールを共有したい場合

不要な場合: グローバル設定で十分で、リポジトリに余計なファイルを増やしたくない場合。

### symlink が git にコミットされる点

リポジトリの symlink は git にコミットされるため、追加の注意がある。

- **git は symlink を保存できる**（特殊 blob）。`git add CLAUDE.md` で symlink のままコミットされる。
  実体としてコミットされていないか `git cat-file -p :CLAUDE.md` 等で確認するとよい。
- **Windows 注意**: `core.symlinks=false` の環境では symlink が「リンク先パスを書いた
  ただのテキストファイル」として展開され壊れる。チームに Windows 利用者がいる場合は、
  symlink ではなく各ツールの「他ファイルを読む」設定（例: CLAUDE.md に `@AGENTS.md` を
  記載して取り込む方式）を検討する。
- **CI / 一部ツール**は symlink を追従しないことがある。重要な経路では追従を確認する。
- リポジトリの `.gitignore` / バックアップファイル（`*.bak_*` `*.old_*`）はコミットしない。

## スキル品質の保守

symlink 統一と同じくこの Skill が所有する。原則は「**決定論的チェックは機械が、改善判断は
需要駆動で人間が**」。測定データなしの定期自動改善はやらない — 改悪とチャーンの温床になり、
使っていないスキルの改善にトークンを浪費するため。

### 静的チェック（決定論的・LLM 不使用）

`scripts/skill_lint.sh` を実行する。SKILL.md の存在、frontmatter の name/description、
name とディレクトリ名の一致、同梱リソース参照（scripts/ references/ assets/）の実在、
壊れた symlink を exit code で判定する。

```bash
bash ~/.agents/skills/origin-skill-commonize/scripts/skill_lint.sh
```

実行タイミング: スキルの新規作成・編集・移動・削除の直後（この Skill の作業の一部として）。
サードパーティ由来スキルの FAIL は情報として報告し、勝手に修正しない — 上流更新で上書き
され得るため、直すのは origin-* など自前スキルのみ。

### 摩擦ログ（skill friction log）

スキルの実使用で気づいた摩擦 — 発火しなかった・指示が曖昧で手戻りした・記述が現実と
ズレていた — を `~/.agents/skills/FRICTION.md` に 1 件 1 行で追記する。どのセッションの
どのエージェントでも、摩擦に気づいたらその場で追記してよい（追記は非破壊・承認不要）。

```
- 2026-07-31 | origin-pptx | 画像生成の保存先指示が曖昧で2回手戻り
```

- **記録と改善を分離する。** ログへの追記は自由、スキル本文の改変はこのログを根拠に
  人間が判断してから（このSkillの手順で）行う。
- **改善は需要駆動。** 同一スキルに摩擦が複数件溜まったら、skill-creator の eval 付き
  改善ループを回す。溜まった摩擦行がそのままテストプロンプトの種になる。
- **定期実行するならトリアージまで。** 定期タスク化してよいのは「ログの溜まり具合を
  提示する」ことまでで、改善の自動実行はしない。

## クイックリファレンス

```
正典:   .agents/AGENTS.md          .agents/skills/
        .agents/claude/commands/    .agents/claude/mcp.json（非秘密のみ）
別名:   CLAUDE.md  → AGENTS.md      .claude/skills → .agents/skills
        .codex/AGENTS.md → ...      .codex/skills/<name> → .agents/skills/<name>
        各Claude accountのcommands/mcp.json → .agents/claude/...
分離:   auth / session / history / project state / log / cache
編集:   どの別名を編集しても正典が更新される（symlink を壊さない限り）
禁止:   symlink の実体化 / 別コピー作成 / 正典削除 / secretやaccount stateの共有
```
