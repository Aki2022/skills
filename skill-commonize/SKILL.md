---
name: skill-commonize
description: >-
  複数のコーディングエージェント（Claude Code / Codex CLI / Antigravity・Gemini CLI 等）の
  設定ファイルとスキルを symlink で単一ソースに統一する規約と手順。正典は `.agents/`
  （グローバルは `~/.agents/`、リポジトリ単位はリポジトリルートの `.agents/`）。
  必ずこのスキルを使うこと: CLAUDE.md / AGENTS.md / GEMINI.md / skills.md などのエージェント
  設定ファイルや `.claude/skills` `.agents/skills` `.codex/skills` を**新規作成・編集・移動・削除**
  しようとする時、「設定を共通化／統一」「symlink で揃える」「新しいリポジトリにエージェント設定を入れる」
  と言われた時、あるいは既存の symlink 構成を壊しかねない操作（symlink を実体ファイルに置換、別の場所へ
  コピー作成、正典ディレクトリの削除）をしようとする時。グローバルでもリポジトリ単位でも適用される。
---

# Agent Config Symlink 統一

Claude Code・Codex CLI・Antigravity/Gemini CLI など複数のエージェントは、それぞれ別名の
設定ファイル（`CLAUDE.md` / `AGENTS.md` / `GEMINI.md`）と別ディレクトリのスキル
（`.claude/skills` / `.agents/skills` / `.codex/skills`）を読む。放置すると同じ内容が
複数箇所に分岐し、「どれが最新か分からない」状態になる。

これを防ぐため、**正典を 1 つ決め、他は正典への symlink にする**。これにより
どのパスを編集しても正典が更新され、全エージェントに即時反映される（常にフレッシュ）。

## 正典（single source of truth）の場所

| スコープ                | 正典                                         | symlink で向けるもの                                                 |
| ----------------------- | -------------------------------------------- | -------------------------------------------------------------------- |
| グローバル（home 直下） | `~/.agents/AGENTS.md`、`~/.agents/skills/`   | `~/.claude/CLAUDE.md`, `~/.codex/AGENTS.md`, `~/.claude/skills` ほか |
| リポジトリ単位          | `<repo>/AGENTS.md`、`<repo>/.agents/skills/` | `<repo>/CLAUDE.md`, `<repo>/.claude/skills` ほか                     |

`.agents/` を正典にする理由: Codex CLI がユーザースキルとして `~/.agents/skills` を公式に読み、
かつ `.agents` はツール非依存の中立な名前のため。

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
```

### 2. 正典を決める

上表の正典（`.agents/` 側）を採用する。正典がまだ無ければ、最も内容が充実した実体を
正典の場所へ移して正典にする。

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
```

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

## リポジトリ単位での注意（グローバルとの違い）

リポジトリの symlink は git にコミットされるため、追加の注意がある。

- **git は symlink を保存できる**（特殊 blob）。`git add CLAUDE.md` で symlink のままコミットされる。
  実体としてコミットされていないか `git cat-file -p :CLAUDE.md` 等で確認するとよい。
- **Windows 注意**: `core.symlinks=false` の環境では symlink が「リンク先パスを書いた
  ただのテキストファイル」として展開され壊れる。チームに Windows 利用者がいる場合は、
  symlink ではなく各ツールの「他ファイルを読む」設定（例: CLAUDE.md に `@AGENTS.md` を
  記載して取り込む方式）を検討する。
- **CI / 一部ツール**は symlink を追従しないことがある。重要な経路では追従を確認する。
- リポジトリの `.gitignore` / バックアップファイル（`*.bak_*` `*.old_*`）はコミットしない。

## クイックリファレンス

```
正典:   .agents/AGENTS.md          .agents/skills/
別名:   CLAUDE.md  → AGENTS.md      .claude/skills → .agents/skills
        .codex/AGENTS.md → ...      .codex/skills/<name> → .agents/skills/<name>
編集:   どの別名を編集しても正典が更新される（symlink を壊さない限り）
禁止:   symlink の実体化 / 別コピー作成 / 正典削除
```
