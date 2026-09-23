---
name: own-skill-commonize
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
  スキルの静的チェック（lint）もこのスキルが所有する。「スキルを改善したい」「スキルの品質を
  チェックして」と言われた時もこのスキルを使う。トラブル・摩擦の**記録**は所有しない
  （`own-trouble-log` が所有する）。
---

# Agent Config Symlink 統一

Claude Code・Codex CLI・Antigravity/Gemini CLI など複数のエージェントは、設定ファイルや
skill の探索規則が一致しない。Claude Code v2.1.277 以降はリポジトリの `AGENTS.md` を直接
読めるが、グローバル設定・skill directory・非対応sessionには依然としてtool固有aliasが要る。
さらに同じ製品の複数accountは、commandsやMCP設定を別々のconfiguration directoryに持ち得る。
放置すると同じ内容が複数箇所に分岐し、「どれが最新か分からない」状態になる。

これを防ぐため、**正典を 1 つ決め、他は正典への symlink にする**。これにより
どのパスを編集しても正典が更新され、全エージェントに即時反映される（常にフレッシュ）。

> **重要（2026-08-28 更新）**: グローバルの `~/.agents` **直下の静的ファイル 27 件**は
> nix-darwin / Home Manager の管理下に入り、`/nix/store/...` への **root 所有・`0444` の
> symlink** として描画されるようになった。**この 27 件については上の「どのパスを編集しても
> 正典が更新される」は成り立たない** — エイリアス経由どころか `~/.agents` 側で直接編集しても
> permission error になる。実際の編集元は下表を参照。
> **`~/.agents/skills/` は nix の管理外**（独立した git repo・書き込み可）なので、
> skill については従来どおり `~/.agents/skills/` が正典のまま。

## 正典（single source of truth）の場所

| スコープ                   | 正典                                                                                                                                                              | symlink で向けるもの                                                    |
| -------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------- |
| グローバル・skill          | `~/.agents/skills/`（**nix 管理外・書き込み可**。ここは従来どおり）                                                                                               | `~/.claude/skills`、`~/.codex/skills` ほか                              |
| グローバル・静的設定 27 件 | **nix-darwin flake repo の `home/agent-config/<相対パス>`**（実ファイル・書き込み可・git 管理下）。`~/.agents/AGENTS.md` 等はその**描画先**であって編集元ではない | `~/.claude/CLAUDE.md`、`~/.codex/AGENTS.md`、`~/.gemini/GEMINI.md` ほか |
| グローバル・ツール固有共有 | `~/.agents/<tool>/`配下。Claude例: `commands/`、非秘密の`mcp.json`。Codex例: `hooks.json`、`agents/`（`AGENTS.override.md` は**共有してはいけない** — 下記）      | primary/secondaryを含む各tool configuration directoryの対応path         |
| リポジトリ単位             | `<repo>/AGENTS.md`、`<repo>/.agents/skills/`、必要なら`<repo>/.agents/<tool>/`                                                                                    | `.claude/skills` 等。`CLAUDE.md` は互換性が必要な場合だけ               |

`.agents/` を正典にする理由: `.agents` はツール非依存の中立な名前で、どの道具にも属さないため。

**どの道具も正典を直接は読まない。** 各道具は自分の配布先だけを見る（2026-09-22 に実装で確認）。

| 道具 | 読む場所 | 根拠 |
| --- | --- | --- |
| Claude Code | `~/.claude/skills` | 正典への root symlink 1本 |
| Codex | `$CODEX_HOME/skills` | バイナリ内の記述 `Installs into $CODEX_HOME/skills/<skill-name> (defaults to ~/.codex/skills)`。3席は `shell.nix` が `CODEX_HOME` を切り替える |
| Gemini | `~/.gemini/config/skills` | per-skill symlink |

**配布先の一覧は持たない。** 配線は必ず正典への symlink なので、
`scripts/audit_skill_wiring.py` が**正典を指す symlink を辿って**配線先を見つける。
新しい道具を配線した瞬間から対象になり、一覧に足す作業が要らない。

一覧を持たない理由は実測にある。2026-09-22 に配布先を `alias-roots.txt` に書き出したが、
**人が書くものは書き落とす** — 一覧は6件で、正典を指す配線先は**13件**あった。
落ちていたのは `~/.claude-private/skills`・`~/.claude-seat2/skills`・
`~/.gemini/antigravity-cli/skills`（いずれも生きていた）と古いバックアップ4件。
一覧方式は最初から半分しか見ていなかったので、2026-09-23 に撤去した。

## skill の供給源は正典のみ（第三者 skill のミラー規約）

**skill を各エージェントへ供給する経路は正典 `~/.agents/skills/` の1本だけ**とする
（`ADR-20260906-skill-supply-source-is-canon-only`・biz_ops）。根拠は実測: Codex CLI は `$CODEX_HOME/skills` しか読まず、Claude Code の plugin（`~/.claude/plugins/`）
と CLI 同梱 built-in は **Claude 専用の供給路**である。plugin 由来の skill を使い続けると
Claude と Codex・他 LLM の skill 群は必ずズレる（2026-09-05〜06 に frontend-design 3重・pdf 3重・
plugin/built-in 2重6件を実測）。

1. **使いたい第三者 skill は上流名で正典へミラーする。** ディレクトリ名＝上流の frontmatter `name`。
   中身は上流と**バイト同一**に保ち、自前改変を入れない（改変が要るなら fork として別名の自前 skill にする）。
2. **ミラーには取得メタを付け、`mirrors.yaml`（正典リポジトリ root）に記録する。** 上流の識別子・
   上流の版（commit sha 等）・取得日・ローカルで比較可能な上流コピーのパス（あれば）。
   ミラーディレクトリ自体には余計なファイルを足さない（バイト同一を保つため）。
3. **skill しか供給していない Claude plugin は無効化する。** 無効化できない built-in との2重は残るが、
   正典側を上流に同期し、`scripts/check_mirrors.sh` の同値検査（ローカル上流コピーとのハッシュ比較）と
   鮮度検査（取得日からの経過）で**ズレを検出できる形**にする。Claude の一覧に同内容が2行載るのは
   「ズレ」ではなく「重複表示」で、本規約には反しない。
4. **`~/.agents/<tool>/skills/` に skill を置かない（第三者ミラーも自前も）。**
   ここは**第3の所属**になるが、命名規則は所属を2つ（3語＝グローバル / 4語＝repo 固有）しか
   定義していない。置くと名前から所属が読めなくなり、`skill_lint.sh` の走査対象にも入らないため
   **規則違反ではなく規則の対象外**になる。tool 専用性は**置き場所ではなく description** で表現する
   （発火を決めるのは description であることを 2026-09-21 に実測済み）。
   実例: `origin-august-luna-loop` は `~/.agents/codex/skills/` に居たため、2026-08-30 に退役させても
   実体が git 未追跡のまま残り、実環境の再キャプチャで復活した。2026-09-22 に `own-luna-run` として
   正典へ移し、この root を廃止した。
   `~/.agents/vibe-guard/skills/own-vibeguard-harden` は nix 管理下にも置かれているが、
   **正典にも同名で存在し、3配線先すべてから見える**（2026-09-23 実測）。
   したがって「語数＝所属」の例外ではない。nix 側は配布のための複製で、正典が本体。
5. 上流の更新への追随は**人間が起動する**（鮮度検査が警告したら再取得）。自動追随はしない。

`mirrors.yaml` の形式:

```yaml
mirrors:
  - dir: frontend-design # 正典内のディレクトリ名（＝上流 name）
    upstream: anthropics/skills # 上流の識別子（GitHub owner/repo 等）
    upstream_path: skills/frontend-design
    upstream_version: <commit sha> # 上流の版。取得時点で分かるもの
    fetched_at: 2026-09-06
    local_copy: ~/.claude/plugins/cache/<marketplace>/<plugin>/<version>/skills/frontend-design
    # local_copy が無い（built-in 等）場合は省略。同値検査は skip され鮮度検査のみ
```

### 静的設定 27 件の編集手順（nix 管理下）

1. nix-darwin flake repo の `home/agent-config/<相対パス>` を**直接編集する**（普通のファイル。
   nix コードを書く必要は無い — `home/agents.nix` がそのディレクトリを**再帰的に取り込む**ため、
   取り込み対象の増減を変えるとき以外は `.nix` に触らない）。
2. 同 repo の driver script（`nix-darwin.sh`。repo 直下の `scripts` ディレクトリにある）を
   `switch` 引数付きで実行して描画し直す。**このスキル配下のパスではない**ので、
   相対パス表記で書くと skill lint の S4（参照の実在検査）に引っかかる。
3. `~/.agents/<相対パス>` が新しい nix store のパスを指していることで反映を確認する。

**`~/.agents` 側で 27 件を git 追跡してはいけない**（2026-08-28 に追跡を解除済み）。
nix store のハッシュは `switch` のたびに変わるため、追跡すると `git status` に
typechange が出続けて本物の変更が埋もれる。正典側の repo が既に追跡しているので、
`~/.agents` での追跡は二重持ちにあたる。

## 自前 skill の命名規則

**許容する形は1つだけ。末尾は必ず動詞。** 語数が所属を表す。

```
3語   own-<対象>-<動作>            → グローバル正典 ~/.agents/skills/
4語   own-<repo>-<対象>-<動作>     → リポジトリ固有 <repo>/.agents/skills/
```

4語なら必ずリポジトリ固有、3語なら必ずグローバル。**名前だけ見てどこを編集するか分かる。**
**第3の所属は作らない** — `~/.agents/<tool>/skills/` に置くと名前から所属が読めなくなる（ミラー規約4）。
repo トークンは `trade` / `yorisoi` / `bizops` / `iscore` / `marketing`。

**禁止**: 動作でない名詞を末尾に置く形（`-runtime` `-loop` `-policy` `-style` `-setup`
`-cleanup` `-maintenance` `-routing` `-warehouse` 等）と、`動作-対象` の逆順。

**skill 名は description と並んで発火条件そのもの**で、動作の無い名前は「何をする skill か」を
名前から読めなくする。実害があった — `origin-design-runtime`（当時の名前。現 `own-web-design`） と `own-design-route` の
役割が名前から区別できず、人間が指摘して初めて分かった。名詞末尾が機能を隠していた実例は
ほかにも4件あり、いずれも description を読むと名前と実体が違っていた
（`-policy` の実体は振り分け、`-access` の実体は質問への回答、`-report` の実体は実装）。

### 動詞は allowlist で持つ

`references/naming-verbs.txt` に列挙した語だけを末尾に置ける。**denylist にしない** —
知らない名詞を黙って通すからで、それは「黙って間違う」側。allowlist なら新しい動詞が
要るときに落ちて止まり、人間が判断して足せる。

幅広い動詞（`handle` `work` `manage` `process` 等）は**足さない**。情報を持たない動詞は
「名前から何をするか読める」という目的を満たさず、禁じた形を動詞の面で復活させるだけ。
動作が複数ある skill にも動詞を1つ選び、選べないときは対象の切り方を見直す。

### 移行中の猶予リスト

`references/naming-exceptions.txt` は burn-down リスト。規則を入れた時点で既にあった
自前 skill を載せ、**改名するたび1行消す**。空になったら猶予は終わる。
**新規 skill をここに足してはいけない。**

猶予が要るのは、規則をそのまま適用すると既存の全 skill が即座に赤くなり、
`skill_lint.sh` は skill 編集のたびに走るので**永久に赤い検査**になるため。
赤が情報でなくなると本物の失敗が埋もれる。

### 改名するときは密結合クラスタごと

**skill は他の skill を名前で呼ぶ。** 特に frontmatter の `description` に書かれた名前は
発火の連鎖に直結し、取りこぼすと**エラーにならずに連鎖だけ壊れる**。

2026-09-20 の実測では、SKILL.md 内の相互参照が **178件**（うち frontmatter **42件**）あり、
デザイン系・ループ系・quarto 系が密に絡んでいた。**1本ずつ改名すると途中で参照が壊れる**ので、
互いを参照し合う塊は**一括で改名して1コミットにする**。

改名の1スライスに含めるもの:

1. ディレクトリ名（`git mv`）と frontmatter の `name:`
2. **他の skill からの参照**（frontmatter・body の両方。grep で旧名 0 件を確認）
3. symlink の張り直し（`~/.claude/skills` は正典を丸ごと指すが、Codex は skill ごと・3席ぶん）
4. docs・AGENTS.md・hooks・スラッシュコマンドからの参照
5. `naming-exceptions.txt` から該当行を削除
6. canary で**実際に発火するか**の確認

`skill_lint.sh` の S9 が (1) と規則違反を機械で見る。**(2)〜(4) は lint では見えない**ので、
grep で旧名が 0 件であることを別に確かめる。

### lint

`scripts/skill_lint.sh` の S9 が3点を検査する — 接頭辞が `own-` か、語数が所属と一致するか、
末尾が動詞か。第三者ミラー（`cloudflare-*` `google-*` 等）は対象外。

規則の根拠と却下案は biz_ops の `ADR-20260915-unify-skill-naming-to-object-action`。

`S10` は配線の監査で、`scripts/audit_skill_wiring.py` を lint から呼ぶ。
**孤立したスクリプトは誰も走らせない** — `check_global_topology.py` は実運用で値を埋めて
呼ぶ場所が 0 件のまま存在していた。同じ形にしない。lint は skill を触るたびに走るので、
「commonize を使うときに検査する」がそのまま実現する。

**実正典を lint したときだけ走る。** 配線は git の外にあり、`main` を見ても状態は分からない。
実際 2026-09-22 に `~/.gemini/config/skills` の配線が壊れていたとき、`main` は完全に正しかった
（`own-doc-update` は存在し、配線だけが旧名を指していた）。**中身が正しいことと
配線が正しいことは別々に確かめる必要がある。**

逆に worktree には検査対象が無い。配線が壊れるのは正典が変わった後であって、
worktree で編集している最中ではない。そこを指す symlink も存在しない。
検査しない場合は `S10: skip（…）` と必ず言う — 黙って飛ばすと緑が何を意味するか読めない。

単独でも走らせられる:

```bash
python3 ~/.agents/skills/own-skill-commonize/scripts/audit_skill_wiring.py
```

見るもの:

1. **個数が同じか** — 正典 N 件に対して配線先も N 件か
2. **不足が無いか** — 正典にあって配線先に無い名前
3. **余分が無いか** — 配線先にあって正典に無い名前（改名の置き去りがこれ）
4. **宙を指していないか**
5. **個別対処した skill**（正典の外を指す symlink）が配線先でも解決するか

**中身は照合しない。** 配線は正典の同じ実体を指すので原理的にズレない
（72 skill × 5 root = 360 通りを realpath で照合し別実体 0 件・2026-09-22 実測）。

退役マーカー（`_backup_` `_old_` `.bak_` `.orphaned_` `.disabled`）のついた配線先は
対象にしない。古い中身のまま残っているので、見ると永久に赤い検査になる。

**検査した配線先の数を常に出す。** 緑と「何も見ていない」を出力で区別するため。

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

## symlink が正しくても正典が届かないことがある

**配置の正しさは到達の証拠ではない。** 2026-08-28 に実測した失敗:
`~/.codex/AGENTS.md` は正典への symlink として正しく張られていたのに、
Codex は正典を 19 日間まったく読んでいなかった。

原因は Codex の解決規則。公式仕様は
"Codex reads `AGENTS.override.md` if it exists. Otherwise, Codex reads `AGENTS.md`."
——**override は追加ではなく置換**で、同じ階層の `AGENTS.md` を無効化する。
`~/.codex/AGENTS.override.md`（23 行・期限切れ）が存在したため、正典 50 行は
一度も読まれなかった。Codex の全 account が同じ symlink を共有していたので
全席が同時に影響を受けた。リポジトリ階層の `AGENTS.md` は正常に読まれ続けたため、
壊れているようには見えなかった。

したがって:

- **`AGENTS.override.md` を正典化・共有してはいけない。** 存在自体が
  「正典が読まれていない」ことを意味する。棚卸しでは symlink 先ではなく
  **存在の有無**を検査する。
- 他ツールにも同種の罠がありうる（Cursor は `.cursorrules` を Agent mode で
  読まない、Copilot は `copilot-instructions.md` と AGENTS.md の優先順位が
  未定義、など）。**新しいツールを正典に接続するときは、ファイル名だけでなく
  「置換規則の有無」を公式ドキュメントで確認する。**
- 最終的な確認は **canary** で行う。正典に一意な文字列を仕込み、各エージェントに
  復唱させて初めて到達が証明できる。symlink の検査は代理経路にすぎない。

## 不変条件（これを破ると分岐が復活する）

これは常に守ること。symlink 構成を前提に動く。

1. **symlink 経由の編集が正しいのは、正典が書き込み可能な場合だけ。**
   `.claude/skills/foo/SKILL.md` のように **nix 管理外**の正典を指す alias は、
   編集すれば正典が更新され全エージェントへ反映される。
   **nix 管理下の 12 件は違う。** `~/.claude/CLAUDE.md` の実体は `/nix/store/...` の
   `r--r--r--`（所有者 `nixbld`）で、alias 経由でも直接でも書けない。
   **`#` キーによるメモリ書き戻しもここでは失敗する** — 2026-08-28 に実測。
   これらの編集元は上表のとおり nix-darwin flake repo 側で、反映には `switch` が要る。
   どちらの側かは `ls -lL <path>` で実体の権限を見れば分かる。

   **編集したまま `switch` を忘れると、宣言と実環境がズレる。** これは
   `system_state.py --check` が「active or prospective system state drift」として
   赤で報告する（2026-08-28 に実測で確認）ので、その赤を放置しないこと。

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

# 複数accountの共有静的設定を確認（内容は読まない）。
# account homeの命名規約: 無印=main、-seat2=二席目、-private=旧personal
for home in ~/.claude ~/.claude-seat2 ~/.claude-private; do
  for p in "$home/CLAUDE.md" "$home/skills" "$home/commands" "$home/mcp.json"; do
    if [ -L "$p" ]; then echo "$p: symlink → $(readlink "$p")";
    elif [ -d "$p" ]; then echo "$p: 実体ディレクトリ";
    elif [ -f "$p" ]; then echo "$p: 実体ファイル";
    else echo "$p: 不在"; fi
  done
done
for home in ~/.codex ~/.codex-seat2 ~/.codex-private; do
  for p in "$home/AGENTS.md" "$home/agents" "$home/hooks.json"; do
    if [ -L "$p" ]; then echo "$p: symlink → $(readlink "$p")";
    elif [ -d "$p" ]; then echo "$p: 実体ディレクトリ";
    elif [ -f "$p" ]; then echo "$p: 実体ファイル";
    else echo "$p: 不在"; fi
  done
  # AGENTS.override.md は「あってはいけない」側。存在すれば正典が読まれていない
  [ -e "$home/AGENTS.override.md" ] && echo "$home: AGENTS.override.md が正典を隠している"
  # Codexのskillsはdirectory全体をsymlinkせず、skillごとのsymlinkを並べる
  # （codexが skills/.system をaccount stateとして書き込むため）
  links=$(find "$home/skills" -maxdepth 1 -type l 2>/dev/null | wc -l | tr -d ' ')
  echo "$home/skills: per-skill symlink ${links}本"
done
```

Codexのper-skill symlinkは正典 `~/.agents/skills/` の各スキルへ直接向ける。
本数が `~/.codex`（main）と他accountで食い違えば未統一と判定する。

**棚卸し時の判定ルール（スキル）:**

- `.agents/skills/` に中身があり、`.claude/skills` が **不在または実体ディレクトリ** → symlink 化が必要
- `.claude/skills → .agents/skills` の symlink が存在する → 正常、対応不要
- `.claude/skills` が不在でも `.agents/skills/` が空なら → 対応不要

`.claude/skills` の不在は「問題なし」ではなく、`.agents/skills/` の中身と合わせて判断すること。
中身があるのに symlink がなければ、Claude Code がプロジェクトスキルを読めない可能性がある。

commandsやMCP設定も同様に、複数accountのうち一つが実体で他がそこへのsymlinkなら未統一と判定する。
`~/.agents/<tool>/`正典へ直接向いて初めて統一済みとする。ただしMCPはsecret-safe確認前に移動しない。

**棚卸し時の判定ルール（リポジトリ指示）:**

- `AGENTS.md` があり `CLAUDE.md` が不在 → Claude Code v2.1.277+ の対応sessionでは正常
- `CLAUDE.md → AGENTS.md` → 互換alias。直接読込できるsessionでは冗長だが内容は二重読込されない
- `CLAUDE.md` が別内容の実体 → defaultでは`AGENTS.md`が読まれないため、意図的な分離か確認する
- `CLAUDE.local.md` がある → defaultでは`AGENTS.md`が読まれない。両方必要なら
  `claude-md-and-agents-md`をuser/managed settingsで選ぶ

直接読込の対象はリポジトリの`AGENTS.md`と`.claude/AGENTS.md`で、`~/.agents/AGENTS.md`や
`.agents/`配下ではない。したがってグローバル`~/.claude/CLAUDE.md` aliasは維持する。

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

### 5. 必要な alias だけ作る

リポジトリでは `AGENTS.md` 単独を標準形とする。Claude Code v2.1.277未満、Bedrock等の
feature flagを取得しないsession、telemetry無効、built-in `agents-md` plugin無効なども支える場合だけ、
`CLAUDE.md`から`@AGENTS.md`をimportするかsymlinkを残す。`AGENTS.md`対応とskill探索は別なので、
`.agents/skills`に中身がある場合の`.claude/skills` aliasは引き続き必要である。

`scripts/unify_config.sh` を使うと、バックアップ・分岐検出・symlink 作成・検証を安全に行える:

```bash
# 互換性が必要な場合だけ: CLAUDE.md を AGENTS.md（正典）へ向ける
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

互換aliasがある場合は`readlink CLAUDE.md`と内容一致を確認する。`AGENTS.md`を直接読む構成では、
interactive sessionの`AGENTS.md loaded`起動表示、または一意なcanaryの復唱で確認する。
直接読込された`AGENTS.md`は`/context`のMemory filesに表示されず、`InstructionsLoaded` hookも
発火しないため、これらを不合格の根拠にしない。canaryは確認後に必ず戻し、正典のhashが元と
一致することを確かめる。

複数accountでは、各aliasがprimary account経由ではなく`.agents/<tool>/`正典へ直接解決されることを
`readlink`で確認する。認証・session・history・cacheがsymlinkでないことも確認する。

## リポジトリ単位での注意（グローバルとの違い）

### AGENTS.md の2種類の役割を区別する

リポジトリに `AGENTS.md` がある場合、その内容が何かを確認すること。

- **プロジェクト固有ルール**（ファイル管理方針・ディレクトリ構造・ワークフロー等）→ この
  リポジトリ専用の内容。残す価値がある。
- **全般的な AI 行動設定**（口調・ツール選択・並列実行方針等）→ グローバルの
  `~/.agents/AGENTS.md` で管理すべき内容が誤ってリポジトリに置かれている可能性がある。

### リポジトリ CLAUDE.md は互換用

Claude Code v2.1.277以降の対応sessionは、作業ディレクトリと祖先に`CLAUDE.md`、
`.claude/CLAUDE.md`、`CLAUDE.local.md`が無いとき、リポジトリの`AGENTS.md`を直接読む。
したがってrepo固有ルールの共有だけが目的なら`AGENTS.md`単独でよい。

移行対象repoの構造は、`scripts/check_repo_instruction_topology.py`で機械検査する。
native modeは通常形（正規ファイルの`AGENTS.md`、`CLAUDE.md`なし）を要求し、compat modeは
旧Claudeや非対応provider向けの`CLAUDE.md -> AGENTS.md` symlinkだけを許可する。実体の
`CLAUDE.md`、壊れたlink、`AGENTS.override.md`、`CLAUDE.local.md`はFAILとする。

```bash
python3 ~/.agents/skills/own-skill-commonize/scripts/check_repo_instruction_topology.py \
  --repo <repo-root> --mode native
python3 ~/.agents/skills/own-skill-commonize/scripts/check_repo_instruction_topology.py \
  --repo <compat-repo-root> --mode compat
```

`CLAUDE.md`を残すのは、非対応sessionを支える場合、Claude固有指示を加える場合、または
`InstructionsLoaded` hookや`/context`表示が必要な場合である。単一正典を保つには、Claude固有内容が
無ければsymlink、内容を足すなら先頭の`@AGENTS.md` importを使う。

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
canonical skill を指す cross-skill 参照の実在、壊れた symlink を exit code で判定する。
name とディレクトリ名の不一致は自前・第三者を問わず
`FAIL` とする — 第三者 skill は上流名でミラーする規約（上記）により不一致は起きない。
加えて、同一 root 内の frontmatter `name` の完全重複と `source-command-*` の対応先を
`WARN S7` として報告し、隣接する `codex/hooks.json` の bash 参照先が存在しない場合は
`FAIL S6` とする。description の意味的類似・発火競合は決定論的 lint の対象外で、トリアージで扱う。

さらに `S8` として、各 skill 直下の `scripts/tests/` にある `test_*.py` を
`python3 -m pytest -q <skill>/scripts/tests` で実行する。テストが赤ければ `FAIL S8`
として exit code に反映する（テストが赤いまま skill を編集させないため、warn ではなく fail）。
`pytest` を持つ skill が 1 件も無ければ「S8: pytest を持つ skill が無い」と明示し、対象0件を
黙って合格扱いにしない。`pytest` 自体が使えない環境、または環境変数
`SKILL_LINT_SKIP_PYTEST=1` を設定した場合は S8 を skip するが、その旨を必ず出力する。

```bash
bash ~/.agents/skills/own-skill-commonize/scripts/skill_lint.sh
```

グローバル alias の形状は `scripts/check_global_topology.py` で別に検査する。
正典 root と各 account root を毎回明示し、推測で別名を増減させない。Claude と
Antigravity は root 自体が正典への directory symlink、Codex と Gemini は regular
directory 内の各 skill が正典への per-skill symlink でなければならない。
Codex の `.system` と、`--local-only` で指定した seat-local path だけが例外で、
Codex-only adapted skill は `--codex-adapted` で明示する。未登録の実体コピー、壊れた
link、root 全体の symlink化は `FAIL` とする。

```bash
python3 ~/.agents/skills/own-skill-commonize/scripts/check_global_topology.py \
  --canonical <global-skills-root> \
  --claude <claude-skills-root> --codex <codex-skills-root> \
  --gemini <gemini-skills-root> --antigravity <antigravity-skills-root> \
  --codex-adapted <codex-adapted-skill-dir> \
  --local-only <seat-local-skill-path>
```

この checker は alias と skill の配置だけを検査し、認証・session・history・cache・
plugin cache・Codex `.system` の中身は読んだり変更したりしない。canary による実読込確認は
別途人間が行う。

per-skill rootの不足を直すときは`scripts/sync_per_skill_aliases.py`を使う。正典と対象rootを
明示し、最初はdry-runする。`--apply`で不足linkを作り、壊れた未登録linkを消す場合だけ
`--prune-stale`も付ける。有効な外部linkや実体entryがあれば上書きせず`CONFLICT`で停止する。
Codexの`.system`やadapted skillなど正当な例外は、直下名を`--ignore-entry`で1件ずつ明示する。

```bash
python3 scripts/sync_per_skill_aliases.py \
  --canonical <global-skills-root> --alias-root <gemini-or-codex-skills-root>
python3 scripts/sync_per_skill_aliases.py \
  --canonical <global-skills-root> --alias-root <gemini-or-codex-skills-root> \
  --apply --prune-stale \
  --ignore-entry <product-owned-or-adapted-entry>
```

実行タイミング: スキルの新規作成・編集・移動・削除の直後（この Skill の作業の一部として）。
サードパーティ由来スキル（ミラー）の FAIL は**中身を手で直さず、上流から再取得して同期する**（ミラー規約: バイト同一）。同値・鮮度は `scripts/check_mirrors.sh` で確認する。第三者本文を手で直さず、修正対象は own-* の自前スキルのみとする。

使用実績を定量化するときは `scripts/measure_skill_usage.py` を使う。Claude/Codex の履歴 root と
正典 root を明示して、実ユーザーメッセージに残る `$skill` / `/skill` だけを account 別に集計する。
本文・path・session ID は出力せず、証拠不足は `unknown` にする。出力保存は `--output` または
`--append` を指定した場合だけで、スケジューラの導入や削除判断はこのコマンドの責務に含めない。
使い方と schema は `docs/guides/GUIDE-skill-usage-metrics.md` を参照する。

### トラブル・摩擦の記録は所有しない

記録先は `own-trouble-log`（保管ルートは `ORIGIN_TROUBLE_LOG_ROOT`）へ移した。
`~/.agents/skills/FRICTION.md` は廃止（ファイル自体は移行の道標として期限付きで残す）。

移した理由: 記録対象は skill 起因に限らず、skill を使っていない場面の作業規律の
欠落も含む。`FRICTION.md` の 1 行形式は `<skill-name>` を必須にしており、
そのようなトラブルを記録できなかった。保管ルートも `~/.agents/skills` の外になるため、
このスキルの scope と一致しない。

**受け渡しの境界。** 「skill の記述が現実とズレていた」型のトラブルは、
**集めるのが `own-trouble-log`・直すのがこのスキル**。境界を書かないと
どちらも動かないケースが生じる。

改善の原則は変わらない。**記録と改善を分離し、改善は需要駆動で人間が判断してから**
このスキルの手順で行う。測定データなしの定期自動改善はやらない。
同一スキルに記録が複数件溜まったら、skill-creator の eval 付き改善ループを回す
（`own-trouble-log` の `skills` フィールドで絞り込める）。

## クイックリファレンス

```
正典:   .agents/AGENTS.md          .agents/skills/
        .agents/claude/commands/    .agents/claude/mcp.json（非秘密のみ）
        .agents/codex/hooks.json    .agents/codex/agents/
repo:   AGENTS.md 単独が標準         CLAUDE.md → AGENTS.md は互換用のみ
別名:   ~/.claude/CLAUDE.md → ~/.agents/AGENTS.md（globalでは引き続き必要）
        .claude/skills → .agents/skills
        .codex/AGENTS.md → ...      .codex/skills/<name> → .agents/skills/<name>
        各Claude accountのcommands/mcp.json → .agents/claude/...
        各Codex accountのhooks.json/agents → .agents/codex/...
禁止:   AGENTS.override.md（置換仕様。あると正典 AGENTS.md が読まれない）
account home: 無印=main、-seat2=二席目、-private=旧personal（Claude/Codex共通）
分離:   auth / session / history / project state / log / cache
        （Claude settings.json・Codex config.tomlはaccount固有。symlinkしない）
編集:   どの別名を編集しても正典が更新される（symlink を壊さない限り）
禁止:   symlink の実体化 / 別コピー作成 / 正典削除 / secretやaccount stateの共有
```
