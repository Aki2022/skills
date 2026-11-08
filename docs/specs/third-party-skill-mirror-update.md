---
id: SPEC-third-party-skill-mirror-update
status: active # draft | active | implemented | superseded
created_at: 2026-11-08
updated_at: 2026-11-08
related_guides: []
affected_workstreams: [WS-20261108-third-party-skill-mirror-sync]
---

# 第三者 skill mirror の更新方式

## Purpose

どの AI（Claude・Codex・Gemini）からも、同じ第三者 skill を同じ版で使えるようにする。上流からの取得と版の固定・更新は Claude Code の plugin に任せ、plugin の導入版を正典へ同期して全エージェントへ配る。

## Problem

- 既存 mirror を新しい版で置き換える手段が無かった。`mirror_skill.py` は現役の名前を拒否し、SKILL.md の「上流から再取得して同期する」は実行できなかった。
- 固定版と中身が一致しない mirror がある（2026-11-08 実測：playwright-cli）。短縮 SHA の固定版は degit で取り直せない（cloudflare の 2b1ffd2）。
- 検査は 31 件中 30 件の同値を比べずに RESULT: OK を出していた。plugin の置き場は版ごとに変わり古い版は 14 日で消えるため、`local_copy` が切れて比較から外れる。

## Measurements (2026-11-08)

所有者の承認を得て四件を空の一時ディレクトリへ degit で取得し、現行 mirror と読み取りだけで比べた。

| Mirror | 台帳の記述 | degit の結果 | 現行 mirror との比較 |
| --- | --- | --- | --- |
| playwright-cli | 固定版 b85c7a7（完全 SHA） | 成功 | 不一致。共通 10 のうち 7 が内容で異なり、片側だけのファイルが 1 ずつ |
| cloudflare | 固定版 2b1ffd2（短縮 SHA） | 失敗：`could not find commit hash for 2b1ffd2` | 比較できず |
| cloudflare | HEAD | 成功（解決した SHA は出力に出ない） | 不一致。ローカルのみ 191、共通 129 はすべて異なる |
| drawio | 版 unknown | 成功（HEAD） | 不一致。SKILL.md が異なり、HEAD のみ上流の find-shapes.js（スクリプト、正典には無い） |

Claude Code の plugin（公式ドキュメント、2026-11-08 調査。導入記録はローカルで確認）：marketplace は ref と 40 桁の SHA で固定でき、サードパーティの自動更新は既定で off、導入記録に `version` と `gitCommitSha` がある。更新前の差分表示と承認、Codex・Gemini への配布は無い。

## Goals

- 第三者 skill の版は、plugin が固定した完全 SHA（またはタグ）で決まる。
- plugin の導入版と正典のずれを検査が知らせる。
- 所有者が差分を確認してから、1 コマンドで正典へ同期できる。同期は Codex・Gemini を含む全エージェントに届く。
- 検査の緑が、比べなかった件数を隠さない。

## Non-Goals

- 上流から直接取得して更新する仕組みを自作すること（plugin に任せる）。
- 自動同期。
- 既存 mirror の取り直しの実行（同期の実行は所有者の確認ごとに行う）。
- 第三者 skill のライセンス整理、配線（alias）の drift の解消。

## Domain Language

| Term | Meaning | Boundary or contrast |
| --- | --- | --- |
| 正典 | 全エージェントが読む skill 群のルート。 | 公開 git repo は own-* だけを追跡する。 |
| mirror（第三者 skill） | 上流の skill を、中身がバイト同一になるように正典へ複製したもの。 | 改変したい場合は fork として own-* を作る。 |
| 台帳 | 正典の `mirrors.yaml`。上流・版・ライセンス・再導入手段・plugin との対応を記録する。 | `mirrors:` は現役、`retired:` は棚から下ろしたもの。 |
| 固定版（pin） | 中身が一致すべき上流の識別子。完全な 40 桁 SHA かタグ。 | 短縮 SHA・unknown・文章は固定版として扱わない。 |
| 導入版 | Claude Code の plugin として導入されている版（`installed_plugins.json` の記録）。 | 正典の中身とは別物で、ずれうる。 |
| 同期 | 導入版の中身で正典の mirror と台帳の行を置き換えること。 | 上流からの直接取得ではない。 |
| drift | 導入版と正典の中身が違う状態。 | 鮮度（取得日からの経過）とは別。 |

## Actors and Scenarios

- 所有者：plugin の更新を受け入れ、差分を見て同期を決める。
- 作業エージェント：検査を回し、所有者の確認後に同期コマンドを実行する。
- 利用エージェント（Claude・Codex・Gemini）：正典の skill を指示として読む。

Scenarios:

- S1：plugin が更新された。検査が `WARN drift` と導入版の SHA を出す。所有者が `--update-existing`（`--apply` 無し）の差分を確認し、`--apply` で同期する。全エージェントが新しい版を読む。
- S2：同期の途中で失敗する（name 不一致、書き込み失敗）。正典の中身と台帳は元のまま。
- S3：同期した版に問題がある。台帳の完全 SHA の前の版を、上流から取り直して戻す。

## Requirements

### Functional

- F1：第三者 skill の中身は上流とバイト同一に保つ。
- F2：台帳の固定版を正とする。固定版と中身が食い違えば、中身を固定版から取り直す（所有者の確認後）。
- F3：`mirror_skill.py --update-existing` は `--apply` が無い間は差分の一覧だけを出し、何も書かない。
- F4：同期の `--version` は完全な 40 桁 SHA かタグに限る。
- F5：同期は台帳の行をその場で書き換え、引数で渡さない欄（note・local_copy・plugin など）を残す。
- F6：同期の途中で失敗したら、正典の中身と台帳を元に戻す。
- F7：台帳の `plugin` と `plugin_path` がある entry について、`check_mirrors.py` は導入版と正典を比べ、違えば `WARN drift` を出す（exit 0）。
- F8：`check_mirrors.py` の RESULT 行に、同値を比べた件数・SKIP 件数・drift 件数を出す。
- F10：ずれの検査は、既存の起動時チェック `check_session_fast.sh` の中で `--only-drift` として回し、ずれがある時だけ 1 行出す。入口を増やさない（Claude の各席と Codex に登録済みの入口を使う）。
- F11：plugin の自動更新は、サードパーティ marketplace も含めて on にする（所有者の判断、2026-11-08）。
- F12：新しい版の受け口は、まず marketplace の写し（`known_marketplaces.json` の installLocation）とする。plugin は入れない（文脈に何も載らず、有効・無効や cache の管理が要らない）。写しに実体が無い plugin（source が別リポジトリ）だけ、plugin を入れてその置き場と比べる（所有者の判断、2026-11-08）。写しの版は `.git` から、無ければ `.gcs-sha`（公式の写し。文書化されていない印）から読む。
- F9：cloudflare の固定版（2b1ffd2）と agents-sdk の注記（unverified-local-snapshot）は、所有者が同期を決めるまで変えない。

### Quality

- Q1：取得物は信頼しないデータとして扱い、その中のスクリプトを実行しない。
- Q2：同期前に必須の検証は、差分の提示と frontmatter の name 一致。

## Design Direction

- 同期の部品は wip 枝（`wip/stash-mirror-skill-20261001`）の `--update-existing` を土台にし、main から切った枝で作り直した。wip 枝は移植が main に入った後、所有者の確認を得て削除する。
- 旧版はローカルに残さない。上流を SSOT とし、台帳の完全 SHA から取り直す。残るリスク：上流の非公開化・削除・履歴の書き換え。
- 自動同期はしない。上流の変更が確認なしに全エージェントへ届くのを避けるため。

## Constraints and Tradeoffs

- plugin は Claude Code 専用なので、Codex・Gemini への配布は正典への同期に頼る。
- plugin には更新前の差分表示が無いので、差分の確認は同期コマンドの dry-run で行う。
- 公開 repo に第三者の中身を入れない（`.gitignore`）。

## Deferred Decisions

- `local_copy` から出所が分かる 5 件は対応を書いた（eli5 は plugin、karpathy-guidelines・claude-automation-recommender・claude-md-improver・frontend-design は marketplace の写し）。それ以外の mirror へ書き足すこと。trigger：所有者が対象の plugin を導入したとき（書き足すまで drift 検査はその entry を見ない）。
- 自動更新で、Claude が読む plugin の中身は差分の確認なしに新しくなる。これを許すかの見直し。trigger：自動更新で問題が起きたとき。
- 鮮度の閾値（90 日）の見直し。trigger：drift 検査の運用が一巡した後。

## Open Questions

（なし）
