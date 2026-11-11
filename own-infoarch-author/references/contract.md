# design/info_arch/ の契約

方針は、利用側の repo の spec `information-architecture-standardization`（v4）に従う。

## ディレクトリ

- `info_arch.yaml`：伝え方の SSOT（7 つの原子と id 台帳）。承認の状態は持たない。変更は他の変更と同じく、検査が緑なら自律マージする（人のレビューは置かない）。記録は git の履歴が持つ。
- `channels/<media>.yaml`：成果物の置き場所と出し方の宣言だけ。中身は `media`（ファイル名と一致）・`outputs`（repo 相対の glob）・`ship`・`ship_command`（`ship: true` なら必須。出荷の正式の入口が実行する）・`merge_base`（pre-push の予測マージの相手。出荷チャネルで無いと doctor が失格）・`viewports`。
- `outlines/<成果物 id>.yaml`：成果物ごとの構造化アウトライン。AI が書き、人は保守しない。`id` はファイル名と一致する。
- `INFO_ARCH.md`：散文。原子を id で参照するだけで、値を持たない。

それ以外のファイルやサブディレクトリが `channels/` `outlines/` にあると失格にする（黙って無視しない）。`.` で始まるファイル（`.DS_Store`・`.gitkeep`）だけは読まずに飛ばす。`.` で始まるサブディレクトリは失格にする。

## 原子（7 種）

id は種別の接頭辞＋連番。連番は 4 桁（0001〜9999。0000 は不可）か、10000 以上なら先頭を 0 にしない。`msg-00001` や末尾に改行の付いたキーのような、同じ番号の別表記は失格（id は常に全体一致で判定する）。`ledger.next` 未満でなければ「台帳が発行していない」。改名・削除は `ledger.tombstones`（旧 → 新、削除は null）。

| 種別 | 接頭辞 | proposed で必須 | 参照 |
| --- | --- | --- | --- |
| 対象者 | `aud` | label・situation・desired_action・priorities（メッセージを優先順に。先頭が 1 位） | priorities → msg |
| 課題 | `prb` | text・audience | audience → aud |
| メッセージ | `msg` | text・role（benefit / proof / how / objection / cta / vision） | addresses → prb |
| 証拠 | `prf` | text・source・supports・as_of か expires | supports → msg |
| 固定値 | `fct` | label・value・as_of か expires か evergreen: true（evergreen と日付は併存させない） | ― |
| 正式名称 | `ent` | type・name | ― |
| 声 | `vce` | traits（3〜4 個。trait・is・is_not） | proposed はプロジェクトに 1 つ |

メッセージの優先順は、対象者ごとに `priorities` で持つ（同じメッセージでも、対象者によって 1 位が違ってよい）。

## 構造化アウトライン（版 1）

- `brief`：purpose（目的）・action（取ってほしい行動）・medium・constraints・cut（この成果物では言わないメッセージ）。cut も参照として検査し、墓標をたどって改名先と比べる。
- `audience`：対象者 id。
- `outputs`：成果物のファイル（チャネルの outputs の glob に一致すること）。
- `pages[].sections[]`：role・message・problem・proofs・copy（文案）。固定値は copy の中で `{{fct-NNNN}}` として参照する（括弧の中に空白や改行を入れない。入れたものは形の崩れた参照として失格）。

## 終了コード

0 合格（info を含む）/ 1 警告 / 2 失格。クラッシュは必ず 2（1 は警告専用）。JSON の報告に、評価日・環境・件数・対象数を載せる。

## 検査（決定論・トークン 0）

| code | 内容 |
| --- | --- |
| IA_MISSING / IA_YAML / IA_SCHEMA | ファイルが無い・YAML として読めない（重複キーを含む。日付は文字列として読む）・スキーマ違反（`state: approved` を含む） |
| IA_ID_NOT_ISSUED / IA_ID_TOMBSTONED / IA_TOMBSTONE_TARGET | 台帳と墓標 |
| IA_REF_BROKEN / IA_EMPTY_REFERENCED | SSOT 内の参照切れ・proposed の原子が空欄を参照（参照の表は `check.ssot_refs` の 1 か所） |
| IA_VOICE_MULTIPLE | proposed の声が 2 つ以上 |
| IA_NO_CHANNELS / IA_NO_OUTLINES / IA_CHANNEL_NAME / IA_CHANNEL_FILE / IA_OUTLINE_FILE / IA_OUTLINE_ID | 対象 0 件・名前の不一致・読まれないファイル |
| IA_OUTLINE_CHANNEL / IA_OUTLINE_OUTPUT_UNDECLARED | 宣言の無いチャネル・チャネルの outputs に一致しない成果物 |
| IA_OUTLINE_REF_BROKEN / IA_OUTLINE_REF_EMPTY | アウトラインの参照切れ（墓標なら先を案内）・空欄の原子を参照（参照の表は `check.outline_refs` の 1 か所） |
| IA_OUTLINE_BAD_PLACEHOLDER / IA_OUTLINE_FACT_LITERAL | `{{fct-NNNN}}` の形でない `{{…}}`（改行をまたぐものを含む）・固定値の値を参照せず直接書いた |
| IA_IMPACT_UNKNOWN_ID | `impact --id` に SSOT にも墓標にも無い id を渡した（空の結果を合格にしない） |
| IA_OUTLINE_CUT_USED | brief.cut で削ると決めたメッセージを使った |
| IA_OUTLINE_OFF_PRIORITY（warn） | 対象者の priorities に無いメッセージを使った |
| IA_EXPIRED_SHIPPED / IA_EXPIRED_UNSHIPPED（info） | 期限切れの証拠・固定値を、出荷チャネル（`ship: true`）のアウトラインが参照している／参照していない |

## 期限

`expires` があればその日、無ければ `as_of` に日数を足した日を期限とする。期限の日の当日までは有効で、翌日から期限切れ。日数の初期値は `references/expiry.json`、プロジェクトの
`info_arch.yaml` の `policy.expiry_days` で上書きできる。`evergreen: true` の固定値は期限を持たない。

## 工場への受け渡し（`skeleton`）

`infoarch-author skeleton --outline <id>` は、check が緑のアウトラインから骨組みの HTML を出す（失格していれば出さない）。
セクションごとに `data-ia-role`・`data-ia-msg`／`data-ia-problem` を付け、文案の `{{fct-NNNN}}` を SSOT の値の
`<span data-ia-fact>` に展開し、文案は HTML エスケープする。pages と outputs の数が同じなら 1 ページずつ対応するファイルに、
違えば全ページを各ファイルに入れる（`dom` も同じ対応で比べる）。`--out` を付けるとファイルに書き、既にあるファイルは上書きしない。
工場（own-web-design）は、文言・順序・`data-ia-*` を変えずにスタイリングし、描画後に `dom` を走らせる。

## 描画物の確認（`dom`）

`infoarch-author dom --root . --outline <id>` は、アウトラインの outputs のうち HTML を、チャネルの viewports（無ければ 375×812 と 1280×800）でブラウザに描画し、DOM を確かめる。

- 解析と描画はブラウザ（Playwright の Chromium）が行う。成果物のソースを自前のパーサで読まない。
- 工場はアンカーを埋める：メッセージは `data-ia-msg`、課題は `data-ia-problem`、固定値の表示は `data-ia-fact`。
- Playwright は `INFOARCH_PLAYWRIGHT_DIR`、無ければ対象 repo（実行ディレクトリ）から解決する。自動では入れない。解決できなければ IA_DOM_PLAYWRIGHT で失格。
- 読み取る前に、ページが落ち着くのを待つ：network idle の後に 1 秒（表示を始めるタイマーのため）、さらに実行中のアニメーションとトランジションが終わるまで最大 5 秒。終わらないアニメーションがあっても読み取り、2 回の描画の比較で揺れを報告する。
- 同じ入力を 2 回描画し、結果が違えば IA_DOM_UNSTABLE（判定に使わない）。
- 最初の画面の文言（各画面幅）を報告の `first_screens` に出す。強調の判定（Jev の判定台帳）の入力になる。

| code | 内容 |
| --- | --- |
| IA_DOM_NO_TARGETS | 描画されたテキストが 0 件 |
| IA_DOM_ANCHOR_MISSING / IA_DOM_ANCHOR_HIDDEN | アウトラインのメッセージ・課題のアンカーが無い・見えていない（display / visibility / opacity / 大きさ 0） |
| IA_DOM_ANCHOR_UNKNOWN | アウトラインに無いメッセージ・課題のアンカー |
| IA_DOM_ORDER | 見えているアンカーの描画順（上から）に、アウトラインの順が部分列として現れない（行動喚起の繰り返しは可） |
| IA_DOM_FACT_MISSING / IA_DOM_FACT_VALUE | 文案が参照する固定値の表示が無い・SSOT の値と違う（空白は正規化して比べる） |
| IA_DOM_UNSTABLE | 2 回の描画で結果が違う |
| IA_DOM_OUTPUT_MISSING / IA_DOM_UNCOVERED（warn） | 成果物のファイルが無い・HTML 以外（未カバー） |
| IA_DOM_PLAYWRIGHT | Playwright を解決できない・描画に失敗した |

見えるかどうかは、要素とその祖先の display / visibility / opacity と、大きさで判定する。コントラストや画面外への配置（clip-path・transform）は、デザイン側の L2 が持つ。

## 意味の判定（`judge`）

意味の判定は Jev が行い、その結果を判定台帳に固定する。検査は台帳だけを読み、その場で Jev を呼ばない。
方針ファイル（モデル・θ・プロンプト）・台帳（`design/judgments.json`）・使用量の台帳（累計 2,000 判定）は、own-design-author と共有する
（own-design-author の `references/judgments-and-modules.md`）。

| 種類 | 単位 | 入力の本文 | 文脈 |
| --- | --- | --- | --- |
| `ia-name-variant` | セクション（正式名称が 1 つ以上ある時） | 固定値を展開した文案 | 正式名称の一覧 |
| `ia-unregistered-claim` | セクション | 固定値を展開した文案 | メッセージ・課題・証拠・使った固定値 |
| `ia-audience-fit` | アウトライン | 全セクションの役割と文案（順に） | 対象者（label・situation・desired_action・優先順位のメッセージを順に）と brief。優先順位を入れ替えると判定のキーが変わる |
| `ia-emphasis` | 成果物 × 画面幅 | `dom` が切り出した最初の画面の文言 | 対象者と、その 1 位のメッセージ |

score は「問題がある確率」。θ 以上は IA_JUDGE_REVIEW（要確認）として人に回し、score だけで失格にしない。
強調の判定でゲートを止めるかは、最初の実回答で精度を測ってから決める（spec）。

```sh
infoarch-author dom --root . --outline lp-xxx > dom.json
infoarch-author judge --root . --dom-report dom.json                       # 台帳の読み取り（本文は出さない）
infoarch-author judge --root . --dom-report dom.json --include-candidates > candidates.json   # Jev に渡す候補（文案を含む）
python3 <own-design-author>/scripts/judgment_run.py --candidates candidates.json --answers <dir>/judgment-answers.json --limit 50 [--execute]
infoarch-author judge --root . --dom-report dom.json --answers <dir>/judgment-answers.json --replacement-out replacement.json
python3 <own-design-author>/scripts/judgment_ledger.py record|refresh --repo . --replacement replacement.json [--write]
```

`--replacement-out` は共有の台帳の置き換えを作る。design の行は全部残し、IA は現在の候補に対する実回答だけを入れ、古い IA の行は落とす。回答ファイルは `judgment_run.py` と同じ読み手で検査する（形・policy・候補のハッシュ・score と raw_answer の一致）。出力先は新規作成だけで、既存のファイル・symlink（親ディレクトリを含む）には書かない。
holdout はそのまま引き継ぐ。回答の policy が今の共有 policy と違えば止める（migrate が要る）。既にあるファイルは上書きしない。

| code | 内容 |
| --- | --- |
| IA_JUDGE_REFUSED | check が失格している（先に直す） |
| IA_JUDGE_NO_TARGETS | 判定の候補が 0 件 |
| IA_JUDGE_LEDGER | 台帳が読めない・policy と合わない・symlink |
| IA_JUDGE_DOM_FAILED / IA_JUDGE_DOM_REPORT | 合格していない描画を判定に使おうとした・最初の画面の行が無いか読めない・アウトラインの成果物と画面幅に無い行か重複（行の本文は報告に出さない） |
| IA_JUDGE_UNCOVERED（warn） | 台帳に無い判定がある（未カバー） |
| IA_JUDGE_REVIEW（warn） | score が θ 以上（要確認） |
| IA_JUDGE_EMPHASIS_MISSING（出荷される HTML なら warn、他は info） | 最初の画面の入力が無く、強調を判定していない。dom の報告が一部の成果物・画面幅だけなら、足りない組ごとに出る |

報告は `uncovered`（未カバー）・`screened`（台帳にあり。「問題なし」の保証ではない）・`missing`（入力が無く判定していない）を分けて出す。

## 出荷ゲートと維持（`ship`・`maintain`）

出荷ゲートは、作業ツリーではなく、出荷されるコミット（clean な HEAD、または予測マージ木）に対して走る。順番は次のとおりで、失格の段で止まる。
1. `check`（決定論）。出荷チャネル（`ship: true`）が 0 件、または出荷されるアウトラインが 0 件なら失格（空のゲートを合格にしない）。
2. `dom`：出荷されるアウトラインのうち HTML の成果物を持つものを、実ブラウザで描画して確かめる。HTML が無いアウトラインは `skipped` に記録する。
3. `judge`：`dom` の最初の画面を渡して、共有の判定台帳を読む。出荷されるアウトラインの判定が台帳に無ければ `IA_SHIP_UNCOVERED`（失格）。出荷しないアウトラインの未カバーは `IA_JUDGE_UNCOVERED`（警告）のまま。θ 以上は `IA_JUDGE_REVIEW`（警告。人に回す。強調でゲートを止めるかは未決）。

| 入口 | 何をするか |
| --- | --- |
| `infoarch-author ship --root .` | HEAD を検査するだけ。出荷しない |
| `infoarch-author ship --root . --channel <media>` | 正式の入口。緑なら、同じコミットのまま `ship_command` を実行する（shlex で argv に分け、シェルは通さない。`INFOARCH_SHIP_COMMIT`・`INFOARCH_SHIP_TREE`・`INFOARCH_SHIP_CHANNEL` を渡す）。実行の後に作業ツリーが変わっていたら失格 |
| `infoarch-author ship --root . --pre-push --refs-file F` | push される各コミットを、チャネルの `merge_base` と予測マージした木（一時 worktree）で検査する。`merge_base` は push されるコミット自身のチャネル定義から読む。無ければ止まる |
| `infoarch-author ship --root . --doctor` | 設定の診断。`ship_command` か `merge_base` の無い出荷チャネル、マーカー（`.vibe-guard/infoarch-ship.enabled`）の欠落、スキャナを保つ pre-push の dispatcher と拡張（`hooks/pre-push.d/61-infoarch-ship` と同じ中身）の欠落を失格にする |

`--pre-push` の終了コードは 0 か 2 だけ（dispatcher は 0 以外で push を止めるので、警告は `warnings` に載せて 0 で返す）。他の入口は 0 / 1（警告）/ 2。
pre-push の拡張は、マーカーのある repo だけで動く（opt-in）。マーカーがあれば、壊れたマーカー・refs の欠落・検査の失格はすべて push を止める。
拡張は `INFOARCH_SHIP_COMMAND`（既定は PATH の `infoarch-author`）を呼ぶ。予測マージ木には node_modules が無いので、`INFOARCH_PLAYWRIGHT_DIR` が無ければ push している checkout を使う。
出力は 1 行の JSON。CI での出荷ゲートは持たない（spec の Deferred）。

`infoarch-author maintain --root .` は own-session-close の Step 2.5 から呼ぶ。`design/info_arch/` が無ければ 1 行で skip する。
あれば `check` と `judge`（台帳を読むだけ。ブラウザは使わない）を走らせ、次を 1 行の JSON で返す。
- `due`：期限が `references/expiry.json` の `due_within_days` 日以内に来る証拠・固定値と、それを参照するアウトライン（出荷されるか）。
- `changed_paths`：出荷チャネルの `merge_base` からのコミット済みの差分と、未コミットの差分（`design/info_arch/` の中だけ）。
- `impacted_outlines`：変わった原子を参照するアウトライン（正式名称と声が変わったら全部）。
- `uncovered_judgments`：台帳に無い判定の数（Jev で取得する候補）。

終了コードは 0（整合）・1（期限が近い・未カバー・警告）・2（失格）。台帳もアウトラインも書き換えない。
