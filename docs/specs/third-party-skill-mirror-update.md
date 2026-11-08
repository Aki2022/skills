---
id: SPEC-third-party-skill-mirror-update
status: draft # draft | active | implemented | superseded
created_at: 2026-11-08
updated_at: 2026-11-08
related_guides: []
affected_workstreams: [WS-20261108-third-party-skill-mirror-sync]
---

# 第三者 skill mirror の更新方式

## Purpose

正典に取り込んだ第三者 skill（mirror）を、上流の新しい版へどう更新するかの方針を決める。この spec は方針だけを扱い、`mirror_skill.py`・`mirrors.yaml`・SKILL.md は変更しない。

## Problem

- 2026-11-08 時点で、台帳の現役 mirror は 31 件ある。取得日から 90 日を超えるものが 13 件あり、`check_mirrors.py` は WARN を出す。ただし鮮度は取得日からの経過だけで、上流に新版があるかは見ていない（検査はオフライン）。
- 同値検査が実際にバイトを比べたのは 31 件中 1 件（eli5）だけで、残り 30 件は SKIP のまま RESULT は OK になる。
- 既存 mirror を上流の新版で置き換える経路が main に無い。`mirror_skill.py` は台帳に現役の名前があると拒否する。一方 `own-skill-commonize/SKILL.md` は「FAIL は上流から再取得して同期する」と書いており、文書と実行手段が合っていない。
- 更新の実装は wip 枝（`wip/stash-mirror-skill-20261001`）にだけある。その基点（2026-09-20）から main は 171 commit 進んでいて、枝の `mirrors.yaml` は main の cloudflare の固定版（2b1ffd2）を古い版（d8c1937）へ戻す。
- 台帳の版情報は揃っていない。現役 31 件のうち、commit・タグ・短縮 SHA の識別子があるのは 21 件。unknown が 4 件、plugin の版表記が 2 件、ローカル snapshot の未検証が 1 件、文章だけが 3 件。
- 公開 repo は第三者の中身を追跡しない（`.gitignore`）。そのため更新で旧版を上書きすると、台帳の再導入手段以外に戻す経路がない。

## Measurements (2026-11-08)

所有者の承認を得て、四件を空の一時ディレクトリへ degit で取得し、現行の mirror と読み取りだけで比べた。正典・台帳・wip 枝は変えていない。取得物のスクリプトは実行していない。

| Mirror | 台帳の記述 | degit の結果 | 現行 mirror との比較 |
| --- | --- | --- | --- |
| playwright-cli | 固定版 b85c7a7（完全 SHA）。reinstall は `#<sha>` | 成功（完全 SHA を受理） | 不一致。どちらも 11 ファイル。ローカルのみ 1、固定版のみ 1、共通 10 のうち 7 が内容で異なる（CRLF の差ではない） |
| cloudflare | 固定版 2b1ffd2（短縮 SHA）。reinstall は固定なし | 失敗：`could not find commit hash for 2b1ffd2` | 比較できず（固定版の中身を取れなかった） |
| cloudflare | 同上（HEAD を取得） | 成功（`#HEAD`。解決した SHA は出力に出ない） | 不一致。ローカル 320 ファイル、HEAD 130 ファイル。ローカルのみ 191、HEAD のみ 1、共通 129 はすべて内容が異なる |
| drawio | 版 unknown。取得から 102 日 | 成功（`#HEAD`） | 不一致。ローカル 1 ファイル、HEAD 2 ファイル。SKILL.md が異なる。HEAD のみ `scripts/find-shapes.js`（3,264 バイト） |

- degit は完全 SHA を受け付け、短縮 SHA は受け付けない。HEAD 取得の出力には解決された commit SHA が出ない。取得した版の SHA を記録するには、別の手段が要る（通信を伴う）。
- yomiyasu の注記（tip 以外は解決できない）は、この実測では肯定も否定もできない。b85c7a7 が tip かどうかを確認していないため。
- 台帳の固定版（playwright-cli）と現行の中身は一致しない。F1 の「上流とバイト同一」に反する状態。新しい版なのか、古い版や改変なのかは、この実測だけでは分からない。
- cloudflare の HEAD は現行と大きく異なる。wip 枝の `--update-existing` は ref の引数を持たず HEAD を取るので（コード読解）、cloudflare に使うと現行の 191 ファイルが消え、129 ファイルが変わる。wip は実行していない。

## Goals

- 既存 mirror を、承認を経た上流の版へ置き換える方式の方針を記録する。
- 更新で台帳の注記・local_copy・固定版を黙って失わないことを前提にする。
- 更新が途中で失敗しても、正典の中身と台帳が変わらないことを前提にする。
- 検査の緑が、比較していない項目を隠さないようにする。

## Non-Goals

- この spec で `mirror_skill.py`・`mirrors.yaml`・SKILL.md を変更すること。
- wip 枝をマージすること。
- 既存の mirror を今回更新すること。
- 第三者 skill のライセンス整理（記載なし・矛盾のある取り扱い）。別の判断として扱う。
- 配線（alias）の drift の解消。

## Domain Language

| Term | Meaning | Boundary or contrast |
| --- | --- | --- |
| 正典 | 全エージェントが読む skill 群のルート。 | 公開 git repo と同一ではない。公開 repo は own-* だけを追跡する。 |
| mirror（第三者 skill） | 上流の skill を、中身がバイト同一になるように正典へ複製したもの。 | own-*（自作）とは別。改変したい場合は fork として own-<対象>-<動作> を作る。 |
| 台帳 | 正典の `mirrors.yaml`。上流・取得パス・版・ライセンス・再導入手段を記録する。 | `mirrors:` は現役、`retired:` は棚から下ろしたもの。 |
| 固定版（pin） | 中身が一致すると主張する上流の識別子（commit SHA やタグ）。 | 文章による説明や unknown は固定版ではない。実測で一致しない場合がある。 |
| 取得（fetch） | 上流の内容を空の一時領域へ取り出すこと。まだ採用しない。 | 取得物は信頼しないデータとして扱う。 |
| 更新（update） | 取得した内容で、現役 mirror の中身と台帳の行を置き換えること。 | 新規追加や再導入（retired からの復帰）とは別の操作。 |
| 同値検査（parity） | 台帳の local_copy と mirror のバイト一致を見る検査。 | 上流は見ない。local_copy が無いものは比較しない。 |
| 鮮度（freshness） | fetched_at からの経過日数。 | 上流の新旧ではない。 |
| 承認 | 取得した上流の版を正典へ採用する前に、所有者が認めて記録すること。 | 誰が行うかは未決（Open Questions）。 |

## Actors and Scenarios

Actors:

- 所有者：上流の変更を採用するかを決める。
- 更新を実行する者（所有者、または作業エージェント）：取得と置き換えを行う。
- 機械検査（`check_mirrors.py` など）：同値と鮮度を判定する。上流の新旧は判定しない。
- 利用エージェント：skill の本文を指示として読む。上流の変更は、そのまま実行権限の範囲で働く。

Scenarios:

- S1：鮮度 WARN の mirror（取得から 99 日）に、上流の新版がある。所有者は差分を見て採否を決め、台帳の注記と固定版を残したまま更新したい。
- S2：更新の途中で失敗する（取得の失敗、frontmatter の name 不一致）。正典の中身と台帳は変わらないこと。
- S3：採用後に問題が見つかる。旧版へ戻せること。
- S4：固定版を持つ mirror（例：cloudflare の 2b1ffd2）は、所有者が上げると決めた時だけ動く。上げた後は、取得した内容の完全な識別子を記録する（unknown で書かない）。
- S5：検査が RESULT: OK を出しても、比較していない項目が残る（Problem を参照）。

## Requirements

### Functional

- F1：第三者 skill の中身は上流とバイト同一に保つ（既存規約。変更しない）。
- F2：cloudflare の固定版（2b1ffd2）と agents-sdk の注記（unverified-local-snapshot）は、この spec の判断では変えない（所有者の前提）。
- F3：wip 枝の内容を丸ごとマージしない（台帳が古い版へ戻るため）。

### Quality

- Q1：更新方式の採否は、一次証拠（実行結果・取得物の差分・台帳の実物）で判定する。台帳や文書の記述だけでは決めない。

## Design Direction

未決。方式が決まってから書く。現時点で設計を縛る制約だけを Constraints and Tradeoffs に記す。

## Constraints and Tradeoffs

- 公開 repo に第三者の中身を入れない（`.gitignore`）。そのため、正典内の旧版だけが巻き戻しの元になる。
- 取得物は信頼しないデータとして扱う。新しい空のディレクトリに取得し、その中のスクリプトは実行しない。
- ネットワークやパッケージ取得（`npx` や degit など）を伴う操作は、取得元・概算サイズ・コマンドを所有者に示し、承認を得てから行う。
- 自動更新は、上流の指示文を利用エージェントへ即座に届ける。承認の無い自動更新は採らない（供給網の危険）。
- 上流の新しいスクリプト（例：drawio HEAD の `scripts/find-shapes.js`）は、取り込み前に所有者が中身を確認する。利用エージェントが読む skill に実行可能なコードが入るため。
- 上流を見る検査はネットワークを使う。現在の検査（オフライン）は上流の新旧を判定できない。
- degit の ref は完全 SHA だけを受け付け（実測）、HEAD 取得は解決した SHA を返さない（実測）。固定版を上げるときは、完全 SHA の解決手段が要る。

## Deferred Decisions

- 鮮度の閾値（現在 90 日）の見直し。trigger：更新方式が決まった後。
- 上流の新旧を検査に入れるかどうか。trigger：完全 SHA の解決手段が決まった後。

## Open Questions

（draft 中、この節に面接の足場を次のタグで置く。【次の質問】は常に 1 件。【影響】は現行の文書・実装・作業との食い違い。【受入】は実装側の作業が示すべき証拠。承認時に【影響】と【受入】は作業単位へ移し、この節から消す。）

- 同期の時機：settled（検査が plugin の導入版と正典のずれを知らせ、同期は所有者の確認後に 1 コマンドで行う。自動同期はしない。所有者の回答、2026-11-08）。
- merge の方針（ws へ移す）：settled（ローカルのゲート `mise run verify` と docs 検証の通過で自動 merge。所有者の回答、2026-11-08）。
- 自動と確認の分け方（ws へ移す）：settled（所有者の回答、2026-11-08）。自動＝新しい枝・テスト先行の実装・SKILL.md・ローカル検査・ws 更新・作業枝の push と PR と、ゲート通過後の merge。確認＝ネットワーク取得、plugin の導入と更新、正典の実ファイルの置き換え（同期の実行）、wip 枝の削除。
- 【次の質問】更新前に必須にする検証は何か（plugin の版と正典の差分の提示、frontmatter の name 一致、ライセンス記載の確認、配線の監査のうちどれを必須にするか）。
  - 調査（2026-11-08、公式ドキュメント。ローカルの記録で一部を確認）：plugin の marketplace は git の ref と 40 桁の SHA で固定できる。更新は version の変化で検知され、自動更新はサードパーティ marketplace では既定で off。導入記録（installed_plugins.json）には version と gitCommitSha がある（ローカルで確認）。更新前の差分表示と承認、単体 skill の git からの取り込み、Codex・Gemini への配布は文書に無い。

Branches (status):

- 更新方式の選択：settled（(3) 実測してから決める。取得は承認を得て実施、結果は Measurements 節）。
- 位置付け：settled（所有者の判断、2026-11-08）。目的は、どの AI からも同じ第三者 skill を同じように使えること。上流からの取得と版の固定・更新は Claude Code の plugin（marketplace と SHA 固定）に任せる。plugin が更新されたら、その中身を正典へ同期（commonize）し、Codex・Gemini を含む全エージェントへ配る。自作するのは「plugin の版 → 正典」の同期だけで、上流から直接取る仕組みは作らない。
  - 既存の手段との関係：`mirror_skill.py --source-dir` は plugin の中身から正典へ取り込める。台帳の `local_copy` は plugin の置き場を指し、`check_mirrors.py` はそれと正典のバイト一致を見る。足りないのは、既に正典にある skill を新しい版で置き換えること（main は拒否する）と、plugin の置き場が版ごとに変わり古い版は 14 日で消えるため `local_copy` が切れること（実測で 4 件が SKIP）。
  - wip 枝の `--update-existing`（`--source-dir` で置き換える）は、この同期に使える部品になる。ただし注記・版を消す欠陥（コード読解）は直す必要がある。
- 正（固定版か中身か）：settled（固定版を正とする。中身は固定版の完全 SHA から取り直し、差分を所有者が確認してから置き換える。所有者の回答、2026-11-08）。
- 固定版の上げ方：open。完全 SHA での取得は実測済み。HEAD から SHA を得る手段と、短縮 SHA の解決手段は未決（通信を伴うので承認が要る）。
- 承認者と承認の記録：settled（同期の実行前に所有者が確認する。記録の置き場所は実装の issue で決める）。
- 旧版の保管：settled（ローカルには残さない。上流を SSOT とみなし、旧版は台帳に記録した完全 SHA から取り直す。所有者の判断、2026-11-08）。前提：台帳には完全 SHA を必ず記録する（短縮 SHA は degit で取り直せないと実測済み）。残るリスク：上流が非公開化・削除・履歴の書き換えをすると取り直せない。
- 必須検証：open。実測の 3 件すべてで、台帳の版と現行の中身が一致しなかった。取得後の比較と差分の提示は必須の候補。frontmatter の name 一致、ライセンス記載の確認、配線の監査も候補。
- 検査の緑の扱い：open。`check_mirrors.py` は 31 件中 30 件を SKIP しても RESULT: OK を出す。SKIP を結果に出す扱いをどうするか。
- wip 枝の処分：settled（`--update-existing` を同期の部品として、main から切った新しい枝へ選択的に移し、欠陥を直す。移植が main に入った後、所有者の確認を得て wip 枝を削除する）。
- ライセンス未解決の mirror の扱い：out-of-scope（別の判断。例：retired の pdf の取り扱い、eli5 の LICENSE 記載の矛盾）。
- 配線（alias）の drift：out-of-scope（別の作業単位）。

Interview findings (to reconcile):

- 【影響】`own-skill-commonize/SKILL.md` の「FAIL は上流から再取得して同期する」は、main の `mirror_skill.py` が有効な名前を拒否するため、実行できる手段がない。
- 【影響】wip 枝の `own-skill-commonize/SKILL.md` の変更は、枝全体を移すと main にあって wip に無い段落（棚の数え方、中身照合、CLI の扱い）を消す。移すなら段落単位にする。
- 【影響】wip 枝の `mirrors.yaml` は cloudflare の固定版を 2b1ffd2 から d8c1937 へ戻す（agents-sdk の注記も変わる）。
- 【影響】wip 枝の更新テストは `--source-dir`（ローカル取得）だけを通す。上流からの取得は、偽の npx を使う再導入テストでしか試していない。
- 【影響】wip 枝の更新は、指定されなかった注記・local_copy を台帳から消し、版を既定値 unknown に戻し、成功時は一時領域の旧版も消える。これはコード読解による結論で、実行による再現は未実施。
- 【影響】台帳の playwright-cli の固定版（b85c7a7）と現行の中身が一致しない（Measurements 節）。F1 に反する状態。
- 【影響】wip 枝の `--update-existing` は HEAD を取るので、固定版を持つ mirror に使うと固定版と無関係の内容へ置き換わる（Measurements 節の cloudflare HEAD）。

Acceptance proofs (to move to the work unit that does the work):

- 【受入】方式 (3) の場合：実測済み（Measurements 節）。残りは、HEAD から解決 SHA を得る手段の確認（承認が要る）。
- 【受入】方式 (2) の場合：wip 相当の更新で、台帳の注記・local_copy・固定版が指定なしでは消えないことを、実行で再現する（赤から緑）。
- 【受入】どの方式でも：更新の途中失敗（取得失敗、name 不一致）で、正典の中身と台帳が変わらないことを、失敗を注入して確かめる。
