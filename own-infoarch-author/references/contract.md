# design/info_arch/ の契約

方針は、利用側の repo の spec `information-architecture-standardization`（v4）に従う。

## ディレクトリ

- `info_arch.yaml`：伝え方の SSOT（7 つの原子と id 台帳）。**承認の対象**。CODEOWNERS でオーナーに割り当てる。
- `channels/<media>.yaml`：成果物の置き場所と出し方の宣言だけ。中身は `media`（ファイル名と一致）・`outputs`（repo 相対の glob）・`ship`・`ship_command`（`ship: true` なら必須）・`viewports`。
- `outlines/<成果物 id>.yaml`：成果物ごとの構造化アウトライン。AI が書き、人は保守しない。`id` はファイル名と一致する。
- `INFO_ARCH.md`：散文。原子を id で参照するだけで、値を持たない。

それ以外のファイルやサブディレクトリが `channels/` `outlines/` にあると失格にする（黙って無視しない）。`.` で始まるファイル（`.DS_Store`・`.gitkeep`）だけは読まずに飛ばす。

## 原子（7 種）

id は種別の接頭辞＋連番。連番は 4 桁（0001〜9999）か、10000 以上なら先頭を 0 にしない（`msg-00001` のような同じ番号の別表記は失格）。`ledger.next` 未満でなければ「台帳が発行していない」。改名・削除は `ledger.tombstones`（旧 → 新、削除は null）。

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
- `pages[].sections[]`：role・message・problem・proofs・copy（文案）。固定値は copy の中で `{{fct-NNNN}}` として参照する。

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

## この版に無いもの（ISSUE-02 以降）

6 段の手順とノウハウ、描画物の DOM の確認、Jev の判定台帳（表記ゆれ・強調）、出荷ゲート、工場の接続。
