# design/info_arch/ の契約

方針は、利用側の repo の spec `information-architecture-standardization`（v4）に従う。

## ディレクトリ

- `info_arch.yaml`：伝え方の SSOT（7 つの原子と id 台帳）。承認の状態は持たない。変更は他の変更と同じく、検査が緑なら自律マージする（人のレビューは置かない）。記録は git の履歴が持つ。
- `channels/<media>.yaml`：成果物の置き場所と出し方の宣言だけ。中身は `media`（ファイル名と一致）・`outputs`（repo 相対の glob）・`ship`・`ship_command`（`ship: true` なら必須。出荷の正式の入口が実行する）・`merge_base`（pre-push の予測マージの相手。出荷チャネルで無いと doctor が失格）・`viewports`。
- `outlines/<成果物 id>.yaml`：成果物ごとの構造化アウトライン。AI が書き、人は保守しない。`id` はファイル名と一致する。
- `INFO_ARCH.md`：散文。原子を id で参照するだけで、値を持たない。

それ以外のファイルやサブディレクトリが `channels/` `outlines/` にあると失格にする（黙って無視しない）。`.` で始まるファイル（`.DS_Store`・`.gitkeep`）だけは読まずに飛ばす。`.` で始まるサブディレクトリは失格にする。

## 原子（スキーマ版 4）

id は種別の接頭辞＋連番。連番は 4 桁（0001〜9999。0000 は不可）か、10000 以上なら先頭を 0 にしない。`msg-00001` や末尾に改行の付いたキーのような、同じ番号の別表記は失格（id は常に全体一致で判定する）。`ledger.next` 未満でなければ「台帳が発行していない」。改名・削除は `ledger.tombstones`（旧 → 新、削除は null）。

| 種別 | 接頭辞 | proposed で必須 | 参照 |
| --- | --- | --- | --- |
| 対象者 | `aud` | label・situation・desired_action・priorities（メッセージを優先順に。先頭が 1 位）・segment | priorities → msg、segment → seg |
| メッセージ | `msg` | text・role（benefit / proof / how / objection / cta / vision） | addresses → pan、serves → val・job・pan・gan（立つ positioning の原子） |
| 証拠 | `prf` | text・source・as_of か expires・supports か segment のどちらか一方 | supports → msg（outline 側）。segment があれば positioning の証拠で、支える先は矢印 `supports` |
| 固定値 | `fct` | label・value・as_of か expires か evergreen: true（evergreen と日付は併存させない） | ― |
| 正式名称 | `ent` | type・name | ― |
| 声 | `vce` | traits（3〜4 個。trait・is・is_not） | proposed はプロジェクトに 1 つ |

対象者とメッセージは、ISSUE-14 で outline の定数（target）と outline の原子へ移る。それまではここに残る。

### positioning（SPEC-information-architecture-standardization v6 要件 2・4・5・8）

原子は 1 つのセグメントにだけ属する（`segment`）。写しの出どころ（`copied_from`）と書いた人・材料（`provenance`）は中身ではないので、どの id と承認のハッシュにも入らない。別のセグメントで同じ代替などを使うときは写しを作り、出どころを `copied_from` に書く（同じ種類・別のセグメントの原子）。v5 の約束（`prm`）とカテゴリーとコンセプト（`cpt`）は持たない（スキーマ違反）。重要度は持たない。

| 種別 | 接頭辞 | proposed で必須 |
| --- | --- | --- |
| セグメント | `seg` | qualifier・situation |
| ジョブ | `job` | text・segment |
| 痛み | `pan` | text・segment |
| 得たいもの | `gan` | text・segment |
| 代替 | `alt` | text・segment |
| カテゴリー | `cat` | text・segment |
| 強み | `str` | text・segment |
| 提供価値 | `val` | text・segment・kind（difference：隙間を満たす／parity：同等であるべき点を満たす） |

矢印は `edges` に 1 本ずつ書く（`kind`・`from`・`to`・`state`・`provenance`）。キーは `<kind>:<from>:<to>` で、同じキーは 1 本だけ。矢印のセグメントは始点の原子のもの。

| kind | 始点 → 終点 | 意味 |
| --- | --- | --- |
| `has_job` | seg → job | セグメントのジョブ |
| `need` | job → pan・gan | ジョブの周りの痛み・得たいもの |
| `fulfills` | alt → pan・gan | 満たす度合い（`degree`：full／partial／none。この種類にだけ必ず書く） |
| `recalls` | job → cat | この場面で思い出すカテゴリー |
| `includes` | cat → alt | カテゴリーに入れている代替 |
| `member_of` | seg → cat | 所属カテゴリー（ちょうど 1 本） |
| `lacks` | str → alt | 代替に無い強み |
| `produces` | str → val | 強みが生む提供価値 |
| `satisfies` | val → pan・gan | 提供価値が満たす隙間か同等であるべき点 |
| `supports` | prf → 原子・矢印のキー | 証拠が支える |

**承認**：オーナーが原子 1 つ・矢印 1 本ずつ、`design/approvals.yaml` に `ia:positioning:<原子 id>`・`ia:positioning:<kind>:<from>:<to>` で記録する（AI は `design-author approve --prepare` まで）。ハッシュ（`positioning.element_hash`）はその要素自身の中身だけ（`copied_from`・`provenance` を除く）を束縛するので、1 つ直すとその要素の承認だけが外れる。承認かどうかは台帳のハッシュの一致だけで決まり、`state` には持たない。承認は採用の判断だけを表し、顧客での検証を意味しない（ISSUE-13 で、要素ごとの --prepare・契約の版・出荷の報告を足す）。

**導出**（`positioning.derive`、記録しない）：positioning ＝ そのセグメントの承認済みの原子と、承認済みの原子どうしを結ぶ承認済みの矢印すべて（下書きは入らない）。隙間 ＝ 痛み・得たいもののうち、そのセグメントの承認済みの代替すべてのどれも `full` でないもの。同等であるべき点 ＝ 所属カテゴリーが `includes` で入れている代替のすべてが `full` のもの（カテゴリーの外の代替は見ない。ADR の 5）。比べる代替の一部だけが満たすものは、どちらにもならない。満たす度合いの矢印が 1 本でも欠ければ導かない（未調査を `none` とみなさない）。

**id**：positioning の id は `pos-` ＋ そのセグメントの承認済みの原子と矢印の中身（`copied_from`・`provenance` を除く）を正規化した JSON の SHA-256。ファイルの中で書いた順番・写しの出どころ・provenance は入らない。証拠（`supports` の矢印）は、支える先が positioning に入っているときだけ数える。`infoarch-author positioning --root .` が、セグメントごとの id・隙間・同等であるべき点・所属カテゴリー・承認待ちの要素を出す。

**失格**（承認済みの positioning だけを見る。下書きは対象外）と、形の失格（proposed の全要素）。条件と赤のフィクスチャの対応表は `fixtures/positioning-red-cases.json`（対応の無い条件 0 件をテストが確かめる）。

| code | 内容 |
| --- | --- |
| IA_POSITIONING_ISOLATED_ATOM | 承認済みの矢印にどれも繋がらない承認済みの原子 |
| IA_POSITIONING_JOB_WITHOUT_NEED | 痛みも得たいものも持たないジョブ |
| IA_POSITIONING_INCOMPLETE | 承認済みのジョブ・痛みか得たいもの・提供価値のどれかが 0 件（Quality 2：中身の無い positioning を合格にしない） |
| IA_POSITIONING_CATEGORY_COUNT | 所属カテゴリーが 0 本か 2 本以上 |
| IA_POSITIONING_NO_ALTERNATIVES | 隙間の比較に使うセグメントの代替が 0 件 |
| IA_POSITIONING_NO_CATEGORY_ALTERNATIVES | 同等であるべき点の比較に使う所属カテゴリーの代替が 0 件 |
| IA_POSITIONING_FULFILLMENT_MISSING | 代替と痛み・得たいものの組で、満たす度合いの承認済みの矢印が欠ける（`detail.pairs`） |
| IA_POSITIONING_VALUE_UNANCHORED | difference の提供価値が隙間を、parity の提供価値が同等であるべき点を 1 つも満たさない |
| IA_POSITIONING_STRENGTH_UNCOMPARED | 代替と比べない（`lacks` の無い）強み |
| IA_POSITIONING_EDGE_SHAPE | 矢印の始点・終点の種類が表と違う、`degree` が fulfills に無い・ほかの種類にある |
| IA_POSITIONING_EDGE_DUPLICATE | 同じキーの矢印が 2 本 |
| IA_POSITIONING_CROSS_SEGMENT | 矢印が別のセグメントの原子を結ぶ（共有しない。写しを作る） |
| IA_POSITIONING_COPY_SOURCE | 写しの出どころが無い・別の種類・同じセグメント |

### 立っている positioning の経路（`check.positioning_path`）

アウトラインのメッセージ（と痛みのセクション）ごとに、対象者のセグメントの中だけで矢印をたどる。提供価値からは、満たす痛み・得たいもの（`satisfies`）と、それを生む強み（`produces`）とその強みが比べる代替（`lacks`）。痛み・得たいものからはジョブ（`need`）、ジョブからはセグメント（`has_job`）。経路は原子と矢印のキーで、中身を正規化した JSON の SHA-256 を経路のハッシュとし、意味の判定の入力に使う。経路の要素に承認の無いものがあれば IA_POSITIONING_UNAPPROVED（ISSUE-14 で、メッセージが提供価値ちょうど 1 つを指す形に替わる）。

### v2 からの移行（`infoarch-author migrate --root . [--write]`）

猶予なしの一括移行。機械で決まることだけを移す：課題（prb）を同じ番号の痛み（pan。対象者のセグメントに属する）へ、対象者ごとに同じ番号のセグメント（seg）と空のジョブ（job）を作る、メッセージの addresses を追う、アウトラインを版 2 へ（problem → pain）、出荷チャネルの HTML の `data-ia-problem`・`data-ia-role="problem"` を改名、`design/design.yaml` の `modules.ia` を 0.3.0 へ。判断が要るもの（ジョブ・代替・カテゴリー・強み・提供価値・矢印・メッセージの serves・アウトラインの strategy）は `drafting_required` に並べ、作らない。報告は、旧→新の id の対応（`id_map`）・移しそこね（`missed`。0 件でなければ exit 2）・HTML のコメントに残る旧 id（`comment_mentions`。属性以外は書き換えない）・外れる IA の判定の数（`invalidated_judgments`）・戻し方（`recovery`）。`--write` が無ければ何も書かない。コミットはしない。版 4 の木では何もしない（再実行できる）。版 3（v5 の形）は断る（v5 から v6 への移行は ISSUE-17 の道具）。

### 対応表（`infoarch-author mapping --outline <id>`）

外部の実験の道具が「どの版の、どこを、何に変えたか」を指すための表。成果物 id・版のハッシュ（アウトライン・使う原子の中身・成果物が立つ positioning の要素の中身・成果物のファイルから作る）・成果物ごとのファイルのハッシュ・枠の実体ごとの種類と、使う原子の id と中身のハッシュ（並びの枠は原子の並び）を出す。比べる案の管理・割り付け・実験の記録・配信・計測は持たない（ISSUE-14 で set の id に置き換える）。

### 承認の対象（`ia:positioning:<key>`）

- `infoarch-author review --root . --request <束>`：束が束縛する原子（種類ごと）と矢印の中身を Markdown で出す（オーナーが読む）。今のハッシュと合わない行があれば `stale` に並べて失格（作り直す）。
- `infoarch-author targets --root .`：proposed の全部の原子と矢印の承認のハッシュ（`approval_hashes`）と、出荷チャネルのアウトラインの経路に乗る要素のキー（`must_approval_keys`）を出す。`ia_module.py`（契約 `ia-git-history-v3`、`modules.ia: 0.3.0`）がこれを own-design-author に返す。
- 出荷の報告（`ship`）は `positioning` に、`owner_adopted`（approved / unapproved）・`semantic_fit`（意味の判定の状態。今は not-judged）・`customer_validation`（IA は判定しない：not-judged-by-ia）を分けて出す。

## 構造化アウトライン（版 2）

- `brief`：purpose（目的）・action（取ってほしい行動）・medium・constraints・cut（この成果物では言わないメッセージ）。cut も参照として検査し、墓標をたどって改名先と比べる。
- `audience`：対象者 id。
- `outputs`：成果物のファイル（チャネルの outputs の glob に一致すること）。
- `strategy`（必須。ISSUE-14 の outline まで、この名前のまま）：段 2〜4 で選んだ positioning の原子（`choose`：job・pan・gan・val）と構成の理由（`rationale`）。既定の優先順から外れる理由は `deviation`。`infoarch-author compose --audience <aud> --positioning <ids>` が、選んだ原子に立つメッセージを対象者の優先順に並べて出す。
- `frames`（任意）：差し替えできる枠の実体。`media`（例 lp）と `instances`（`id`・`type`）。種類と多重度は `references/media/frames.json`。実体 id は成果物の中で一意。要素の枠はセクションの `frame` で 1 つのセクションに付け、並びの枠（節の順序）はアウトラインの上だけで持つ。
- `pages[].sections[]`：role（v1 の problem は pain）・message・pain・proofs・frame・copy（文案）。固定値は copy の中で `{{fct-NNNN}}` として参照する（括弧の中に空白や改行を入れない。入れたものは形の崩れた参照として失格）。

## 終了コード

0 合格（info を含む）/ 1 警告 / 2 失格。クラッシュは必ず 2（1 は警告専用）。JSON の報告に、評価日・環境・件数・対象数を載せる。

## 検査（決定論・トークン 0）

| code | 内容 |
| --- | --- |
| IA_MISSING / IA_YAML / IA_SCHEMA | ファイルが無い・YAML として読めない（重複キーを含む。日付は文字列として読む）・スキーマ違反（`state: approved` を含む） |
| IA_ID_NOT_ISSUED / IA_ID_TOMBSTONED / IA_TOMBSTONE_TARGET | 台帳と墓標 |
| IA_REF_BROKEN / IA_EMPTY_REFERENCED | SSOT 内の参照切れ・proposed の原子が空欄を参照（参照の表は `check.ssot_refs` の 1 か所） |
| IA_VOICE_MULTIPLE | proposed の声が 2 つ以上 |
| IA_POSITIONING_*（要件 5・要件 2） | 上の positioning の表 |
| IA_POSITIONING_UNTRACED | アウトラインのメッセージが、どの positioning の原子にも立っていない（出荷チャネルは失格、ほかは警告） |
| IA_POSITIONING_SEGMENT_MISMATCH | メッセージ・痛みの positioning が、対象者のセグメントにたどり着かない（出荷チャネルは失格、ほかは警告） |
| IA_POSITIONING_UNAPPROVED | 経路の原子・矢印に、オーナーの承認が無いか外れたものがある（`detail.atoms`。出荷チャネルは失格、ほかは警告） |
| IA_APPROVALS_INVALID | `design/approvals.yaml` が読めない・形が違う |
| IA_COMPOSITION_POSITIONING_INVALID | `strategy.choose` の原子が無いか、対象者のセグメントにたどり着かない |
| IA_COMPOSITION_OFF_POSITIONING（警告） | セクションのメッセージが、選んだ positioning の原子に立っていない |
| IA_COMPOSITION_DEVIATION_UNEXPLAINED（警告） | 対象者の 1 位のメッセージを使わず、`brief.cut` にも `strategy.deviation` にも理由が無い |
| IA_FRAME_UNKNOWN_MEDIA / IA_FRAME_UNKNOWN_TYPE | 枠の媒体・種類がノウハウに無い |
| IA_FRAME_DUPLICATE_ID / IA_FRAME_MULTIPLICITY | 枠の実体 id の重複・種類ごとの実体の数が多重度に反する |
| IA_FRAME_UNDECLARED / IA_FRAME_SEQUENCE_ANCHORED / IA_FRAME_DUPLICATE_USE / IA_FRAME_UNUSED | 宣言していない枠を使う・並びの枠をセクションに付ける・要素の枠を 2 つ以上のセクションに付ける・宣言した要素の枠をどこにも付けない |
| IA_PAGES_OUTPUT_NOT_INVENTORIED（警告） | 出荷チャネルの HTML の成果物が、design/validation.json の pages（covered・excluded）に無い（design の検査が見ていない） |
| IA_PAGES_OUTPUT_EXCLUDED_BY_DESIGN（info） | design が対象外にした成果物。IA の検査は続ける（ページの宣言は一本化しない。ADR-20261115-design-ia-two-skills-shared-base） |
| IA_NO_CHANNELS / IA_NO_OUTLINES / IA_CHANNEL_NAME / IA_CHANNEL_FILE / IA_OUTLINE_FILE / IA_OUTLINE_ID | 対象 0 件・名前の不一致・読まれないファイル |
| IA_OUTLINE_CHANNEL / IA_OUTLINE_OUTPUT_UNDECLARED | 宣言の無いチャネル・チャネルの outputs に一致しない成果物 |
| IA_OUTLINE_REF_BROKEN / IA_OUTLINE_REF_EMPTY | アウトラインの参照切れ（墓標なら先を案内）・空欄の原子を参照（参照の表は `check.outline_refs` の 1 か所） |
| IA_OUTLINE_BAD_PLACEHOLDER / IA_OUTLINE_FACT_LITERAL | `{{fct-NNNN}}` の形でない `{{…}}`（改行をまたぐものを含む）・固定値の値を参照せず直接書いた |
| IA_IMPACT_UNKNOWN_ID | `impact --id` に SSOT にも墓標にも無い id を渡した（空の結果を合格にしない） |
| IA_OUTLINE_TEXT_KEY_MISSING / IA_OUTLINE_TEXT_KEY_DUPLICATE / IA_OUTLINE_TEXT_KEY_MISPLACED | message も pain も無い navigation・other のセクションに `key` が無い・`key` の重複・message か pain のあるセクションか、navigation・other 以外の role に `key` を書いた |
| IA_OUTLINE_CUT_USED | brief.cut で削ると決めたメッセージを使った |
| IA_OUTLINE_OFF_PRIORITY（warn） | 対象者の priorities に無いメッセージを使った |
| IA_PRIORITY_UNSUPPORTED（warn） | proposed の対象者の priorities の先頭のメッセージを、proposed の証拠がどれも `supports` していない（`role` が cta・vision のものは主張ではないので対象外。2 番目以降は見ない） |
| IA_EXPIRED_SHIPPED / IA_EXPIRED_UNSHIPPED（info） | 期限切れの証拠・固定値を、出荷チャネル（`ship: true`）のアウトラインが参照している／参照していない |

## 期限

`expires` があればその日、無ければ `as_of` に日数を足した日を期限とする。期限の日の当日までは有効で、翌日から期限切れ。日数の初期値は `references/expiry.json`、プロジェクトの
`info_arch.yaml` の `policy.expiry_days` で上書きできる。`evergreen: true` の固定値は期限を持たない。

## 工場への受け渡し（`skeleton`）

`infoarch-author skeleton --outline <id>` は、check が緑のアウトラインから骨組みの HTML を出す（失格していれば出さない）。
セクションごとに `data-ia-role`・`data-ia-msg`／`data-ia-pain`（宣言した文言なら `data-ia-text`、要素の枠があれば `data-ia-frame`）を付け、文案の `{{fct-NNNN}}` を SSOT の値の
`<span data-ia-fact>` に展開し、文案は HTML エスケープする。pages と outputs の数が同じなら 1 ページずつ対応するファイルに、
違えば全ページを各ファイルに入れる（`dom` も同じ対応で比べる）。`--out` を付けるとファイルに書き、既にあるファイルは上書きしない。
工場（own-web-design）は、文言・順序・`data-ia-*` を変えずにスタイリングし、描画後に `dom` を走らせる。

## 描画物の確認（`dom`）

`infoarch-author dom --root . --outline <id>` は、アウトラインの outputs のうち HTML を、チャネルの viewports（無ければ 375×812 と 1280×800）でブラウザに描画し、DOM を確かめる。

- 解析と描画はブラウザ（Playwright の Chromium）が行う。成果物のソースを自前のパーサで読まない。
- 工場はアンカーを埋める：メッセージは `data-ia-msg`、痛みは `data-ia-pain`、固定値の表示は `data-ia-fact`、正式名称（SSOT の `ent-`）の表示は `data-ia-entity`、宣言した文言は `data-ia-text`。v4 の `data-ia-problem` が残っていれば IA_DOM_LEGACY_ATTRIBUTE で失格（互換は持たない）。要素の枠は `data-ia-frame`（既存の `data-ia-msg` 等には相乗りしない）。宣言した実体の印が無い・2 つ以上ある・宣言していない印がある、は IA_DOM_FRAME_MISSING / IA_DOM_FRAME_DUPLICATE / IA_DOM_FRAME_UNDECLARED で失格。
- Playwright は `INFOARCH_PLAYWRIGHT_DIR`、無ければ対象 repo（実行ディレクトリ）から解決する。自動では入れない。解決できなければ IA_DOM_PLAYWRIGHT で失格。
- 読み取る前に、ページが落ち着くのを待つ：network idle の後に 1 秒（表示を始めるタイマーのため）、さらに実行中のアニメーションとトランジションが終わるまで最大 5 秒。終わらないアニメーションがあっても読み取り、2 回の描画の比較で揺れを報告する。
- 同じ入力を 2 回描画し、結果が違えば IA_DOM_UNSTABLE（判定に使わない）。
- 最初の画面の文言（各画面幅）を報告の `first_screens` に出す。強調の判定（Jev の判定台帳）の入力になる。
- 各アンカーの表示の文言（描かれている文字だけ。`unanchored` と同じ読み方で、opacity: 0 などの隠れた文字は入らない）を、アウトラインの文案（固定値を SSOT の値に展開）と比べる。NFKC で幅をそろえ、空白と改行をすべて除いてから比べる（改行は見た目で、文言ではない）。同じアンカーが繰り返されるときは、見えているうちの 1 つが文案と一致すればよい。報告に表示の文言は出さない。
- **主張でない文言の宣言**：ナビ・節のラベル・見出し・法務やフッターのように主張でない文言は、アウトラインに `role: navigation` か `other` のセクションとして、安定した `key` と文言そのもの（`copy`）を書き、工場は `data-ia-text="<key>"` を付ける。主張になりうる文言（副の行動喚起など）は宣言ではなくメッセージにする。`dom` は、その key の見えている印の**すべて**が文案と一致することを確かめる（メッセージの繰り返しのような緩めは無い）。ページとの対応はメッセージと同じ（別のページの key は IA_DOM_ANCHOR_UNKNOWN、このページの key の印が無ければ IA_DOM_ANCHOR_MISSING）。`other` は順序の検査に入り、`navigation` は入らない（ヘッダーとフッターのどちらに置いてもよい）。`navigation` は画面幅によって畳んでよい（見えていなくても IA_DOM_ANCHOR_HIDDEN にしない。印が無ければ失格、見えている所は文案と一致）。
- **正式名称**：`data-ia-entity="ent-NNNN"` の見えている印は、すべて全文が SSOT の `name` と一致する（NFKC・空白を除いて比べる）。名前以外の文言を包めない。
- 宣言した文言・正式名称の印の中に、ほかの IA の印（メッセージ・痛み・宣言した文言・枠。正式名称は固定値と正式名称も）を入れると IA_DOM_ANCHOR_NESTED。宣言した文言は固定値と正式名称を中に持てる。
- `dom` の 0 は「描画物がアウトラインの宣言どおりに表示している」までで、主張の妥当性は確かめない。宣言した文言の意味は、ほかのセクションと同じく判定台帳（`ia-unregistered-claim`。role にかかわらず全セクションの文案を回す）が見る。
- どの印（メッセージ・痛み・固定値・正式名称・宣言した文言）にも属さない、見えている文言を報告の `unanchored`（成果物 × 画面幅）に出す。文案と違う繰り返しのアンカーの文言も、ここに足す。判定台帳では、支える根拠の無い主張として `ia-unregistered-claim` で判定する。
- このページの文案が参照しない固定値のアンカーは、SSOT にある id で、その値を表示している時だけ認める（任意の文言を固定値のアンカーで包んで、検査を逃がさない）。
- 文言が見えているかは、文字そのものが描かれているか（テキストの範囲の大きさ）で判定する。大きさ 0 の箱でも、はみ出して文字を描いていれば見えているとみなす。
- 限界：CSS の `::before`・`::after` の `content` で描く文言は読まない。`judge` の判定は Jev の助言で、score が θ 未満でも IA_DOM_UNANCHORED_TEXT の警告は消えない（アンカーの外の文言は、必ず人の目に届く）。

| code | 内容 |
| --- | --- |
| IA_DOM_NO_TARGETS | 描画されたテキストが 0 件 |
| IA_DOM_ANCHOR_MISSING / IA_DOM_ANCHOR_HIDDEN | アウトラインのメッセージ・痛み・宣言した文言のアンカーが無い・見えていない（display / visibility / opacity / 大きさ 0） |
| IA_DOM_ANCHOR_UNKNOWN | アウトラインに無いメッセージ・痛み・宣言した文言のアンカー（別のページの key を含む）・SSOT に無い固定値・正式名称のアンカー |
| IA_DOM_ANCHOR_NESTED | 宣言した文言・正式名称の印の中に、ほかの IA の印がある |
| IA_DOM_ENTITY_NAME | 正式名称の印の表示の全文が、SSOT の名前と違う |
| IA_DOM_ORDER | 見えているアンカーの描画順（上から）に、アウトラインの順が部分列として現れない（行動喚起の繰り返しは可） |
| IA_DOM_FACT_MISSING / IA_DOM_FACT_VALUE | 文案が参照する固定値の表示が無い・固定値のアンカーの表示が SSOT の値と違う（空白は正規化して比べる。文案が参照しない固定値のアンカーも対象） |
| IA_DOM_UNSTABLE | 2 回の描画で結果が違う |
| IA_DOM_COPY_DRIFT | 見えているアンカーのどれも、アウトラインの文案と一致しない（工場が文言を変えた）。宣言した文言は、1 か所でも違えば失格 |
| IA_DOM_COPY_EXTRA（warn） | 繰り返したアンカーのうち、文案と違うものがある（文案と一致するものは別にある） |
| IA_DOM_UNANCHORED_TEXT（warn） | どの印にも属さない、見えている文言がある＝IA の誰も管理していない文言（文字数だけを報告する。主張ならメッセージに、主張でなければ宣言する） |
| IA_DOM_OUTPUT_MISSING / IA_DOM_UNCOVERED（warn） | 成果物のファイルが無い・HTML 以外（未カバー） |
| IA_DOM_PLAYWRIGHT | Playwright を解決できない・描画に失敗した |

見えるかどうかは、要素とその祖先の display / visibility / opacity と、大きさ（箱が 0 でも文字を描いていれば見える）で判定する。コントラストや画面外への配置（clip-path・transform）は、デザイン側の L2 が持つ。

## 意味の判定（`judge`）

意味の判定は Jev が行い、その結果を判定台帳に固定する。検査は台帳だけを読み、その場で Jev を呼ばない。
方針ファイル（モデル・θ・プロンプト）・台帳（`design/judgments.json`）・使用量の台帳（累計 2,000 判定）は、own-design-author と共有する
（own-design-author の `references/judgments-and-modules.md`）。

| 種類 | 単位 | 入力の本文 | 文脈 |
| --- | --- | --- | --- |
| `ia-name-variant` | セクション（正式名称が 1 つ以上ある時） | 固定値を展開した文案 | 正式名称の一覧 |
| `ia-unregistered-claim` | セクション。加えて、成果物 × 画面幅ごとのアンカーの外の文言（id の末尾が `#unanchored`） | 固定値を展開した文案。アンカーの外は `dom` の `unanchored` | メッセージ・課題・証拠・使った固定値。アンカーの外は、すべて空（支える根拠が無い） |
| `ia-strategy-fit`（ISSUE-16 で「文案 × 提供価値 1 件」に置き換える） | セクション（メッセージがあるもの） | 固定値を展開した文案 | メッセージと、対象者のセグメントで立っている positioning の経路（原子と矢印のキーと文）と経路のハッシュ。positioning が変わると判定のキーが変わる。θ 以上は要確認で、出荷は止めない（精度を測るまで。D8）。`ship` の `positioning.semantic_fit` は judged / needs-review / not-judged |
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
holdout はそのまま引き継ぐ。回答の policy が今の共有 policy と違えば止める（migrate が要る）。既にあるファイルは上書きしない。置き換えは schema 3（判定の鍵が種類ごと）で作る。共有の台帳が schema 2 以前なら止める（先に `judgment_ledger.py rekey` で移す）。

| code | 内容 |
| --- | --- |
| IA_JUDGE_REFUSED | check が失格している（先に直す） |
| IA_JUDGE_NO_TARGETS | 判定の候補が 0 件 |
| IA_JUDGE_LEDGER | 台帳が読めない・policy と合わない・symlink |
| IA_JUDGE_DOM_FAILED / IA_JUDGE_DOM_REPORT | 失格した（終了コード 0・1 以外の）描画を判定に使おうとした・最初の画面の行が無いか読めない・`first_screens` か `unanchored` に、アウトラインの成果物と画面幅に無い行か重複がある（行の本文は報告に出さない）。警告だけの描画（アンカーの外の文言など）は判定に使う |
| IA_JUDGE_UNCOVERED（warn） | 台帳に無い判定がある（未カバー） |
| IA_JUDGE_REVIEW（warn） | score が θ 以上（要確認） |
| IA_JUDGE_EMPHASIS_MISSING（出荷される HTML なら warn、他は info） | 最初の画面の入力が無く、強調を判定していない。dom の報告が一部の成果物・画面幅だけなら、足りない組ごとに出る |

報告は `uncovered`（未カバー）・`screened`（台帳にあり。「問題なし」の保証ではない）・`missing`（入力が無く判定していない）を分けて出す。

## 作った物の評価（`evaluate`）

`infoarch-author evaluate --root . [--outline <id> ...]` は、作った物の評価の入口。工場は描画した後に必ず走らせる。

1. `check`。失格なら描画しない（`dom` と `judge` は `skipped`）。
2. 指定したアウトライン（無ければ全部）ごとに `dom`。1 つでも失格なら判定しない（`judge` は `skipped`）。
3. `dom` の報告（`first_screens`・`unanchored`）を渡して `judge`。

報告は 1 つで、同じ指摘は 1 回だけ出す。終了コードは全段の最大（0 / 1 / 2）。`stages` に段ごとの終了コードを載せる。
判定の候補（文案を含む）は出さない（必要なら `judge --include-candidates`）。

| code | 内容 |
| --- | --- |
| IA_EVALUATE_OUTLINE | 指定したアウトラインが無い |
| IA_EVALUATE_ACQUIRE | Jev の判定を取れない（上限超過・認証・通信・台帳の記録の失敗） |

`--acquire --answers-dir DIR --limit N` は、台帳に無い判定を Jev で取る。

- `--execute` が無ければ見積もりだけ（件数と要求数。キーを読まず、通信せず、何も書かない）。
- `--execute` があれば、own-design-author の `judgment_run`（共有の使用量台帳。累計の上限を超えるなら 1 件も送らない）で取り、回答を `DIR/judgment-answers.json` に足す（次の実行で続きから取る）。
- 取った回答から台帳の置き換えを作り、`judgment_ledger.py` で記録する（台帳が無ければ record、あれば refresh）。その後、評価をやり直す。
- 評価が失格している時は取らない（`acquisition.status: refused`）。
- 従量課金なので、`--execute` は実行の前に件数と上限を確かめる。

## 出荷ゲートと維持（`ship`・`maintain`）

出荷ゲートは、作業ツリーではなく、出荷されるコミット（clean な HEAD、push されるコミット、または予測マージ木）に対して走る。中身は、そのコミットを一時 worktree に取り出したもの（追跡されたファイルだけ）から読む。無視されたファイル・未コミットの変更・検査の後の書き換えは読まない。順番は次のとおりで、失格の段で止まる。
1. `tree`：`design/info_arch/` の下と `design/judgments.json` にシンボリックリンクがあれば `IA_SHIP_SYMLINK`（失格。`check` の前に見る）。
2. `check`（決定論）。出荷チャネル（`ship: true`）が 0 件、または出荷チャネルのどれか 1 つでもアウトラインが 0 件なら失格（空のゲートを合格にしない。他のチャネルが出荷していても、空のチャネルは合格にしない）。
3. `tree`（続き）：出荷されるアウトラインの成果物は、HTML かどうかに関わらず、すべてコミットの中に通常ファイルとして無ければならない。無い・ディレクトリ等は `IA_SHIP_OUTPUT_MISSING`、成果物か親ディレクトリがシンボリックリンクなら `IA_SHIP_SYMLINK`（どちらも失格）。git のオブジェクトから判定するので、手元にある無視されたファイルは数えない。
4. `dom`：出荷されるアウトラインのうち HTML の成果物を持つものを、実ブラウザで描画して確かめる。HTML が無いアウトラインは `skipped` に記録する。
5. `judge`：`dom` の最初の画面を渡して、共有の判定台帳を読む。出荷されるアウトラインの判定が台帳に無ければ `IA_SHIP_UNCOVERED`（失格）。出荷しないアウトラインの未カバーは `IA_JUDGE_UNCOVERED`（警告）のまま。θ 以上は `IA_JUDGE_REVIEW`（警告。人に回す。強調でゲートを止めるかは未決）。

| 入口 | 何をするか |
| --- | --- |
| `infoarch-author ship --root .` | HEAD を検査するだけ。出荷しない。作業ツリーは clean で、検査の後も同じコミットでなければ失格 |
| `infoarch-author ship --root . --channel <media>` | 正式の入口。緑なら、同じコミットのまま `ship_command` を実行する（shlex で argv に分け、シェルは通さない。`INFOARCH_SHIP_COMMIT`・`INFOARCH_SHIP_TREE`・`INFOARCH_SHIP_CHANNEL` を渡す）。実行の後に作業ツリーが変わっていたら失格 |
| `infoarch-author ship --root . --pre-push --refs-file F` | refs の各行（`<local ref> <local sha> <remote ref> <remote sha>`）を全部読む。削除（local sha が全部 0）は飛ばす。push される各コミットについて、(1) そのコミット自身の木、(2) チャネルの `merge_base` との予測マージ木、(3) 同じ push が `merge_base` の枝（`origin/main` なら `refs/heads/main`）も更新するなら、その新しいコミットとの予測マージ木、を検査する。同じ木は 1 回だけ検査する。`merge_base` は push されるコミット自身のチャネル定義から読む。無ければ止まる |
| `infoarch-author ship --root . --doctor` | 設定の診断。`ship_command` か `merge_base` の無い出荷チャネル、マーカー（`.vibe-guard/infoarch-ship.enabled`）の欠落、スキャナを保つ pre-push の dispatcher と拡張（`hooks/pre-push.d/61-infoarch-ship` と同じ中身）の欠落、拡張の終了コードを伝えない dispatcher を失格にする。dispatcher は、`ROOT=` の行を一時ディレクトリに差し替えた写しを、失敗する拡張と成功する拡張で実際に走らせて確かめる（本物のスキャナと導入済みの拡張は走らせない） |

`--pre-push` の終了コードは 0 か 2 だけ（dispatcher は 0 以外で push を止めるので、警告は `warnings` に載せて 0 で返す）。他の入口は 0 / 1（警告）/ 2。
pre-push の拡張は、マーカーのある repo だけで動く（opt-in）。マーカーは、作業ツリーのファイルに加えて、push されるコミット・push で置き換えられる側のコミット・HEAD・remote-tracking の枝のどれかにコミットされていれば有効とみなす。手元のファイルを消しても止まらない。コミットでマーカーを消す push は、置き換えられる側か `merge_base` にマーカーが残る間は検査され、消したことを stderr と `warnings` に出す（黙って外れない）。コミットされたマーカーが `1` の通常ファイルでなければ止まる。マーカーがあれば、壊れたマーカー・refs の欠落・検査の失格はすべて push を止める。
拡張は `INFOARCH_SHIP_COMMAND`（既定は PATH の `infoarch-author`）を呼ぶ。一時 worktree には node_modules が無いので、`INFOARCH_PLAYWRIGHT_DIR` が無ければ検査している checkout を使う。
出力は 1 行の JSON。CI での出荷ゲートは持たない（spec の Deferred）。

`infoarch-author maintain --root .` は own-session-close の Step 2.5 から呼ぶ。`design/info_arch/` が無ければ 1 行で skip する。
あれば `check` と `judge`（台帳を読むだけ。ブラウザは使わない）を走らせ、次を 1 行の JSON で返す。
- `due`：期限が `references/expiry.json` の `due_within_days` 日以内に来る証拠・固定値と、それを参照するアウトライン（出荷されるか）。
- `changed_paths`：出荷チャネルの `merge_base` からのコミット済みの差分と、未コミットの差分（`design/info_arch/` の中だけ）。
- `impacted_outlines`：変わった原子を参照するアウトライン（正式名称と声が変わったら全部）。
- `uncovered_judgments`：台帳に無い判定の数（Jev で取得する候補）。

終了コードは 0（整合）・1（期限が近い・未カバー・警告）・2（失格）。台帳もアウトラインも書き換えない。
