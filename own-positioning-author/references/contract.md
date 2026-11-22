# positioning の契約（design/info_arch/info_arch.yaml）

方針は、利用側の repo の spec `information-architecture-standardization`（v6）と `deliverable-production-layers`（layer の切り方）に従う。positioning（なぜ・誰に・何を）は repo に 1 つで、オーナーが提供価値ごとの経路を承認する。outline（成果物ごとの infoarch）は own-infoarch-author が持ち、positioning の id を引く。

## ファイル

`design/info_arch/info_arch.yaml`（スキーマ版 5。JSON Schema は `references/schema/info_arch.v5.schema.json`）。承認の状態は持たない（承認台帳 `design/approvals.yaml` が持つ）。変更は他の変更と同じく、検査が緑なら自律マージする（人のレビューは置かない）。記録は git の履歴が持つ。`design/` の下にあるのは名前と合わないが、パスを check と hook が参照しているので変えない（SPEC-deliverable-production-layers の Deferred Decisions）。

スキーマ版 4（固定値・正式名称・声を SSOT に持つ形）は、警告（IA_SSOT_LEGACY_V4）つきで読む。固定値と正式名称は辞書の値として扱い、声は読まない。`infoarch-author migrate-dictionary` で版 5 と辞書に移す。他の版は失格（IA_SCHEMA）。

## 原子（スキーマ版 5）

id は種別の接頭辞＋連番。連番は 4 桁（0001〜9999。0000 は不可）か、10000 以上なら先頭を 0 にしない。`msg-00001` や末尾に改行の付いたキーのような、同じ番号の別表記は失格（id は常に全体一致で判定する）。`ledger.next` 未満でなければ「台帳が発行していない」。改名・削除は `ledger.tombstones`（旧 → 新、削除は null）。

| 種別 | 接頭辞 | proposed で必須 | 参照 |
| --- | --- | --- | --- |
| 証拠 | `prf` | text・source・segment・as_of か expires | 支える原子・矢印は、矢印 `supports` で結ぶ（positioning の原子として承認する） |

対象者とメッセージは SSOT に持たない。outline の target（定数）とメッセージ（outline の原子）に書く（own-infoarch-author）。固定値と正式名称も持たない：`docs/dictionary/` の辞書が持つ（SPEC-deliverable-production-layers 要件 9）。声は廃止した（要件 8）。

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

**承認**：オーナーが提供価値ごとに、その経路をまとめて承認する（ADR-20261118-ia-approval-per-value-path、biz_ops）。`design/approvals.yaml` の行のキーは `ia:path:<提供価値 id>`、ハッシュは経路の原子と矢印すべての中身（`copied_from`・`provenance` を除く）を id でそろえ重複を除いて正規化した JSON の SHA-256（`positioning.extract`）。経路は次の 4 つの役からなり、下書きを含む proposed の全要素から毎回抜き出す（承認済みのものだけから抜き出さない。代替や度合いの矢印の追加・削除が必ずハッシュに映る）。

| 役 | 原子と矢印 |
| --- | --- |
| 主張（claim） | その提供価値、それを生む強みと `produces`、その強みの `lacks` とその先の代替 |
| 差別化の根拠（basis） | `satisfies` とその先の痛み・得たいもの、比べる集合の代替すべてと、その代替から満たす先への `fulfills` すべて。比べる集合は difference ならセグメントの代替すべて、parity なら所属カテゴリーが `includes` で入れている代替すべて（`includes` の矢印も入る） |
| 前提（premise） | セグメント、満たす先を持つジョブと `need`・`has_job`、所属カテゴリーと `member_of`、そのジョブの `recalls` とその先のカテゴリー |
| 証拠（evidence） | 上のどれかを `supports` する証拠とその矢印（`supports` を支える `supports` もたどる） |

経路の中身が 1 つでも変わると、その経路の承認が外れる。経路どうしで共有する原子（セグメント・ジョブ・所属カテゴリー・思い出す場面など）を直すと、それを含む経路の承認がまとめて外れる。ある経路の承認は別の経路に使い回せない（キーが提供価値を名指す）。原子と矢印ごとの旧い行（`ia:positioning:`）は承認として数えず、IA だけの repo では IA_APPROVALS_LEGACY（警告）、design.yaml のある repo では design の検査が「approval target absent」で止まる。旧い行は `infoarch-author migrate-approvals` で一度だけ経路の行へ移す（下）。承認の判断はオーナー、記録は AI（`design-author approve --request --owner-words --approved-on`。オーナー自身の承認の言葉が無ければ道具が断る。ADR-20261118-ai-records-owner-approvals、biz_ops）。positioning は outline より先に承認するので、outline 0 件でも束を作れる。design.yaml の無い IA だけの repo でも記録できる。承認かどうかは台帳のハッシュの一致だけで決まり、`state` には持たない。承認は採用の判断だけを表し、顧客での検証を意味しない。台帳は own-design-author と同じ `approvals.verify_rows` で読み（項目・キーの名前空間・ハッシュ・承認者・ISO の日付・キーの重複）、symlink も断る。どれかに反すれば IA_APPROVALS_INVALID で、何も承認済みとして数えない。束は `design-author approve --prepare` が経路ごとの行で作る。

**移行**（`infoarch-author migrate-approvals --root . [--write]`）：経路の原子と矢印がすべて旧い行とハッシュ（`positioning.element_hash`）一致で承認済みのときだけ、経路の行を作る。日付は元の行の日付、承認者は元の承認者（1 人でなければ作らない）。根拠のオーナーの言葉は、台帳を書いた commit の `Owner-Approval:` trailer から機械で読み、出力の `commit_trailers`（`Owner-Approval-Migrated: 「…」 (日付)`）に引用する。言葉が見つからない経路は作らない（`no-owner-words`）。1 つでも欠ける経路は未承認のまま（`unmigrated` に理由と欠けたキー）。`--write` で旧い行をすべて消し、経路の行と ds: などほかの行を残す（旧い行は git の履歴に残り、移行の commit を revert すれば戻る）。`--write` が無ければ書かない。言葉の使い回しはこの移行に限る。

**導出**（`positioning.derive`、記録しない）：positioning ＝ そのセグメントで承認済みの経路をすべて合わせたもの。どの承認済みの経路にも入らない原子と矢印は下書きのままで、positioning に入らず、outline から指せない。隙間と同等であるべき点は経路ごとに導く。隙間 ＝ 満たす先のうち、比べる集合（セグメントの代替すべて）のどれも `full` でないもの。同等であるべき点 ＝ 比べる集合（所属カテゴリーの代替）のすべてが `full` のもの。一部の代替だけが満たすものは、どちらにもならない。比べる集合の度合いの矢印が 1 本でも欠ければ導かない（未調査を `none` とみなさない）。

**id**：経路の id は `pth-` ＋ 経路のハッシュ。positioning の id は `pos-` ＋ そのセグメントで承認済みの経路の原子と矢印を合わせた中身の SHA-256（報告にだけ使い、set の id には入れない）。ファイルの中で書いた順番・写しの出どころ・provenance は入らない。`infoarch-author positioning --root .` が、セグメントごとの positioning（承認済みの経路・承認待ちの提供価値・隙間・同等であるべき点・所属カテゴリー）と、提供価値ごとの経路（`paths`：比べる集合・満たす先・失格・承認済みか）を出す。

**失格**は経路ごとに見る（要件 3・5）。出荷チャネルの outline が指す経路は失格（fail）、下書きの outline が指すか承認済みの経路は警告、どちらでもない下書きの経路は見ない。形の失格は proposed の全要素。条件と赤のフィクスチャの対応表は `fixtures/positioning-red-cases.json`（対応の無い条件 0 件をテストが確かめる）。

| code | 内容 |
| --- | --- |
| IA_POSITIONING_ISOLATED_ATOM | 経路の矢印にどれも繋がらない経路の原子（矢印の無い代替、何にも繋がらない提供価値） |
| IA_POSITIONING_JOB_WITHOUT_NEED | 満たす先の痛み・得たいものを持つジョブが無い（ジョブ→痛み・得たいものの矢印が無い） |
| IA_POSITIONING_INCOMPLETE | 経路のジョブか痛み・得たいものが 0 件、またはジョブがセグメントに `has_job` で繋がらない（Quality 2：中身の無い経路を合格にしない） |
| IA_POSITIONING_CATEGORY_COUNT | 所属カテゴリーが 0 本か 2 本以上 |
| IA_POSITIONING_NO_ALTERNATIVES | 隙間の比較に使うセグメントの代替が 0 件 |
| IA_POSITIONING_NO_CATEGORY_ALTERNATIVES | 同等であるべき点の比較に使う所属カテゴリーの代替が 0 件 |
| IA_POSITIONING_FULFILLMENT_MISSING | 比べる集合の代替と満たす先の組で、満たす度合いの矢印が欠ける（`detail.pairs`）。比べる集合の外の組は見ない |
| IA_POSITIONING_VALUE_UNANCHORED | difference の提供価値が隙間を、parity の提供価値が同等であるべき点を 1 つも満たさない |
| IA_POSITIONING_STRENGTH_UNCOMPARED | 代替と比べない（`lacks` の無い）強み |
| IA_POSITIONING_EDGE_SHAPE | 矢印の始点・終点の種類が表と違う、`degree` が fulfills に無い・ほかの種類にある |
| IA_POSITIONING_EDGE_DUPLICATE | 同じキーの矢印が 2 本 |
| IA_POSITIONING_CROSS_SEGMENT | 矢印が別のセグメントの原子を結ぶ（共有しない。写しを作る） |
| IA_POSITIONING_COPY_SOURCE | 写しの出どころが無い・別の種類・同じセグメント |
| IA_POSITIONING_PATH_PENDING | 承認済みの経路があるセグメントに、どの outline も指さない未承認の経路が残る（警告。止めない） |

## 承認の依頼と読み戻し（`positioning-author request`）

オーナーに見せる承認依頼のページを、承認の対象と同じデータから機械で作る（SPEC-deliverable-production-layers 要件 12）。

- `positioning-author request --root . --out DIR`：提供価値の経路ごとに 1 枚のカード（主張→差別化の根拠→前提→証拠。比べる集合・満たす先・失格＝「決めてほしいこと」・ハッシュ先頭 12 桁）を持つ自己完結の HTML と、`design-author approve --prepare` が束縛する行（`approval-request.json`）を出す。既にあるファイルは上書きしない。経路が 0 件なら出さない（exit 2）。
- ページは非公開の Artifact として出す（呼んだエージェントが `capabilities: {db: {}}` で公開する。skill は HTML を一時フォルダに書くまで）。各カードに ○（承認する）・×（差し戻す）とメモがあり、選ぶと Artifact の保存先 `decisions/<提供価値 id>`（`{key, value_hash, decision, note, decided_at}`）に残る。download リンクには頼らない。保存できない画面では、その旨と、コピーして AI に渡す JSON の控えを出す。
- `positioning-author request --root . --read-decisions <rows.json>`：保存先から読み戻した行（`ArtifactData` の list の結果、または控えの JSON）を、今のカードと突き合わせる。`ok`・`ng`・`stale`（経路の中身が変わって、判断した時のハッシュと合わない。数えない）・`undecided`・`unknown` を返す。stale と unknown があれば exit 1。
- ページは承認台帳に書かない。台帳に書くのは、オーナーの承認の言葉を受けた AI だけ（ADR-20261118-ai-records-owner-approvals）。○ は承認の言葉の代わりにならない。
- `--must-file`：`infoarch-author targets` の出力を渡すと、出荷が要る経路に「出荷に必要」の印を付ける（outline を知っているのは infoarch の側）。

## 読む・移す

- `positioning-author review --root . --request <束>`：束が束縛する経路を、経路ごとに全部（主張→差別化の根拠→前提→証拠の順、比べる集合・満たす先・失格つき）Markdown で出す。見せる範囲とハッシュの範囲は同じ。今のハッシュと合わない行があれば `stale` に並べて失格（作り直す）。
- `positioning-author positioning --root .`：セグメントごとの positioning と、提供価値ごとの経路。
- `positioning-author targets --root .`：提供価値ごとの経路のハッシュ。出荷が要るキーは `infoarch-author targets`。
- `positioning-author migrate-approvals --root . [--write]`：原子と矢印ごとの旧い承認の行を、経路の行へ一度だけ移す。
- `positioning-author check --root . [--strict]`：positioning だけの検査。要件 5 の失格は、出荷の outline が指す経路だけを止めるのが正規の動きで、outline を知っているのは `infoarch-author check`。`--strict` は、すべての提供価値の経路が出荷の outline から指されているものとして失格にする（赤ケースの表 `fixtures/positioning-red-cases.json` を outline 無しで走らせる入口）。要件 15（出荷の outline が未承認の経路を指す）は outline が要るので、infoarch の検査だけが出す。提供価値が 0 件なら IA_NO_VALUES で失格。

## 検査の code（positioning）

IA_MISSING / IA_YAML / IA_SCHEMA、IA_ID_NOT_ISSUED / IA_ID_TOMBSTONED / IA_TOMBSTONE_TARGET、IA_REF_BROKEN / IA_EMPTY_REFERENCED、IA_POSITIONING_*（上の表）、IA_APPROVALS_LEGACY / IA_APPROVALS_INVALID、IA_SSOT_LEGACY_V4（警告）、IA_NO_VALUES（positioning-author の check だけ）。条件と赤のフィクスチャの対応表は `fixtures/positioning-red-cases.json`。
