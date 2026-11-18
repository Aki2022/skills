# design/info_arch/ の契約

方針は、利用側の repo の spec `information-architecture-standardization`（v6）に従う。

## ディレクトリ

- `info_arch.yaml`：人が保守する SSOT（positioning の原子と矢印、固定値・正式名称・声、id 台帳）。承認の状態は持たない（承認台帳が持つ）。変更は他の変更と同じく、検査が緑なら自律マージする（人のレビューは置かない）。記録は git の履歴が持つ。
- `channels/<media>.yaml`：成果物の置き場所と出し方の宣言だけ。中身は `media`（ファイル名と一致）・`outputs`（repo 相対の glob）・`ship`・`ship_command`（`ship: true` なら必須。出荷の正式の入口が実行する）・`merge_base`（pre-push の予測マージの相手。出荷チャネルで無いと doctor が失格）・`viewports`。
- `outlines/<成果物 id>.yaml`：成果物ごとの outline（版 3）。1 ページ＝1 outline。AI が書き、人は保守しない。`id` はファイル名と一致する。
- `INFO_ARCH.md`：散文。原子を id で参照するだけで、値を持たない。

それ以外のファイルやサブディレクトリが `channels/` `outlines/` にあると失格にする（黙って無視しない）。`.` で始まるファイル（`.DS_Store`・`.gitkeep`）だけは読まずに飛ばす。`.` で始まるサブディレクトリは失格にする。

## 原子（スキーマ版 4）

id は種別の接頭辞＋連番。連番は 4 桁（0001〜9999。0000 は不可）か、10000 以上なら先頭を 0 にしない。`msg-00001` や末尾に改行の付いたキーのような、同じ番号の別表記は失格（id は常に全体一致で判定する）。`ledger.next` 未満でなければ「台帳が発行していない」。改名・削除は `ledger.tombstones`（旧 → 新、削除は null）。

| 種別 | 接頭辞 | proposed で必須 | 参照 |
| --- | --- | --- | --- |
| 証拠 | `prf` | text・source・segment・as_of か expires | 支える原子・矢印は、矢印 `supports` で結ぶ（positioning の原子として承認する） |
| 固定値 | `fct` | label・value・as_of か expires か evergreen: true（evergreen と日付は併存させない） | ― |
| 正式名称 | `ent` | type・name | ― |
| 声 | `vce` | traits（3〜4 個。trait・is・is_not） | proposed はプロジェクトに 1 つ |

対象者とメッセージは SSOT に持たない。outline の target（定数）とメッセージ（outline の原子）に書く。

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

### outline が指す経路（`check.referenced_values`・`check.positioning_path`）

outline のメッセージが指す提供価値のうち、target のセグメントのものが、その outline の指す経路。出荷チャネルの outline が指す経路が 1 つでも承認されていない（ハッシュが合わない）と出荷で止まる（IA_POSITIONING_UNAPPROVED、`detail.values`。下書きは警告）。比べる相手（`compares`）は、その提供価値の経路の比べる集合の中に無ければ失格（IA_OUTLINE_COMPARE_OUTSIDE_SET）。比べる代替は経路と一緒に承認される。指していない経路が未承認でも止めない（IA_POSITIONING_PATH_PENDING の警告）。`check.positioning_path` は提供価値の経路を合わせたキーとハッシュを返し、意味の判定（未登録の主張）の証拠の入力にも使う。一番のメッセージの経路に証拠が 1 つも無ければ警告（IA_LEAD_UNSUPPORTED）。

成果物が立つもの（期限・逆引き・maintain の影響）は、指す経路の原子と矢印すべて：隙間と同等であるべき点は経路の比べる集合の代替と度合いの矢印から決まるので、その証拠の期限切れや変更も成果物に届く。

### 承認の対象（`ia:path:<提供価値 id>`）

- `infoarch-author review --root . --request <束>`：束が束縛する経路を、経路ごとに全部（主張→差別化の根拠→前提→証拠の順、比べる集合・満たす先・失格つき）Markdown で出す（オーナーが読む）。見せる範囲とハッシュの範囲は同じ。今のハッシュと合わない行があれば `stale` に並べて失格（作り直す）。
- `infoarch-author targets --root .`：proposed の全部の提供価値の経路のハッシュ（`approval_hashes`）と、出荷チャネルの outline が指す経路のキー（`must_approval_keys`）を出す。`ia_module.py`（契約 `ia-git-history-v4`、`modules.ia: 0.4.0`、governance の schema_version 4）がこれを own-design-author に返す。キーの形は `ia:path:val-NNNN` だけ。
- 出荷の報告（`ship`）は `positioning` に、`owner_adopted`（approved / unapproved）・`semantic_fit`（意味の判定の状態。今は not-judged）・`customer_validation`（IA は判定しない：not-judged-by-ia）を分けて出す。

## outline（版 3。SPEC v6 要件 6〜9・13）

- 置き場所：`channel`・`outputs`（チャネルの outputs の glob に一致すること）。`id`・`channel`・`outputs` は outline の id に入らない。
- 定数：`target`（`segment`＋`condition`。outline はこのセグメントの positioning にだけ立つ）・`touchpoint`（接点）・`tone`。目的変数：`objective`（`id`・`action`）。
- `messages`（`msg-NNNN`）：`value`（提供価値ちょうど 1 つ）・`rank`（lead / support）・`type`（normal / cta）・`compares`（比べる代替）・`achieves`（CTA が指す目的変数）。`no_cta: true` で CTA を置かないことを宣言できる。
- `sections`（`sec-NNNN`）：`order`・`role`（aware / pain / compare / trust / act）・`places`（通常のメッセージ 1 つと CTA 0〜1 つ）。
- `copies`（`cpy-NNNN`）：`section`・`order`（セクションの中の順番）・`message`（同じセクションに置かれたもの）・`text`。固定値は `{{fct-NNNN}}` で参照する（括弧の中に空白や改行を入れない。入れたものは形の崩れた参照として失格）。
- `texts`（任意）：主張でない文言の宣言。key ごとに `kind`（navigation / label）・`text`・（label なら）`section`。
- 並びは `order` で持ち、ファイルの中で書いた順番は意味を持たない。outline の原子の id は outline の中で一意で、台帳は持たない（別案は同じ id を引き継ぐ）。

### id と set（`infoarch-author set --outline <id>`）

- outline の id：`out-` ＋ outline の中身（定数・目的変数・`no_cta`・メッセージ・セクションの順序と置き方・文案の順番と文言・`texts`）を正規化した JSON の SHA-256。立つ positioning の id は入れない。メッセージは提供価値を原子の id で指すので、提供価値の文面だけを直しても outline の id は変わらない。
- set の id：`<outline の id>+<経路の id>+…`（outline が指す経路の id を提供価値の id の順に並べる）。中身から決まり、承認の有無では変わらない（承認は出荷ゲートが見る）。指していない経路の手直しでは変わらず、指している経路の中（共有する原子を含む）を直すと変わる。指す経路が無ければ None。`set` は `paths`（提供価値 → 経路の id）、報告用の positioning の id、要素（messages・sections・copies・texts）ごとの中身のハッシュも出す。v5 の対応表（mapping）はこれに置き換えた。

### 差分（`infoarch-author diff --a <id> --b <id> [--root-b DIR] [--dom-a A --dom-b B]`）

2 つの set の差分を、変わった要素（`messages/<id>` などの outline の要素と、経路の原子・矢印）と属性で列挙する。指す経路は提供価値の id で突き合わせ、両方にあってハッシュが違えば positioning の編集（`layer: positioning`、`path` に提供価値の id）、片方にしか無ければ outline の選択（`layer: outline`、`element: paths/<提供価値 id>`、`kind: selection`）として別々に出す。両方の文案が参照する固定値の値が違えば、IA_DIFF_UNASSIGNED で失格（固定値はどちらの id にも入らない。値は報告に出さない）。`changed_layers` で、どちらの層が変わったかが分かる（両方が変わる別案も認める）。定数か目的変数が違えば `same_experiment: false`（別の試験として扱う。失格ではない）。両方の dom の報告を渡すと、印の外の文を画面幅ごとに比べ、違えば IA_DIFF_UNASSIGNED で失格（どの要素にも割り当てられない差分）。渡さなければ `not_checked` に比べていないことを出す。A/B の実行・指標・統計は持たない。

## v5 からの移行（`infoarch-author migrate-v6 --root . [--write]`。SPEC v6 要件 17）

v5 の木（`info_arch.yaml` 版 3・outline 版 2・描画物の v5 の印・`modules.ia` 0.3.0）を、一括で v6 に移す。`--write` が無ければ何も書かない（見積もりだけ）。移しそこねが 1 件でもあれば終了コード 2 で、`--write` でも何も書かない（一括。半分だけ移さない）。commit はしない。v6 の木（版 4）では何もせず終了コード 0（再実行できる）。版 3 でも 4 でもなければ失格。

機械で決まることだけを移す。

| v5 | v6 |
| --- | --- |
| 約束（`prm`）・コンセプト（`cpt`） | 無くす（`dropped`）。コンセプトのカテゴリーは、カテゴリーを下書きする材料として `drafting_required` に出す |
| ジョブの `segment`・痛みと得たいものの `job`・提供価値の `relieves` / `creates`・強みの `versus` / `produces` | 矢印 `has_job`・`need`・`satisfies`・`lacks`・`produces`。痛み・得たいもの・提供価値・強みには、たどった先のセグメントを書く。ジョブの 4 つの力は無くす |
| 証拠の `supports`（メッセージ） | positioning の証拠。そのメッセージが担っていた提供価値・ジョブ・痛み・得たいもの（約束は `values` をたどる）への `supports`。2 つのセグメントに届く証拠は、セグメントごとに写しを作る（証拠は `copied_from` を持たないので、出どころは `provenance` の `copy of prf-…` に書く） |
| outline の対象者（`aud`） | outline の定数 `target`（セグメントと、対象者の名前を足す条件に写す）と `tone`。`brief.action` は目的変数 `objective` |
| outline が使うメッセージ（`msg`） | outline の原子（同じ id）。target のセグメントの提供価値をちょうど 1 つ指す。v5 の役割が cta なら CTA（`achieves: obj-0001`）。一番は対象者の `priorities` の先頭（無ければ最初の通常のメッセージ） |
| セクション（ページの中の並び） | セクション 1 つと文案 1 つ（`sec-NNNN`・`cpy-NNNN`、並びの順）。役割は表で読み替える（pain→pain、benefit・vision→aware、proof・how→trust、objection→compare、cta→act）。痛みだけのセクションは、その痛みを扱うこの outline のメッセージ（`addresses`、無ければ提供価値の満たす先）が 1 つに決まれば置く。navigation・other の宣言（`key`）は `texts` |
| 枠（`frames`）と対応表 | セクションの id（並びの枠は `sections.order`）と set の id（`set` で出す） |
| 複数ページの outline | ページ数と outputs の数が同じならページごとの outline（`<id>-<ページ id>`）に分け、元のファイルを消す |
| 描画物の `data-ia-role`＋`data-ia-msg`・`data-ia-pain`・`data-ia-frame` | セクションの要素に `data-ia-section`（と v6 の役割）。文案は、中身が要素 1 つならその要素に `data-ia-copy`、そうでなければ中身を `<span data-ia-copy style="display:contents">` で包む（見た目を変えない）。印は outline のセクションと順に、数と id で突き合わせる |
| `design/design.yaml` の `modules.ia: 0.3.0` | v6 の契約の版（`ia_module.VERSION`） |

判断が要るもの（提供価値の `kind`、所属カテゴリーと `member_of`・`recalls`・`includes`、代替 × 痛み・得たいものの `fulfills` の度合い、CTA だけのセクションに置く通常のメッセージ、役割・touchpoint・target の条件の確かめ）は `drafting_required` に並べ、作らない。移した直後の木は、これを済ませるまで `check` が失格する。

出力：`id_map`（旧 id → 新 id。outline の要素は `<outline>:<id>`、v5 のセクションは `<outline>:<ページ>[<番号>]`、枠は `<outline>#<枠>`。`[]` は無くしたもの。載っていない id は同じ id のまま）・`unmigrated` と `unmigrated_count`・`dropped`・`drafting_required`・`role_map`・`stops_applying`・`module_pin`・`files`（write・create・delete）・`rollback`（戻す手順）。

`stops_applying` は、外れる承認と判定の一覧。`design/approvals.yaml` の `ia:` の行のうち `ia:path:` でないもの（v5 の `ia:strategy:` は経路の承認に移らない。オーナーが経路を新たに承認する。`ia:positioning:` の行は `migrate-approvals` で移せるか確かめる）と、`design/judgments.json` の IA の判定（`ia-` の種類。v6 の鍵は v6 の outline の id と文案・提供価値の中身から作るので読まれない）。どちらの台帳も読むだけで、書かない。

移しそこね（`unmigrated`。条件と赤のフィクスチャの対応表は `fixtures/migrate-red-cases.json`。対応の無い条件 0 件をテストが確かめる）：

| reason | 内容 |
| --- | --- |
| message-value-none / message-value-several | outline が使うメッセージが、target のセグメントの提供価値を 0 個か 2 個以上指す |
| message-unused | どの outline も使わない proposed のメッセージ（移す先が無い。v5 で墓標にしてから移す） |
| audience-without-segment | outline の対象者が無いか、セグメントを持たない |
| atom-segment | 痛み・得たいもの・提供価値・強みのセグメントが 1 つに決まらない |
| proof-without-target | proposed の証拠が positioning の原子に 1 つも届かない |
| section-without-message | メッセージの決まらないセクション |
| pages-outputs-mismatch | 複数ページで、ページ数と outputs の数が違う |
| html-unmatched | 描画物の v5 の印が outline のセクションと合わない、書き換えた後も旧い印が残る、outline に属さない描画物に旧い印がある |
| module-pin-unknown | `modules.ia` が 0.3.0 でも v6 の版でもない |

戻す手順：commit 前なら `git restore` で `files.write` と `files.delete` を戻し、`files.create` を消す。commit 後なら移行の commit を `git revert` する。承認台帳と判定台帳は書いていないので戻すものは無い。

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
| IA_POSITIONING_UNAPPROVED | outline が指す経路に、オーナーの承認が無いか外れたものがある（`detail.values`。出荷チャネルは失格、ほかは警告） |
| IA_APPROVALS_LEGACY | 承認台帳に原子と矢印ごとの旧い行（`ia:positioning:`）がある。数えない（警告。`migrate-approvals` で移す） |
| IA_APPROVALS_INVALID | `design/approvals.yaml` が読めない・symlink・design の `verify_rows` を通らない（重複キー・項目の過不足・承認者・日付・ハッシュ）。何も承認済みとして数えない |
| IA_PAGES_OUTPUT_NOT_INVENTORIED（警告） | 出荷チャネルの HTML の成果物が、design/validation.json の pages（covered・excluded）に無い（design の検査が見ていない） |
| IA_PAGES_OUTPUT_EXCLUDED_BY_DESIGN（info） | design が対象外にした成果物。IA の検査は続ける（ページの宣言は一本化しない。ADR-20261115-design-ia-two-skills-shared-base） |
| IA_NO_CHANNELS / IA_NO_OUTLINES / IA_CHANNEL_NAME / IA_CHANNEL_FILE / IA_OUTLINE_FILE / IA_OUTLINE_ID | 対象 0 件・名前の不一致・読まれないファイル |
| IA_OUTLINE_CHANNEL / IA_OUTLINE_OUTPUT_UNDECLARED | 宣言の無いチャネル・チャネルの outputs に一致しない成果物 |
| IA_OUTLINE_REF_BROKEN / IA_OUTLINE_REF_EMPTY | アウトラインの参照切れ（墓標なら先を案内）・空欄の原子を参照（参照の表は `check.outline_refs` の 1 か所） |
| IA_OUTLINE_BAD_PLACEHOLDER / IA_OUTLINE_FACT_LITERAL | `{{fct-NNNN}}` の形でない `{{…}}`（改行をまたぐものを含む）・固定値の値を参照せず直接書いた |
| IA_IMPACT_UNKNOWN_ID | `impact --id` に SSOT にも墓標にも無い id を渡した（空の結果を合格にしない） |
| IA_OUTLINE_TARGET | target のセグメントが無いか空欄（target が無い・形が違うのは IA_SCHEMA） |
| IA_OUTLINE_VALUE_UNAPPROVED | メッセージが、経路の承認が無いか外れた提供価値を指す（出荷チャネルは失格、ほかは警告） |
| IA_OUTLINE_VALUE_OTHER_POSITIONING | メッセージが指す提供価値の経路のセグメントが、target のセグメントと違う |
| IA_OUTLINE_COMPARE_INVALID | 比べる相手が立っている positioning の代替でない |
| IA_OUTLINE_COMPARE_OUTSIDE_SET | 比べる相手が、その提供価値の経路の比べる集合に無い（parity なら所属カテゴリーの代替だけ） |
| IA_OUTLINE_LEAD_COUNT / IA_OUTLINE_LEAD_UNPLACED | 一番のメッセージが 1 つでない・どのセクションにも置かれていない |
| IA_OUTLINE_CTA_MISSING / IA_OUTLINE_CTA_CONTRADICTS / IA_OUTLINE_CTA_NO_ACTION | 宣言なしで CTA を置いていない・`no_cta` と書いたのに CTA がある・CTA が目的変数を指さない |
| IA_OUTLINE_NO_SECTIONS / IA_OUTLINE_SECTION_NO_MESSAGE / IA_OUTLINE_SECTION_OVERLOADED | セクションが 0 件・通常のメッセージを置かないセクション・通常 2 つ以上か CTA 2 つ以上のセクション |
| IA_OUTLINE_COPY_SECTION / IA_OUTLINE_COPY_MESSAGE / IA_OUTLINE_MESSAGE_UNCONVEYED | セクションに属さない文案・同じセクションに無いメッセージを伝える文案・置かれたメッセージを伝える文案が無い |
| IA_OUTLINE_ORDER_DUPLICATE | セクションの順序、同じセクションの文案の順番が重なる |
| IA_OUTLINE_TEXT_MISPLACED | ナビの文言をセクションに置いた・ラベルを置くセクションが無い |
| IA_OUTLINE_MESSAGE_UNPLACED（warn） | 一番以外のメッセージが、どのセクションにも置かれていない |
| IA_OUTLINE_NO_CTA_INTENDED（info） | `no_cta: true` の outline。報告に必ず「CTA なし（意図）」を出す |
| IA_LEAD_UNSUPPORTED（warn） | 一番のメッセージの提供価値の経路に、証拠が `supports` する原子・矢印が無い（CTA は対象外） |
| IA_DIFF_UNASSIGNED | 別案どうしで、印の外の文か参照する固定値の値が違う（`diff`） |
| IA_EXPIRED_SHIPPED / IA_EXPIRED_UNSHIPPED（info） | 期限切れの証拠（経路から受け継ぐもの）・固定値（文案が参照するもの）に、出荷チャネル（`ship: true`）の outline が立っている／立っていない |

条件と赤のフィクスチャの対応表：`fixtures/red-cases.json`（参照と形）・`positioning-red-cases.json`（要件 5）・`outline-red-cases.json`（要件 7・9）。

## 期限

`expires` があればその日、無ければ `as_of` に日数を足した日を期限とする。期限の日の当日までは有効で、翌日から期限切れ。日数の初期値は `references/expiry.json`、プロジェクトの
`info_arch.yaml` の `policy.expiry_days` で上書きできる。`evergreen: true` の固定値は期限を持たない。

## 工場への受け渡し（`skeleton`）

`infoarch-author skeleton --outline <id>` は、check が緑のアウトラインから骨組みの HTML を出す（失格していれば出さない）。
セクションを `order` の順に `<section data-ia-section="sec-…" data-ia-role="…">` で、その中の文案を `order` の順に `data-ia-copy="cpy-…"` で出す。
ナビの宣言は `<nav data-ia-text>`、ラベルはそのセクションの先頭に `data-ia-text` で出す。文案の `{{fct-NNNN}}` は SSOT の値の
`<span data-ia-fact>` に展開し、文案は HTML エスケープする。1 ページ＝1 outline なので、どの output にも outline 全体を出す。`--out` を付けるとファイルに書き、既にあるファイルは上書きしない。
工場（own-web-design）は、文言・順序・`data-ia-*` を変えずにスタイリングし、描画後に `dom` を走らせる。

## 描画物の確認（`dom`）

`infoarch-author dom --root . --outline <id>` は、アウトラインの outputs のうち HTML を、チャネルの viewports（無ければ 375×812 と 1280×800）でブラウザに描画し、DOM を確かめる。

- 解析と描画はブラウザ（Playwright の Chromium）が行う。成果物のソースを自前のパーサで読まない。
- 工場はアンカーを埋める：セクションは `data-ia-section`、文案は `data-ia-copy`、固定値の表示は `data-ia-fact`、正式名称（SSOT の `ent-`）の表示は `data-ia-entity`、宣言した文言は `data-ia-text`。v4・v5 の `data-ia-problem`・`data-ia-msg`・`data-ia-pain`・`data-ia-frame` が残っていれば IA_DOM_LEGACY_ATTRIBUTE で失格（互換は持たない。v5 の描画物は `infoarch-author migrate-v6` が読み替える）。
- **画面の適合（SPEC v6 要件 10）**：セクションの印は outline の順序どおりにちょうど並ぶ（IA_DOM_ORDER）。文案の印は、outline で属するセクションの印の中にあり（IA_DOM_COPY_SECTION）、セクションの中で outline の順番どおりに並ぶ（IA_DOM_COPY_ORDER）。同じ文案の印は 1 つだけ（IA_DOM_COPY_REPEATED。繰り返したい行動喚起は、outline に別の文案として置く）。条件と赤のフィクスチャの対応表は `fixtures/screen-red-cases.json`（DOM の写しと実ブラウザの両方で確かめる）。
- Playwright は `INFOARCH_PLAYWRIGHT_DIR`、無ければ対象 repo（実行ディレクトリ）から解決する。自動では入れない。解決できなければ IA_DOM_PLAYWRIGHT で失格。
- 読み取る前に、ページが落ち着くのを待つ：network idle の後に 1 秒（表示を始めるタイマーのため）、さらに実行中のアニメーションとトランジションが終わるまで最大 5 秒。終わらないアニメーションがあっても読み取り、2 回の描画の比較で揺れを報告する。
- 同じ入力を 2 回描画し、結果が違えば IA_DOM_UNSTABLE（判定に使わない）。
- 最初の画面の文言（各画面幅）を報告の `first_screens` に出す。強調の判定（Jev の判定台帳）の入力になる。
- 各文案のアンカーの表示の文言（描かれている文字だけ。`unanchored` と同じ読み方で、opacity: 0 などの隠れた文字は入らない）を、アウトラインの文案（固定値を SSOT の値に展開）と比べる。NFKC で幅をそろえ、空白と改行をすべて除いてから比べる（改行は見た目で、文言ではない）。見えている印はすべて一字一句一致すること（IA_DOM_COPY_DRIFT）。報告に表示の文言は出さない。
- **主張でない文言の宣言**：ナビ・節のラベル・見出し・法務やフッターのように主張でない文言は、outline の `texts` に安定した key と文言そのもの（`text`）を書き、工場は `data-ia-text="<key>"` を付ける。主張になりうる文言（副の行動喚起など）は宣言ではなくメッセージにする。`dom` は、その key の見えている印の**すべて**が文言と一致することを確かめる（繰り返しの緩めは無い）。outline に無い key は IA_DOM_ANCHOR_UNKNOWN、key の印が無ければ IA_DOM_ANCHOR_MISSING。`navigation` は画面幅によって畳んでよい（見えていなくても IA_DOM_ANCHOR_HIDDEN にしない。印が無ければ失格、見えている所は文言と一致）。
- **正式名称**：`data-ia-entity="ent-NNNN"` の見えている印は、すべて全文が SSOT の `name` と一致する（NFKC・空白を除いて比べる）。名前以外の文言を包めない。
- 文案・宣言した文言・正式名称の印の中に、ほかの IA の印（セクション・文案・宣言した文言。正式名称は固定値と正式名称も）を入れると IA_DOM_ANCHOR_NESTED。文案と宣言した文言は固定値と正式名称を中に持てる。
- `dom` の 0 は「描画物が outline の宣言どおりに表示している」までで、主張の妥当性は確かめない。文案の意味は判定台帳（`ia-unregistered-claim`。role にかかわらず全部の文案を回す）が見る。
- どの印（文案・固定値・正式名称・宣言した文言）にも属さない（セクションは入れ物なので、セクションの中でも文案の外なら属さない）見えている文言があれば、IA_DOM_UNANCHORED_TEXT で失格（要件 10。v5 の D12 では警告だった。文言は報告に出さない）。その文言は報告の `unanchored`（成果物 × 画面幅）に出す（文案と違う印の文言も、ここに足す）。dom が失格した描画は判定に使わないので、まず outline へ戻す。
- このページの文案が参照しない固定値のアンカーは、SSOT にある id で、その値を表示している時だけ認める（任意の文言を固定値のアンカーで包んで、検査を逃がさない）。
- 文言が見えているかは、文字そのものが描かれているか（テキストの範囲の大きさ）で判定する。大きさ 0 の箱でも、はみ出して文字を描いていれば見えているとみなす。
- 限界：CSS の `::before`・`::after` の `content` で描く文言は読まない。印の外の文言は、判定の結果にかかわらず失格のまま（必ず outline に戻す）。

| code | 内容 |
| --- | --- |
| IA_DOM_NO_TARGETS | 描画されたテキストが 0 件 |
| IA_DOM_ANCHOR_MISSING / IA_DOM_ANCHOR_HIDDEN | outline のセクション・文案・宣言した文言の印が無い・見えていない（display / visibility / opacity / 大きさ 0） |
| IA_DOM_ANCHOR_UNKNOWN | outline に無いセクション・文案・宣言した文言の印・SSOT に無い固定値・正式名称の印 |
| IA_DOM_ANCHOR_NESTED | 文案・宣言した文言・正式名称の印の中に、ほかの IA の印がある |
| IA_DOM_ENTITY_NAME | 正式名称の印の表示の全文が、SSOT の名前と違う |
| IA_DOM_ORDER | 見えているセクションの印の描画順（上から）が、outline の順序とちょうど同じでない |
| IA_DOM_COPY_ORDER / IA_DOM_COPY_SECTION / IA_DOM_COPY_REPEATED | セクションの中の文案の順番が違う・文案が別のセクションの中にある・同じ文案の印が 2 つ以上 |
| IA_DOM_LEGACY_ATTRIBUTE | v4・v5 の印（`data-ia-problem`・`data-ia-msg`・`data-ia-pain`・`data-ia-frame`）が残っている |
| IA_DOM_FACT_MISSING / IA_DOM_FACT_VALUE | 文案が参照する固定値の表示が無い・固定値のアンカーの表示が SSOT の値と違う（空白は正規化して比べる。文案が参照しない固定値のアンカーも対象） |
| IA_DOM_UNSTABLE | 2 回の描画で結果が違う |
| IA_DOM_COPY_DRIFT | 見えている文案の印の文言が、outline の文案と一字一句同じでない（工場が文言を変えた）。宣言した文言も、1 か所でも違えば失格 |
| IA_DOM_UNANCHORED_TEXT | どの印にも属さない、見えている文言がある＝IA の誰も管理していない文言（文字数だけを報告する。主張なら outline のメッセージと文案に、主張でなければ texts に宣言する） |
| IA_DOM_OUTPUT_MISSING / IA_DOM_UNCOVERED（warn） | 成果物のファイルが無い・HTML 以外（未カバー） |
| IA_DOM_PLAYWRIGHT | Playwright を解決できない・描画に失敗した |

見えるかどうかは、要素とその祖先の display / visibility / opacity と、大きさ（箱が 0 でも文字を描いていれば見える）で判定する。コントラストや画面外への配置（clip-path・transform）は、デザイン側の L2 が持つ。

## 意味の判定（`judge`）

意味の判定は Jev が行い、その結果を判定台帳に固定する。検査は台帳だけを読み、その場で Jev を呼ばない。
方針ファイル（モデル・θ・プロンプト）・台帳（`design/judgments.json`）・使用量の台帳（累計 2,000 判定）は、own-design-author と共有する
（own-design-author の `references/judgments-and-modules.md`）。

| 種類 | 単位 | 入力の本文 | 文脈 |
| --- | --- | --- | --- |
| `ia-name-variant` | 文案（正式名称が 1 つ以上ある時。id は `<outline>/<sec>/<cpy>`） | 固定値を展開した文案 | 正式名称の一覧 |
| `ia-unregistered-claim` | 文案。加えて、成果物 × 画面幅ごとのアンカーの外の文言（id の末尾が `#unanchored`） | 固定値を展開した文案。アンカーの外は `dom` の `unanchored` | メッセージとその提供価値・経路から受け継ぐ証拠・使った固定値。アンカーの外は、すべて空（支える根拠が無い） |
| `ia-value-fit` | 文案（承認済みの提供価値を指すメッセージのもの） | 固定値を展開した文案 | その文案が伝えるメッセージが指す提供価値 1 件（id と文面）。文案か提供価値の中身が変われば判定のキーが変わる。正解表で精度を測るまでは警告だけで、θ 以上も、台帳に無いことも、出荷を止めない（v6 要件 11）。`ship` の `positioning.semantic_fit` は judged / needs-review / not-judged。v5 の `ia-strategy-fit`（経路を丸ごと渡す）は生成しない（既存の台帳の行は読める） |
| `ia-audience-fit` | outline | 全部の文案と役割（セクションと文案の順に） | target・接点・トーン・目的変数・一番のメッセージ。並びや一番を替えると判定のキーが変わる |
| `ia-emphasis` | 成果物 × 画面幅 | `dom` が切り出した最初の画面の文言 | target と、一番のメッセージ（とその提供価値） |

### 正解表（`infoarch-author answer-key`。v6 要件 12）

- `answer-key --root . --out DIR`：承認済みの提供価値に立つ文案ごとに、自分の提供価値との組（own）と、同じ positioning のほかの承認済みの提供価値に付け替えた組（swap。提供価値が 1 つしか無い positioning では、別の positioning の提供価値と組む）を作る。`answer-key.json`（組と判定のハッシュ、正解表の id）・`candidates.json`（`judgment_run.py` に渡す `ia-value-fit` の候補）・`answer-key-page.html`（オーナーが ○ 伝わる／× 伝わらない／？ 分からないを付けて labels.json を保存するページ。外部の読み込み無し）を書く。既にあるファイルは上書きしない。機械が作るのは外れの候補まで。正解はオーナーのラベルで決まる。
- `answer-key --root . --measure labels.json --key DIR/answer-key.json`：ラベルと共有の判定台帳の答え（θ 以上＝伝わらない）を比べ、一致の割合・見逃し（オーナーが × で Jev が通した組）・誤検出（オーナーが ○ で Jev が止めた組）を出す。？ は数えない。台帳に無い組は `uncovered` に並べて数えない。数えた組に ○ か × のどちらかが 0 件なら、数字を出さず `status: unmeasurable`。ラベルが別の正解表のもの（正解表の id が違う）・知らない組を含む・outline や提供価値が作った後で変わった、は失格（作り直してラベルを付け直す）。
- モデルは呼ばない。本物の Jev は `judgment_run.py`（上限つき、`--execute` のときだけ）で取る。

score は「問題がある確率」。θ 以上は IA_JUDGE_REVIEW（要確認）として人に回し、score だけで失格にしない。
オーナーが原文を見て認めた要確認は、`accept-reviews` でオーナーの言葉と日付を付けて `design/info_arch/accepted_reviews.yaml`（`schema_version` 1、行は `judgment_hash`・`id`・`approver`・`date`）に記録する。記録は判定のハッシュに結び付くので、文案か文脈が変われば当たらなくなり、要確認に戻る。今の要確認に無い id・言葉や ISO の日付が無いものは記録しない。`--write` が無ければ書かない。共有の判定台帳の形は変えない。
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
| IA_JUDGE_ACCEPTED（info） | θ 以上だが、オーナーが原文を見て認めた（`design/info_arch/accepted_reviews.yaml` に判定のハッシュで記録がある） |
| IA_JUDGE_ACCEPTED_STALE（info） | 「認めた」記録が今の判定に当たらない（文案・文脈が変わったか、θ 未満になった）。要確認は記録が当たらない限り戻る |
| IA_JUDGE_ACCEPTED_FILE | 認めた要確認の記録が読めない・形が違う・symlink。要確認は隠さずに全部出す |
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
