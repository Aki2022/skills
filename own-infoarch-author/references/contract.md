# ux/infoarch/ の契約

方針は、利用側の repo の spec `information-architecture-standardization`（v6）と `deliverable-production-layers`（layer の切り方）に従う。

## ディレクトリ

`ux/infoarch/` の中は次のとおり。positioning の SSOT（`ux/positioning/positioning.yaml`）はここに無い。positioning と infoarch は別のフォルダで、positioning は own-positioning-author が持つ（スキーマ版 5。原子と矢印、id 台帳。固定値・正式名称は持たない（辞書）。承認の状態は持たない（承認台帳 `ux/approvals.yaml` が持つ）。記録は git の履歴が持つ）。

- `channels/<media>.yaml`：成果物の置き場所と出し方の宣言だけ。中身は `media`（ファイル名と一致）・`outputs`（repo 相対の glob）・`ship`・`ship_command`（`ship: true` なら必須。出荷の正式の入口が実行する）・`merge_base`（pre-push の予測マージの相手。出荷チャネルで無いと doctor が失格）・`viewports`。
- `outlines/<成果物 id>.yaml`：成果物ごとの outline（版 3）。1 ページ＝1 outline。AI が書き、人は保守しない。`id` はファイル名と一致する。
- `INFO_ARCH.md`：散文。原子を id で参照するだけで、値を持たない。`check` は引く id を確かめる：positioning の原子・プロダクトの項目（`val-0001`・`sto-0001` など）、画面（`scr-<slug>`）、outline（`outline:<id>` と書く）が無ければ `IA_INFO_ARCH_REF_BROKEN`（失格。行番号つき）、取り下げた id（墓標）なら `IA_INFO_ARCH_REF_TOMBSTONED`（警告）。ファイルが無い repo は何も確かめない。v5 の id（`msg-…`・`aud-…`）は読まない（`migrate-v6` の `prose_mentions` が持つ）。

それ以外のファイルやサブディレクトリが `channels/` `outlines/` にあると失格にする（黙って無視しない）。`.` で始まるファイル（`.DS_Store`・`.gitkeep`）だけは読まずに飛ばす。`.` で始まるサブディレクトリは失格にする。

## positioning と辞書（この skill が読むもの）

- positioning（原子・矢印・証拠・id 台帳・承認）は own-positioning-author の持ち物：`ux/positioning/positioning.yaml`（スキーマ版 5）。契約は `own-positioning-author/references/contract.md`。outline は、その id（提供価値・代替）を引く。引けなければ失格。
- 固定値と正式名称は `docs/dictionary/<分野>.json` の辞書が持つ（own-doc-update の `dictionary.py` が形を検査し、この skill が同じファイルを読む）。辞書の `value` が固定値、`name` が正式名称。outline の文案は固定値を `{{value-<辞書の id>}}` で参照する。辞書の形が壊れていれば IA_DICTIONARY_INVALID で失格。
- 固定値の期限は辞書の `expires`（日付）か `evergreen: true`。`as_of` から日数を足す暗黙の期限は辞書には無い。
- 声は廃止した。成果物ごとのトーンは outline の `tone` が持つ。

### outline が指す経路（`check.referenced_values`・`check.positioning_path`）

outline のメッセージが指す提供価値のうち、target のセグメントのものが、その outline の指す経路。出荷チャネルの outline が指す経路が 1 つでも承認されていない（ハッシュが合わない）と出荷で止まる（IA_POSITIONING_UNAPPROVED、`detail.values`。下書きは警告）。比べる相手（`compares`）は、その提供価値の経路の比べる集合の中に無ければ失格（IA_OUTLINE_COMPARE_OUTSIDE_SET）。比べる代替は経路と一緒に承認される。指していない経路が未承認でも止めない（IA_POSITIONING_PATH_PENDING の警告）。`check.positioning_path` は提供価値の経路を合わせたキーとハッシュを返し、意味の判定（未登録の主張）の証拠の入力にも使う。一番のメッセージの経路に証拠が 1 つも無ければ警告（IA_LEAD_UNSUPPORTED）。

成果物が立つもの（期限・逆引き・maintain の影響）は、指す経路の原子と矢印すべて：隙間と同等であるべき点は経路の比べる集合の代替と度合いの矢印から決まるので、その証拠の期限切れや変更も成果物に届く。

### 承認の対象（`ia:path:<提供価値 id>`）

- 承認依頼の表示と、束の読み取り（`review`）、承認の台帳の移行（`migrate-approvals`）は own-positioning-author（`positioning-author request` / `review` / `migrate-approvals`）。
- `infoarch-author targets --root .`：proposed の全部の提供価値の経路のハッシュ（`approval_hashes`）と、出荷チャネルの outline が指す経路のキー（`must_approval_keys`）を出す。`ia_module.py`（契約 `ia-git-history-v4`、`modules.ia: 0.4.0`、governance の schema_version 4）がこれを own-design-author に返す。キーの形は `ia:path:val-NNNN` だけ。
- 出荷の報告（`ship`）は `positioning` に、`owner_adopted`（approved / unapproved）・`semantic_fit`（意味の判定の状態。今は not-judged）・`customer_validation`（IA は判定しない：not-judged-by-ia）を分けて出す。

## outline（版 3。SPEC v6 要件 6〜9・13）

- 置き場所：`channel`・`outputs`（チャネルの outputs の glob に一致すること）。`id`・`channel`・`outputs` は outline の id に入らない。
- 定数：`target`（`segment`＋`condition`。outline はこのセグメントの positioning にだけ立つ）・`touchpoint`（接点）・`tone`。目的変数：`objective`（`id`・`action`）。
- `messages`（`msg-NNNN`）：`value`（提供価値ちょうど 1 つ）・`rank`（lead / support）・`type`（normal / cta）・`compares`（比べる代替）・`achieves`（CTA が指す目的変数）。`no_cta: true` で CTA を置かないことを宣言できる。
- `sections`（`sec-NNNN`）：`order`・`role`（aware / pain / compare / trust / act）・`places`（通常のメッセージ 1 つと CTA 0〜1 つ）。
- `copies`（`cpy-NNNN`）：`section`・`order`（セクションの中の順番）・`message`（同じセクションに置かれたもの）・`text`。固定値は `{{value-<辞書の id>}}` で参照する（括弧の中に空白や改行を入れない。入れたものは形の崩れた参照として失格）。
- `texts`（任意）：主張でない文言の宣言。key ごとに `kind`（navigation / label）・`text`・（label なら）`section`。
- 並びは `order` で持ち、ファイルの中で書いた順番は意味を持たない。outline の原子の id は outline の中で一意で、台帳は持たない（別案は同じ id を引き継ぐ）。

### id と set（`infoarch-author set --outline <id>`）

- outline の id：`out-` ＋ outline の中身（定数・目的変数・`no_cta`・メッセージ・セクションの順序と置き方・文案の順番と文言・`texts`）を正規化した JSON の SHA-256。立つ positioning の id は入れない。メッセージは提供価値を原子の id で指すので、提供価値の文面だけを直しても outline の id は変わらない。
- set の id：`<outline の id>+<経路の id>+…`（outline が指す経路の id を提供価値の id の順に並べる）。中身から決まり、承認の有無では変わらない（承認は出荷ゲートが見る）。指していない経路の手直しでは変わらず、指している経路の中（共有する原子を含む）を直すと変わる。指す経路が無ければ None。`set` は `paths`（提供価値 → 経路の id）、報告用の positioning の id、要素（messages・sections・copies・texts）ごとの中身のハッシュも出す。v5 の対応表（mapping）はこれに置き換えた。

### 差分（`infoarch-author diff --a <id> --b <id> [--root-b DIR] [--dom-a A --dom-b B]`）

2 つの set の差分を、変わった要素（`messages/<id>` などの outline の要素と、経路の原子・矢印）と属性で列挙する。指す経路は提供価値の id で突き合わせ、両方にあってハッシュが違えば positioning の編集（`layer: positioning`、`path` に提供価値の id）、片方にしか無ければ outline の選択（`layer: outline`、`element: paths/<提供価値 id>`、`kind: selection`）として別々に出す。両方の文案が参照する固定値の値が違えば、IA_DIFF_UNASSIGNED で失格（固定値はどちらの id にも入らない。値は報告に出さない）。`changed_layers` で、どちらの層が変わったかが分かる（両方が変わる別案も認める）。定数か目的変数が違えば `same_experiment: false`（別の試験として扱う。失格ではない）。両方の dom の報告を渡すと、印の外の文を画面幅ごとに比べ、違えば IA_DIFF_UNASSIGNED で失格（どの要素にも割り当てられない差分）。渡さなければ `not_checked` に比べていないことを出す。A/B の実行・指標・統計は持たない。

## v5 からの移行（`infoarch-author migrate-v6 --root . [--write]`。SPEC v6 要件 17）

v5 の木（旧い名前 `info_arch.yaml` の版 3（今の置き場所は `ux/positioning/positioning.yaml`）・outline 版 2・描画物の v5 の印・`modules.ia` 0.3.0）を、一括で v6 に移す。`--write` が無ければ何も書かない（見積もりだけ）。移しそこねが 1 件でもあれば終了コード 2 で、`--write` でも何も書かない（一括。半分だけ移さない）。commit はしない。v6 の木（版 4）では何もせず終了コード 0（再実行できる）。版 3 でも 4 でもなければ失格。

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
| 描画物の `data-ia-role`＋`data-ia-msg`・`data-ia-pain`・`data-ia-frame` | セクションの要素に `data-ia-section`（と v6 の役割）。文案は、中身が要素 1 つならその要素に `data-ia-copy`、そうでなければ中身を包む（中身に段落などのブロックがあれば `<div data-ia-copy>`、無ければ `<span data-ia-copy>`。dom は文案を描かれた位置で並べるので、箱を持たない `display: contents` や style の書き足しは使わない）。印は outline のセクションと順に、数と id で突き合わせる |
| `ux/design/design.yaml` の `modules.ia: 0.3.0` | v6 の契約の版（`ia_module.VERSION`） |

判断が要るもの（提供価値の `kind`、所属カテゴリーと `member_of`・`recalls`・`includes`、代替 × 痛み・得たいものの `fulfills` の度合い、CTA だけのセクションに置く通常のメッセージ、役割・touchpoint・target の条件の確かめ）は `drafting_required` に並べ、作らない。移した直後の木は、これを済ませるまで `check` が失格する。

出力：`id_map`（旧 id → 新 id。outline の要素は `<outline>:<id>`、v5 のセクションは `<outline>:<ページ>[<番号>]`、枠は `<outline>#<枠>`。`[]` は無くしたもの。載っていない id は同じ id のまま）・`unmigrated` と `unmigrated_count`・`dropped`・`drafting_required`・`role_map`・`stops_applying`・`module_pin`・`files`（write・create・delete）・`rollback`（戻す手順）。

`stops_applying` は、外れる承認と判定の一覧。`ux/approvals.yaml` の `ia:` の行のうち `ia:path:` でないもの（v5 の `ia:strategy:` は経路の承認に移らない。オーナーが経路を新たに承認する。`ia:positioning:` の行は `migrate-approvals` で移せるか確かめる）と、`ux/judgments.json` の IA の判定（`ia-` の種類。v6 の鍵は v6 の outline の id と文案・提供価値の中身から作るので読まれない）。どちらの台帳も読むだけで、書かない。

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

### 報告の 5 区分（SPEC 情報設計の Quality 2）

`infoarch-author check`・`infoarch-author blueprint`・`positioning-author check` の JSON は、次の 5 つの鍵を空でも必ず出す（`ia_common.CATEGORIES`。実装は `ia_common.Report` の `categorized`。`check`・`blueprint` は `check.categorize` が埋める）。

| 区分 | 中身 | check・blueprint・positioning check では |
| --- | --- | --- |
| `executed` | 走らせた検査 `{id, targets}` | `ia-positioning`（提供価値の件数）・`ia-outlines`（channels・outlines・sections）・`ia-blueprint`（screens・items）。対象があるものだけ |
| `skipped` | このコマンドが走らせない検査 `{id, reason}`。別のコマンドが持つので走らせない | `ia-dom`（`dom`／`ship`）・`ia-semantic-judgments`（`judge`／`evaluate`）。outline の木で blueprint が無ければ `ia-blueprint`、プロダクト UI で outline が無ければ `ia-outlines`、blueprint と positioning check は outline 側 |
| `missing` | 要るのに無い入力 `{id, reason, code}`。失格の code から作る（`IA_MISSING`→`ia-positioning`、`IA_NO_VALUES`→`ia-values`、`IA_NO_CHANNELS`→`ia-channels`、`IA_NO_OUTLINES`→`ia-outlines`、`IA_BLUEPRINT_NO_SCREENS`→`ia-blueprint-screens`、`IA_BLUEPRINT_PRODUCT_MISSING`→`ia-blueprint-product`） | 対象 0 件を合格にしない。失格と同じ事実を、入力の名前で言い直す |
| `uncovered` | あるのに見ていない対象（台帳に無い判定・HTML 以外の成果物など） | 決定論の検査は見ないものを持たないので常に `[]`（`ia_module` はこれを要求する） |
| `screened` | 台帳にあって読んだ判定（「問題なし」の保証ではない） | 台帳を読まないので常に `[]` |

`judge`・`evaluate`・`ship` は、この 5 区分の鍵を足さず、これまでの形のまま（`uncovered`・`screened`・`missing` を台帳の言葉で持つ。`ship` の段は `_stage` が決めた鍵だけ）。`ia_module.resolve`（design の L0）が上流の `check` を読むときの鍵の集合には、この 4 つの鍵（executed・skipped・missing・screened）が入っている。

## 検査（決定論・トークン 0）

| code | 内容 |
| --- | --- |
| IA_MISSING / IA_YAML / IA_SCHEMA | ファイルが無い・YAML として読めない（重複キーを含む。日付は文字列として読む）・スキーマ違反（`state: approved` を含む） |
| IA_ID_NOT_ISSUED / IA_ID_TOMBSTONED / IA_TOMBSTONE_TARGET | 台帳と墓標 |
| IA_REF_BROKEN / IA_EMPTY_REFERENCED | SSOT 内の参照切れ・proposed の原子が空欄を参照（参照の表は `check.ssot_refs` の 1 か所） |
| IA_SSOT_LEGACY_V4（警告） | スキーマ版 4 の SSOT（固定値・正式名称・声を持つ）。読むが、辞書へ移す（`migrate-dictionary`） |
| IA_DICTIONARY_INVALID | `docs/dictionary/` が辞書の形でない（固定値と正式名称を読めない。evergreen と日付の併存など） |
| IA_POSITIONING_*（要件 5・要件 2） | 上の positioning の表 |
| IA_POSITIONING_UNAPPROVED | outline が指す経路に、オーナーの承認が無いか外れたものがある（`detail.values`。出荷チャネルは失格、ほかは警告） |
| IA_APPROVALS_LEGACY | 承認台帳に原子と矢印ごとの旧い行（`ia:positioning:`）がある。数えない（警告。`positioning-author migrate-approvals` で移す） |
| IA_APPROVALS_INVALID | `ux/approvals.yaml` が読めない・symlink・design の `verify_rows` を通らない（重複キー・項目の過不足・承認者・日付・ハッシュ）。何も承認済みとして数えない |
| IA_PAGES_OUTPUT_NOT_INVENTORIED（警告） | 出荷チャネルの HTML の成果物が、ux/design/validation.json の pages（covered・excluded）に無い（design の検査が見ていない） |
| IA_PAGES_OUTPUT_EXCLUDED_BY_DESIGN（info） | design が対象外にした成果物。IA の検査は続ける（ページの宣言は一本化しない。ADR-20261115-design-ia-two-skills-shared-base） |
| IA_NO_CHANNELS / IA_NO_OUTLINES / IA_CHANNEL_NAME / IA_CHANNEL_FILE / IA_OUTLINE_FILE / IA_OUTLINE_ID | 対象 0 件・名前の不一致・読まれないファイル |
| IA_OUTLINE_CHANNEL / IA_OUTLINE_OUTPUT_UNDECLARED | 宣言の無いチャネル・チャネルの outputs に一致しない成果物 |
| IA_OUTLINE_REF_BROKEN / IA_OUTLINE_REF_EMPTY | アウトラインの参照切れ（墓標なら先を案内）・空欄の原子を参照（参照の表は `check.outline_refs` の 1 か所） |
| IA_OUTLINE_BAD_PLACEHOLDER / IA_OUTLINE_FACT_LITERAL | `{{value-<辞書の id>}}` の形でない `{{…}}`（改行をまたぐものを含む）・固定値の値を参照せず直接書いた |
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

条件と赤のフィクスチャの対応表：`fixtures/red-cases.json`（参照と形）・`own-positioning-author/fixtures/positioning-red-cases.json`（要件 5。positioning の持ち物で、outline を持つこの検査と、outline 無しの `positioning-author check --strict` の両方で走らせる）・`fixtures/outline-red-cases.json`（要件 7・9）。

## 期限

証拠は、`expires` があればその日、無ければ `as_of` に日数を足した日を期限とする。期限の日の当日までは有効で、翌日から期限切れ。日数の初期値は `references/expiry.json`、プロジェクトの
`ux/positioning/positioning.yaml` の `policy.expiry_days` で上書きできる。固定値は辞書の `expires`（日付）か `evergreen: true`（期限なし）で、日数を足す暗黙の期限は無い（スキーマ版 4 の SSOT の固定値だけは、`as_of` から日数の読み方が残る）。出荷の成果物が期限切れの値を参照すると IA_EXPIRED_SHIPPED で失格。

## 工場への受け渡し（`skeleton`）

`infoarch-author skeleton --outline <id>` は、check が緑のアウトラインから骨組みの HTML を出す（失格していれば出さない）。
セクションを `order` の順に `<section data-ia-section="sec-…" data-ia-role="…">` で、その中の文案を `order` の順に `data-ia-copy="cpy-…"` で出す。
ナビの宣言は `<nav data-ia-text>`、ラベルはそのセクションの先頭に `data-ia-text` で出す。文案の `{{value-<辞書の id>}}` は辞書の値の
`<span data-dict="<辞書の id>">` に展開し、文案は HTML エスケープする。1 ページ＝1 outline なので、どの output にも outline 全体を出す。`--out` を付けるとファイルに書き、既にあるファイルは上書きしない。
工場（own-web-build）は、文言・順序・`data-ia-*` を変えずにスタイリングし、描画後に `dom` を走らせる。

## 描画物の確認（`dom`）

`infoarch-author dom --root . --outline <id>` は、アウトラインの outputs のうち HTML を、チャネルの viewports（無ければ 375×812 と 1280×800）でブラウザに描画し、DOM を確かめる。

- 解析と描画はブラウザ（Playwright の Chromium）が行う。成果物のソースを自前のパーサで読まない。
- 工場はアンカーを埋める：セクションは `data-ia-section`、文案は `data-ia-copy`、固定値と正式名称の表示は `data-dict="<辞書の id>"`（`value-` の id は固定値、`name-` の id は正式名称。`term-` の id は印にできない）、宣言した文言は `data-ia-text`。v4・v5 の `data-ia-problem`・`data-ia-msg`・`data-ia-pain`・`data-ia-frame` が残っていれば IA_DOM_LEGACY_ATTRIBUTE で失格（互換は持たない。v5 の描画物は `infoarch-author migrate-v6` が読み替える）。
- **画面の適合（SPEC v6 要件 10）**：セクションの印は outline の順序どおりにちょうど並ぶ（IA_DOM_ORDER）。文案の印は、outline で属するセクションの印の中にあり（IA_DOM_COPY_SECTION）、セクションの中で outline の順番どおりに並ぶ（IA_DOM_COPY_ORDER）。同じ文案の印は 1 つだけ（IA_DOM_COPY_REPEATED。繰り返したい行動喚起は、outline に別の文案として置く）。条件と赤のフィクスチャの対応表は `fixtures/screen-red-cases.json`（DOM の写しと実ブラウザの両方で確かめる）。
- Playwright は own-design-author の `scripts/playwright_runtime.mjs` で解決する：`OWN_PLAYWRIGHT_DIR` → 対象 repo（`INFOARCH_PLAYWRIGHT_DIR`、無ければ実行ディレクトリとその祖先）→ 端末の共通置き場（対象 repo の `ux/design/render-environment.json` が固定する版）。`NODE_PATH` は見ない。自動では入れない。解決できなければ IA_DOM_PLAYWRIGHT で失格。
- 読み取る前に、ページが落ち着くのを待つ：network idle の後に 1 秒（表示を始めるタイマーのため）、さらに実行中のアニメーションとトランジションが終わるまで最大 5 秒。終わらないアニメーションがあっても読み取り、2 回の描画の比較で揺れを報告する。
- 同じ入力を 2 回描画し、結果が違えば IA_DOM_UNSTABLE（判定に使わない）。
- 最初の画面の文言（各画面幅）を報告の `first_screens` に出す（描画の事実の記録。どの判定の入力にもしない）。
- 各文案のアンカーの表示の文言（描かれている文字だけ。`unanchored` と同じ読み方で、opacity: 0 などの隠れた文字は入らない）を、アウトラインの文案（固定値を SSOT の値に展開）と比べる。NFKC で幅をそろえ、空白と改行をすべて除いてから比べる（改行は見た目で、文言ではない）。見えている印はすべて一字一句一致すること（IA_DOM_COPY_DRIFT）。報告に表示の文言は出さない。
- **主張でない文言の宣言**：ナビ・節のラベル・見出し・法務やフッターのように主張でない文言は、outline の `texts` に安定した key と文言そのもの（`text`）を書き、工場は `data-ia-text="<key>"` を付ける。主張になりうる文言（副の行動喚起など）は宣言ではなくメッセージにする。`dom` は、その key の見えている印の**すべて**が文言と一致することを確かめる（繰り返しの緩めは無い）。outline に無い key は IA_DOM_ANCHOR_UNKNOWN、key の印が無ければ IA_DOM_ANCHOR_MISSING。`navigation` は画面幅によって畳んでよい（見えていなくても IA_DOM_ANCHOR_HIDDEN にしない。印が無ければ失格、見えている所は文言と一致）。
- **正式名称**：`data-ia-entity="ent-NNNN"` の見えている印は、すべて全文が SSOT の `name` と一致する（NFKC・空白を除いて比べる）。名前以外の文言を包めない。
- 文案・宣言した文言・正式名称の印の中に、ほかの IA の印（セクション・文案・宣言した文言。正式名称は固定値と正式名称も）を入れると IA_DOM_ANCHOR_NESTED。文案と宣言した文言は固定値と正式名称を中に持てる。
- `dom` の 0 は「描画物が outline の宣言どおりに表示している」までで、主張の妥当性は確かめない。文案の意味は判定台帳の `ia-value-fit`（文案 × 提供価値）が見る。ほかの種類の意味の判定は作らない。
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
方針ファイル（モデル・θ・プロンプト）・台帳（`ux/judgments.json`）・使用量の台帳（累計 2,000 判定）は、own-design-author と共有する
（own-design-author の `references/judgments-and-modules.md`）。

意味の判定は `ia-value-fit`（文案 × その文案が指す提供価値）の **1 種類だけ**（情報設計の要件 11・ADR-20261123）。出荷を止める判定は作らない。表記ゆれは辞書の検査（`check`）が持つ。ほかの種類（未登録の主張・target との相性・最初の画面の強調）が要ると分かったら、精度を測ってから spec に足す。

| 種類 | 単位 | 入力の本文 | 文脈 |
| --- | --- | --- | --- |
| `ia-value-fit` | 文案（承認済みの提供価値を指すメッセージのもの。役割が trust（事実・証拠）と act（行動喚起）のセクションの文案は、提供価値を伝える役ではないので聞かない。正解表にも入れない。id は `<outline>/<sec>/<cpy>`） | 固定値を展開した文案 | その文案が伝えるメッセージが指す提供価値 1 件（id と文面）。文案か提供価値の中身が変われば判定のキーが変わる。正解表で精度を測るまでは警告だけで、θ 以上も、台帳に無いことも、出荷を止めない（要件 11）。`ship` の `positioning.semantic_fit` は judged / needs-review / not-judged。v5 の `ia-strategy-fit`（経路を丸ごと渡す）は生成しない |

v5 までの `ia-name-variant`・`ia-unregistered-claim`・`ia-audience-fit`・`ia-emphasis`・`ia-strategy-fit` の行が repo の台帳に残っていても、`judge`・`evaluate`・`ship` は**読まずに無視する**（候補にも報告にも出さず、台帳は書き換えない。台帳の検証は種類の名前を今も受け付けるので、残っていても壊れない）。

### 正解表（`infoarch-author answer-key`。v6 要件 12）

- `answer-key --root . --out DIR`：承認済みの提供価値に立つ文案ごとに、自分の提供価値との組（own）と、同じ positioning のほかの承認済みの提供価値に付け替えた組（swap。提供価値が 1 つしか無い positioning では、別の positioning の提供価値と組む）を作る。`answer-key.json`（組と判定のハッシュ、正解表の id）・`candidates.json`（swap の組だけの `ia-value-fit` の候補。own の組は出荷の判定そのものなので入れない）・`answer-key-page.html`（オーナーが ○ 伝わる／× 伝わらない／？ 分からないを付けて labels.json を保存するページ。外部の読み込み無し）を書く。既にあるファイルは上書きしない。機械が作るのは外れの候補まで。正解はオーナーのラベルで決まる。
- **Jev の答えの置き場所**（ADR-20261123）：own の組は、`judge` が聞くのと同じ候補（id が `<outline>/<sec>/<cpy>`。`@提供価値` が付かない）なので、共有の判定台帳（`ux/judgments.json`）の答えをそのまま使う（測ったものと出荷で使うものが同じ）。swap の組（id が `<outline>/<sec>/<cpy>@<提供価値>`）は測るためだけにあるので、答えを**測定用の台帳**（`DIR/measurement-judgments.json`。共有の台帳と同じ形。`answer-key.json` と同じ場所）に記録する。出荷（`ship`・`judge`・`evaluate`）は測定用の台帳を読まず、`answer-key --record` は共有の台帳を書かない。
- `answer-key --root . --record --key DIR/answer-key.json --limit N [--execute]`：swap の組のうち測定用の台帳に答えが無いものを、上限つきで Jev に聞いて記録する。`--execute` が無ければ見積もりだけ（件数と要求数。キーを読まず、通信せず、何も書かない）。`--execute` があれば own-design-author の `judgment_run`（共有の使用量台帳。累計の上限を超えるなら 1 件も送らない）で取り、回答を `DIR/judgment-answers.json` に足して（次の実行で続きから取る）、測定用の台帳を更新する。own の組は聞かない（出荷の判定と同じ問いなので、`evaluate --acquire` で共有の台帳へ）。
- `answer-key --root . --measure labels.json --key DIR/answer-key.json`：ラベルと答え（own は共有の台帳、swap は測定用の台帳。θ 以上＝伝わらない）を比べ、一致の割合・見逃し（オーナーが × で Jev が通した組）・誤検出（オーナーが ○ で Jev が止めた組）を出す。？ は数えない。答えが無い組（置き場所が違う答えも含む）は `uncovered` に並べて数えない。数えた組に own か swap のどちらかが 0 件、または ○ か × のどちらかが 0 件なら、数字を出さず `status: unmeasurable`（`reason` に足りない側）。ラベルが別の正解表のもの（正解表の id が違う）・知らない組を含む・outline や提供価値が作った後で変わった・測定用の台帳が読めない（symlink・今の policy と合わない）は失格（作り直す・記録し直す）。
- モデルを呼ぶのは `--record --execute` だけ。

score は「問題がある確率」。θ 以上は IA_JUDGE_REVIEW（要確認）として人に回し、score だけで失格にしない。
オーナーが原文を見て認めた要確認は、`accept-reviews` でオーナーの言葉と日付を付けて `ux/infoarch/accepted_reviews.yaml`（`schema_version` 1、行は `judgment_hash`・`id`・`approver`・`date`）に記録する。記録は判定のハッシュに結び付くので、文案か文脈が変われば当たらなくなり、要確認に戻る。今の要確認に無い id・言葉や ISO の日付が無いものは記録しない。`--write` が無ければ書かない。共有の判定台帳の形は変えない。
文案 × 提供価値の判定でゲートを止めるかは、正解表で精度を測ってから決める（spec）。

```sh
infoarch-author judge --root .                                              # 台帳の読み取り（本文は出さない）
infoarch-author judge --root . --include-candidates > candidates.json      # Jev に渡す候補（文案を含む）
python3 <own-design-author>/scripts/judgment_run.py --candidates candidates.json --answers <dir>/judgment-answers.json --limit 50 [--execute]
infoarch-author judge --root . --answers <dir>/judgment-answers.json --replacement-out replacement.json
python3 <own-design-author>/scripts/judgment_ledger.py record|refresh --repo . --replacement replacement.json [--write]
```

`--replacement-out` は共有の台帳の置き換えを作る。design の行は全部残し、IA は現在の候補に対する実回答だけを入れる。作り直された（古くなった）IA の行と、無くなった outline の行は落とす。v5 までの古い種類の行は、outline が残る間はそのまま残す（repo のデータは直さない）。回答ファイルは `judgment_run.py` と同じ読み手で検査する（形・policy・候補のハッシュ・score と raw_answer の一致）。出力先は新規作成だけで、既存のファイル・symlink（親ディレクトリを含む）には書かない。
holdout はそのまま引き継ぐ。回答の policy が今の共有 policy と違えば止める（migrate が要る）。既にあるファイルは上書きしない。置き換えは schema 3（判定の鍵が種類ごと）で作る。共有の台帳が schema 2 以前なら止める（先に `judgment_ledger.py rekey` で移す）。

| code | 内容 |
| --- | --- |
| IA_JUDGE_REFUSED | check が失格している（先に直す） |
| IA_JUDGE_NO_TARGETS（warn） | 判定の候補が 0 件（trust・act の文案だけなど。判定していないという印で、止めない） |
| IA_JUDGE_LEDGER | 台帳が読めない・policy と合わない・symlink |
| IA_JUDGE_UNCOVERED（warn） | 台帳に無い判定がある（未カバー） |
| IA_JUDGE_REVIEW（warn） | score が θ 以上（要確認） |
| IA_JUDGE_ACCEPTED（info） | θ 以上だが、オーナーが原文を見て認めた（`ux/infoarch/accepted_reviews.yaml` に判定のハッシュで記録がある） |
| IA_JUDGE_ACCEPTED_STALE（info） | 「認めた」記録が今の判定に当たらない（文案・文脈が変わったか、θ 未満になった）。要確認は記録が当たらない限り戻る |
| IA_JUDGE_ACCEPTED_FILE | 認めた要確認の記録が読めない・形が違う・symlink。要確認は隠さずに全部出す |

報告は `uncovered`（未カバー）・`screened`（台帳にあり。「問題なし」の保証ではない）を分けて出す（`missing` は今は常に空）。

## 作った物の評価（`evaluate`）

`infoarch-author evaluate --root . [--outline <id> ...]` は、作った物の評価の入口。工場は描画した後に必ず走らせる。

1. `check`。失格なら描画しない（`dom` と `judge` は `skipped`）。
2. 指定したアウトライン（無ければ全部）ごとに `dom`。1 つでも失格なら判定しない（`judge` は `skipped`）。
3. `judge`（`dom` の報告は渡さない。判定は文案 × 提供価値だけ）。

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
1. `tree`：`ux/infoarch/` と `ux/positioning/` の下と `ux/judgments.json` にシンボリックリンクがあれば `IA_SHIP_SYMLINK`（失格。`check` の前に見る）。
2. `check`（決定論）。出荷チャネル（`ship: true`）が 0 件、または出荷チャネルのどれか 1 つでもアウトラインが 0 件なら失格（空のゲートを合格にしない。他のチャネルが出荷していても、空のチャネルは合格にしない）。
3. `tree`（続き）：出荷されるアウトラインの成果物は、HTML かどうかに関わらず、すべてコミットの中に通常ファイルとして無ければならない。無い・ディレクトリ等は `IA_SHIP_OUTPUT_MISSING`、成果物か親ディレクトリがシンボリックリンクなら `IA_SHIP_SYMLINK`（どちらも失格）。git のオブジェクトから判定するので、手元にある無視されたファイルは数えない。
4. `dom`：出荷されるアウトラインのうち HTML の成果物を持つものを、実ブラウザで描画して確かめる。HTML が無いアウトラインは `skipped` に記録する。
5. `judge`：共有の判定台帳を読む。**どの判定も出荷を止めない**。台帳に無い判定は `IA_JUDGE_UNCOVERED`（警告）、θ 以上は `IA_JUDGE_REVIEW`（警告。人に回す）。台帳が読めない・symlink は `IA_JUDGE_LEDGER`（失格。判定ではなく入力の破損）。

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

`infoarch-author maintain --root .` は own-session-close の Step 2.5 から呼ぶ。`ux/infoarch/` も `ux/positioning/` も無ければ 1 行で skip する。
あれば（`ux/infoarch/` が無くても `ux/positioning/` があれば走る）`check` と `judge`（台帳を読むだけ。ブラウザは使わない）を走らせ、次を 1 行の JSON で返す。
- `due`：期限が `references/expiry.json` の `due_within_days` 日以内に来る証拠・固定値と、それを参照するアウトライン（出荷されるか）。
- `changed_paths`：出荷チャネルの `merge_base` からのコミット済みの差分と、未コミットの差分（`ux/infoarch/` と `ux/positioning/` の中だけ。辞書の変更も含む）。
- `impacted_outlines`：変わった原子を参照するアウトライン（辞書（`docs/dictionary/`）が変わったら全部）。
- `uncovered_judgments`：台帳に無い判定の数（Jev で取得する候補）。

終了コードは 0（整合）・1（期限が近い・未カバー・警告）・2（失格）。台帳もアウトラインも書き換えない。

## スキーマ版 4 からの移行（`infoarch-author migrate-dictionary --root . --map <map.json> [--write]`）

固定値（`fct-NNNN`）を辞書の `value`、正式名称（`ent-NNNN`）を `name` に移し、声（`voices`）を捨て、SSOT をスキーマ版 5 にする。outline の `{{fct-NNNN}}` は `{{<辞書の id>}}` に、描画物の `data-ia-fact`・`data-ia-entity` は `data-dict` に書き換える。SSOT のコメントは残す（ブロックを文面から除き、読み戻した結果がスキーマ版 4 を持ち上げた形と一致しなければ何も書かない）。

- `--map`：`{旧 id: {"id": "<kind>-<slug>", "expires": "YYYY-MM-DD" | "evergreen": true, "variants": [...]}}`。辞書の id は人が読む名前なので道具は付けない。値には `expires` か `evergreen` のどちらか 1 つを必ず決める（暗黙の期限は無い。どちらにするかはその repo の中身の判断で、既定値は置かない）。
- 移せないもの（対応の無い固定値・正式名称、期限の決まっていない値、辞書の id でない名前、SSOT に無い固定値を指す文案・印、既にある辞書のファイル）は `unmigrated` に並べて exit 2、何も書かない。
- 既定は dry run。散らかさない scratch の複製で移した木の検査も走らせ、移す前の失格の code と並べて出す（`check`）。新しい失格が増えるなら移さない。`--write` で書く。スキーマ版 5 の木は何もしない（再実行できる）。承認台帳は触らない（承認のハッシュは positioning の経路にだけ掛かり、固定値・正式名称・声はどの経路にも載っていない）。

## blueprint（プロダクト UI の infoarch。SPEC-product-ui-infoarch）

`ux/infoarch/blueprint/product.yaml`（`references/schema/blueprint-product.v1.schema.json`）と `screens/<scr-…>.yaml`（`blueprint-screen.v1.schema.json`）。`blueprint/` が無い repo では何も読まない。画面だけの repo（outline もチャネルも無い）は `IA_NO_OUTLINES` で失格にしない。

- **項目と id**：ストーリー `sto`・原則 `pri`・制約 `con`・オブジェクト `ojc`・入口 `ntr`。id は `product.yaml` の台帳が発行し（`ledger.next`）、`ledger.tombstones` にある id は再利用しない（画面 id `scr-…` も載せられる）。承認の状態は持たない（state は empty / proposed）。
- **承認**：`ia:item:<項目 id>`。ハッシュは `blueprint.item_hash`（項目 id と、書かれた内容から provenance を除いたものの JCS 正規化の SHA-256）。ほかの項目に依らない。`infoarch-author targets` の `approval_hashes` に proposed の項目すべて、`must_approval_keys` に出荷する画面が引く項目が出る。`ia_module.py` の鍵の形は `ia:path:val-NNNN` と `ia:item:<種別>-NNNN`。
- **画面の領域**：`priority`（小さいほど上位、画面の中で重ならない）、`shape`（`value`・`list`・`comparison`・`trend`・`distribution`・`status`・`form`・`text`。見せるものの種類。どの部品で見せるかは design）。状態は `loading`・`empty`・`error`・`partial`・`offline`・`denied` から宣言する。主操作は `primary: true` の操作がちょうど 1 つ（持たない画面は `no_primary: true`）。
- **検査の code**：IA_BLUEPRINT_NO_STORIES / NO_SCREENS / NO_REGIONS / PRODUCT_MISSING / LEAD_COUNT / VALUE_NONE / VALUE_UNAPPROVED（警告）/ CONSTRAINT_FORM / ID_NOT_ISSUED / ID_REUSED / REF_BROKEN / REF_EMPTY / ITEM_UNAPPROVED / SCREEN_NO_STORY / SCREEN_ID / SCREEN_FILE / GROUP_NO_PRINCIPLE / PRIORITY_MISSING / PRIORITY_DUPLICATE / PRIMARY_MANY / PRIMARY_NONE / PRIMARY_CONTRADICTS / STATES_MISSING、IA_BLUEPRINT_CHECKED（info。画面と項目の件数）。形の違反（shape が決まった種類でない等）は IA_SCHEMA。
- **縦の順の確かめ（型を変えるときだけ・手作業）**：blueprint の型（項目の種類と持つ情報）を変えるときだけ、その変更の中で一度、実配置を見ない別の AI が、画面の優先度と design の割り付け規則だけから縦の順を導けるかを手で一度試し、結果を作業の記録に残す（食い違いの理由が型のどの項目で説明できるか。説明できない食い違いがあれば型を直す）。毎回の検査にはしない。この skill に道具は無い。描画した画面が優先度どおりに並んでいるかは、design の割り付けの検査が毎回確かめる。
- **ia_module の targets の 5 つの数**：outline 側の `channels`・`outlines`・`sections` と、blueprint 側の `screens`（画面）・`items`（product.yaml の項目すべて）。鍵は固定の 5 つで、`check.py`（数える）・`ia_module.py`（`TARGET_KEYS`・`zero_targets`）・`design_modules._git_contract`（受ける）・解決済みマニフェストの `modules.ia.resolved.targets` と `media.*.targets`（記録する）が同じ 5 つを持つ。
  - **対象 0 件の判定**：どちらの側も「全部が 1 以上」か「全部が 0」のどちらかで、半分だけ 0 の側（チャネルだけあって outline が無い、画面だけあって項目が無い）は 0 件扱い。両方の側が空でも 0 件で、`zero IA check targets` の失格になる。片方の側が丸ごと埋まっていれば足りる（outline の無いプロダクト UI は blueprint が、blueprint の無い成果物は outline が対象）。
  - **既存への影響**：outline だけの repo は `screens: 0, items: 0` が増えるだけで、判定は従来と同じ。マニフェストの `modules.ia.resolved.targets` を検証して読むコードは無く（読み込みで縛るのは `ia_module.py` のハッシュ）、`ia_module.py` が変わると既存のマニフェストは「IA resolver changed; rerun L0」で止まるので、`design-author check` を再実行すると 5 つの鍵で書き直される（`VERSION`・`CONTRACT` は据え置き）。3 つの鍵だけの結果を返す古い resolver は受けない。
  - 出荷ゲート（`infoarch-author ship`）は数に依らない。ship:true のチャネルが無ければ、blueprint だけの repo も `no ship:true IA channel` で止まる（空のゲートは合格にしない）。
