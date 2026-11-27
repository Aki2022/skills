---
name: own-infoarch-author
description: "成果物ごとの infoarch（台本：target・メッセージ・まとまり・文案）を、承認済みの positioning（own-positioning-author）と docs/dictionary/ の辞書の id を引いて読み書きし、Web サイト・LP・チラシ・ポスター・資料を作るときに、誰に・何を・どの順で・どう話すかをそろえる。描画後の確認・意味の判定・出荷ゲートも持つ。「情報設計を作って」「伝え方をそろえて」「SSOT から LP の構成を作って」「メッセージを整理して」「infoarch を検査して」で使う。positioning（セグメント・提供価値・承認）は own-positioning-author、見た目は own-design-author と工場 skill（own-web-build 等）。"
---

# own-infoarch-author — 台本（成果物ごとの infoarch）

infoarch は 3 つの layer の真ん中で、成果物ごとに決める「何を・どの順で」。この skill は、承認済みの positioning（own-positioning-author。`ux/positioning/positioning.yaml`）と辞書（`docs/dictionary/`）から、成果物ごとの
**台本**（script。1 ページ＝1 台本。v6 の outline を置き換えた）を作り（作る側）、台本の失格・id・別案の差分を検査する（保守する側）。positioning の検査もこの skill の `check` が一緒に走らせる（台本が指す経路だけが出荷を止めるため）。
工場 skill は台本を受け取って描画するだけで、情報設計の判断はしない。描画した物は `evaluate` で、
設計どおりか（文言・順序・固定値は決定論、意味は Jev の判定台帳）を確かめ、warn / error を出す。

## 重要ルール

- **承認の状態を書かない。** `ux/positioning/positioning.yaml` の `state` は `empty` / `proposed` だけ。positioning の変更も、検査が緑なら
  他の変更と同じく自律マージする（人のレビューは置かない）。値の出所は `provenance` に必ず書く。
- **制作中に positioning を書き換えない。** 足りない・合わないと分かったら、台本を直すか、positioning の変更をオーナーに提案する（承認はオーナーが提供価値ごとの経路で。依頼のページは own-positioning-author が作る）。
- **固定値と正式名称は辞書の id で書く。** 台本の文案では、数値・連絡先・免責などは `{{value-<辞書の id>}}` で参照し、値を直接書かない。製品名などは辞書の `name` の表記で書く。声（repo 全体の話し方の正典）は無い。トーンは台本の `tone`。
- **対象 0 件は失格。** チャネル 0 件・台本 0 件は終了コード 2。廃止した `ux/infoarch/outlines/` にファイルが残っていても失格（`IA_OUTLINE_RETIRED`）。outline は機械で写さず、positioning から台本を作り直す。
- **引き先の無い文は提案の印を付けて残す。** 主張・事実・証拠と声の文案は、引き先（伝えるメッセージ・positioning の証拠 `refs`・辞書の値）か提案の印（`proposal: true`）を持つ。制作中は警告、出荷する台本に残っていれば `ship` が止める（`IA_SCRIPT_PROPOSAL_AT_SHIP`）。
- **仕上がりの承認は台本・組み方の計画・成果物のファイルをハッシュで縛る。** 出荷する台本ごとに承認台帳の `ia:release:<台本 id>` の行が要り、承認の後に 1 字でも変われば外れる（カンプの画像は縛らない）。提案の印が残る台本は承認の対象を作らない。`infoarch-author release` が今のハッシュと状態を出す。承認を記録するのはオーナーの言葉があるときだけ（AI は台帳に推測で書かない）。

## プロダクト UI は blueprint で書く

ログイン後の画面・ダッシュボード・操作が主目的の成果物（プロダクト UI）の infoarch は、台本ではなく **blueprint**（`ux/infoarch/blueprint/`）。型は SPEC-product-ui-infoarch（biz_ops の `docs/specs/product-ui-infoarch.md`）、契約は `references/contract.md` の「blueprint」。

- `product.yaml`（プロダクトに 1 つ）：ストーリー・原則（並べ方の原則を含む）・制約（やらないこと）・オブジェクトと主役・入口。**オーナーが 1 件ずつ承認する**（行のキーは `ia:item:<項目 id>`、ハッシュは項目 1 件の中身だけ。1 件足しても、ほかの承認は外れない）。AI は案を作ってよいが、承認の台帳に記録するのは、オーナーの承認の言葉と日付があるものだけ。
- `screens/<scr-…>.yaml`（画面ごと。AI が作る）：応えるストーリー、グループ（並べ方の原則が正当化する）と領域（優先度・中身の種類 `shape`）、主操作（ちょうど 1 つ）、状態、移動先。承認済みの項目の id だけを引く。位置・幅・部品は書かない（design の割り付け規則。部品の選択も design）。
- 失格：どのストーリーにも応えない画面・引く id が無い・承認の無い項目を引く（出荷する画面は失格、下書きは警告）・主操作が 2 つ以上か 0 個・優先度が欠けるか重なる・状態の宣言が無い・並べ方の原則が無いグループ・shape が決まった種類でない・id の再利用 など。条件と赤のフィクスチャの対応表は `fixtures/blueprint-red-cases.json`。
- **縦の順の確かめは、blueprint の型を変えるときだけ**（毎回の検査にはしない。道具も無い）：項目の種類や持つ情報を変えるときは、その変更の中で一度、実配置を見ない別の AI が、画面の優先度と design の割り付け規則だけから縦の順を導けるかを手で一度試し、結果を作業の記録（WS・issue の Log）に残す。食い違いの理由が型のどの項目で説明できるかも残し、説明できない食い違いがあれば型を直す。描画した画面が優先度どおりに並んでいるかは、design の割り付けの検査が毎回確かめる（SPEC-product-ui-infoarch の Quality、ADR-20261123-ia-skill-completion-judgments-and-measurement）。
- `infoarch-author blueprint --root .`（検査。`check` の報告から台本側の finding だけを除く。positioning の失格・id の参照切れも `check` と同じに出て、同じ終了コードになる）／`infoarch-author request --root . --out DIR`（オーナーに見せる承認依頼のページ。プロダクト全体の関係図と、項目ごとのカードに ○ / ×。Artifact の使い方は own-positioning-author と同じ）／`infoarch-author targets`（`ia:item:` の行と、出荷が要る行を含む）／`infoarch-author impact --id <項目 id>`（引く画面）。

## 作る：依頼から台本までの 6 段

依頼を受けたら、次の 6 段を順に行い、`ux/infoarch/scripts/<成果物 id>.yaml`（台本）を書く。各段では、その段のファイルと、作る媒体のファイルだけを読む。

| 段 | 読むファイル | 台本に書くもの |
| --- | --- | --- |
| 1 依頼を固める | `references/steps/01-brief.md` | objective（取ってほしい行動）・touchpoint・channel・outputs |
| 2 target を決める | `references/steps/02-audience.md` | target（セグメント＋足す条件）・tone |
| 3 選ぶ・一番・削る | `references/steps/03-select-and-cut.md` | messages（提供価値を 1 つ・rank・type）・no_cta |
| 4 構造を決める | `references/steps/04-structure.md` と媒体のファイル | blocks（まとまり：order・function＝内容の働き・places）・texts |
| 5 文案を書く | `references/steps/05-copy.md` | copies（引き先：message・refs・`{{value-<辞書の id>}}`。無ければ proposal） |
| 6 見直して渡す | `references/steps/06-review.md` | `check` を緑にし、`set` の id を控えて工場へ渡す |

媒体のファイル：LP は `references/media/lp.md`、複数ページのサイトは `references/media/site.md`、チラシ・ポスターは
`references/media/flyer-poster.md`、資料は `references/media/deck.md`、アプリの文言は `references/media/app-copy.md`。

- 小さな依頼（単発のチラシなど）は `references/light-process.md` に従って簡略にしてよい。段 1 と段 6 は省かない。
- 新しい repo で SSOT が空なら、先に `references/bootstrap.md` に従って既存の資料から下書きを起こし、検査を緑にしてマージする。
- 例：`references/examples/ux/` に、2 つのセグメントの positioning と、それぞれに立つ台本（LP と資料）がある。

ノウハウの数値に「出典なし・目安」とあるものは、検査に使わない。

## 使い方

```sh
infoarch-author check --root .              # SSOT・チャネル・台本の決定論の検査。終了コード 0 / 1（警告）/ 2（失格）。報告は executed・skipped・missing・uncovered・screened の 5 区分を空でも出す（走らせなかった dom・判定は skipped に名前が出る）
infoarch-author impact --root . --id val-0001  # その原子に立つ台本と成果物を列挙（positioning は矢印をたどって推移的に）
infoarch-author skeleton --root . --script lp-xxx [--out DIR]  # 台本から骨組みの HTML（アンカー・順序・固定値）を出す。工場はこれを崩さずにスタイリングする
infoarch-author dom --root . --script lp-xxx # 描画物をブラウザで開き、アンカー・順序・固定値の表示を確かめる
infoarch-author judge --root .                 # 意味の判定（文案 × 提供価値の 1 種類だけ）を共有の判定台帳から読む。Jev は呼ばず、出荷を止めない
infoarch-author accept-reviews --root . --id <判定の id> --owner-words '<オーナーの言葉>' --approved-on YYYY-MM-DD [--write]  # オーナーが原文を見て認めた要確認を記録する（言葉が無ければ書かない）
infoarch-author evaluate --root . [--script lp-xxx]   # 作った物の評価の入口。check → dom → judge を順に走らせ、報告と終了コードを 1 つにする（工場は描画後に必ず走らせる）
infoarch-author evaluate --root . --script lp-xxx --acquire --answers-dir DIR --limit 50 [--execute]  # 足りない Jev 判定を上限つきで取って台帳に記録し、評価し直す（--execute が無ければ見積もりだけ）
infoarch-author ship --root . --channel web   # 出荷の正式の入口。出荷されるコミット（git の中身）で tree → check → dom → judge が失格でない時だけ ship_command を実行する（judge は警告だけで、止めない）
infoarch-author ship --root . --doctor        # ship_command・merge_base の無い出荷チャネル、pre-push の未導入、拡張の失格を伝えない dispatcher を失格にする
infoarch-author maintain --root .             # セッション終了の Step 2.5。ux/infoarch/ も ux/positioning/ も無ければ 1 行で skip
infoarch-author set --root . --script lp-xxx   # set の id（台本の id `spt-…`＋指す経路の id）と、台本の要素ごとのハッシュ（`--outline` も同じ）
infoarch-author diff --root . --a lp-a --b lp-b [--root-b DIR] [--dom-a A.json --dom-b B.json]  # 2 つの set の差分（変わった要素・層・同じ試験か）。印の外の文の違いは失格
infoarch-author answer-key --root . --out DIR      # 意味の判定（文案 × 提供価値）の正解表：組・swap の候補・オーナーが ○× を付けるページ
infoarch-author answer-key --root . --record --key DIR/answer-key.json --limit 50 [--execute]  # swap の組の答えを上限つきで Jev に聞き、測定用の台帳（DIR/measurement-judgments.json。出荷は読まない）に記録する（--execute が無ければ見積もりだけ）
infoarch-author answer-key --root . --measure labels.json --key DIR/answer-key.json  # 一致の割合・見逃し・誤検出（own は共有の台帳、swap は測定用の台帳から）。own か swap か、○ か × が 0 件なら「測れない」
infoarch-author release --root .               # 仕上がりの承認：出荷する台本ごとのキー `ia:release:<台本 id>`・承認が持つべきハッシュ（台本・組み方の計画・成果物のファイルの中身）・状態。出荷する台本に承認が現在のものでないものがあれば終了コード 1。`ship` も同じ理由で止める（IA_RELEASE_UNAPPROVED）
infoarch-author answer-key --root . --carry-over OLD_LABELS --from-key OLD_KEY --key DIR/answer-key.json --carry-out FILE  # v6 の正解表のラベルを、文案の本文と提供価値が同じ組へ移す（id は変わるので id では対応しない）。移さなかったものと、ラベルが付かない今の組を出す
infoarch-author targets --root .               # 承認の対象（ia:path:<提供価値 id> と経路のハッシュ）と、出荷が要るキー（承認依頼のページは positioning-author request）
infoarch-author migrate-dictionary --root . --map map.json [--write]  # スキーマ版 4 の木（固定値・正式名称・声を SSOT に持つ）を、辞書とスキーマ版 5 へ一度だけ移す。dry run は移した木の検査も走らせる
infoarch-author migrate-v6 --root . [--write]  # v5 の木（SSOT・outline・描画物の印・modules.ia）を一括で移す（outline は台本に）。旧→新の id・移しそこね（1 件でも終了コード 2、何も書かない）・外れる承認と判定・戻す手順を出す
```

出荷は `ship --channel` を通す（`ship_command` を直接叩かない）。push は、マーカー（`templates/infoarch-ship.enabled` を
`.vibe-guard/` に写してコミットする）のある repo で、vibe-guard の pre-push に `hooks/pre-push.d/61-infoarch-ship` を置くと、
push される各コミットそのものと、その予測マージ木で止まる。手元のマーカーを消しても、コミットされたマーカーで動く。
導入（マーカー・拡張・PATH）は repo ごとにオーナーの確認を取ってから行う。

`scripts/infoarch-author` を PATH に置き、PyYAML と jsonschema のある Python を `DESIGN_AUTHOR_PYTHON` で選ぶ
（自動インストールしない。無ければ終了コード 2）。repo の設定にローカルパスを書かない。
新しい repo では `templates/infoarch/`（チャネルと散文）を `ux/infoarch/` に、own-positioning-author の `templates/positioning/positioning.yaml` を `ux/positioning/` に写す。

## 参照

- `references/contract.md` — ディレクトリ・台本の契約、辞書との関係、検査の code 一覧、期限の扱い、描画物の確認（工場が埋めるアンカー）、意味の判定（判定台帳）。positioning の契約は own-positioning-author の `references/contract.md`。
- `references/schema/` — `channels/*.yaml`（版 2）・`scripts/*.yaml`（台本、版 1）の JSON Schema。positioning（`ux/positioning/positioning.yaml`）の schema（ファイル名は `info_arch.v5.schema.json` のまま）は own-positioning-author。
- `references/expiry.json` — 証拠の期限の初期値（固定値の期限は辞書が持つ）。
- `fixtures/red-cases.json`・`script-red-cases.json`・`migrate-red-cases.json`（と own-positioning-author の `positioning-red-cases.json`。ここでも台本付きで走らせる）— 必須の赤ケースと、失格条件（移行では移しそこねの条件）との対応表。自己テストは `uv run --no-project --with-requirements <skill>/scripts/tests/requirements.txt -- pytest -q <skill>/scripts/tests`。
