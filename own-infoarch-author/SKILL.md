---
name: own-infoarch-author
description: "repo 直下 design/info_arch/ の「伝え方の SSOT」（オーナーが承認する positioning：セグメントごとの原子と矢印＝ジョブ・痛み・代替・カテゴリー・強み・提供価値・証拠、と固定値・正式名称・声）と、成果物ごとの outline（target・メッセージ・セクション・文案）を読み書きし、Web サイト・LP・チラシ・ポスター・資料を作るときに、誰に・何を・どの順で・どう話すかをそろえる。「情報設計を作って」「伝え方をそろえて」「SSOT から LP の構成を作って」「メッセージを整理して」「info_arch を検査して」で使う。見た目は own-design-author と工場 skill（own-web-design 等）。"
---

# own-infoarch-author — positioning と outline

`design/info_arch/` は、repo の伝え方の唯一の正典。視覚の `design/` の兄弟モジュールで、スタイリングではなく
「誰に・何を・どの順で・どれだけ強く・どんな話し方で伝えるか」を持つ。
人が保守するのは positioning（オーナーが原子と矢印を 1 つずつ承認）だけ。この skill は、承認済みの positioning から成果物ごとの
**outline**（1 ページ＝1 outline）を作り（作る側）、positioning と outline の失格・id・別案の差分を検査する（保守する側）。
工場 skill はアウトラインを受け取って描画するだけで、情報設計の判断はしない。描画した物は `evaluate` で、
設計どおりか（文言・順序・固定値は決定論、意味は Jev の判定台帳）を確かめ、warn / error を出す。

## 重要ルール

- **承認の状態を書かない。** `info_arch.yaml` の `state` は `empty` / `proposed` だけ。SSOT の変更も、検査が緑なら
  他の変更と同じく自律マージする（人のレビューは置かない）。値の出所は `provenance` に必ず書く。
- **制作中に SSOT を書き換えない。** 足りない・合わないと分かったら、outline を直すか、positioning の変更をオーナーに提案する（承認はオーナーが原子と矢印ごとに）。
- **固定値は参照で書く。** アウトラインの文案では、数値・連絡先・免責などは `{{fct-NNNN}}` で参照し、値を直接書かない。
- **対象 0 件は失格。** チャネル 0 件・アウトライン 0 件は終了コード 2。

## 作る：依頼からアウトラインまでの 6 段

依頼を受けたら、次の 6 段を順に行い、`design/info_arch/outlines/<成果物 id>.yaml` を書く。各段では、その段のファイルと、作る媒体のファイルだけを読む。

| 段 | 読むファイル | アウトラインに書くもの |
| --- | --- | --- |
| 1 依頼を固める | `references/steps/01-brief.md` | objective（取ってほしい行動）・touchpoint・channel・outputs |
| 2 target を決める | `references/steps/02-audience.md` | target（セグメント＋足す条件）・tone |
| 3 選ぶ・一番・削る | `references/steps/03-select-and-cut.md` | messages（提供価値を 1 つ・rank・type）・no_cta |
| 4 構造を決める | `references/steps/04-structure.md` と媒体のファイル | sections（order・role・places）・texts |
| 5 文案を書く | `references/steps/05-copy.md` | copies（固定値は `{{fct-NNNN}}`） |
| 6 見直して渡す | `references/steps/06-review.md` | `check` を緑にし、`set` の id を控えて工場へ渡す |

媒体のファイル：LP は `references/media/lp.md`、複数ページのサイトは `references/media/site.md`、チラシ・ポスターは
`references/media/flyer-poster.md`、資料は `references/media/deck.md`、アプリの文言は `references/media/app-copy.md`。

- 小さな依頼（単発のチラシなど）は `references/light-process.md` に従って簡略にしてよい。段 1 と段 6 は省かない。
- 新しい repo で SSOT が空なら、先に `references/bootstrap.md` に従って既存の資料から下書きを起こし、検査を緑にしてマージする。
- 例：`references/examples/design/info_arch/` に、2 つのセグメントの positioning と、それぞれに立つ outline（LP と資料）がある。

ノウハウの数値に「出典なし・目安」とあるものは、検査に使わない。

## 使い方

```sh
infoarch-author check --root .              # SSOT・チャネル・アウトラインの決定論の検査。終了コード 0 / 1（警告）/ 2（失格）
infoarch-author impact --root . --id fct-0001  # その原子に立つアウトラインと成果物を列挙（positioning は矢印をたどって推移的に）
infoarch-author skeleton --root . --outline lp-xxx [--out DIR]  # アウトラインから骨組みの HTML（アンカー・順序・固定値）を出す。工場はこれを崩さずにスタイリングする
infoarch-author dom --root . --outline lp-xxx # 描画物をブラウザで開き、アンカー・順序・固定値の表示を確かめる
infoarch-author judge --root . [--dom-report dom.json]  # 意味の判定（表記ゆれ・強調・未登録の主張・文案 × 提供価値・妥当性）を共有の判定台帳から読む。Jev は呼ばない
infoarch-author evaluate --root . [--outline lp-xxx]   # 作った物の評価の入口。check → dom → judge を順に走らせ、報告と終了コードを 1 つにする（工場は描画後に必ず走らせる）
infoarch-author evaluate --root . --outline lp-xxx --acquire --answers-dir DIR --limit 50 [--execute]  # 足りない Jev 判定を上限つきで取って台帳に記録し、評価し直す（--execute が無ければ見積もりだけ）
infoarch-author ship --root . --channel web   # 出荷の正式の入口。出荷されるコミット（git の中身）で tree → check → dom → judge が緑の時だけ ship_command を実行する
infoarch-author ship --root . --doctor        # ship_command・merge_base の無い出荷チャネル、pre-push の未導入、拡張の失格を伝えない dispatcher を失格にする
infoarch-author maintain --root .             # セッション終了の Step 2.5。design/info_arch/ が無ければ 1 行で skip
infoarch-author set --root . --outline lp-xxx   # set の id（positioning の id＋outline の id）と、outline の要素ごとのハッシュ
infoarch-author diff --root . --a lp-a --b lp-b [--root-b DIR] [--dom-a A.json --dom-b B.json]  # 2 つの set の差分（変わった要素・層・同じ試験か）。印の外の文の違いは失格
infoarch-author answer-key --root . --out DIR      # 意味の判定（文案 × 提供価値）の正解表：組・Jev の候補・オーナーが ○× を付けるページ
infoarch-author answer-key --root . --measure labels.json --key DIR/answer-key.json  # 一致の割合・見逃し・誤検出。○ か × が 0 件なら「測れない」
infoarch-author positioning --root .           # セグメントごとの positioning の id・隙間・同等であるべき点・所属カテゴリー・承認待ち
infoarch-author targets --root .               # 承認の対象（ia:positioning:<原子 id か矢印のキー> のハッシュ）と、出荷の経路が要るキー
infoarch-author review --root . --request design/approval-request-ia-positioning.json  # 承認の束の中身を読める形で。古い束は失格
```

出荷は `ship --channel` を通す（`ship_command` を直接叩かない）。push は、マーカー（`templates/infoarch-ship.enabled` を
`.vibe-guard/` に写してコミットする）のある repo で、vibe-guard の pre-push に `hooks/pre-push.d/61-infoarch-ship` を置くと、
push される各コミットそのものと、その予測マージ木で止まる。手元のマーカーを消しても、コミットされたマーカーで動く。
導入（マーカー・拡張・PATH）は repo ごとにオーナーの確認を取ってから行う。

`scripts/infoarch-author` を PATH に置き、PyYAML と jsonschema のある Python を `DESIGN_AUTHOR_PYTHON` で選ぶ
（自動インストールしない。無ければ終了コード 2）。repo の設定にローカルパスを書かない。
新しい repo では `templates/design/info_arch/` を写す。

## 参照

- `references/contract.md` — ディレクトリ・原子・アウトラインの契約、検査の code 一覧、期限の扱い、描画物の確認（工場が埋めるアンカー）、意味の判定（判定台帳）。
- `references/schema/` — `info_arch.yaml`（版 4）・`channels/*.yaml`（版 2）・`outlines/*.yaml`（版 3）の JSON Schema。
- `references/expiry.json` — 証拠・固定値の期限の初期値。
- `fixtures/red-cases.json`・`positioning-red-cases.json`・`outline-red-cases.json` — 必須の赤ケースと、失格条件との対応表。自己テストは `uv run --no-project --with-requirements <skill>/scripts/tests/requirements.txt -- pytest -q <skill>/scripts/tests`。
