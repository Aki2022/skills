---
name: own-design-author
description: リポジトリの ux/design/ を作成・更新し、ブランドの値・方針・アセットを一本化する。デザインシステムを定義する時、ux/design/ を起こす時、トークンや項目カタログを移行する時に使う。成果物の生成・UIレビュー・ブランド戦略の策定は各担当へ渡す。
---

# Own Design Author

承認記録とJCSハッシュ照合、固定環境のL0〜L2出荷ゲート、維持モードを提供する。署名必須は所有者の判断でDeferred。実パイロットの承認台帳・保存済み判定・全画面検証は別に揃える。人工repoの合格を本番の合格にしない。導入境界は `references/approval-upgrade.md` と `references/ship-gates.md`。

repo の git toplevel にある `ux/design/` を唯一の正典にする（layer の置き場所は `references/layer-paths.json` が唯一の規則。positioning は `ux/positioning/`、infoarch は `ux/infoarch/`、承認台帳 `ux/approvals.yaml` と判定台帳 `ux/judgments.json` は `ux/` の直下）。上流のカタログは項目の定義で、ブランドの既定値ではない。

## 入口

1. `references/contract.md` のファイル契約と承認境界を読む。
2. PATH にこの skill の `scripts/` があれば `design-author` を使う。無ければ skill の `scripts/design_author.py` を既存の Python 環境から実行する。repo の設定へ個別端末のパスを書かない。
3. `design-author check --repo <repo>` を実行する。無い場合だけ `init` する。生成工場へは、L0 が終了コード2なら渡さない。1は警告を残して制作できる。0だけを無条件の整合合格とする。

## 新規作成と充填

```sh
design-author init --repo <repo> --preset product-ui --catalog-version 2026-10-14
design-author check --repo <repo>
```

プリセットは `lp` / `product-ui` / `presentation`。一括状態は `deferred` のみ。must が空なら最初の L0 が失格になるのが正常。`n-a` は一括で付けず、項目ごとの理由を残す。

カタログ全体を LLM のコンテキストに入れない。機械可読カタログをスクリプトで絞り、作業対象の差分・未定義mustと必要な原文だけを読む。`catalog/2026-10-14.json` の must 判定を編集者の勘で置き換えない。

ブランドの値は、オーナーの指定またはAI提案でトークンへ入れ、出所を `owner` / `proposal` / `reference:<id>` として記録する。値の状態は承認台帳から導く。AIが `approved` フラグを立てない。

参照元を指定された場合だけ参照する。許可された参照元・ライセンスの確認は `own-brand-resolve` の現行方針に従う。その方針を本 skill から拡張しない。ロゴ・商標・固有ビジュアルは写さない。本家の英語UIからは和文 typography を写さず、未決定を報告する。参照値には `NOTICE` の出典行を付ける。

`DESIGN.md` は方針と `token:group.name` 参照を持つ。値を二重記載しない。項目状態は `design.yaml` でプリセットとの差分だけを持つ。`defined` は実在する token / doc anchor / 宣言的 rule を指す。

## 承認と検査

`ux/approvals.yaml` はオーナーが承認した対象のkey・hash・承認者・日付を記録する。承認の判断はオーナー、記録はAI（ADR-20261118-ai-records-owner-approvals、biz_ops）。AIが記録してよいのは、会話やPRにオーナー自身の承認の言葉がある時だけで、`approve --request` に `--owner-words`（言葉の引用）と `--approved-on`（その日付）が無ければ記録しない。言葉と日付は、結果の `commit_trailer` として commit メッセージと PR 本文に書く。署名は要求せず、承認者の真正性やAIによる偽造防止は保証しない。`approve --prepare` は対象hashの確認用リクエストだけを返す。positioning は outline より先に承認するので、outline 0 件でも prepare できる。`ux/design/design.yaml` の無い IA だけの repo でも prepare と記録ができる。値・理由・散文・alias先・媒体profileの変更は失効する。確認・記録の手順と運用上の限界は `references/approval-upgrade.md`。

```sh
design-author check --repo <repo>
design-author check --repo <repo> --ship
```

L0 はモデルを使わず、対象項目数・充足率・未判定・提案値・未カバーを出す。整合が通った時だけ `ux/design/.resolved/<media>.json` と CSS、`assets.json` を生成する。失格時は古いマニフェストを残さない。L1/L2 はマニフェストだけを読む。L0だけで成果物の準拠合格を名乗らない。

`--ship` は未承認mustと未カバー規則を失格にする（規則の未カバーはmust以外も失格）。hook/miseへの接続は `references/ship-gates.md`。既存hookを置き換えない。
`defined` の全項目は、skillの対応表 `references/item-checks.json` で3つの確かめ方のどれかに乗る。値・構造で言えるものはコードで画面ごとに確かめる。repoごとに決めること（どの型をいつ使うか、文章の方針など）は、repoが `checks/*.yaml` に宣言した検査（数・文・計算済みスタイル）で画面ごとに確かめ、オーナーの承認も要る。DESIGN.mdに書いて承認しただけでは合格しない。画面を見ないと判断できないものだけを、部品がある画面に限って端末内のApple Foundation Models（AFM）で判定する。どれにも乗らない項目は出荷で止まる。AFMの違反は精度を測るまで警告。詳細は `references/web-validation.md`。
push時の検査は、`ux/design/design.yaml` を持つ全repoで pre-push の拡張が必ず回す。repoごとの有効化は要らない。
出荷には `validation.json` の `pages`（画面の棚卸し）が必須。repo の画面はすべて、検査するか、オーナーの承認つきで対象外にする。URL で検査する画面は、URL のパスが画面ファイルの経路に合うことを L0 で照らし合わせる。詳細は `references/web-validation.md` の「画面の棚卸し」。
クリックや送信の後にだけ現れる部品（ダイアログ・メニュー・送信後のエラー文など）は、`validation.json` の URL に `states`（click/focus/fill/submit の宣言的な手順）を置き、その状態を描いて確かめる。手順が要素に合わなければ失格。詳細は同じ文書の「操作の後の状態」。
最初に外れた画面で止まらずに全画面・全状態の外れを並べたい時は `design-check --report-all`（診断用、出荷の証拠にはならない）。Google Fonts など CDN のフォントは、`render-environment.json` の `remote_fonts` で URL と SHA-256 を固定すれば描画に使える（`scripts/pin_remote_fonts.mjs`）。詳細は `references/web-validation.md`。

`design-author sync` はL0通過後に宣言したTailwindとassetコピーを配置する。`design-check` はL0/L1を通ったら画面ごとにL2とITEMSを実行し、全画面が通ってからAFM段を実行するためL2実施確認後に呼ぶ。
L1/L2と生成Tailwindの使い方・制約は `references/web-validation.md` を読む。

LP など台本から作る成果物の組み方の型（分割 2 列・帯・行の表など）の語彙は `catalog/layout-patterns.json`。repo は `design.yaml` の `layout_patterns` に差分（使わない型・足した型）だけを書き、足した型はオーナーが 1 つずつ承認する。カンプの大きさの既定は `references/comp-sizes.json`。どちらも L0 がマニフェストに解決する（`references/contract.md`）。

プロダクト UI の画面の位置と幅は、割り付け規則 `references/allocation-rules.md`（機械が読む形は `allocation-rules.json`）が、infoarch（blueprint）の優先度・グループ・shape から導く。グリッド・breakpoint・画面の型・部品の選択（表かカードか、どのグラフか）・規則に反する配置を人が指定するときの書き方（`ux/design/layout-overrides.yaml`）はここにある。作るのは own-web-build の `build_screen.py`、確かめるのは `audit_screen.py`。

規則に反する配置は、1 行（id・画面・領域の組・幅・理由。幅は省けない）ずつオーナーが承認する。理由だけでは通らない。承認は台帳の `ds:web:override:<id>` で、行の中身と指す領域の宣言にハッシュで縛られる（行を変える・領域が変わる・写すと引き継がない）。承認依頼のページは `design-author override-request --repo <repo> --out <dir>`（プロダクト単位の項目と同じ台帳・同じページの形。`--read-decisions` で答えを読み戻す）で作り、記録はオーナーの言葉と日付を受けた `design-author approve --prepare --target ds:web:override:<id>` → `approve --request …` で行う。詳細は `references/allocation-rules.md`。

## 判定台帳と共有runner

散文のdoc証拠とn-a理由は、保存済みの判定だけを読む。未登録入力は未カバーとして報告し、全候補の現在の記録が揃わなければ出荷失格。
doc証拠/n-a理由の保存済み判定はscreenedとして報告し、高scoreは原文照合用の警告にする。値だけで出荷を止めず、意味の十分さを保証しない。実ラベルが無ければ精度未測定と明示し、recallを作らない。判定の未登録・古い入力・不正台帳は出荷停止。モデルを検査中に呼ばない。
回答取得は検査の外でjudgment_run.pyを明示実行し、既存分類と共通の使用量上限を守る。未ラベル回答を合格へ昇格しない。
記録・移行と5区分報告、IA接続の詳細は `references/judgments-and-modules.md`。
IA入力を見つけても対応実装を黙って無視しない。IA本体は兄弟WSが担当する。

## 上流更新

セッション終了時は `design-author maintain --repo <repo>` を使う。前回成功時からの変更・期限・実装/カタログ/policyの変更で再生成とL0を実行し、失効した承認と候補を列挙する。承認は書かず、モデルは呼ばない。`ux/design/` 不在は1行skip。期限設定・出力と失敗時の扱いは `references/maintenance.md`。

上流の再取得は人間が起動する。`references/catalog-maintenance.md` の手順でカタログを生成し、差分と判定の使用量を確認する。推測で見出しを改名扱いにしない。元の版を残したまま新しい版を追加する。

```sh
design-author migrate --repo <repo> --catalog-version <new-version>
design-author check --repo <repo>
```

`migrate` 前のrepoのpinは変えない。旧状態と承認台帳を保存し、新しい項目は `deferred` として評価する。新mustはL0で可視化する。

## design/ から ux/ への移行

旧い並び（repo 直下の `design/` に design・承認台帳・判定台帳・`design/info_arch/` が入る）の repo は、どのコマンドも同じヒントを付けて断る（`LAYER_LAYOUT_FOREIGN_ROOT`）。`design/` を読み飛ばして「0 件」にしないため。移すのは一度だけ。

```sh
design-author migrate-layout --repo .          # dry run: 何がどこへ動くかを出すだけで、何も動かさない
design-author migrate-layout --repo . --apply  # 実際に動かす
```

動かすのは置き場所だけ。中身のバイトは変えない。git の追跡ファイルは rename で動かし（履歴を辿れる）、`.resolved/` などの無視された生成物も一緒に運ぶ。`.gitignore` の行と、アセット規則の対象（`checks/*.yaml` の `assets/` への参照）は新しい置き場所に追従する。repo に無い layer は作らない（IA だけの repo に `ux/design/` は作らない）。`design/` の中の散文（`DESIGN.md` など）に旧い置き場所の名前が残っていれば、承認が文面に結び付いているので書き換えず、報告だけする。途中で失敗したら全部元に戻る。

## 検証と範囲

`references/contract.md` に runtime・ハッシュ・対応範囲、`fixtures/red-cases.json` に必須の赤ケースがある。自己テストは `uv run --no-project --with-requirements <skill>/scripts/tests/requirements.txt -- pytest -q <skill>/scripts/tests`。実上流の全件照合は `DADS_TEST_ARCHIVE` に手元のzipを指定する。未指定のテストskipを照合済みと報告しない。

pptx 工場は従属WSが終わるまで `ux/design/` を読まない。資料presetは現在webプレビュー用で、pptxの準拠を保証しない。L0/L1/L2をLLMで代替しない。必要な非決定論レビューは決定論検査の後に行う。
