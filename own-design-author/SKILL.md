---
name: own-design-author
description: リポジトリの design/ を作成・更新し、ブランドの値・方針・アセットを一本化する。デザインシステムを定義する時、design/ を起こす時、トークンや項目カタログを移行する時に使う。成果物の生成・UIレビュー・ブランド戦略の策定は各担当へ渡す。
---

# Own Design Author

repo の git toplevel にある `design/` を唯一の正典にする。上流のカタログは項目の定義で、ブランドの既定値ではない。

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

`approvals.yaml` は人間のチャット確認が唯一の根拠。AIが自律的に書かない。L0 のマニフェストにある `value_hash` を提示し、未承認と承認済みを混同しない。値・理由・散文の対象部分が変われば台帳のハッシュ不一致で止まる。

```sh
design-author check --repo <repo>
design-author check --repo <repo> --ship
```

L0 はモデルを使わず、対象項目数・充足率・未判定・提案値・未カバーを出す。整合が通った時だけ `design/.resolved/<media>.json` と CSS、`assets.json` を生成する。失格時は古いマニフェストを残さない。L1/L2 はマニフェストだけを読む。L0だけで成果物の準拠合格を名乗らない。

`--ship` は未承認mustと未カバーmust規則を失格にする。hook/miseへの接続は `references/ship-gates.md`。既存hookを置き換えない。
L1/L2と生成Tailwindの使い方・制約は `references/web-validation.md` を読む。

## 上流更新

上流の再取得は人間が起動する。`references/catalog-maintenance.md` の手順でカタログを生成し、差分と判定の使用量を確認する。推測で見出しを改名扱いにしない。元の版を残したまま新しい版を追加する。

```sh
design-author migrate --repo <repo> --catalog-version <new-version>
design-author check --repo <repo>
```

`migrate` 前のrepoのpinは変えない。旧状態と承認台帳を保存し、新しい項目は `deferred` として評価する。新mustはL0で可視化する。

## 検証と範囲

`references/contract.md` に runtime・ハッシュ・対応範囲、`fixtures/red-cases.json` に必須の赤ケースがある。自己テストは `python3 -m pytest -q <skill>/scripts/tests`。実上流の全件照合は `DADS_TEST_ARCHIVE` に手元のzipを指定する。未指定のテストskipを照合済みと報告しない。

pptx 工場は従属WSが終わるまで `design/` を読まない。資料presetは現在webプレビュー用で、pptxの準拠を保証しない。L0/L1/L2をLLMで代替しない。必要な非決定論レビューは決定論検査の後に行う。
