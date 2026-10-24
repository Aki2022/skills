---
name: own-brand-resolve
description: "指示されたブランドやデザインの参照ライブラリ。自社ブランド・デジタル庁風・指定サイトの方針や値を調べ、出所と権利情報を own-design-author に渡す時に使う。repo の design/ が正典で、値の承認と成果物の生成は担当しない。"
---

# Own Brand Resolve

倉庫は指示された時だけ引く参照ライブラリ。repo の `design/` が唯一の正典で、
工場は倉庫から値を継承しない。ブランドの企画・命名・戦略や成果物の生成は扱わない。

## 参照の手順

1. ユーザーが指定した参照元と必要な部分だけを読む。自律探索・全件スキャンはしない。
2. 第三者 DESIGN.md の許可元は `kzhrknt/awesome-design-md-jp` と `VoltAgent/awesome-design-md`。
   `plugins/awesome-design-md-jp-import/import-policy.md` と `mapping.md` に従う。
3. 取得した版のライセンス実物を確認し、色・余白・角・影と方針を `own-design-author` に渡す。
   repo の `design/` に `proposed` として写し、値の出所を `reference:<id>`、
   URL・commit・ライセンス・著作権表示・編集有無を同じrepoの `design/NOTICE` に残す。
4. 本家の英語UIから typography の値は写さない。空のまま「和文フォント未決定」を報告する。
   和文フォントの提案元は -jp またはオーナー指定だけ。
5. L0 の結果と未決定を報告する。AIが `approvals.yaml` を自律的に書かない。

ロゴ・商標・固有ビジュアル、実アセット、既存サービスの画面・実装は写さない。
`allow_asset_copy: false` は画像・アイコン・フォント等の実ファイルを指し、数値提案の許可とは別。
出所と権利確認はブランド値の承認を代替しない。

## 同梱物と取得メタ

| パス | 役割 |
| --- | --- |
| `plugins/my_company` | owned の参照箱（自社ブランド） |
| `plugins/digital-agency` | observed の参照箱（公式tokenあり） |
| `plugins/notion` | observed の参照箱（散文のみ） |
| `plugins/awesome-design-md-jp-import` | 許可された第三者DESIGN.mdの参照入口 |
| `references/plugin-spec.md` | 箱のスキーマと参照時の権利確認 |
| `references/resolver.md` | 旧Resolverの互換仕様。web工場へは適用しない |

等級は `owned` / `observed`。各 `plugin.yaml` は `id` 直後に `tier` と `source`
（upstream / upstream_version / fetched_at）を持つ。版が無いなら `unversioned`、
取得時に記録していなければ `unrecorded` と書き、版を推測しない。再取得は人間が起動する。

## 互換範囲

既存箱・Resolverはpptxの段階移行中の互換のため保持する。旧Resolverの優先度・1階層extendsは
既存利用者に限り、repoの `design/` に持ち込まない。`own-pptx-build` は従属WSが完了するまで
`design/` を読まず、倉庫との同値検査を維持する。既存の実アセットの許可は
`plugin.yaml` と `references.md` の記録を確認し、参照トークンの許可で拡大しない。
