# DESIGN.md — Reference Import Policy

ユーザーが指示した個別サイトだけを参照する。許可元は `kzhrknt/awesome-design-md-jp` と
`VoltAgent/awesome-design-md`。自律的な探索・全件スキャンはしない。

## 写す手順

1. 指定された DESIGN.md の URL・commit と、その版のライセンス実物・著作権表示を確認する。
   確認できなければ値を写さず、不足を報告する。過去の取得メタを今回の確認の代用にしない。
2. `mapping.md` に従い、方針と色・余白・角・影の値を repo の `design/` へ写す。
   値は DTCG トークンにし、出所を `reference:<id>`、状態を `proposed` とする。
3. 同じ repo の `design/NOTICE` に id・URL・commit・ライセンス・著作権表示・編集有無を記録する。
4. -jp またはオーナー指定から和文フォントを提案できる。本家の英語UIからは typography の値を
   写さず、空のまま「和文フォント未決定」と報告する。
5. `own-design-author` の L0 で整合を確認する。ブランド値の承認は人間が行う。
   `approvals.yaml` にAIが自律的に書き込まない。

参照は初期提案の材料。repo の `design/` が唯一の正典で、倉庫や参照元との継承はしない。
工場の Context Composer へ直接値を渡さない。

## コピーしないもの

ロゴ・商標・固有ビジュアル、画像・アイコン・フォント等の実アセット、既存サービスの画面・実装、
固有のコピー文言を写さない。フォント名の提案はフォントファイルの再配布許可を意味しない。
`allow_asset_copy: false` は実アセットを指す。数値トークンの提案は別の `allow_token_copy` で表す。
ライセンス記録とブランド値の承認は別の確認である。
