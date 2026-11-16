# 段 4 構造を決める

目的：段 3 で選んだメッセージを、媒体の器にセクションとして並べる。セクションは「どこで・何番目に言うか」で、LP のセクション、資料のスライド、チラシの面にあたる。

## やること

1. 媒体のファイル（`media/` の該当ファイル）を読み、その媒体の構成の型を選ぶ。型は出発点で、依頼に合わせて変えてよい。
2. セクションごとに `order`（何番目か）と `role`（aware 気づく／pain 困りごと／compare 比べる／trust 信じる／act 行動）を付ける。結論 → 根拠 → 証拠 → 行動 を基本にし、課題から入る型（困りごと → 効用 → 証拠 → 行動）は、読み手が自分の課題に気づいていない時に使う。
3. 各セクションに、通常のメッセージをちょうど 1 つと、CTA を 0〜1 つ置く（`places`）。同じメッセージを複数のセクションに置いてよい。一番のメッセージは必ずどこかに置く。
4. ナビ・節のラベルのように主張でない文言は、`texts` に安定した key と文言で宣言する（ナビはセクションの外、ラベルは `section` で置くセクションを指す）。宣言しない文言は、描画後に IA_DOM_UNANCHORED_TEXT（誰も管理していない文言）になる。主張になりうる文言は宣言せず、メッセージにする。
5. A/B の別案は、outline の要素（セクションの順序、一番のメッセージ、文案など）を変えた別の outline として作る。`infoarch-author diff` が変わった要素を列挙し、`infoarch-author set` が set の id を出す。定数（target・接点・トーン）か目的変数が違えば、別の試験になる。

## よくある失敗

- 媒体の型を、依頼の目的を見ずにそのまま当てはめる。
- ラベルに社内の言葉（部署名・機能名の略称）を使う。読み手が探す語で書く。
- 1 つのセクションに通常のメッセージを 2 つ置く（IA_OUTLINE_SECTION_OVERLOADED）。

## アウトラインへの書き方

`sections`（`order`・`role`・`places`）・`texts`。

## 出典

- 結論 → 根拠 → 証拠（ピラミッド原則）、導入の SCQA：Barbara Minto の Pyramid Principle の要約（https://thinkinsights.net/insights/minto-pyramid-principle）。原典は未確認。
- ラベルと構造の検証（ツリーテスト）：Nielsen Norman Group（https://www.nngroup.com/articles/tree-testing/）。
