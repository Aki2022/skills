# 段 4 構造を決める

目的：段 3 で選んだメッセージと、ジャンルに要る中身を、媒体の器にまとまりとして並べる。まとまりは「どこで・何番目に言うか」で、LP の区画、資料のスライド、チラシの面にあたる。

## やること

1. 媒体のファイル（`media/` の該当ファイル）を読み、その媒体・ジャンルに要るまとまりを選ぶ。型は出発点で、依頼に合わせて変えてよい。
2. まとまりごとに `order`（何番目か）と `function`（内容の働き。辞書 `docs/dictionary/deliverables.json` の 7 つから 1 つ：identify 名乗り／claim 主張／fact 事実／evidence 証拠と声／offer 提供条件／action 行動／source 出典と運営者）を付ける。まとまりの数・順序は自由で、同じ働きを何回使ってもよい。説得の順序（v6 のセクションの役割 aware・pain・compare・trust・act）は持たない。
3. メッセージを置けるのは主張と行動のまとまりだけ（`places`）。主張のまとまりは通常のメッセージをちょうど 1 つ、行動のまとまりは CTA を 0〜1 つ置く。ほかの働きのまとまりはメッセージを置かない。一番のメッセージは必ずどれかの主張のまとまりに置く。
4. ナビ・節のラベルのように主張でない文言は、`texts` に安定した key と文言で宣言する（ナビはまとまりの外、ラベルは `block` で置くまとまりを指す）。宣言しない文言は、描画後に IA_DOM_UNANCHORED_TEXT（誰も管理していない文言）になる。主張になりうる文言は宣言せず、メッセージにする。
5. A/B の別案は、台本の要素（まとまりの順序、一番のメッセージ、文案など）を変えた別の台本として作る。`infoarch-author diff` が変わった要素を列挙し、`infoarch-author set` が set の id を出す。定数（target・接点・トーン）か目的変数が違えば、別の試験になる。

## よくある失敗

- 媒体の型を、依頼の目的を見ずにそのまま当てはめる。
- ラベルに社内の言葉（部署名・機能名の略称）を使う。読み手が探す語で書く。
- 主張のまとまりに通常のメッセージを 2 つ置く（IA_OUTLINE_SECTION_OVERLOADED）。事実・証拠と声のまとまりにメッセージを置く（IA_SCRIPT_MESSAGE_MISPLACED）。
- 申込み・声・事実のまとまりを、主張のまとまりに無理に詰める（働きの違うまとまりは分ける）。

## 台本への書き方

`blocks`（`order`・`function`・`places`）・`texts`。

## 出典

- 結論 → 根拠 → 証拠（ピラミッド原則）、導入の SCQA：Barbara Minto の Pyramid Principle の要約（https://thinkinsights.net/insights/minto-pyramid-principle）。原典は未確認。
- ラベルと構造の検証（ツリーテスト）：Nielsen Norman Group（https://www.nngroup.com/articles/tree-testing/）。
