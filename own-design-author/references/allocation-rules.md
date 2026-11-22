# 割り付け規則（プロダクト UI）

infoarch（blueprint）は、領域の**優先度**・**グループ**・**中身の種類（shape）** だけを書く。この決まりが、そこから**位置と幅と部品**を導く。画面幅を変えても変わらないものは infoarch、変わるものはここ（design）。機械が読む形は `allocation-rules.json`。

## 規則

1. **縦の順は優先度の小さい順。** 同じ行に並べるときは、優先度の小さい方を左に置く。優先度の順は、画面幅を変えても保つ（PC でもスマホでも、優先度の高い領域が低い領域より下に来ない）。
2. **グループは連続して置く。** グループどうしの順は、そのグループの中でいちばん小さい優先度で決める。
3. **幅は shape で決まる。** 12 列のグリッド（間隔 16px、最大幅 1200px）で、`value`・`status` は 3 列（横に 4 つまで並べてよい）、`comparison`・`distribution`・`trend` は 6 列、`list`・`form`・`text` は 12 列。スマホ（639px 以下）は全部 1 列で、順は PC と同じ。
4. **部品は shape から選ぶ。** `value` は数値の表示、`status` はバッジ、`comparison`・`distribution`・`trend` はグラフ、`list` は表、`form` はフォーム、`text` は文章。表かカードか、どのグラフかは design が決める（infoarch は書かない）。
5. **状態は領域の中に出る。** 宣言した状態（loading・empty・error・partial・offline・denied）は、その領域の中の表示で、縦の順を変えない。
6. **主操作は 1 つだけ塗る。** `primary: true` の操作だけを塗りつぶしにし、ほかの操作は枠だけ・文字だけにする。

## 画面の型（決め方）

画面ごとに 1 つ選び、領域の置き方の基準にする。`dashboard`（読むだけ）・`detail`（1 つの対象の中身）・`list`（一覧が主役）・`form`（入力が主役）。迷ったら、応えるストーリーの動詞が「知りたい・把握したい」なら dashboard、「確かめたい」なら detail、「探したい」なら list、「登録したい」なら form。

## 規則に反する配置を人が指定するとき

規則では下になる領域を上に置くときは、`ux/design/layout-overrides.yaml` に **1 行ずつ** 書き、**オーナーが 1 行ずつ承認する**。理由を書くだけでは何も通らない（理由は誰でも書けるので、AI が理由を書くだけでどの配置も通ってしまうため）。

```yaml
- id: ovr-0001          # 行の名前。承認はこの名前に付く。2 行で同じ名前を使わない
  screen: scr-overview
  above: rgn-0006       # 規則では rgn-0002 より下
  below: rgn-0002
  viewports: [desktop]  # 許す幅。省けない（省くと全部の幅に広がるため）。desktop・tablet・mobile から選ぶ
  reason: 「部署別の一覧」を先に見たいと、オーナーが指定した（2026-11-20）
```

- **1 行が許すのは、書いた画面・領域の組・幅だけ。** ほかの画面、ほかの組、ほかの幅は、規則どおりに置かれていなければ失格のまま。
- **承認はその範囲ごとハッシュで縛る。** ハッシュは、行の中身（id・画面・組・幅・理由）と、行が指す 2 つの領域の blueprint の宣言を束ねる。行のどこか 1 か所を変えると、その行だけが未承認に戻る。指す領域の宣言が変わっても、領域や画面が消えても、承認は引き継がれない。行を足しても、ほかの行の承認は外れない。
- **写した行は引き継がない。** 行を写すと、同じ id（`OVERRIDE_DUPLICATE`）か同じ範囲（`OVERRIDE_DUPLICATE_SCOPE`）の行が 2 つになり、両方が失格になる。新しい id には承認が無い。
- **失格になる行**：承認の無い行（`OVERRIDE_UNAPPROVED`）、理由だけの行・画面か領域が無い行（`OVERRIDE_NO_TARGET`）、幅が無い・知らない幅・重複した幅の行（`OVERRIDE_NO_WIDTH`・`OVERRIDE_WIDTH_UNKNOWN`）、理由が無い行（`OVERRIDE_NO_REASON`）、画面や領域が無い行（`OVERRIDE_SCREEN_MISSING`・`OVERRIDE_REGION_MISSING`）、知らない欄のある行。`audit_screen.py` はその画面の承認の無い行と、ファイル全体の失格を失敗にする。承認のある行が許した逆転は、報告（info）に残る（合格にはなるが、消えない）。
- **承認の台帳の key** は `ds:web:override:<id>`（プロダクト単位の項目と同じ台帳 `ux/approvals.yaml`、同じ行の形）。`design-author check` は承認の無い行を警告にし、`--ship` で失格にする。must の項目と一緒には承認しない（`approve --prepare --must` に入らない。`--target ds:web:override:<id>` で 1 行ずつ指す）。
- **承認依頼の表示**は own-design-author が作る：`design-author override-request --repo <repo> --out <dir>`（1 行 1 枚のカード。○ と × を Artifact の db に保存する、プロダクト単位の項目と同じページ）。オーナーの答えは `--read-decisions` で読み戻す。台帳に書くのは、オーナーの承認の言葉と日付を受けた AI の `design-author approve` だけ。
