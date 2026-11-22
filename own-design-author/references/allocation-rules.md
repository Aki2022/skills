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

`ux/design/layout-overrides.yaml` に、規則では下になる領域を上に置く理由を書く。理由の無い逆転は `audit_screen.py` が失格にし、書いてあれば報告に残る（合格にはなるが、消えない）。

```yaml
- screen: scr-overview
  above: rgn-0006        # 規則では rgn-0002 より下
  below: rgn-0002
  reason: 「部署別の一覧」を先に見たいと、オーナーが指定した（2026-11-20）
  viewports: [desktop]
```
