---
title: 既存デッキの編集モード（部分修正・スライド追加）
---

# 編集モード — 既存PPTXの部分修正・追加

ゼロベース生成（⓪〜⑤）ではなく、**既に存在するデッキの一部を直したい/足したい**場合の手順。
最初にやることは修正作業ではなく、**「正（source of truth）がどこにあるか」の判定**。
これを誤ると人間の手作業を消す事故になる（このスキル最大の破壊的失敗モード）。

## 0. 正の所在判定（必須・最初に実行）

```bash
# デッキの process/ にビルド資産があるか
ls <デッキdir>/process/slides/*.js <デッキdir>/process/build_deck.js
# 成果物pptxが「最後のビルド出力そのまま」か（= 人間が直接編集していないか）
cmp -s <デッキdir>/process/output.pptx <デッキdir>/<デッキ名>.pptx && echo SCRIPT-CANONICAL || echo HUMAN-EDITED
```

| 判定                                            | 正         | 使うルート  |
| ----------------------------------------------- | ---------- | ----------- |
| ビルド資産あり ＋ cmp一致                       | スクリプト | **ルートA** |
| cmp不一致（or output.pptx欠落なのに成果物あり） | 成果物pptx | **ルートB** |
| ビルド資産なし（他所で作られたpptx）            | 成果物pptx | **ルートB** |

判定結果と根拠を人間に一言報告してから作業に入る（「手を入れた版が正ですよね？」の確認を兼ねる）。

## ルートA: スクリプトが正 — 差分ビルド（最も効率的）

ゼロベースの③④をスライド単位に縮めたもの。**モックアップ再生成（②）は原則不要**。

1. 対象の `slides/sNN.js` だけを編集（文言・座標修正はメインループ直編集可、構図変更はビルダーサブエージェント）。
   新スライドは `sNN.js` を追加（entry が `slides/` を走査する構成なら追加だけで組み込まれる。
   ページ番号 `footer(N)` の繰り下がりに注意）
2. 再ビルド → 再レンダリング → **変更・追加したスライドのみ** ④のペア検証
   （構図を変えた場合は mockup が旧仕様になる——「preview↔mockup比較」ではなく
   「preview↔修正指示」の照合に切り替えるか、そのスライドだけ②を再実施する）
3. 成果物を上書きコピーする**前に必ず 0. の cmp ガードを再実行**（作業中に人間が触った可能性を潰す）

構図ごと変えたい1枚だけ②から回す（mockup 1枚生成→承認→③④）ことも可能で、これもデッキ全体の再走よりはるかに安い。

## ルートB: 人間が編集したpptxが正 — 外科的編集（再ビルド絶対禁止）

**`node build_deck.js` の出力で成果物を上書きしてはいけない**。手貼りスクショ・手修正テキストが全て消える。
原本は必ずコピーしてから作業する（`cp 成果物.pptx process/work.pptx`）。

### B-1. 小修正（文言・画像差し替え・スライド削除/並べ替え）→ python-pptx でその場編集

```python
from pptx import Presentation
from pptx.util import Emu
prs = Presentation("process/work.pptx")

# 文言修正（runを保てば書式維持）
for sh in prs.slides[8].shapes:
    if sh.has_text_frame:
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                if "旧文言" in r.text: r.text = r.text.replace("旧文言", "新文言")

# 画像差し替え（位置・サイズを保持して置換）
old = next(s for s in prs.slides[5].shapes if s.shape_type == 13)  # PICTURE
pos = (old.left, old.top, old.width, old.height)
old._element.getparent().remove(old._element)
prs.slides[5].shapes.add_picture("input/new.png", *pos)

prs.save("process/work.pptx")
```

- スライド削除/並べ替えは `xml_slides = prs.slides._sldIdLst` の要素操作（既知のイディオム）
- 編集後は該当スライドだけ soffice→pdftoppm でレンダリングし目視/VLM検証（④の縮小版）

### B-2. スライド追加 → 「1枚ミニデッキ」を別ファイルで納品し、人間がコピペで合流

python-pptx でのプレゼン間スライド複製は rels 追跡が壊れやすく**非推奨**。確実な手順:

1. 追加分だけを `deck_helpers.js` ＋ tokens で **1枚（〜数枚）のミニデッキ**としてビルド
   （スタイルは既存と自動で揃う。フッターのページ番号は合流先の番号を指定）
2. `additions.pptx` として成果物と並べて納品し、**PowerPoint上でのコピペ合流は人間に依頼**する
   （「デザインを保持する」貼り付けを案内）。これが最も安全で、人間の他の手修正とも干渉しない
3. どうしても自動合流が必要な場合のみ XML レベルの slide copy に踏み込む（工数と破損リスクを人間に説明して合意を取る）

### B-3. 事後処理 — スクリプト資産の扱いを明示する

ルートBを実施したら process/ のスクリプトは**旧仕様**になる。次のどちらかを必ず行う:

- 以後もスクリプト運用を続けたい → 人間の変更とB-1/B-2の変更を `sNN.js` に**バックポート**し、
  再ビルド結果が成果物と一致することを確認して SCRIPT-CANONICAL に戻す
- 単発対応で終わり → `process/STALE.md` に「YYYY-MM-DD以降は成果物pptxが正。再ビルド禁止」と記録
  （次のセッションの誰かが `build_deck.js` を無邪気に叩く事故の防止）

## 効率の目安

| シナリオ               | 再利用できるもの         | 走る工程                        |
| ---------------------- | ------------------------ | ------------------------------- |
| ルートA 文言・座標修正 | mockup/アセット/他13枚   | sNN.js編集→ビルド→該当1枚の検証 |
| ルートA 1枚構図変更    | アセット/他スライド      | ②1枚→③1枚→④1枚                  |
| ルートB 小修正         | 成果物そのもの           | python-pptx→該当1枚レンダ検証   |
| ルートB 追加           | style-guide/deck_helpers | ミニデッキ③④→人間がコピペ合流   |
