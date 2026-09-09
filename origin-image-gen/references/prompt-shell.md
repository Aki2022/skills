---
title: プロンプトの器 — 10フィールドの固定順序とスタイルアンカー
---

# プロンプトの器

公式スキーマに従ったラベル付き短文ブロックの**固定順序**。長い一段落にしない。
この器は媒体をまたいで共有する — 中身の値だけが媒体とブランドで変わる。

## 順序と各フィールドの役割

| # | ラベル | 何を入れるか | 誰が値を持つか |
| --- | --- | --- | --- |
| 1 | `Use case` | `infographic-diagram` / `ui-mockup` / `productivity-visual` 等の分類 | 呼び出し側（工程の目的） |
| 2 | `Asset type` | `illustration` / `icon` / `background` / `full mockup` ＋ アスペクト比 | 呼び出し側 |
| 3 | `Style references` | スタイルアンカー画像。**edit-mode で渡す**（下記） | ブランド軸 |
| 4 | `Style` | スタイル記述を**逐語**貼付。セット内で一字も変えない | ブランド軸 |
| 5 | `House structure` | ハウススタイル構造（要素の役割分担） | ブランド軸 |
| 6 | `Composition` | 図解デバイスの指示。**内容と候補を渡し、選択はモデルに委ねる** | 呼び出し側 |
| 7 | `Content` | この1枚の固有の内容。日本語は引用符で括る | 呼び出し側 |
| 8 | `Color palette` | hex と言語表現の**両方**（hex だけだとモデルが外す） | ブランド軸 |
| 9 | `Constraints` | テキスト量・余白・アスペクト比等 | 媒体軸 |
| 10 | `Avoid` | 負の制約。**毎回必ず付ける** | 媒体軸 |

**1と2を先頭に置く理由**: 用途宣言でモデルの「モード」と仕上げ水準が決まる。後ろに置くと
一般的な仕上げになる。

**なぜ順序を検査で守るか**: この器は `origin-pptx/style-guide/imagegen-prompt-convention.md`
にも同じものがある（当面は二重に存在する。段階移行のため）。ラベルと順序がズレると
「同じ器を使っている」という前提が黙って崩れる。`scripts/check_shell_parity.py` が
**骨（ラベルと順序）だけ**を比較する。中身のズレは検出しない。

## 反復修正

**1ターン1変更**とし、**変えない要素を毎回明示的に再宣言する**。再宣言を省くとドリフトする
（前ターンで指定した要素が次のターンで勝手に変わる）。

## スタイルアンカーを edit-mode で渡す

文章でスタイルを説明するより、**実物の見本画像を参照として渡す**ほうが一致する。
アンカー画像は呼び出し側（ブランド倉庫）が持ち、本 skill はパスを受け取って器の
`Style references` に載せる。

プロンプト側の書き方:

```
First view style_ref images with view_image, then generate in edit-mode:
Image 1-N: style references.
Match their visual style exactly — same treatment, same annotation devices,
same density, same color roles. New content: <内容>
```

- **`view_image` を先に踏ませる**。踏ませないと参照画像を無視して生成することがある。
- アンカーは**人間が承認した実物**から採る。生成物をアンカーにすると劣化が累積する。
- **edit-mode は枚数がズレることがある**（参照画像が出力に混ざる等）。窓口は
  `--take-latest` 相当のフラグ1つだけでこれに手当てする。フラグを増やさない。

## 大量生成の一貫性

同じ `Style` / `House structure` / `Color palette` を**逐語**で再利用する。言い換えると
セット内で見た目が割れる。1枚目をスモークとして直列で通し、様式が合っていることを
確認してから残りを並列にする（references/execution.md の完了判定3原則）。
