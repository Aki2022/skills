---
title: Chart Rules
version: 2.0.0
---

# Chart Rules — オンブランドチャートの生成規則

会社チャート様式を本スキルのパイプラインで使える形に翻訳した規則。
v2 (2026-07-13): 系列色を tokens.json v3 のグレーグラデーション体系に更新
（グレー濃淡＋強調1系列のみ #44546A、ネガティブ強調 #C00000）。

> **大原則: 実データのチャートはネイティブに描く。imagegenで描かせない。**
> imagegenは数値・軸・比率を捏造する（全モデルで報告: IGenBench等）。
> **②のモックアップでも実データチャートは数値を描かない**——チャート領域は「空のラベル付き
> プレースホルダ枠」（例: 「チャート領域: 7系列折れ線 / 数値は③でoutlineから描画」）にする。
> **理由（2026-07-10 実証）**: ②で"それらしい"数値をimagegenに描かせたところ、人間レビューで
> 捏造数値をそのまま見て「数字が全く違う」とrejectされた。モックアップは構図の仕様であって
> 数値の仕様ではない。数値は常に outline.md が唯一の真実源（画像からの読み取り禁止）。

> **経路の選び方（実務優先度・2026-07-10 ユーザー要件で確定）**:
>
> 1. **主経路 = ネイティブOOXMLチャート＋ネイティブ・オーバーレイのハイブリッド**。
>    データ本体は `addChart`（環境があれば crtx）で描き、**「データの編集」でExcel編集できる状態を
>    必ず保つ**。チャート機能で表現できないカスタム装飾（系列末尾の国名/国旗ラベル・赤枠コール
>    アウト・軸外注記・リーダー線）は、**チャートに焼き込まずネイティブのテキスト/図形/小画像を
>    チャートの上・横に重ねて再現**する（§2.5の実装レシピ参照）。
> 2. **crtx直接適用（§2）**: テンプレ細部（系列別データラベル色等）まで要る場合の経路だったが、
>    依存していた旧`old-pptx`スキルは廃止済み。現状は主経路（1）または最終fallback（3）のみ。
> 3. **手書きSVG＋soffice（`scripts/build_chart_svg.py`）は最終fallback**。チャートが画像になり
>    **Excel編集可能性が失われるため、ユーザーが編集可能性を明示的に放棄した場合に限る**
>    （2026-07-10: SVG全面画像で一度納品→「Excelで編集できなければならない。国旗など非グラフ部分も
>    再現すること」と差し戻された実失敗）。使う場合も図として addImage し、文字はネイティブ別置き。

## 1. オンブランドチャートの定義（crtx/style.yaml から抽出）

| 要素               | 規則                                                                       | トークン                        |
| ------------------ | -------------------------------------------------------------------------- | ------------------------------- |
| 値軸（縦軸）       | **非表示**。数値はデータラベルで示す                                       | `chart.valueAxisVisible: false` |
| グリッド線         | **なし**（major/minorとも）                                                | `chart.gridlines: none`         |
| カテゴリ軸（横軸） | 表示。軸線0.75pt・ごく薄い色、目盛りなし                                   | `chart.categoryAxis`            |
| 軸ラベル           | 11pt・#7F7F7F                                                              | `chart.categoryAxis.fontSizePt` |
| 凡例               | **下配置**・11pt                                                           | `chart.legend`                  |
| データラベル       | 表示・11pt・#404040（系列末尾のみ非表示可）                                | `chart.dataLabels`              |
| チャートタイトル   | 14pt・**非bold**・#7F7F7F（通常はスライドタイトルで代替し省略）            | `chart.title`                   |
| 系列色             | **グレー濃淡戦略**: #404040 → #7F7F7F → #BFBFBF → #D9D9D9                  | `color.dataViz.monochrome`      |
| 強調               | 強調1系列のみ **#44546A**（ネガティブ強調は **#C00000**）、他はグレー濃淡  | `color.dataViz.accentPositive`  |
| 多系列（4+）       | それでもグレー濃淡で段階を付けるのが第一選択。判別不能な場合のみ人間に相談 | —                               |

## 2. 生成経路

上の「経路の選び方」で 1（ネイティブ＋オーバーレイ）→ 2（crtx）→ 3（SVG fallback）の順に検討する。

### 2.5 主経路の実装レシピ: `addChart` ＋ ネイティブ・オーバーレイ（2026-07-10 確立）

「Excelで編集できるデータ本体」と「チャート機能を超える装飾」を分離して両立する:

1. **チャート本体**: `slide.addChart` に**実データ（outline.md直書き）**を渡す。装飾要件は
   チャートオプションで寄せられるだけ寄せる:
   - 系列強調: `chartColors` で強調1系列のみ #44546A（ネガ文脈は #C00000）、他は同一グレー（#BFBFBF等）
   - 凡例の廃止: `showLegend: false`（代わりに末尾ラベルをオーバーレイ）
   - 軸クリップ（外れ値対策）: `valAxisMaxVal` を設定（超過系列はプロット領域で自然にクリップ）
   - **プロット領域の固定**: `layout: { x, y, w, h }`（チャート枠内の0-1比率）を明示すると、
     オーバーレイの座標計算が決定的になる（`endY = chartY + (ly + (1 - v/vmax) * lh) * chartH`）
2. **オーバーレイ（すべてネイティブ or 小画像）**: 系列末尾の値・国名ラベル＝`addText`、
   国旗等の小アイコン＝正規化済み小PNGを `addImage`、赤枠コールアウト＝`roundRect`＋text、
   リーダー線＝`line`、軸外注記＝`addText`。**チャート画像化は一切しない**
3. **検証**: ④のレンダリングでオーバーレイと線端の位置ズレを確認し、`layout` 比率か
   オーバーレイ座標を1回校正する。編集可能性の決定的チェックとして
   `unzip -l output.pptx | grep -E "charts/chart|embeddings/.*xlsx"` でOOXMLチャートと
   埋め込みワークブックの存在を確認する

> **crtx直接適用の位置付け（廃止）**: テンプレ細部（系列別データラベル色等）まで再現できたが、
> 依存していた `old-pptx`＋python-pptx＋crtxテンプレ一式は廃止済み。必要になった場合は
> git履歴（旧パス `pptx/`）から `crtx_utils.py` とテンプレを復元して再導入する。

### 最終fallback: 手書きSVG ＋ soffice（`scripts/build_chart_svg.py`）

**ユーザーがExcel編集可能性を明示的に放棄した場合のみ**。チャートは画像になる。
`build_chart_svg.py` を `process/` にコピーして改造（数値はoutline直書き）→
`soffice --headless --convert-to png` → 図として `addImage`。文字はネイティブ別置き。
編集元スクリプトを必ず版管理する。

### crtx直接適用（廃止）

デッキ本体はPptxGenJSで組み、チャートだけ後段のpython-pptx／crtxパスで挿入する2段ビルドが
かつて存在した。依存先の `old-pptx`（`crtx_utils.py`・`template.crtx` 一式）は廃止済みのため、
この経路は現在使えない。テンプレ細部（系列別データラベル色等）の再現が必要になった場合は、
git履歴（旧パス `pptx/`、コミット `e517f78`・`9c26b09`）から該当ファイルを復元して再導入する。

### 代替: PptxGenJS `addChart`（簡易チャート向け）

```js
slide.addChart(pptx.ChartType.bar, data, {
  x,
  y,
  w,
  h,
  chartColors: ["404040", "7F7F7F", "BFBFBF"], // gray-gradation series
  valAxisHidden: true, // 値軸非表示
  valGridLine: { style: "none" }, // グリッド線なし
  catGridLine: { style: "none" },
  catAxisLineColor: "D9D9D9", // 薄い軸線
  catAxisLabelColor: "7F7F7F",
  catAxisLabelFontSize: 11,
  showLegend: true,
  legendPos: "b", // 凡例下
  legendFontSize: 11,
  legendColor: "7F7F7F",
  showValue: true, // データラベル
  dataLabelColor: "404040",
  dataLabelFontSize: 11,
  showTitle: false,
  fontFace: F, // mk({font}) の値（ハードコードしない。set_fonts.pyが欧文/和文を最終調整）
});
```

- 棒/横棒/折れ線/円をおおむねオンブランド化できるが、**PoCで確認された翻訳漏れ**がある
  （系列別データラベル色は不可・カテゴリ間に目盛り線が残る）。細部が問われないチャートのみに使う
- どちらの経路でもデータはワークシートとして埋め込まれ、PowerPointで「データの編集」が可能

## 3. imagegen での「チャート風」表現（限定許可）

imagegenが許されるのはチャートの**装飾的・例示的表現**のみ:

- **許可**: モックアップ内のチャートの見た目仕様（後でネイティブに置換される前提）/
  大きな数字＋簡易バー等のKPIカード風ビジュアル / カテゴリ≤5の概念的なチャート風モチーフ
- **禁止**: 実データを正確に表現させること / 生成された軸・数値をそのまま成果物に使うこと
- プロンプトには「illustrative values, not real data」を明記し、検収では数値の正確性を評価しない

## 4. 表（table）について

表は**必ず**ネイティブで作る（`slide.addTable` / python-pptx）。**枠線・セルを図形や画像で
模造しない**（PowerPoint上で表として編集できることが要件・2026-07-13確定）。
スタイルは tokens.json v3 `table.*`: 罫線#D9D9D9（外枠1.0pt・内枠0.75pt）/
ヘッダー行=#F2F2F2塗り・#404040太字12pt / 本文=#7F7F7F 12pt・ゼブラ（#FFFFFF/#FAFAFA）。
