---
title: origin-pptx — Material Symbols アイコン標準
---

# アイコンは Material Symbols 標準（image_gen 生成は廃止・2026-08-27）

③ネイティブビルドの**セマンティック・ラインアイコン**（人物・書類・お金・警告等の意味アイコン）は、
image_gen で独自生成せず **Google Material Symbols**（Apache 2.0）を使う。
**image_gen のアセット生成はイラスト専用**（人物・いらすとや調・UIモック等「密度が品質に直結する」もの）。

**根拠（2026-08-27 kanpu_kaigo/20260831_nissay での A/B 対照実験・同一3スライド15アイコン）**:

| 指標              | Material Symbols                      | image_gen 従来法           |
| ----------------- | ------------------------------------- | -------------------------- |
| トークン/アイコン | ≈0（HTTP取得＋機械処理）              | ≈4.4万（参照画像込み実測） |
| 工程数・介入      | 2工程・介入0                          | 5工程・介入2回             |
| 正規化やり直し    | 0周（決定的）                         | 2周の実績（淡線・白箱）    |
| 決定性            | 決定的（再現可能）                    | 非決定的（毎回別物）       |
| 濃色セル          | 白fill直置きで可（白チップ不要）      | 白チップ必須               |
| 色差し替え        | fill書き換えのみ・数秒                | 不可（焼き込み）           |
| ④.5 人間判定      | 採用（wght300×opsz48 条件で品質同等） | 基準                       |

## 取得

- ソース: `google/material-design-icons`（GitHub raw）の `symbols/web/<symbol>/materialsymbolsoutlined/`
- **バリアント標準: `wght300 × opsz48`**（ファイル名 `<symbol>_wght300_48px.svg`）。
  既定の wght400×opsz24 はスライドの 60〜104pt 表示に拡大すると**太く黒く見える**ため不採用
  （2026-08-27 人間ラダー選定）。24px 系は使わない。
- `scripts/fetch_material_icons.py` が vocab_map.json から一括取得し、`process/icons_std/` に
  LICENSE（Apache 2.0）ごとキャッシュする。**ネットワーク取得はリポジトリ規約に従い人間確認を
  取ってから**（デッキごとに1回。数百KB）。ライセンス前例: origin-design-runtime の
  digital-agency プラグインで Apache 2.0 利用可を確認済み。

## vocab_map.json 規約

デッキの `process/vocab_map.json` に「意味 → symbol → ファイル」の対応を持つ:

```json
{
  "_meta": { "weight": "wght300×opsz48", "embed_path": "SVG直接 addImage" },
  "vocab": {
    "人物・世帯・利用者": { "symbol": "group", "file": "group_wght300_48px" },
    "金銭・保険料・還付": {
      "symbol": "payments",
      "file": "payments_wght300_48px"
    }
  }
}
```

- ビルダーは割当に無いアイコンのファイル名を**発明せず**「不足アセット: <意味の説明>」と報告する
  （image_gen 時代と同じ規律）。集約後に fetch を1回流す。
- **グリフ内の文字・記号のロケールを確認する**: `currency_exchange` は $ 入りで円建て資料と矛盾した
  （2026-08-27 実測）→ `currency_yen` へ。通貨・単位・文字を含む symbol は選定時に必ず目視する。

## 再着色と埋め込み

- 再着色: `scripts/recolor_svg.py <icons_dir>` — fill 書き換えで
  #404040（既定）/ #44546A（positive）/ #C00000（negative）/ #FFFFFF（濃色セル用）の4色を機械生成
  （冪等・数秒。image_gen 産に必要だった normalize の輝度→アルファ変換は不要）。
- 埋め込み: **SVG を `addImage` に直接渡す**（K1 実証 2026-08-26: sanitize --check・template_v3
  注入後・soffice レンダ・PowerPoint 実機すべて合格）。pptxgenjs は svgBlip＋フォールバックPNG
  （中身は SVG と同一バイトの偽PNG）を書くが、PowerPoint も LibreOffice も svgBlip を読む。
- **濃色セルには白 fill（`*_cFFFFFF.svg`）を直置き**する。image_gen 時代の「白チップを敷く」
  ハックは不要（gotchas §8 旧記述の置き換え）。
- **サイズは mockup の占有率に合わせる**: 既定の 60pt は mockup の半分以下だった実測あり。
  カード内占有率を mockup と見比べて 1.4〜2倍化する。
- フォールバック（svgBlip 非対応の配布先が出た場合のみ）: `scripts/svg_to_png.py` —
  soffice 黒レンダ→輝度→アルファ＋着色の透過PNG。**soffice の SVG→PNG は不透明白背景で出力される**
  （白アイコンが消える）ため、直接ラスタライズではなくこの経路を使う。

## 単一概念グリフの限界と対策（2026-08-27 実測・④で有効性確認済み）

Material Symbols は**1シンボル=1概念**。image_gen モックアップが描く合成モチーフ
（書類+AI文字・ZIP+クラウド+矢印・領収書+還付矢印）は1つの symbol で再現できない。対策の型:

1. **より意味の近い複合 symbol を探す**（例: lock → mail_lock）
2. **ネイティブ文字バッジで補完**（例: document_scanner ＋「AI」badge —— ④検証で有効確認済み）
3. **主役級・人物は image_gen イラストの役割のまま**（S14 中心円で実証: アイコン化すると
   「組織相関図」化して差し戻し → イラスト復帰で合格）

## ④ 検証への影響

- ペア比較プロンプトに**「アイコンは意味一致で判定（線の癖・作画スタイル差は不問）」**の但し書きを
  入れる。mockup のアイコンは image_gen の解釈画・build は Material Symbols で線が必ず違うため。
  但し書きの有効性は実測済み（線スタイル差の偽陽性ゼロ・検出されたのは真の意味不一致のみ）。
- icon-legibility チェックは継続（サイズ過小・濃色セルでの視認）。

## normalize_icons.py の残る役割

`scripts/normalize_icons.py` は **image_gen 産イラストの白背景透過（`--illustrations`）専用**になる。
ラインアイコンには使わない（Material Symbols に輝度→アルファ変換は不要かつ有害）。
