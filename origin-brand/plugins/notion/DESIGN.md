# DESIGN.md — Notion-inspired（抽象化・非公式）

> **ステータス: 下書き（抽象参照）。** awesome-design-md-jp の `notion/DESIGN.md`
> （https://github.com/kzhrknt/awesome-design-md-jp/blob/main/design-md/notion/DESIGN.md）を
> 参照元に、**雰囲気・密度・トーンのみを抽象化**した「Notion 風」の参考実装である。
> **Notion 公式ではない。ロゴ・商標・固有ビジュアル・実ブランドアセットは使用しない。**
> 参照元 DESIGN.md の権利は提供元に帰属する（`references.md` 参照、`allow_asset_copy: false`）。

## Visual Theme

- ミニマルで「道具的（tool-like）」。装飾を削ぎ、コンテンツとテキストを主役にする。
- ダークを主テーマとする落ち着いた低コントラスト背景の上に、明快なテキスト階層。
- 余白を広く取り、境界は控えめな細い罫線で示す。

## Color Roles（役割ベース。抽象値）

- Background: 深いニュートラルダーク基調。Surface はそれより一段明るいニュートラル。
- Text: 高不透明の白系を本文に、セカンダリはより低い不透明度で階層を作る（色のみに依存しない）。
- Accent: 単一のブルー系をリンク・フォーカス・インライン強調に限定使用。CTA はモノクロ（白面/ダーク面）主体。
- Semantic: success / warning / error / info を色 **と** 記号・テキストの併用で伝える。
- Border: 背景に溶ける極薄の境界線。過度な区切りを避ける。

## Typography

- Sans-serif グロテスク系（Inter 等）を主体に、日本語フォント（Noto Sans JP 等）へフォールバック。
- ディスプレイは大きく太く、本文は行間広めで可読性優先。数値は等幅・整列（tabular / lnum 相当）。
- **フォントファイルは同梱・再配布しない**。Web フォント参照または利用者環境に委ねる。

## Spacing / Layout

- 4〜64px の一定スケールで余白を統一。gutter は 24px 目安。最大幅は広めのコンテナ。
- 情報密度は中庸。カードは控えめな角丸とやわらかい影で軽く浮かせる。

## Components

- Button: primary は面（モノクロ）主体、secondary は透明背景＋細罫線。角丸は中庸。
- Input: Surface 背景＋極薄罫線、フォーカスで accent を明示。
- Card: Surface 背景・中庸の角丸・やわらかな影。

## Iconography

- 意味の明確な線アイコンを最小限。装飾目的では使わない。

## Motion

- 過度なアニメーションを避け、`prefers-reduced-motion` を尊重。

## Accessibility

- ダーク基調でも本文コントラストは WCAG AA（4.5:1 目安）を確保。フォーカス表示を明示。
- 色のみで情報を伝えない。キーボード操作・代替テキストを担保。

## Content Tone

- 簡潔でフラット。道具としての明快さを優先し、過剰な修飾を避ける。

## Do / Don't

- Do: ミニマル・広い余白・控えめな罫線・単一アクセント・非色依存・可読性。
- Don't: Notion のロゴ/商標/固有ビジュアルの再利用、公式サービスと誤認される表現、
  参照元からの具体資産の直接コピー、装飾過多。

## 参照の扱い

- awesome-design-md-jp からの**抽象参照のみ**（Context Composer 優先度6相当）。具体値はコピーせず、
  方針・密度・トーンを反映した独自トークンとして `tokens/` に格納する。
- ライセンスゲート未承認（`license_gate_approved: false`）。実ブランドアセットは出力に使わない。
