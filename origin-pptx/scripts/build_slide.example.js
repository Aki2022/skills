/**
 * ⚠️ v2時代の参照実装（2026-07-13注記）: 本例のキャンバス(720x405pt)・色(navy #4F4F70系)・
 * フォント(Noto Sans CJK JP)・eyebrow+主張文タイトルは旧v2仕様。v3では 960x540pt・
 * グレーグラデーション・Meiryo・tracker+1行キーメッセージが正（tokens.json v3 /
 * deck_helpers.js v3）。**構成・組み方の参考としてのみ読む**こと。
 */
/**
 * ============================================================================
 * origin-pptx REFERENCE IMPLEMENTATION (from the PoC) — read before adapting.
 * ============================================================================
 *
 * This is the proven "hybrid" PptxGenJS build script for slide
 * "導入までの3ステップ" (layout pattern B / 図解型), copied verbatim from the
 * PoC that validated the origin-pptx pipeline (see references/pipeline.md
 * step ③ and the design spec this skill was built from).
 *
 * What it demonstrates (use as a template — copy and adapt, don't rewrite
 * the underlying logic):
 *   - Native title + subtitle text, fully editable (Japanese stays native —
 *     never baked into a generated image).
 *   - Native card shapes (pptx.ShapeType.roundRect) with native label bars
 *     and native description text inside each card.
 *   - Embedded text-free icon images (icon_1.png / icon_2.png / icon_3.png,
 *     produced via Codex image_gen — see references/image_gen.md) placed
 *     into each card with slide.addImage(). Icons are the ONLY raster part
 *     of the slide; everything else is a native PPTX object.
 *   - Filled pptx.ShapeType.rightArrow connectors between cards (bolder and
 *     more visible than a thin line stroke — deliberate choice from the PoC).
 *   - 16:9 canvas defined via pptx.defineLayout({ width: 10, height: 5.625 })
 *     (720x405pt, 72pt = 1in).
 *   - fontFace: "Noto Sans CJK JP" for CJK-safe rendering (requires
 *     `brew install --cask font-noto-sans-cjk`; otherwise falls back to a
 *     substituted font and may render tofu/boxes for Japanese text).
 *
 * How to adapt for a new slide:
 *   1. Copy this file, don't edit it in place.
 *   2. Pull the color/font constants from style-guide/tokens.json (do not
 *      hardcode new values — keep them as named variables so the script
 *      tracks token changes).
 *   3. Re-derive x/y/w/h from style-guide/layout-grammar.md's % region table
 *      for the layout pattern you're using (A/B/C), converted to inches
 *      against the 10in x 5.625in canvas.
 *   4. Swap in your own outline.md text and icon image paths (icons must be
 *      generated text-free via image_gen; see references/image_gen.md for
 *      the save-path gotcha before assuming a generated icon file exists).
 *   5. Render-check with soffice + pdftoppm (references/pipeline.md step ④)
 *      before treating the output as final.
 *
 * Colors/typography from tokens.json. Canvas: 720x405pt (16:9) = 10in x 5.625in.
 */
const pptxgen = require("pptxgenjs");

// --- tokens.json values ---
const COLOR_PRIMARY = "1F3A5F";
const COLOR_ACCENT = "A67C2E";
const COLOR_TEXT_BODY = "1A1A1C";
const COLOR_TEXT_SECONDARY = "55555A";
const COLOR_ON_PRIMARY = "FFFFFF";
const COLOR_BG_DEFAULT = "FFFFFF";
const COLOR_CARD_BORDER = "1F3A5F";

const FONT_FACE = "Noto Sans CJK JP";

// --- canvas: 720pt x 405pt = 10in x 5.625in (72pt = 1in) ---
const CANVAS_W_PT = 720;
const CANVAS_H_PT = 405;
const CANVAS_W_IN = CANVAS_W_PT / 72; // 10
const CANVAS_H_IN = CANVAS_H_PT / 72; // 5.625

const pptx = new pptxgen();
pptx.defineLayout({
  name: "PPTX_V2_16x9",
  width: CANVAS_W_IN,
  height: CANVAS_H_IN,
});
pptx.layout = "PPTX_V2_16x9";

const slide = pptx.addSlide();
slide.background = { color: COLOR_BG_DEFAULT };

// --- Title: top of slide, navy bold, its own line ---
slide.addText("導入までの3ステップ", {
  x: 0.6,
  y: 0.32,
  w: 8.8,
  h: 0.6,
  fontFace: FONT_FACE,
  fontSize: 28,
  bold: true,
  color: COLOR_TEXT_BODY,
  align: "left",
  valign: "middle",
});

// --- Subtitle: clearly below the title (no overlap), smaller/gray ---
slide.addText("お申込みから最短5営業日でご利用開始いただけます", {
  x: 0.6,
  y: 0.92,
  w: 8.8,
  h: 0.4,
  fontFace: FONT_FACE,
  fontSize: 14,
  color: COLOR_TEXT_SECONDARY,
  align: "left",
  valign: "middle",
});

// --- Card layout ---
const CARD_Y = 1.55;
const CARD_W = 2.55;
const CARD_H = 3.55;
const CARD_GAP = 0.55; // gap between cards, where the arrow sits
const TOTAL_CARDS_W = CARD_W * 3 + CARD_GAP * 2;
const START_X = (CANVAS_W_IN - TOTAL_CARDS_W) / 2;

const cardXs = [
  START_X,
  START_X + CARD_W + CARD_GAP,
  START_X + 2 * (CARD_W + CARD_GAP),
];

const ICON_SIZE = 1.2; // inches, square
const ICON_Y = CARD_Y + 0.28;

const LABEL_H = 0.5;
const LABEL_Y = ICON_Y + ICON_SIZE + 0.22;

const DESC_Y = LABEL_Y + LABEL_H + 0.18;
const DESC_H = CARD_H - (DESC_Y - CARD_Y) - 0.15;

const steps = [
  {
    icon: "icon_1.png",
    heading: "申込み",
    desc: "オンラインで申込書を送信",
  },
  {
    icon: "icon_2.png",
    heading: "審査",
    desc: "書類確認と与信審査",
  },
  {
    icon: "icon_3.png",
    heading: "承認",
    desc: "契約締結とアカウント発行",
  },
];

steps.forEach((step, i) => {
  const x = cardXs[i];

  // Card: white rounded rectangle with thin navy border
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y: CARD_Y,
    w: CARD_W,
    h: CARD_H,
    rectRadius: 0.12,
    fill: { color: COLOR_BG_DEFAULT },
    line: { color: COLOR_CARD_BORDER, width: 1 },
  });

  // Icon image, centered horizontally within the card
  const iconX = x + (CARD_W - ICON_SIZE) / 2;
  slide.addImage({
    path: step.icon,
    x: iconX,
    y: ICON_Y,
    w: ICON_SIZE,
    h: ICON_SIZE,
  });

  // Navy label bar with step name, white bold, centered
  const labelW = CARD_W - 0.4;
  const labelX = x + 0.2;
  slide.addShape(pptx.ShapeType.roundRect, {
    x: labelX,
    y: LABEL_Y,
    w: labelW,
    h: LABEL_H,
    rectRadius: 0.06,
    fill: { color: COLOR_PRIMARY },
    line: { type: "none" },
  });
  slide.addText(step.heading, {
    x: labelX,
    y: LABEL_Y,
    w: labelW,
    h: LABEL_H,
    fontFace: FONT_FACE,
    fontSize: 20,
    bold: true,
    color: COLOR_ON_PRIMARY,
    align: "center",
    valign: "middle",
  });

  // Description text, dark gray
  slide.addText(step.desc, {
    x: x + 0.15,
    y: DESC_Y,
    w: CARD_W - 0.3,
    h: DESC_H,
    fontFace: FONT_FACE,
    fontSize: 13,
    color: COLOR_TEXT_SECONDARY,
    align: "center",
    valign: "top",
    lineSpacingMultiple: 1.3,
  });
});

// --- Gold filled right-arrows between cards ---
const ARROW_W = 0.42;
const ARROW_H = 0.34;
const arrowCenterY = ICON_Y + ICON_SIZE / 2; // align with icon vertical center

for (let i = 0; i < 2; i++) {
  const gapCenterX = cardXs[i] + CARD_W + CARD_GAP / 2 - ARROW_W / 2;
  slide.addShape(pptx.ShapeType.rightArrow, {
    x: gapCenterX,
    y: arrowCenterY - ARROW_H / 2,
    w: ARROW_W,
    h: ARROW_H,
    fill: { color: COLOR_ACCENT },
    line: { type: "none" },
  });
}

pptx.writeFile({ fileName: "output_hybrid.pptx" }).then(() => {
  console.log("Wrote output_hybrid.pptx");
});
