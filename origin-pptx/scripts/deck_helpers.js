/**
 * deck_helpers.js — ③ネイティブビルドの共通ヘルパ集。
 *
 * v3 (2026-07-13): test3.pptx 実測仕様に全面改訂。
 *   - キャンバス: 960x540pt（33.87x19.05cm、PowerPoint標準ワイド画面）
 *   - グレーグラデーション体系（tokens.json v3 が正典）: 本文7F7F7F / 見出し404040 /
 *     パネルF2F2F2(枠線なし) / neutral 404040 / positive 44546A / negative C00000
 *   - フォント: 全デッキ一律 BIZ UDPGothic（2026-07-28 ユーザー決定・2プロファイル制は廃止。
 *     mk({font}) 必須＋ビルド後 set_fonts.py で全XML統一）
 *   - クローム層（日付・ページ番号・縦書きコピーライト）は defineSlideMaster に集約
 *   - chrome() は title=tracker（14pt灰・名詞句or疑問形）+ keyMessage（28pt・tone色・1行）
 *   - numUnit(): 数字と単位の50%ルール（例: 20pt数字 + 10pt単位）
 *   ※ v2（720x405pt・navy #4F4F70・eyebrow+主張文タイトル）とは非互換。
 *
 * 使い方: build_deck.js の冒頭で `const { mk } = require("./deck_helpers");` のように読み込み、
 *   const { pptx, ST, T, box, rct, arrow, homePlate, line, imgFit, chrome, numUnit,
 *           badge, source, sourceLinks, newSlide, C, IN } = mk({ date: "July 13, 2026", font: "BIZ UDPGothic" });
 *   （font必須。ビルド直後と inject後の最終ファイルに set_fonts.py を実行して全XMLを統一する）
 * を得て各スライドを組む。座標はすべて pt（960x540pt キャンバス、IN() で inch 変換）。
 */
const P = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

// tokens.json v3 準拠（正典は style-guide/tokens.json）
const C = {
  heading: "404040", // 見出し・強調・枠線/背景以外の着色全般
  body: "7F7F7F", // 本文・ラベル・タイトル(tracker)
  panel: "F2F2F2", // カード/オブジェクト背景（塗ったら枠線なし）
  footer: "D9D9D9", // フッター文字・薄い罫線
  copyright: "BFBFBF", // 縦書きコピーライト
  neutral: "404040",
  pos: "44546A",
  neg: "C00000",
  white: "FFFFFF",
  border: "D9D9D9",
};
const IN = (pt) => pt / 72;

// フォントは全デッキ一律 BIZ UDPGothic（tokens.json fontFamily・2026-07-28決定）。
// mk({font: "BIZ UDPGothic"}) で渡し、ビルド後に scripts/set_fonts.py で
// テーマ・テンプレ由来部品含む全XMLの a:latin/a:ea/a:cs を統一する。
function mk({ date = "", year = new Date().getFullYear(), font } = {}) {
  if (!font)
    throw new Error(
      'mk({font}) が未指定。標準フォント "BIZ UDPGothic" を渡す（tokens.json fontFamily が正典）',
    );
  const F = font;
  const pptx = new P();
  pptx.defineLayout({ name: "WIDE", width: 13.333, height: 7.5 });
  pptx.layout = "WIDE";
  const ST = pptx.ShapeType; // 落とし穴: ShapeType はインスタンス側。P.ShapeType は undefined

  // クローム層: コピーライトのみマスター（inject_template.py で注入後は template_v3 の
  // レイアウト側コピーライトに置き換わる想定＝スライド直描きだと二重になるためマスターに置く）。
  // 日付・ページ番号は newSlide() がスライド直描き（注入でレイアウトが差し替わっても消えない）。
  pptx.defineSlideMaster({
    title: "HOUSE",
    background: { color: C.white },
    objects: [
      {
        text: {
          text: `Copyright© ${year} オフィスオハナ合同会社 All Rights Reserved`,
          options: {
            x: IN(953),
            y: IN(268),
            w: IN(7),
            h: IN(272),
            fontFace: F,
            fontSize: 6,
            color: C.copyright,
            align: "left",
            valign: "bottom",
            vert: "vert270", // 右端縦書き
          },
        },
      },
    ],
    // ⚠️ pptxgenjs の slideNumber オプションは使わない: 不正 idx のプレースホルダを吐く
    // （修復トリガーとしては未確定だが、静的連番テキストの方が確実。gotchas §15）。
  });

  function T(s, text, x, y, w, h, o = {}) {
    s.addText(text, {
      x: IN(x),
      y: IN(y),
      w: IN(w),
      h: IN(h),
      fontFace: F,
      fontSize: o.size || 14,
      color: o.color || C.body,
      bold: !!o.bold,
      italic: !!o.italic,
      align: o.align || "left",
      valign: o.valign || "top",
      lineSpacingMultiple: o.lh || 1.3,
      ...(o.extra || {}),
    });
  }
  // 数字と単位の50%ルール: numUnit(s, "6,000", "件/年", x, y, w, h, {size:20})
  // → 数字20pt bold + 単位10pt bold を同一ボックス内ランで混在
  function numUnit(s, num, unit, x, y, w, h, o = {}) {
    const sz = o.size || 20;
    s.addText(
      [
        { text: String(num), options: { fontSize: sz, bold: true } },
        { text: String(unit), options: { fontSize: sz / 2, bold: true } },
      ],
      {
        x: IN(x),
        y: IN(y),
        w: IN(w),
        h: IN(h),
        fontFace: F,
        color: o.color || C.heading,
        align: o.align || "left",
        valign: o.valign || "middle",
      },
    );
  }
  // カード: panel塗り+枠線なしが既定（v3原則: 背景があるなら枠線不要）。
  // 枠線が要るのは fill:null（塗りなし）の注釈枠など明示時のみ。
  function box(s, x, y, w, h, o = {}) {
    const filled = o.fill !== null;
    s.addShape(ST.roundRect, {
      x: IN(x),
      y: IN(y),
      w: IN(w),
      h: IN(h),
      rectRadius: IN(o.r == null ? 13 : o.r),
      fill: filled ? { color: o.fill || C.panel } : { type: "none" },
      line: o.line ? { color: o.line, width: o.lw || 1 } : { type: "none" },
    });
  }
  function rct(s, x, y, w, h, o = {}) {
    const filled = o.fill !== null;
    s.addShape(ST.rect, {
      x: IN(x),
      y: IN(y),
      w: IN(w),
      h: IN(h),
      fill: filled ? { color: o.fill || C.panel } : { type: "none" },
      line: o.line ? { color: o.line, width: o.lw || 1 } : { type: "none" },
    });
  }
  const arrow = (s, x, y, w, h, color) =>
    s.addShape(ST.rightArrow, {
      x: IN(x),
      y: IN(y),
      w: IN(w),
      h: IN(h),
      fill: { color: color || C.heading },
      line: { type: "none" },
    });
  // 落とし穴: chevron は左の切れ込みが左寄せテキストを食う。工程ステップは homePlate 推奨。
  const chevron = (s, x, y, w, h, color) =>
    s.addShape(ST.chevron, {
      x: IN(x),
      y: IN(y),
      w: IN(w),
      h: IN(h),
      fill: { color: color || C.panel },
      line: { type: "none" },
    });
  const homePlate = (s, x, y, w, h, color) =>
    s.addShape(ST.homePlate, {
      x: IN(x),
      y: IN(y),
      w: IN(w),
      h: IN(h),
      fill: { color: color || C.panel },
      line: { type: "none" },
    });
  // 左上→右下以外の向き（上向き・左向き）は flipH/flipV で表現する。
  // 負の w/h をそのまま渡すと <a:ext> が負値になり OOXML違反 → PowerPointが修復を要求
  // （2026-07-13 実証。LibreOfficeは寛容なので④では検出できない。gotchas §17）
  const line = (s, x1, y1, x2, y2, o = {}) => {
    const w = x2 - x1;
    const h = y2 - y1;
    s.addShape(ST.line, {
      x: IN(Math.min(x1, x2)),
      y: IN(Math.min(y1, y2)),
      w: IN(Math.abs(w)),
      h: IN(Math.abs(h)),
      flipH: w < 0,
      flipV: h < 0,
      line: {
        color: o.color || C.border,
        width: o.w || 1,
        dashType: o.dash || "solid",
      },
    });
  };

  const missing = [];
  // path/aspect(w/h) を渡すとボックス内にアスペクト保持でフィット配置（歪み防止）。
  // v3: 枠線なしが既定（必要なら o.border に色を渡す）。
  function imgFit(s, imgPath, aspect, bx, by, bw, bh, o = {}) {
    if (!fs.existsSync(path.resolve(imgPath))) {
      missing.push(imgPath);
      return;
    }
    let w, h;
    if (bw / bh > aspect) {
      h = bh;
      w = bh * aspect;
    } else {
      w = bw;
      h = bw / aspect;
    }
    const x = bx + (bw - w) / 2,
      y = by + (bh - h) / 2;
    s.addImage({ path: imgPath, x: IN(x), y: IN(y), w: IN(w), h: IN(h) });
    if (o.border)
      s.addShape(ST.rect, {
        x: IN(x),
        y: IN(y),
        w: IN(w),
        h: IN(h),
        fill: { type: "none" },
        line: { color: o.border === true ? C.border : o.border, width: 1 },
      });
  }
  // v3クローム: title=tracker（名詞句or疑問形・14pt灰・目立たせない）+
  // keyMessage（28pt・tone色・必ず1行。タイトルが疑問形ならその回答）。
  // 塗りつぶしタイトルバンドは作らない。下段のtakeaway box・バンパーステートメントも禁止
  // （説明が要るときは右1/3の箇条書きセクション＝ContentText型を使う）。
  function chrome(s, { title, keyMessage, tone, size }) {
    if (title)
      T(s, title, 12, 18.4, 936, 21.6, { size: 14, color: C.body, lh: 0.9 });
    if (keyMessage) {
      // 1行厳守の自動縮小: 全角28ptは幅936ptに約33字まで。超える場合はフォントを絞る
      // （2026-07-14 実測: 34字で折返しが発生し3スライドがMAJOR差し戻しになった）
      const autoSize =
        keyMessage.length > 33
          ? Math.max(20, Math.floor(920 / keyMessage.length))
          : 28;
      T(s, keyMessage, 12, 39.7, 936, 64.9, {
        size: size || autoSize,
        color: tone === "neg" ? C.neg : tone === "pos" ? C.pos : C.neutral,
        valign: "middle",
        lh: 0.9,
        extra: { wrap: false },
      });
    }
  }
  // 出所表記: 本文最下端(y≤500)より下・フッター帯より上。色はフッターと同じ D9D9D9（2026-07-13確定）
  const source = (s, txt) =>
    T(s, txt, 42, 496, 850, 14, { size: 9, color: C.footer });
  // 出所（リンク付き）: sourceLinks(s, [{label, url}, ...])
  // 「出所：」＋各ラベルをPowerPoint上でクリック可能なハイパーリンクにする。
  // source() と同一座標・同一体裁。元資料のURLを出所として残す用途（2026-08-05追加）。
  function sourceLinks(s, items) {
    const runs = [{ text: "出所：", options: {} }];
    (items || []).forEach((it, i) => {
      if (i > 0) runs.push({ text: "／", options: {} });
      runs.push({
        text: it.label,
        options: { hyperlink: { url: it.url, tooltip: it.url } },
      });
    });
    s.addText(runs, {
      x: IN(42),
      y: IN(496),
      w: IN(850),
      h: IN(14),
      fontFace: F,
      fontSize: 9,
      color: C.footer,
      align: "left",
      valign: "top",
    });
  }
  // 吹き出し(尻尾つき)。描画順が肝: 本体→尻尾三角(枠線つき)→付け根を白矩形で開口。
  // 三角を先に描くと本体の枠線が付け根を横切り「分離した浮遊三角」に見える
  // (2026-07-11に2回失敗した実証。tailCenterYは尻尾中心のy、尻尾は左向き)。
  // v3: 注釈用途（白地+枠線）が既定。警告は line: C.neg を明示。
  function speechBubble(s, x, y, w, h, tailCenterY, o = {}) {
    const tailLen = o.tailLen || 20;
    const baseHalf = o.baseHalf || 12;
    const overlap = 4;
    const lc = o.line || C.heading;
    const lw = o.lw || 1.25;
    s.addShape(ST.roundRect, {
      x: IN(x),
      y: IN(y),
      w: IN(w),
      h: IN(h),
      rectRadius: IN(o.r || 11),
      fill: { color: o.fill || C.white },
      line: { color: lc, width: lw },
    });
    s.addShape(ST.triangle, {
      x: IN(x + overlap - baseHalf - tailLen / 2),
      y: IN(tailCenterY - tailLen / 2),
      w: IN(2 * baseHalf),
      h: IN(tailLen),
      rotate: 270,
      fill: { color: o.fill || C.white },
      line: { color: lc, width: lw },
    });
    s.addShape(ST.rect, {
      x: IN(x - 1),
      y: IN(tailCenterY - baseHalf + 3),
      w: IN(overlap + 2),
      h: IN(2 * baseHalf - 6),
      fill: { color: o.fill || C.white },
      line: { type: "none" },
    });
  }
  function badge(s, cx, cy, d, n, o = {}) {
    s.addShape(ST.ellipse, {
      x: IN(cx - d / 2),
      y: IN(cy - d / 2),
      w: IN(d),
      h: IN(d),
      fill: { color: o.fill || C.heading },
      line: { type: "none" },
    });
    T(s, String(n), cx - d / 2, cy - d / 2, d, d, {
      size: o.size || 13,
      bold: true,
      color: C.white,
      align: "center",
      valign: "middle",
    });
  }
  let pageNo = 0;
  function newSlide() {
    const s = pptx.addSlide({ masterName: "HOUSE" });
    pageNo += 1;
    // 日付・ページ番号はスライド直描き（テンプレ注入でレイアウトが差し替わっても保たれる）
    if (date)
      s.addText(date, {
        x: IN(21),
        y: IN(512.8),
        w: IN(224),
        h: IN(25.5),
        fontFace: F,
        fontSize: 12,
        color: C.footer,
        align: "left",
        valign: "bottom",
      });
    s.addText(String(pageNo), {
      x: IN(725.4),
      y: IN(512.8),
      w: IN(224),
      h: IN(25.5),
      fontFace: F,
      fontSize: 12,
      color: C.footer,
      align: "right",
      valign: "bottom",
    });
    return s;
  }

  return {
    pptx,
    ST,
    C,
    IN,
    T,
    numUnit,
    box,
    rct,
    arrow,
    chevron,
    homePlate,
    line,
    imgFit,
    chrome,
    source,
    sourceLinks,
    speechBubble,
    badge,
    newSlide,
    missing,
  };
}

module.exports = { mk, C, IN };
