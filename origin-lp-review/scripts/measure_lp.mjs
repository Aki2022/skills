#!/usr/bin/env node
/**
 * LP の審美・レイアウトを決定論的に実測する。
 *
 * 設計方針: ここに主観を入れない。「美しいか」は判定しない。
 * 判定するのは「意図した設計が実際の画面で成立しているか」だけで、すべて数値で出す。
 * 主観の採点は references/marketing-rubric.md 側（独立レビュアー）の仕事。
 *
 * 存在チェック（lang があるか、alt があるか）は既存の validator が持っている。
 * このスクリプトが足すのは **描画結果の実測** — 存在しても効いていない、
 * 通っているのに見えていない、を捕まえる層。
 *
 * 画素の評価は必ず **スクリーンショット**（合成後）に対して行う。
 * img を canvas に drawImage すると元画像を読むだけで、scrim・filter・重なりが
 * 反映されず「実際にどう見えているか」を測れない（v1 でこの誤りを踏んだ）。
 *
 * 使い方:
 *   node measure_lp.mjs <url> [--viewports 375,1280] [--schemes dark,light] [--json out.json]
 *
 * 終了コード: 0 = 失格なし / 1 = 失格あり / 2 = 実行不能
 */

import { writeFileSync } from "node:fs";
import { createRequire } from "node:module";

const require = createRequire(`${process.cwd()}/`);

function loadPlaywright() {
  try {
    return require("playwright");
  } catch {
    console.error(
      "playwright が解決できない。playwright を持つディレクトリで実行するか、" +
        "`npm i -D playwright && npx playwright install chromium` を実行する。",
    );
    process.exit(2);
  }
}

const args = process.argv.slice(2);
const url = args.find((a) => !a.startsWith("--"));
if (!url) {
  console.error("使い方: node measure_lp.mjs <url> [--viewports 375,1280]");
  process.exit(2);
}
const flag = (name, fallback) => {
  const i = args.indexOf(`--${name}`);
  return i === -1 ? fallback : args[i + 1];
};
const viewports = flag("viewports", "375,1280")
  .split(",")
  .map((v) => Number(v.trim()))
  .filter(Boolean);
const jsonOut = flag("json", null);
const schemes = flag("schemes", "dark,light").split(",").filter(Boolean);

/**
 * 閾値。根拠と、緩めた履歴は references/aesthetic-thresholds.md に残すこと。
 * 黙って緩めると検査が意味を失う。
 */
const T = {
  fontSizeKinds: 8,
  lineHeightKinds: 4,
  radiusKinds: 5,
  bodyLineHeightMin: 1.5,
  bodyFontSizeMin: 15, // 本文（最頻サイズ）
  textFontSizeFloor: 12, // 注記も含めた絶対下限
  measureMaxCh: 48, // 和文の全角基準
  contrastMin: 4.5,
  tapTargetMin: 44,
  orphanMinChars: 3,
  imageStandoutArea: 200, // 焦点画素の面積(CSS px²)。14x14 相当を下回れば絵が沈んでいる
  imageCropMax: 0.45, // object-fit: cover の切り取り率。意図した構図が壊れる境
};

// ---- ページ内で走る DOM/CSSOM 計測（画素は扱わない） ----
function collect(T) {
  const px = (v) => parseFloat(v) || 0;
  const vis = (el) => {
    const s = getComputedStyle(el);
    if (s.display === "none" || s.visibility === "hidden" || s.opacity === "0")
      return false;
    const r = el.getBoundingClientRect();
    return r.width > 0 && r.height > 0;
  };
  const all = [...document.querySelectorAll("body *")].filter(vis);

  const srgb = (c) => {
    const m = c.match(/[\d.]+/g);
    if (!m) return null;
    const [r, g, b, a = 1] = m.map(Number);
    return { r, g, b, a };
  };
  const lum = ({ r, g, b }) => {
    const f = (v) => {
      v /= 255;
      return v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4;
    };
    return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b);
  };
  const ratio = (a, b) => {
    const [x, y] = [lum(a), lum(b)].sort((m, n) => n - m);
    return (x + 0.05) / (y + 0.05);
  };
  // 透明を遡って合成する。透明のまま比べると常に緑になる。
  // 不透明な背景色を遡って探す。途中に背景画像・グラデ・重ねた <img> がある場合は
  // CSS だけでは合成結果が分からないので null を返し、Node 側の画素評価へ回す。
  const imgBoxes = [...document.querySelectorAll("img")]
    .filter(vis)
    .map((im) => ({ el: im, r: im.getBoundingClientRect() }));
  const overlapsImage = (el) => {
    const a = el.getBoundingClientRect();
    return imgBoxes.some(
      ({ el: im, r: b }) =>
        !el.contains(im) &&
        a.left < b.right &&
        a.right > b.left &&
        a.top < b.bottom &&
        a.bottom > b.top,
    );
  };
  const effectiveBg = (el) => {
    let node = el;
    while (node && node !== document.documentElement) {
      const st = getComputedStyle(node);
      // 自分（または祖先）が不透明な面を持つならそれが背景。背後に画像が
      // 重なっていても関係ない（ボタンなど）。この順序を逆にすると誤検知する。
      const c = srgb(st.backgroundColor);
      if (c && c.a > 0.95) return c;
      if (st.backgroundImage && st.backgroundImage !== "none") return null;
      if (overlapsImage(node)) return null;
      node = node.parentElement;
    }
    const c = srgb(getComputedStyle(document.body).backgroundColor);
    return c && c.a > 0.95 ? c : { r: 255, g: 255, b: 255, a: 1 };
  };
  const hasText = (el) =>
    [...el.childNodes].some(
      (n) => n.nodeType === 3 && n.textContent.trim().length > 0,
    );
  // 横スクロールできる祖先の中の要素は「はみ出し」ではない（意図した内部スクロール）
  const inScroller = (el) => {
    let node = el.parentElement;
    while (node && node !== document.body) {
      const ox = getComputedStyle(node).overflowX;
      if (ox === "auto" || ox === "scroll") return true;
      node = node.parentElement;
    }
    return false;
  };

  const kinds = {
    fontSize: new Set(),
    lineHeight: new Set(),
    radius: new Set(),
  };
  const spacingSeen = new Map();
  const spacingOutliers = [];
  const contrastFails = [];
  const pendingContrast = [];
  const tapFails = [];
  const overflow = [];
  const tooSmall = [];
  const bodySizes = [];
  let bodyLineHeightMin = Infinity;

  for (const el of all) {
    const s = getComputedStyle(el);
    kinds.fontSize.add(s.fontSize);
    if (px(s.borderTopLeftRadius) > 0) kinds.radius.add(s.borderTopLeftRadius);

    for (const p of [
      "marginTop",
      "marginBottom",
      "paddingTop",
      "paddingBottom",
      "paddingLeft",
      "paddingRight",
      "rowGap",
      "columnGap",
    ]) {
      const v = px(s[p]);
      if (v > 0) spacingSeen.set(v, (spacingSeen.get(v) || 0) + 1);
    }

    if (hasText(el)) {
      const size = px(s.fontSize);
      const lh = px(s.lineHeight) / size;
      if (Number.isFinite(lh) && lh > 0) kinds.lineHeight.add(lh.toFixed(2));
      if (size < T.textFontSizeFloor) {
        tooSmall.push({
          tag: el.tagName.toLowerCase(),
          text: el.textContent.trim().slice(0, 24),
          fontSize: size,
        });
      }
      // 本文＝長文を持つ p/li。最頻サイズを本文サイズとみなす（注記に引きずられない）
      if (["p", "li", "dd"].includes(el.tagName.toLowerCase())) {
        const chars = el.textContent.trim().length;
        if (chars > 40) {
          bodySizes.push(size);
          if (Number.isFinite(lh))
            bodyLineHeightMin = Math.min(bodyLineHeightMin, lh);
        }
      }

      const fg = srgb(s.color);
      const bg = effectiveBg(el);
      if (fg && !bg) {
        const rect = el.getBoundingClientRect();
        pendingContrast.push({
          tag: el.tagName.toLowerCase(),
          text: el.textContent.trim().slice(0, 28),
          fg: [fg.r, fg.g, fg.b],
          fontSize: size,
          bold: Number(s.fontWeight) >= 700,
          x: Math.round(rect.left + window.scrollX),
          y: Math.round(rect.top + window.scrollY),
          w: Math.round(rect.width),
          h: Math.round(rect.height),
        });
      }
      if (fg && bg) {
        const r = ratio(fg, bg);
        const bold = Number(s.fontWeight) >= 700;
        const large = size >= 24 || (size >= 18.66 && bold);
        const min = large ? 3 : T.contrastMin;
        if (r < min) {
          contrastFails.push({
            tag: el.tagName.toLowerCase(),
            text: el.textContent.trim().slice(0, 28),
            ratio: Number(r.toFixed(2)),
            min,
            fontSize: size,
          });
        }
      }
    }

    // タップ領域。文中のインラインリンクは WCAG 2.5.8 の例外なので除く。
    if (el.matches("a,button,[role=tab],[role=button],summary,input,select")) {
      const display = s.display;
      const inlineInText =
        el.tagName === "A" &&
        (display === "inline" || display === "inline-block") &&
        el.parentElement &&
        hasText(el.parentElement);
      if (!inlineInText) {
        const r = el.getBoundingClientRect();
        if (r.height < T.tapTargetMin || r.width < T.tapTargetMin) {
          tapFails.push({
            tag: el.tagName.toLowerCase(),
            label: (el.innerText || el.getAttribute("aria-label") || "").slice(
              0,
              24,
            ),
            w: Math.round(r.width),
            h: Math.round(r.height),
          });
        }
      }
    }

    const r = el.getBoundingClientRect();
    if (
      r.width > 0 &&
      (r.left < -1 || r.right > window.innerWidth + 1) &&
      !inScroller(el)
    ) {
      overflow.push({
        tag: el.tagName.toLowerCase(),
        cls: (el.className || "").toString().slice(0, 30),
        left: Math.round(r.left),
        right: Math.round(r.right),
      });
    }
  }

  // 余白のスケール外検出。
  // 旧実装は「出現回数上位 12 値をトークンとみなし、残りを外れ値」としていたが、
  // 使われる余白の種類がそもそも 12 以下のページでは **何も外れ値にならず、
  // 検査が常に空振りしていた**（実測: 13/27/41px を仕込んでも 0 件）。
  // 実際のトークン体系は「ある基準単位の倍数」なので、基準単位を推定して
  // 割り切れない値を外れ値にする。これなら種類数に依存しない。
  const spacingRanked = [...spacingSeen.entries()].sort((a, b) => b[1] - a[1]);
  const base = (() => {
    // 4 と 8 のうち、より多くの値を説明できる方を基準にする。
    // どちらも説明率が低いならトークン体系が無いとみなし、検査を見送る。
    let best = null;
    for (const unit of [8, 4]) {
      const hit = spacingRanked.filter(([v]) => Math.abs(v % unit) < 0.51 || Math.abs((v % unit) - unit) < 0.51);
      const ratio = spacingRanked.length ? hit.length / spacingRanked.length : 0;
      if (ratio >= 0.6 && (!best || ratio > best.ratio)) best = { unit, ratio };
    }
    return best;
  })();
  if (base) {
    for (const [v, n] of spacingRanked) {
      const m = v % base.unit;
      if (Math.min(m, base.unit - m) > 0.51) {
        spacingOutliers.push({ value: v, count: n, base: base.unit });
      }
    }
  }

  // 本文サイズは最頻値で見る
  const freq = new Map();
  for (const s of bodySizes) freq.set(s, (freq.get(s) || 0) + 1);
  const bodyFontSize =
    [...freq.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] ?? null;

  // 行長（全角基準の概算 ch）
  const measures = [...document.querySelectorAll("p,li,dd")]
    .filter((el) => vis(el) && el.textContent.trim().length > 80)
    .map((el) => {
      const s = getComputedStyle(el);
      return el.getBoundingClientRect().width / px(s.fontSize);
    });

  const heads = [...document.querySelectorAll("h1,h2,h3,h4,h5,h6")]
    .filter(vis)
    .map((h) => Number(h.tagName[1]));
  const headJumps = [];
  for (let i = 1; i < heads.length; i++) {
    if (heads[i] - heads[i - 1] > 1)
      headJumps.push(`h${heads[i - 1]} → h${heads[i]}`);
  }

  // フォントのサイレントフォールバック。宣言の第一候補が実際に使える状態か。
  const fontFallbacks = [];
  const seen = new Set();
  for (const el of all) {
    if (!hasText(el)) continue;
    const first = getComputedStyle(el)
      .fontFamily.split(",")[0]
      .trim()
      .replace(/^["']|["']$/g, "");
    if (
      !first ||
      seen.has(first) ||
      /^(sans-serif|serif|monospace|system-ui|ui-\w+|-apple-system)$/i.test(
        first,
      )
    )
      continue;
    seen.add(first);
    // document.fonts.check() は CJK webfont（unicode-range で ~120 分割）では
    // 当てにならない。必要な subset だけが loaded になるため、任意の文字での
    // 判定が false を返す。確実に判るのは「その family の FontFace が
    // 1 つも loaded でない」= まったく読めていない場合だけ。
    const faces = [...document.fonts].filter((f) => f.family === first);
    if (faces.length && !faces.some((f) => f.status === "loaded")) {
      fontFallbacks.push(first);
    }
  }

  // 孤立折り返し（見出しの最終行が数文字）
  // Range.getClientRects() は「視覚的な行」でなく「インラインボックス」を返す。
  // 見出しに <span> や <br> があると 2 行の見出しが 6 矩形になり、最後の矩形＝
  // 最終行と誤認する（実測: 最終行 20.3 文字を 3 文字と判定していた）。
  // top 座標でグルーピングすると視覚的な行に戻せる。行数は要素高さ÷行高でも
  // 検算し、食い違うときは判定を見送る（誤検知より未検証を選ぶ）。
  const orphans = [];
  for (const h of document.querySelectorAll("h1,h2,h3")) {
    if (!vis(h)) continue;
    const st = getComputedStyle(h);
    const lh = px(st.lineHeight);
    const fs = px(st.fontSize);
    if (!lh || !fs) continue;
    const range = document.createRange();
    range.selectNodeContents(h);
    const rects = [...range.getClientRects()].filter((r) => r.width > 1);
    if (!rects.length) continue;
    // 同じ行に属する矩形は top がほぼ一致する。行高の半分を許容幅にする。
    const rows = new Map();
    for (const r of rects) {
      const key = Math.round(r.top / (lh / 2));
      rows.set(key, (rows.get(key) || 0) + r.width);
    }
    const lines = [...rows.entries()].sort((a, b) => a[0] - b[0]);
    if (lines.length < 2) continue;
    const expected = Math.round(h.getBoundingClientRect().height / lh);
    if (expected !== lines.length) continue; // 数え方が一致しないなら判定しない
    const lastChars = lines[lines.length - 1][1] / fs;
    if (lastChars < T.orphanMinChars) {
      orphans.push({
        text: h.textContent.trim().slice(0, 30),
        lastLineChars: Number(lastChars.toFixed(1)),
        lines: lines.length,
      });
    }
  }

  // ファーストビュー
  const vh = window.innerHeight;
  const inFold = (el) => {
    if (!el) return null;
    const r = el.getBoundingClientRect();
    return r.top + window.scrollY < vh && r.bottom + window.scrollY <= vh;
  };
  const sticky = [...document.querySelectorAll("a,button")].some(
    (el) =>
      getComputedStyle(el).position === "fixed" ||
      (el.parentElement &&
        getComputedStyle(el.parentElement).position === "fixed"),
  );

  // 画像領域（画素評価は Node 側のスクリーンショットで行う）
  const imageRegions = [];
  for (const img of document.querySelectorAll("img")) {
    if (!vis(img) || !img.complete || !img.naturalWidth) continue;
    const r = img.getBoundingClientRect();
    if (r.width * r.height < 20000) continue;
    const s = getComputedStyle(img);
    const natural = img.naturalWidth / img.naturalHeight;
    const shown = r.width / r.height;
    const cropped =
      s.objectFit === "cover"
        ? 1 - Math.min(natural, shown) / Math.max(natural, shown)
        : 0;
    imageRegions.push({
      src: (img.currentSrc || img.src).slice(-40),
      x: Math.round(r.left + window.scrollX),
      y: Math.round(r.top + window.scrollY),
      w: Math.round(r.width),
      h: Math.round(r.height),
      objectFit: s.objectFit,
      croppedRatio: Number(cropped.toFixed(2)),
    });
  }

  return {
    kinds: {
      fontSize: kinds.fontSize.size,
      lineHeight: kinds.lineHeight.size,
      radius: kinds.radius.size,
      lineHeightValues: [...kinds.lineHeight].sort(),
    },
    spacingOutliers,
    contrastFails,
    pendingContrast,
    tapFails,
    overflow,
    tooSmall,
    horizontalScroll:
      document.documentElement.scrollWidth >
      document.documentElement.clientWidth,
    measureMaxCh: measures.length
      ? Number(Math.max(...measures).toFixed(1))
      : null,
    bodyFontSize,
    bodyLineHeight: Number.isFinite(bodyLineHeightMin)
      ? Number(bodyLineHeightMin.toFixed(2))
      : null,
    headJumps,
    fontFallbacks,
    orphans,
    fold: {
      viewportH: vh,
      h1Visible: inFold(document.querySelector("h1")),
      primaryCtaVisible: inFold(
        document.querySelector("[data-cta]:not([data-cta*=sticky])") ||
          document.querySelector("main a[href]"),
      ),
      hasStickyCta: sticky,
    },
    imageRegions,
  };
}

/** 合成後のスクリーンショットから、領域の知覚明度（L*）分位を測る */
async function measureRegion(page, region) {
  const shot = await page.screenshot({
    fullPage: true, // clip はページ座標。fullPage 無しだと折り返し下が撮れず undefined になる
    clip: { x: region.x, y: region.y, width: region.w, height: region.h },
  });
  const b64 = shot.toString("base64");
  return page.evaluate(async (b64) => {
    const img = new Image();
    img.src = `data:image/png;base64,${b64}`;
    await img.decode();
    const c = document.createElement("canvas");
    const w = (c.width = Math.min(320, img.width));
    const h = (c.height = Math.min(320, img.height));
    const ctx = c.getContext("2d", { willReadFrequently: true });
    ctx.drawImage(img, 0, 0, w, h);
    const d = ctx.getImageData(0, 0, w, h).data;
    const f = (v) => {
      v /= 255;
      return v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4;
    };
    const ls = [];
    for (let i = 0; i < d.length; i += 4) {
      const Y = 0.2126 * f(d[i]) + 0.7152 * f(d[i + 1]) + 0.0722 * f(d[i + 2]);
      // Y → L*（知覚明度 0..100）。線形輝度のまま比べると暗部の差を潰す。
      const t = Y > 0.008856 ? Math.cbrt(Y) : 7.787 * Y + 16 / 116;
      ls.push(116 * t - 16);
    }
    // 最頻値（5 刻みのヒストグラム peak）。文字が乗った狭い領域でも、
    // 面積で勝る背景側が peak を取る。中央値だと文字色を拾ってしまう。
    const bins = new Map();
    for (const l of ls) {
      const k = Math.round(l / 5) * 5;
      bins.set(k, (bins.get(k) || 0) + 1);
    }
    const mode = [...bins.entries()].sort((a, b) => b[1] - a[1])[0][0];
    // 焦点画素 = 支配的な明度から 20 以上離れた画素。灯りのような小さな焦点は
    // 全体の 1% 未満しか占めないので、分位ではなく「面積」で見る。
    const standout = ls.filter((l) => Math.abs(l - mode) > 20).length;
    ls.sort((a, b) => a - b);
    const q = (p) => Number(ls[Math.floor((ls.length - 1) * p)].toFixed(1));
    return {
      p5: q(0.05),
      p50: q(0.5),
      p95: q(0.95),
      mode,
      standoutRatio: Number((standout / ls.length).toFixed(4)),
    };
  }, b64);
}

const { chromium } = loadPlaywright();
const browser = await chromium.launch();
const report = {
  url,
  measuredAt: new Date().toISOString(),
  runs: [],
  failures: [],
};

for (const scheme of schemes) {
  for (const width of viewports) {
    const context = await browser.newContext({
      viewport: { width, height: width < 500 ? 812 : 900 },
      colorScheme: scheme,
      deviceScaleFactor: 1,
    });
    const page = await context.newPage();
    const consoleErrors = [];
    page.on(
      "console",
      (m) => m.type() === "error" && consoleErrors.push(m.text()),
    );
    page.on("pageerror", (e) => consoleErrors.push(String(e)));
    // console のテキストには URL が出ないため、favicon の 404 まで欠陥に見える。
    // 実レスポンスを見て、ページが本当に要求した資源だけを対象にする。
    const badRequests = [];
    const thirdPartyIssues = [];
    const pageOrigin = new URL(url).origin;
    page.on("response", (res) => {
      const u = res.url();
      if (res.status() < 400 || /\/favicon\.\w+$/.test(u)) return;
      const line = `${res.status()} ${u.slice(-70)}`;
      // 自ホストの資源だけを失格にする。第三者 CDN の失敗（headless 特有の
      // font subset 404 など）で常に赤くすると、本物の欠陥が埋もれる。
      if (u.startsWith(pageOrigin)) badRequests.push(line);
      else thirdPartyIssues.push(line);
    });

    await page.goto(url, { waitUntil: "networkidle", timeout: 60000 });
    // 遅延読み込みを確定させる（未読込を「壊れている」と誤判定しないため）
    await page.evaluate(() =>
      document.querySelectorAll("img").forEach((i) => (i.loading = "eager")),
    );
    await page.evaluate(() => document.fonts.ready);
    await page.waitForTimeout(700);

    const data = await page.evaluate(collect, T);
    // 404 等はテキストに URL が出ないので、favicon を除いた実リクエストで置き換える。
    data.consoleErrors = [
      ...consoleErrors.filter((t) => !/Failed to load resource/.test(t)),
      ...badRequests,
    ];
    data.thirdPartyIssues = thirdPartyIssues;

    // 画像・グラデの上に乗った文字は、背後の実画素からコントラストを出す。
    // CSS の背景色だけで判定すると body の地と比べてしまい誤検知になる。
    for (const t of data.pendingContrast) {
      try {
        const q = await measureRegion(page, t);
        const L = q.mode; // 背後の代表明度（最頻値）
        const Y = L > 8 ? ((L + 16) / 116) ** 3 : L / 903.3;
        const f = (v) => {
          v /= 255;
          return v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4;
        };
        const fgY =
          0.2126 * f(t.fg[0]) + 0.7152 * f(t.fg[1]) + 0.0722 * f(t.fg[2]);
        const [hi, lo] = [fgY, Y].sort((a, b) => b - a);
        const ratio = (hi + 0.05) / (lo + 0.05);
        const large = t.fontSize >= 24 || (t.fontSize >= 18.66 && t.bold);
        const min = large ? 3 : T.contrastMin;
        if (ratio < min) {
          data.contrastFails.push({
            tag: t.tag,
            text: t.text,
            ratio: Number(ratio.toFixed(2)),
            min,
            fontSize: t.fontSize,
            measuredBy: "pixel",
          });
        }
      } catch {
        data.contrastFails.push({
          tag: t.tag,
          text: t.text,
          skipped: "画素を取得できず未検証（緑にはしない）",
        });
      }
    }
    delete data.pendingContrast;

    // 合成後の画素で画像領域を評価する
    data.imageLightness = [];
    for (const region of data.imageRegions) {
      try {
        const q = await measureRegion(page, region);
        if (typeof q.standoutRatio !== "number") {
          throw new Error("standoutRatio が返っていない（計測の破損）");
        }
        // サンプル比から実寸の面積（CSS px²）へ戻す
        data.imageLightness.push({
          ...region,
          ...q,
          standoutArea: Math.round(q.standoutRatio * region.w * region.h),
        });
      } catch (e) {
        data.imageLightness.push({
          ...region,
          skipped: String(e).slice(0, 60),
        });
      }
    }
    report.runs.push({ scheme, width, ...data });
    await context.close();
  }
}
await browser.close();

const fail = (id, detail) => report.failures.push({ id, ...detail });
// 校正が足りない / 環境差が出る項目は警告に置く。失格に混ぜると
// 「常に赤い検査」になり、本物の欠陥が埋もれる。
report.warnings = [];
const warn = (id, detail) => report.warnings.push({ id, ...detail });
for (const r of report.runs) {
  const at = `${r.scheme}/${r.width}px`;
  if (r.kinds.fontSize > T.fontSizeKinds)
    fail("type-scale", { at, got: r.kinds.fontSize, max: T.fontSizeKinds });
  if (r.kinds.lineHeight > T.lineHeightKinds)
    fail("line-height-scale", {
      at,
      got: r.kinds.lineHeight,
      values: r.kinds.lineHeightValues,
      max: T.lineHeightKinds,
    });
  if (r.kinds.radius > T.radiusKinds)
    fail("radius-scale", { at, got: r.kinds.radius, max: T.radiusKinds });
  if (r.spacingOutliers.length)
    fail("spacing-off-scale", { at, outliers: r.spacingOutliers.slice(0, 6) });
  // CSS で背景が確定できたものは決定論。ここは失格にする。
  const cssContrast = r.contrastFails.filter((c) => c.measuredBy !== "pixel");
  if (cssContrast.length)
    fail("contrast", { at, count: cssContrast.length, sample: cssContrast.slice(0, 4) });
  // 画像の上の文字は背景の代表色を近似するしかなく、狭い箱では文字自身を拾う。
  // v1 では警告に留め、目視で確かめる（校正が済むまで失格にしない）。
  const pixelContrast = r.contrastFails.filter((c) => c.measuredBy === "pixel");
  if (pixelContrast.length)
    warn("contrast-over-image", {
      at,
      count: pixelContrast.length,
      sample: pixelContrast.slice(0, 4),
      note: "画像の上の文字。近似値なので目視で確認する",
    });
  if (r.tapFails.length)
    fail("tap-target", {
      at,
      count: r.tapFails.length,
      sample: r.tapFails.slice(0, 4),
    });
  if (r.tooSmall.length)
    fail("text-too-small", {
      at,
      floor: T.textFontSizeFloor,
      sample: r.tooSmall.slice(0, 3),
    });
  if (r.horizontalScroll) fail("horizontal-scroll", { at });
  if (r.overflow.length)
    fail("element-overflow", { at, sample: r.overflow.slice(0, 4) });
  if (r.measureMaxCh && r.measureMaxCh > T.measureMaxCh)
    fail("line-length", { at, got: r.measureMaxCh, max: T.measureMaxCh });
  if (r.bodyFontSize !== null && r.bodyFontSize < T.bodyFontSizeMin)
    fail("body-font-size", { at, got: r.bodyFontSize, min: T.bodyFontSizeMin });
  if (r.bodyLineHeight !== null && r.bodyLineHeight < T.bodyLineHeightMin)
    fail("body-line-height", {
      at,
      got: r.bodyLineHeight,
      min: T.bodyLineHeightMin,
    });
  if (r.headJumps.length) fail("heading-jump", { at, jumps: r.headJumps });
  if (r.fontFallbacks.length)
    // headless では webfont の subset 読み込みが実ブラウザと揺れる。
    // 実ブラウザで確認するまで失格にしない。
    warn("font-silent-fallback", { at, families: r.fontFallbacks });
  if (r.orphans.length)
    fail("heading-orphan", { at, sample: r.orphans.slice(0, 3) });
  if (r.consoleErrors.length)
    fail("console-error", { at, sample: r.consoleErrors.slice(0, 3) });
  if (r.fold.h1Visible === false) fail("fold-h1", { at });
  if (r.fold.primaryCtaVisible === false && !r.fold.hasStickyCta)
    fail("fold-cta", { at });
  for (const im of r.imageLightness) {
    if (im.skipped) {
      // 未検証を緑にしない。測れなかったこと自体を報告する。
      warn("image-unverified", { at, src: im.src, reason: im.skipped });
      continue;
    }
    if (typeof im.standoutArea !== "number") {
      fail("measurement-broken", {
        at,
        src: im.src,
        note: "standoutArea が数値でない",
      });
      continue;
    }
    if (im.standoutArea < T.imageStandoutArea)
      warn("image-sunk", {
        at,
        src: im.src,
        standoutArea: im.standoutArea,
        min: T.imageStandoutArea,
        mode: im.mode,
        note: "合成後の画像に、支配的な明度から離れた画素がほとんど無い。scrim や重なりで焦点が潰れ、絵が地に沈んでいる",
      });
    if (im.croppedRatio > T.imageCropMax)
      fail("image-over-crop", {
        at,
        src: im.src,
        croppedRatio: im.croppedRatio,
        max: T.imageCropMax,
        note: "object-fit: cover の切り取りが大きく、意図した構図が画面に出ていない",
      });
  }
}

const out = JSON.stringify(report, null, 2);
if (jsonOut) writeFileSync(jsonOut, out);

const byId = report.failures.reduce((acc, f) => {
  (acc[f.id] ||= []).push(f);
  return acc;
}, {});
console.log(`\n■ ${url}`);
console.log(
  `  計測: ${report.runs.map((r) => `${r.scheme}/${r.width}`).join(", ")}`,
);
if (!report.failures.length) {
  console.log("  失格なし\n");
} else {
  console.log(
    `  失格 ${report.failures.length} 件 / ${Object.keys(byId).length} 種\n`,
  );
  for (const [id, list] of Object.entries(byId)) {
    console.log(`  ✗ ${id} (${list.length})`);
    const s = JSON.stringify(list[0]);
    console.log(`    ${s.slice(0, 230)}${s.length > 230 ? "…" : ""}`);
  }
  console.log("");
}
if (report.warnings.length) {
  const wById = report.warnings.reduce((acc, w) => {
    (acc[w.id] ||= []).push(w);
    return acc;
  }, {});
  console.log(`  警告 ${report.warnings.length} 件（要目視・要校正。失格ではない）`);
  for (const [id, list] of Object.entries(wById)) {
    console.log(`  ! ${id} (${list.length}) ${JSON.stringify(list[0]).slice(0, 150)}`);
  }
  console.log("");
}
if (jsonOut) console.log(`  詳細: ${jsonOut}\n`);
process.exit(report.failures.length ? 1 : 0);
