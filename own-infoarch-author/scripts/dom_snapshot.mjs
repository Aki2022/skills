// Snapshot of a rendered page for the information-architecture DOM checks.
// Usage: node dom_snapshot.mjs <html file> <viewports JSON>
// Prints a JSON array, one snapshot per viewport: anchors (data-ia-section / data-ia-copy / data-ia-text, and
// data-dict="<dictionary id>": a value-* id is reported as a `fact`, a name-* id as an `entity`, any other as a
// `term`; the v4/v5 attributes data-ia-problem / data-ia-msg / data-ia-pain / data-ia-frame are
// reported as legacy-<name>) with their rendered box,
// visibility, visible text (painted text nodes only, the same rule as `unanchored`) and whether a declared-text or
// official-name anchor holds another IA anchor (`nested`), its place in document order (`dom_index`), and for a block
// or copy anchor the CSS that draws it in another order than the DOM (`reorder`) and moves this cannot judge
// (`unverified`), the total rendered text length, the text visible in the
// first screen, and the visible text that belongs to no anchor (`unanchored`).
// Playwright is found by own-design-author's playwright_runtime.mjs: OWN_PLAYWRIGHT_DIR, else the project
// (INFOARCH_PLAYWRIGHT_DIR when set, else the current directory), else the per-machine runtime of the version
// pinned in the project's ux/design/render-environment.json. It is never installed automatically; when it cannot
// be resolved this exits 2.
import { pathToFileURL } from "node:url";
import { loadPlaywright, pinnedVersion } from "../../own-design-author/scripts/playwright_runtime.mjs";
import { layerPath } from "../../own-design-author/scripts/layer_paths.mjs";

const projectRoot = process.env.INFOARCH_PLAYWRIGHT_DIR || process.cwd();
const { playwright } = loadPlaywright({ version: pinnedVersion(layerPath(projectRoot, "render_environment")) });

const [file, viewportsJson] = process.argv.slice(2);
if (!file || !viewportsJson) {
  console.error("使い方: node dom_snapshot.mjs <html file> <viewports JSON>");
  process.exit(2);
}
const viewports = JSON.parse(viewportsJson);
const SETTLE_MS = 1000;            // timers that start reveals (setTimeout of ~100 ms is common)
const ANIMATION_TIMEOUT_MS = 5000; // upper bound for running transitions / animations to finish

const browser = await playwright.chromium.launch();
const out = [];
try {
  for (const vp of viewports) {
    const context = await browser.newContext({
      viewport: { width: vp.width, height: vp.height },
      deviceScaleFactor: 1,
      reducedMotion: "reduce",
    });
    const page = await context.newPage();
    await page.goto(pathToFileURL(file).href, { waitUntil: "load" });
    // Settle before reading: pages often reveal content with timers, transitions and CSS animations
    // (a hero shown after a short setTimeout, sections fading in). Reading at "load" would see them as hidden.
    await page.waitForLoadState("networkidle").catch(() => {});
    await page.waitForTimeout(SETTLE_MS);
    await page
      .waitForFunction(() => document.getAnimations().every((a) => a.playState !== "running" && a.playState !== "pending"),
        null, { timeout: ANIMATION_TIMEOUT_MS })
      .catch(() => {}); // endless animations: read anyway; the two-run stability check reports a moving page
    const snap = await page.evaluate((vh) => {
      const shown = (el) => {
        for (let e = el; e && e.nodeType === 1; e = e.parentElement) {
          const cs = getComputedStyle(e);
          if (cs.display === "none" || cs.visibility === "hidden" || Number(cs.opacity) === 0) return false;
        }
        return true;
      };
      const textPainted = (n) => {
        const range = document.createRange();
        range.selectNodeContents(n);
        const r = range.getBoundingClientRect();
        return r.width > 0 && r.height > 0;
      };
      // An element is visible when it is shown and either has a box or paints text: a zero-size box with
      // overflow: visible still paints its text, and must not escape the checks as "hidden".
      const visible = (el) => {
        if (!shown(el)) return false;
        const r = el.getBoundingClientRect();
        if (r.width > 0 && r.height > 0) return true;
        const w = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
        for (let n = w.nextNode(); n; n = w.nextNode()) {
          if (n.textContent.trim() && n.parentElement && shown(n.parentElement) && textPainted(n)) return true;
        }
        return false;
      };
      const paintedText = (n) => {
        const t = n.textContent.trim();
        return t && n.parentElement && shown(n.parentElement) && textPainted(n) ? t : "";
      };
      // Anchor text is what is painted, read the same way as the unanchored text: a hidden child cannot make an
      // anchor read as its copy (innerText keeps opacity: 0 text), and the two never disagree.
      const visibleText = (el) => {
        const out = [];
        const w = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
        for (let n = w.nextNode(); n; n = w.nextNode()) {
          const t = paintedText(n);
          if (t) out.push(t);
        }
        return out.join(" ");
      };
      // Declared text may show facts and official names; an official name holds nothing but itself.
      const NESTED = {
        copy: "[data-ia-section],[data-ia-copy],[data-ia-text]",
        text: "[data-ia-section],[data-ia-copy],[data-ia-text]",
        entity: "[data-ia-section],[data-ia-copy],[data-ia-text],[data-dict]",
      };
      // Order (2026-11-26): the order of the blocks and the copy is the document order (what a screen reader and a
      // keyboard follow), not the drawn position, so a row of cards in the script's order passes. CSS that draws IA
      // items in another order than the DOM is a failure (`reorder`); a move this cannot judge is `unverified`.
      const domIndex = new Map();
      let nth = 0;
      for (const el of document.querySelectorAll("*")) domIndex.set(el, nth++);
      const label = (el) => el.tagName.toLowerCase() + (el.id ? `#${el.id}` : "") + (el.classList[0] ? `.${el.classList[0]}` : "");
      const isFlex = (d) => d === "flex" || d === "inline-flex";
      const isGrid = (d) => d === "grid" || d === "inline-grid";
      const outOfFlow = (cs) => cs.position === "absolute" || cs.position === "fixed";
      const explicit = (v) => v !== "auto" && !v.startsWith("span");
      const placed = (cs) => [cs.gridRowStart, cs.gridColumnStart, cs.gridRowEnd, cs.gridColumnEnd].some(explicit);
      // b is drawn before a in the grid's reading order: rows first (columns first when the grid flows by column).
      // Boxes must be clearly apart (no overlap on the cross axis), so a few px of misalignment never counts.
      const drawnBefore = (b, a, byColumn) => {
        const rb = b.getBoundingClientRect(), ra = a.getBoundingClientRect();
        const [b0, b1, a0, a1, bs, ae] = byColumn
          ? [rb.left, rb.right, ra.left, ra.right, rb.bottom, ra.top]
          : [rb.top, rb.bottom, ra.top, ra.bottom, rb.right, ra.left];
        if (b1 <= a0) return true;
        if (a1 <= b0) return false;
        return bs <= ae;
      };
      // Walk from the anchor up to `stop` (exclusive): each element is an item of its parent. For a copy, `stop` is
      // its block (so the block's own container properties count, its place among blocks does not); for a block, the
      // walk goes up to the root.
      const orderOf = (el, stop) => {
        const reorder = [], unverified = [];
        if (getComputedStyle(el).direction === "rtl") unverified.push(`${label(el)}: direction: rtl`);
        for (let e = el; e && e !== stop && e.parentElement; e = e.parentElement) {
          const cs = getComputedStyle(e), p = e.parentElement, pcs = getComputedStyle(p);
          if (outOfFlow(cs)) unverified.push(`${label(e)}: position: ${cs.position}`);
          if (cs.float === "right" || cs.float === "inline-end") unverified.push(`${label(e)}: float: ${cs.float}`);
          // a transform that is not the identity (an animation that settled at translate(0) is) can move the box
          const moved = (cs.transform !== "none" && !new DOMMatrixReadOnly(cs.transform).isIdentity)
            || [cs.translate, cs.rotate].some((v) => v && v !== "none" && !/^0(px|deg)?( 0(px)?)*$/.test(v));
          if (moved) unverified.push(`${label(e)}: transform`);
          if (pcs.display === "contents") unverified.push(`${label(p)}: display: contents`);
          if (outOfFlow(cs) || !(isFlex(pcs.display) || isGrid(pcs.display))) continue;
          if (Number(cs.order) !== 0) reorder.push(`${label(e)}: order: ${cs.order}`);
          const items = [...p.children].filter((c) => getComputedStyle(c).display !== "none" && !outOfFlow(getComputedStyle(c)));
          if (items.length < 2) continue;
          if (isFlex(pcs.display) && pcs.flexDirection.endsWith("-reverse")) reorder.push(`${label(p)}: flex-direction: ${pcs.flexDirection}`);
          if (isFlex(pcs.display) && pcs.flexWrap === "wrap-reverse") reorder.push(`${label(p)}: flex-wrap: wrap-reverse`);
          if (isGrid(pcs.display) && pcs.gridAutoFlow.includes("dense")) reorder.push(`${label(p)}: grid-auto-flow: ${pcs.gridAutoFlow}`);
          if (isGrid(pcs.display) && items.some((c) => placed(getComputedStyle(c)))) {
            const byColumn = pcs.gridAutoFlow.startsWith("column");
            const earlier = items.filter((c) => domIndex.get(c) < domIndex.get(e));
            if (earlier.some((c) => drawnBefore(e, c, byColumn))) reorder.push(`${label(e)}: グリッドの明示の配置で DOM の前の項目より先に描かれる`);
          }
        }
        return { reorder: [...new Set(reorder)], unverified: [...new Set(unverified)] };
      };
      const anchors = [];
      const dictKind = (id) => ({ value: "fact", name: "entity" })[id.split("-")[0]] ?? "term";
      for (const [attr, base] of [["data-ia-section", "section"], ["data-ia-copy", "copy"], ["data-dict", null], ["data-ia-text", "text"],
        ["data-ia-problem", "legacy-problem"], ["data-ia-msg", "legacy-msg"], ["data-ia-pain", "legacy-pain"], ["data-ia-frame", "legacy-frame"]]) {
        for (const el of document.querySelectorAll(`[${attr}]`)) {
          const r = el.getBoundingClientRect();
          const kind = base ?? dictKind(el.getAttribute(attr));
          const block = el.parentElement?.closest("[data-ia-section]") ?? null;
          const ordered = kind === "section" ? orderOf(el, null) : kind === "copy" ? orderOf(el, block) : { reorder: [], unverified: [] };
          anchors.push({
            kind, id: el.getAttribute(attr),
            top: Math.round(r.top + scrollY), left: Math.round(r.left + scrollX),
            width: Math.round(r.width), height: Math.round(r.height),
            visible: visible(el), text: visibleText(el),
            nested: NESTED[kind] ? el.querySelector(NESTED[kind]) !== null : false,
            // the section a mark is drawn inside (screen conformance: a copy belongs to its own section)
            section: block?.getAttribute("data-ia-section") ?? null,
            dom_index: domIndex.get(el), reorder: ordered.reorder, unverified: ordered.unverified,
          });
        }
      }
      const parts = [], loose = [];
      // A section is a container: its text belongs to the copy and the declared text inside it, not to the section.
      const ANCHORED = "[data-ia-copy],[data-dict],[data-ia-text]";
      const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
      for (let n = walker.nextNode(); n; n = walker.nextNode()) {
        const t = paintedText(n);
        if (!t) continue;
        if (!n.parentElement.closest(ANCHORED)) loose.push(t);
        const range = document.createRange();
        range.selectNodeContents(n);
        const r = range.getBoundingClientRect();
        if (r.bottom > 0 && r.top < vh) parts.push(t);
      }
      return { anchors, text_chars: document.body.innerText.trim().length, first_screen: parts.join(" "), unanchored: loose.join(" ") };
    }, vp.height);
    out.push({ viewport: vp, ...snap });
    await context.close();
  }
} finally {
  await browser.close();
}
console.log(JSON.stringify(out));
