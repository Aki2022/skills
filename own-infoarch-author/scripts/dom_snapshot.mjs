// Snapshot of a rendered page for the information-architecture DOM checks.
// Usage: node dom_snapshot.mjs <html file> <viewports JSON>
// Prints a JSON array, one snapshot per viewport: anchors (data-ia-section / data-ia-copy / data-ia-text, and
// data-dict="<dictionary id>": a value-* id is reported as a `fact`, a name-* id as an `entity`, any other as a
// `term`; the v4/v5 attributes data-ia-problem / data-ia-msg / data-ia-pain / data-ia-frame are
// reported as legacy-<name>) with their rendered box,
// visibility, visible text (painted text nodes only, the same rule as `unanchored`) and whether a declared-text or
// official-name anchor holds another IA anchor (`nested`), the total rendered text length, the text visible in the
// first screen, and the visible text that belongs to no anchor (`unanchored`).
// Playwright is found by own-design-author's playwright_runtime.mjs: OWN_PLAYWRIGHT_DIR, else the project
// (INFOARCH_PLAYWRIGHT_DIR when set, else the current directory), else the per-machine runtime of the version
// pinned in the project's design/render-environment.json. It is never installed automatically; when it cannot
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
      const anchors = [];
      const dictKind = (id) => ({ value: "fact", name: "entity" })[id.split("-")[0]] ?? "term";
      for (const [attr, base] of [["data-ia-section", "section"], ["data-ia-copy", "copy"], ["data-dict", null], ["data-ia-text", "text"],
        ["data-ia-problem", "legacy-problem"], ["data-ia-msg", "legacy-msg"], ["data-ia-pain", "legacy-pain"], ["data-ia-frame", "legacy-frame"]]) {
        for (const el of document.querySelectorAll(`[${attr}]`)) {
          const r = el.getBoundingClientRect();
          const kind = base ?? dictKind(el.getAttribute(attr));
          anchors.push({
            kind, id: el.getAttribute(attr),
            top: Math.round(r.top + scrollY), left: Math.round(r.left + scrollX),
            width: Math.round(r.width), height: Math.round(r.height),
            visible: visible(el), text: visibleText(el),
            nested: NESTED[kind] ? el.querySelector(NESTED[kind]) !== null : false,
            // the section a mark is drawn inside (screen conformance: a copy belongs to its own section)
            section: el.parentElement?.closest("[data-ia-section]")?.getAttribute("data-ia-section") ?? null,
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
