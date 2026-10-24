// Snapshot of a rendered page for the information-architecture DOM checks.
// Usage: node dom_snapshot.mjs <html file> <viewports JSON>
// Prints a JSON array, one snapshot per viewport: anchors (data-ia-msg / data-ia-problem / data-ia-fact) with
// their rendered box, visibility and text, the total rendered text length, and the text visible in the first screen.
// Playwright is resolved from INFOARCH_PLAYWRIGHT_DIR, else from the current directory (the target repo).
// It is never installed automatically; when it cannot be resolved this exits 2.
import { createRequire } from "node:module";
import { pathToFileURL } from "node:url";
import { resolve } from "node:path";

const base = process.env.INFOARCH_PLAYWRIGHT_DIR || process.cwd();
let playwright;
try {
  // An explicit runtime root is authoritative; NODE_PATH must not silently replace it.
  // With no override, retain normal project/monorepo ancestor module resolution.
  playwright = createRequire(`${base}/`)(process.env.INFOARCH_PLAYWRIGHT_DIR
    ? resolve(base, "node_modules/playwright") : "playwright");
} catch {
  console.error(`playwright が解決できない（${base ? "INFOARCH_PLAYWRIGHT_DIR か実行ディレクトリ" : ""}）。自動では入れない。`);
  process.exit(2);
}

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
      const visible = (el) => {
        for (let e = el; e && e.nodeType === 1; e = e.parentElement) {
          const cs = getComputedStyle(e);
          if (cs.display === "none" || cs.visibility === "hidden" || Number(cs.opacity) === 0) return false;
        }
        const r = el.getBoundingClientRect();
        return r.width > 0 && r.height > 0;
      };
      const anchors = [];
      for (const [attr, kind] of [["data-ia-msg", "msg"], ["data-ia-problem", "problem"], ["data-ia-fact", "fact"]]) {
        for (const el of document.querySelectorAll(`[${attr}]`)) {
          const r = el.getBoundingClientRect();
          anchors.push({
            kind, id: el.getAttribute(attr),
            top: Math.round(r.top + scrollY), left: Math.round(r.left + scrollX),
            width: Math.round(r.width), height: Math.round(r.height),
            visible: visible(el), text: el.innerText,
          });
        }
      }
      const parts = [];
      const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
      for (let n = walker.nextNode(); n; n = walker.nextNode()) {
        const t = n.textContent.trim();
        if (!t || !n.parentElement || !visible(n.parentElement)) continue;
        const range = document.createRange();
        range.selectNodeContents(n);
        const r = range.getBoundingClientRect();
        if (r.bottom > 0 && r.top < vh) parts.push(t);
      }
      return { anchors, text_chars: document.body.innerText.trim().length, first_screen: parts.join(" ") };
    }, vp.height);
    out.push({ viewport: vp, ...snap });
    await context.close();
  }
} finally {
  await browser.close();
}
console.log(JSON.stringify(out));
