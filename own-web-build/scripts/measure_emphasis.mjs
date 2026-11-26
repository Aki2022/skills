// Measure the plan's emphasis on the drawn page: for each emphasis, the font size of the emphasized characters and of the
// rest of their target, inside the target's block, at each width.
// Usage: node measure_emphasis.mjs <html file> <emphasis JSON> <viewports JSON>
//   emphasis JSON: [{"block": "blk-…", "target": "cpy-…|value-…", "text": "…"}, …]
// Prints a JSON array, one row per emphasis and viewport. The target is the elements marked data-ia-copy / data-dict with
// its id inside [data-ia-section="<block>"]; characters are read whitespace-free and width-folded (NFKC), the way the copy
// and the dictionary values are compared. Playwright is found by own-design-author's playwright_runtime.mjs (never installed
// automatically; when it cannot be resolved this exits 2), the same way measure_screen.mjs does.
import { pathToFileURL } from "node:url";
import { loadPlaywright, pinnedVersion } from "../../own-design-author/scripts/playwright_runtime.mjs";
import { layerPath } from "../../own-design-author/scripts/layer_paths.mjs";

const projectRoot = process.env.SCREEN_PLAYWRIGHT_DIR || process.cwd();
let renderEnvironment;
try { renderEnvironment = layerPath(projectRoot, "render_environment"); }
catch (error) { console.error(error.message); process.exit(2); }
const [file, emphasisJson, viewportsJson] = process.argv.slice(2);
if (!file || !emphasisJson || !viewportsJson) {
  console.error("使い方: node measure_emphasis.mjs <html file> <emphasis JSON> <viewports JSON>");
  process.exit(2);
}
let playwright;
try { ({ playwright } = loadPlaywright({ version: pinnedVersion(renderEnvironment) })); }
catch (error) { console.error(error.message); process.exit(2); }
const browser = await playwright.chromium.launch();
const out = [];
try {
  for (const vp of JSON.parse(viewportsJson)) {
    const context = await browser.newContext({ viewport: { width: vp.width, height: vp.height }, deviceScaleFactor: 1, reducedMotion: "reduce" });
    const page = await context.newPage();
    await page.goto(pathToFileURL(file).href, { waitUntil: "load" });
    await page.waitForLoadState("networkidle").catch(() => {});
    await page.evaluate(() => document.fonts && document.fonts.ready);
    const rows = await page.evaluate((specs) => {
      const shown = (el) => {
        for (let e = el; e && e.nodeType === 1; e = e.parentElement) {
          const cs = getComputedStyle(e);
          if (cs.display === "none" || cs.visibility === "hidden" || Number(cs.opacity) === 0) return false;
        }
        return true;
      };
      const fold = (s) => s.normalize("NFKC").replace(/\s+/g, "");
      return specs.map((spec) => {
        const scope = `[data-ia-section="${spec.block}"]`;
        const els = [...document.querySelectorAll(`${scope} [data-ia-copy="${spec.target}"], ${scope} [data-dict="${spec.target}"]`)].filter(shown);
        const want = fold(spec.text);
        const each = els.map((el) => {
          const chars = [];
          const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
          for (let n = walker.nextNode(); n; n = walker.nextNode()) {
            if (!n.parentElement || !shown(n.parentElement)) continue;
            const size = parseFloat(getComputedStyle(n.parentElement).fontSize);
            for (const c of fold(n.textContent)) chars.push([c, size]);
          }
          const text = chars.map(([c]) => c).join("");
          const at = text.indexOf(want);
          if (at < 0) return { found: false };
          const inside = chars.slice(at, at + want.length).map(([, s]) => s);
          const rest = chars.filter((_, i) => i < at || i >= at + want.length).map(([, s]) => s);
          const inMin = Math.min(...inside), restMax = rest.length ? Math.max(...rest) : null;
          return { found: true, emphasized_px: inMin, rest_px: restMax, larger: restMax !== null && inMin > restMax };
        });
        return { ...spec, drawn: els.length, elements: each };
      });
    }, JSON.parse(emphasisJson));
    for (const row of rows) out.push({ viewport: vp.name, ...row });
    await context.close();
  }
} finally {
  await browser.close();
}
console.log(JSON.stringify(out));
