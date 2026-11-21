// Measure a product UI screen in a real browser: where each marked region is drawn, which marked actions are drawn as
// primary, and which declared states exist. Usage: node measure_screen.mjs <html file> <viewports JSON>
// Prints a JSON array, one snapshot per viewport. Playwright is found by own-design-author's playwright_runtime.mjs (never
// installed automatically; when it cannot be resolved this exits 2), the same way own-infoarch-author's dom_snapshot.mjs does.
import { pathToFileURL } from "node:url";
import { loadPlaywright, pinnedVersion } from "../../own-design-author/scripts/playwright_runtime.mjs";
import { layerPath } from "../../own-design-author/scripts/layer_paths.mjs";

const projectRoot = process.env.SCREEN_PLAYWRIGHT_DIR || process.cwd();
// layerPath guards first: a repository that moved to a layout this version does not read stops here (exit 2, before any
// browser is looked for), instead of "no render environment pinned".
let renderEnvironment;
try { renderEnvironment = layerPath(projectRoot, "render_environment"); }
catch (error) { console.error(error.message); process.exit(2); }
const { playwright } = loadPlaywright({ version: pinnedVersion(renderEnvironment) });
const [file, viewportsJson] = process.argv.slice(2);
if (!file || !viewportsJson) {
  console.error("使い方: node measure_screen.mjs <html file> <viewports JSON>");
  process.exit(2);
}
const browser = await playwright.chromium.launch();
const out = [];
try {
  for (const vp of JSON.parse(viewportsJson)) {
    const context = await browser.newContext({ viewport: { width: vp.width, height: vp.height }, deviceScaleFactor: 1, reducedMotion: "reduce" });
    const page = await context.newPage();
    await page.goto(pathToFileURL(file).href, { waitUntil: "load" });
    await page.waitForLoadState("networkidle").catch(() => {});
    const snap = await page.evaluate(() => {
      const shown = (el) => {
        for (let e = el; e && e.nodeType === 1; e = e.parentElement) {
          const cs = getComputedStyle(e);
          if (cs.display === "none" || cs.visibility === "hidden" || Number(cs.opacity) === 0) return false;
        }
        return true;
      };
      const box = (el) => { const r = el.getBoundingClientRect(); return { top: Math.round(r.top + scrollY), left: Math.round(r.left + scrollX), width: Math.round(r.width), height: Math.round(r.height) }; };
      const alpha = (c) => { const m = c.match(/rgba?\(([^)]+)\)/); if (!m) return 0; const p = m[1].split(",").map(Number); return p.length > 3 ? p[3] : 1; };
      return {
        screens: [...document.querySelectorAll("[data-ia-screen]")].map((e) => e.getAttribute("data-ia-screen")),
        regions: [...document.querySelectorAll("[data-ia-region]")].map((e) => ({ id: e.getAttribute("data-ia-region"), priority: e.getAttribute("data-ia-priority"), visible: shown(e), ...box(e) })),
        actions: [...document.querySelectorAll("[data-ia-action]")].map((e) => {
          const cs = getComputedStyle(e);
          return { id: e.getAttribute("data-ia-action"), marked_primary: e.getAttribute("data-ia-primary") === "true", visible: shown(e), painted: alpha(cs.backgroundColor) > 0.5, ...box(e) };
        }),
        states: [...document.querySelectorAll("[data-ia-state]")].map((e) => ({ id: e.getAttribute("data-ia-state"), text: e.textContent.trim() })),
      };
    });
    out.push({ viewport: vp, ...snap });
    await context.close();
  }
} finally {
  await browser.close();
}
console.log(JSON.stringify(out));
