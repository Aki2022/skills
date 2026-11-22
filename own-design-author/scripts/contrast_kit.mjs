// The one place a contrast ratio is computed (WCAG 2.x relative luminance and ratio). Everything that judges contrast
// after rendering uses it: the rendered measurement (own-web-build/scripts/measure_render.mjs), the pixel check inside
// it, and the component-boundary rules (resolved_contract.mjs). Nothing else may spell the formula; a test counts the files.
//
// Code that runs inside the page cannot import, so the kit is one factory: `contrastKit` is what Node code calls, and
// `kitSource` defines the same factory's result as window.__contrastKit in the page (inject it before evaluating a collector).
function factory() {
  const channel = (v) => {
    v /= 255;
    return v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4;
  };
  const luminance = (r, g, b) => 0.2126 * channel(r) + 0.7152 * channel(g) + 0.0722 * channel(b);
  const rgb = (c) => (Array.isArray(c) ? c : [c.r, c.g, c.b]);
  const ratioOfLuminances = (a, b) => {
    const [hi, lo] = [a, b].sort((x, y) => y - x);
    return (hi + 0.05) / (lo + 0.05);
  };
  const ratio = (a, b) => ratioOfLuminances(luminance(...rgb(a)), luminance(...rgb(b)));
  return { luminance, ratio, ratioOfLuminances };
}

export const contrastKit = factory();
// A statement, not an expression that returns the kit: evaluating it in the page must not try to send functions back.
export const kitSource = `window.__contrastKit = (${factory.toString()})(); void 0;`;
export const kitMissing = 'contrast kit not injected; evaluate kitSource in the page before the collector';
