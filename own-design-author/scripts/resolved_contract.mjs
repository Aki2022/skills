/** Read an L0 output and measure computed membership; no DTCG resolution. */
import {readFileSync, readdirSync, realpathSync, statSync} from 'node:fs';
import {dirname, basename, join, relative, resolve, sep} from 'node:path';
import {createHash} from 'node:crypto';

export function loadManifest(path) {
  path = realpathSync(path);
  const design = dirname(dirname(path));
  if (basename(dirname(path)) !== '.resolved' || basename(design) !== 'design')
    throw new Error('manifest must be design/.resolved/<media>.json');
  const data = JSON.parse(readFileSync(path, 'utf8'));
  if (data.schema_version !== 1 || data.media !== 'web' || !Object.keys(data.tokens || {}).length)
    throw new Error('unsupported or empty manifest');
  if (data.approval_scheme !== 'hash-jcs-v1') throw new Error('unsupported approval scheme; rerun L0');
  const actual = {};
  function walk(dir) {
    for (const name of readdirSync(dir).sort()) {
      if (name === '.resolved') continue;
      const file = join(dir, name);
      const real = realpathSync(file);
      if (real !== file || !real.startsWith(design + sep))
        throw new Error('symlinked design input is unsupported');
      if (statSync(file).isDirectory()) walk(file);
      else if (!['assets.json','migration.json'].includes(name))
        actual[relative(design, file).split(sep).join('/')] = createHash('sha256').update(readFileSync(file)).digest('hex');
    }
  }
  walk(design);
  const keys = Object.keys(actual).sort();
  if (JSON.stringify(keys) !== JSON.stringify(Object.keys(data.input_hashes || {}).sort()) ||
      keys.some(k => actual[k] !== data.input_hashes[k])) throw new Error('stale manifest; rerun L0');
  for (const key of ['text_contrast','large_text_contrast','ui_contrast','target_size_px','focus_width_px','focus_contrast'])
    if (!Number.isFinite(data.minimums?.[key]) || data.minimums[key] <= 0)
      throw new Error('invalid manifest minimums');
  return data;
}

// This function is serialized into the browser. It must have no module dependencies.
export function collectDesign(manifest) {
  const probe = document.createElement('div');
  probe.style.cssText = 'position:fixed;visibility:hidden;pointer-events:none';
  document.body.append(probe);
  const normalize = (property, value) => {
    probe.style[property] = '';
    probe.style[property] = value;
    if (!probe.style[property]) throw new Error('invalid palette CSS value');
    return getComputedStyle(probe)[property];
  };
  const colors = new Set((manifest.palette.color || []).map(v => normalize('color', v)));
  const fonts = new Set((manifest.palette.fontFamily || []).map(v => normalize('fontFamily', v)));
  const radii = new Set((manifest.palette.dimension || []).map(v => normalize('borderTopLeftRadius', v)));
  probe.remove();
  const elements = [...document.querySelectorAll('body *')].filter(el =>
    !['SCRIPT','STYLE','TEMPLATE','LINK','META'].includes(el.tagName));
  const failures = [];
  for (const [index, el] of [document.body, ...elements].entries()) {
    const style = getComputedStyle(el);
    const check = (property, allowed) => {
      if (!allowed.has(style[property])) failures.push({index, tag:el.tagName.toLowerCase(), property, value:style[property]});
    };
    for (const property of ['color','backgroundColor','borderTopColor','borderRightColor','borderBottomColor','borderLeftColor','outlineColor','textDecorationColor'])
      check(property, colors);
    check('fontFamily', fonts);
    for (const property of ['borderTopLeftRadius','borderTopRightRadius','borderBottomLeftRadius','borderBottomRightRadius'])
      check(property, radii);
  }
  return {targetCount:elements.length, failures,
    uncovered:[...(manifest.coverage?.uncovered || []), 'pseudo-elements', 'shadow-dom', 'iframe-documents']};
}
