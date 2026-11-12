/** Read an L0 output and measure computed membership; no DTCG resolution. */
import {readFileSync, readdirSync, realpathSync, statSync, lstatSync} from 'node:fs';
import {dirname, basename, join, relative, resolve, sep} from 'node:path';
import {createHash} from 'node:crypto';
import {fileURLToPath} from 'node:url';

function iaDigest(file, full) {
  if (!full) return createHash('sha256').update(readFileSync(file)).digest('hex');
  const skill = dirname(dirname(file));
  const inputs = {};
  function walk(directory) {
    if (lstatSync(directory).isSymbolicLink() || realpathSync(directory) !== directory)
      throw new Error('IA trusted dependency symlink or escape');
    for (const name of readdirSync(directory)) {
      if (['tests', '__pycache__'].includes(name) || name.endsWith('.pyc')) continue;
      const path = join(directory, name);
      if (lstatSync(path).isSymbolicLink() || realpathSync(path) !== path)
        throw new Error('IA trusted dependency symlink or escape');
      if (statSync(path).isDirectory()) walk(path);
      else if (statSync(path).isFile())
        inputs[relative(skill, path).split(sep).join('/')] = createHash('sha256').update(readFileSync(path)).digest('hex');
    }
  }
  for (const sub of ['scripts', 'references']) walk(join(skill, sub));
  const keys = Object.keys(inputs).sort((a, b) => Buffer.compare(Buffer.from(a), Buffer.from(b)));
  if (!keys.length) throw new Error('zero IA trusted dependencies');
  return createHash('sha256').update(JSON.stringify(Object.fromEntries(keys.map(key => [key, inputs[key]])))).digest('hex');
}

export function loadManifest(path) {
  path = realpathSync(path);
  const design = dirname(dirname(path));
  if (basename(dirname(path)) !== '.resolved' || basename(design) !== 'design')
    throw new Error('manifest must be design/.resolved/<media>.json');
  const data = JSON.parse(readFileSync(path, 'utf8'));
  if (data.schema_version !== 1 || data.media !== 'web' || !Object.keys(data.tokens || {}).length)
    throw new Error('unsupported or empty manifest');
  if (data.approval_scheme !== 'hash-jcs-v1') throw new Error('unsupported approval scheme; rerun L0');
  const policyFile = fileURLToPath(new URL('../references/judgment-policy.json', import.meta.url));
  if (data.judgment_policy_sha256 !== createHash('sha256').update(readFileSync(policyFile)).digest('hex'))
    throw new Error('judgment policy changed; rerun L0');
  const modules = data.modules || {};
  if (Object.keys(modules).some(key => key !== 'ia')) throw new Error('unsupported resolved module');
  if (modules.ia) {
    const file = fileURLToPath(new URL('../../own-infoarch-author/scripts/ia_module.py', import.meta.url));
    if (realpathSync(file) !== file || iaDigest(file, modules.ia.resolved?.governance?.contract === 'ia-git-history-v2') !== modules.ia.implementation_sha256)
      throw new Error('IA resolver changed; rerun L0');
  }
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
  // Page-wide reduced-motion CSS must not animate the normalization probe.
  probe.style.setProperty('transition', 'none', 'important');
  probe.style.setProperty('animation', 'none', 'important');
  document.body.append(probe);
  const normalize = (property, value) => {
    probe.style[property] = '';
    probe.style[property] = value;
    if (!probe.style[property]) throw new Error('invalid palette CSS value');
    return getComputedStyle(probe)[property];
  };
  const colors = new Set((manifest.palette.color || []).map(v => normalize('color', v)));
  const fonts = new Set((manifest.palette.fontFamily || []).map(v => normalize('fontFamily', v)));
  // A DTCG dimension can be a signed transform/spacing value, not a radius.
  // Validate the length itself, then include only values CSS permits as radii.
  const radii = new Set((manifest.palette.dimension || []).flatMap(value => {
    normalize('left', value);
    return CSS.supports('border-top-left-radius', value)
      ? [normalize('borderTopLeftRadius', value)] : [];
  }));
  probe.remove();
  const elements = [...document.querySelectorAll('body *')].filter(el =>
    !['SCRIPT','STYLE','TEMPLATE','LINK','META'].includes(el.tagName));
  const failures = [];
  let targetCount = 0;
  const transparent = value => {
    const alpha = value.includes('/') ? value.split('/').at(-1)
      : value.startsWith('rgba(') && value.split(',').length === 4 ? value.split(',').at(-1) : null;
    return alpha !== null && parseFloat(alpha) === 0;
  };
  const directText = el => [...el.childNodes].some(node => {
    if (node.nodeType !== Node.TEXT_NODE || !node.textContent.trim()) return false;
    const range = document.createRange();
    range.selectNodeContents(node);
    return [...range.getClientRects()].some(rect => rect.width > 0 && rect.height > 0);
  });
  for (const [index, el] of [document.body, ...elements].entries()) {
    const style = getComputedStyle(el);
    // Unpainted defaults (notably border/outline currentColor on wrappers) do
    // not describe rendered design. Hidden ancestors must hide their children.
    let visible = style.visibility === 'visible';
    for (let node = el; visible && node; node = node.parentElement) {
      const ancestor = getComputedStyle(node);
      visible = ancestor.display !== 'none' && ancestor.contentVisibility !== 'hidden'
        && Number(ancestor.opacity) !== 0;
    }
    const text = directText(el);
    if (!visible || (!text && ![...el.getClientRects()].some(rect => rect.width > 0 && rect.height > 0))) continue;
    if (el !== document.body || text) targetCount++;
    const check = (property, allowed) => {
      if (!allowed.has(style[property])) failures.push({index, tag:el.tagName.toLowerCase(), property, value:style[property]});
    };
    const controlText = ['INPUT','TEXTAREA','SELECT'].includes(el.tagName)
      && (el.value || el.getAttribute('placeholder') || ['submit','reset'].includes(el.type));
    if ((text || controlText) && !transparent(style.color)) {
      check('color', colors);
      check('fontFamily', fonts);
    }
    if (!transparent(style.backgroundColor)) check('backgroundColor', colors);
    for (const side of ['Top','Right','Bottom','Left']) {
      if (!['none','hidden'].includes(style[`border${side}Style`]) && parseFloat(style[`border${side}Width`]) > 0
          && !transparent(style[`border${side}Color`])) check(`border${side}Color`, colors);
    }
    if (!['none','hidden'].includes(style.outlineStyle) && parseFloat(style.outlineWidth) > 0
        && !transparent(style.outlineColor)) check('outlineColor', colors);
    if (style.textDecorationLine !== 'none' && !transparent(style.textDecorationColor)) check('textDecorationColor', colors);
    for (const property of ['borderTopLeftRadius','borderTopRightRadius','borderBottomLeftRadius','borderBottomRightRadius'])
      check(property, radii);
  }
  return {targetCount, failures,
    uncovered:[...(manifest.coverage?.uncovered || []), 'pseudo-elements', 'shadow-dom', 'iframe-documents']};
}

// Serialized into the browser. Measures, without failing L2, the per-item code checks the manifest asks for
// and which component selectors are present; item verification decides per adopted item afterwards.
export function collectItemChecks(manifest) {
  const spec = manifest.item_checks || {applicability:{}, collect:[]};
  const presence = {};
  for (const [doc, selectors] of Object.entries(spec.applicability || {})) {
    let count = 0;
    for (const selector of selectors) {
      try { count += document.querySelectorAll(selector).length; }
      catch { throw new Error('invalid applicability selector: ' + selector); }
    }
    presence[doc] = count;
  }
  const want = new Set(spec.collect || []);
  const results = {};
  const describe = el => {
    const cls = String(el.getAttribute('class') || '').trim().split(/\s+/).filter(Boolean).slice(0, 2);
    return (el.tagName.toLowerCase() + (el.id ? '#' + el.id : '') + (cls.length ? '.' + cls.join('.') : '')).slice(0, 80);
  };
  const add = (id, el, detail) => {
    if (!want.has(id)) return;
    const row = results[id];
    row.count++;
    if (row.sample.length < 4) row.sample.push(detail ? describe(el) + ' ' + detail : describe(el));
  };
  for (const id of want) if (id.startsWith('dom:') || id.startsWith('palette:')) results[id] = {checked:true, count:0, sample:[]};
  const q = selector => [...document.querySelectorAll(selector)];
  q('input[placeholder]:not([placeholder=""]),textarea[placeholder]:not([placeholder=""])').forEach(el => add('dom:placeholder', el));
  q('input[readonly],textarea[readonly]').forEach(el => add('dom:readonly-input', el));
  q('input[disabled],textarea[disabled],select[disabled]').forEach(el => add('dom:disabled-input', el));
  q('input[maxlength],textarea[maxlength]').forEach(el => add('dom:maxlength', el));
  q('ol').forEach(el => add('dom:ol', el));
  q('[tabindex]').filter(el => parseInt(el.getAttribute('tabindex'), 10) > 0).forEach(el => add('dom:positive-tabindex', el));
  q('input[type="date"]').forEach(el => add('dom:date-input', el));
  q('td[colspan],td[rowspan],th[colspan],th[rowspan]').filter(el =>
    parseInt(el.getAttribute('colspan') || '1', 10) > 1 || parseInt(el.getAttribute('rowspan') || '1', 10) > 1)
    .forEach(el => add('dom:table-span', el));
  const vague = /^(ここ|こちら|こちらから|ここをクリック|こちらをクリック|詳しくはこちら|click here|here)$/i;
  q('a[href]').filter(el => vague.test((el.textContent || '').replace(/\s+/g, ' ').trim()))
    .forEach(el => add('dom:vague-link-text', el));
  const probe = document.createElement('div');
  probe.style.cssText = 'position:fixed;visibility:hidden;pointer-events:none';
  probe.style.setProperty('transition', 'none', 'important');
  document.body.append(probe);
  const normalize = (property, value) => {
    probe.style[property] = '';
    probe.style[property] = String(value);
    if (!probe.style[property]) return null;
    return getComputedStyle(probe)[property];
  };
  const set = (property, values, extra = []) =>
    new Set([...extra, ...values.map(v => normalize(property, v)).filter(v => v !== null)]);
  const dims = manifest.palette?.dimension || [];
  const sizes = set('fontSize', dims);
  const spaces = set('paddingTop', dims.filter(v => !String(v).trim().startsWith('-')), ['0px']);
  const weights = set('fontWeight', manifest.palette?.fontWeight || []);
  const shadows = set('boxShadow', manifest.palette?.shadow || [], ['none']);
  probe.remove();
  const directText = el => [...el.childNodes].some(node => {
    if (node.nodeType !== Node.TEXT_NODE || !node.textContent.trim()) return false;
    const range = document.createRange();
    range.selectNodeContents(node);
    return [...range.getClientRects()].some(rect => rect.width > 0 && rect.height > 0);
  });
  const shown = el => {
    for (let node = el; node; node = node.parentElement) {
      const s = getComputedStyle(node);
      if (s.display === 'none' || s.contentVisibility === 'hidden' || Number(s.opacity) === 0) return false;
    }
    return getComputedStyle(el).visibility === 'visible' && [...el.getClientRects()].some(r => r.width > 0 && r.height > 0);
  };
  for (const el of [...document.querySelectorAll('body *')]) {
    if (['SCRIPT','STYLE','TEMPLATE','LINK','META','NOSCRIPT'].includes(el.tagName) || !shown(el)) continue;
    const style = getComputedStyle(el);
    if (directText(el)) {
      if (!sizes.has(style.fontSize)) add('palette:fontSize', el, 'font-size ' + style.fontSize);
      if (!weights.has(style.fontWeight)) add('palette:fontWeight', el, 'font-weight ' + style.fontWeight);
      if (style.textAlign === 'justify') add('dom:text-justify', el);
    }
    for (const side of ['Top','Right','Bottom','Left'])
      if (!spaces.has(style['padding' + side])) add('palette:spacing', el, 'padding-' + side.toLowerCase() + ' ' + style['padding' + side]);
    if (/(flex|grid)/.test(style.display))
      for (const property of ['rowGap','columnGap'])
        if (style[property] !== 'normal' && !spaces.has(style[property])) add('palette:spacing', el, property + ' ' + style[property]);
    if (!shadows.has(style.boxShadow)) add('palette:shadow', el, 'box-shadow');
  }
  return {presence, results, uncovered:['margin (auto centering is indistinguishable after layout)', 'pseudo-elements', 'shadow-dom']};
}
