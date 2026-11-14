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
    if (realpathSync(file) !== file || iaDigest(file, modules.ia.resolved?.governance?.contract === 'ia-git-history-v3') !== modules.ia.implementation_sha256)
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
  for (const id of want) if (!id.startsWith('axe:')) results[id] = {checked:true, count:0, sample:[]};
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
  const shown = el => {
    for (let node = el; node; node = node.parentElement) {
      const s = getComputedStyle(node);
      if (s.display === 'none' || s.contentVisibility === 'hidden' || Number(s.opacity) === 0) return false;
    }
    return getComputedStyle(el).visibility === 'visible' && [...el.getClientRects()].some(r => r.width > 0 && r.height > 0);
  };
  const q = selector => [...document.querySelectorAll(selector)];
  const qs = selector => q(selector).filter(shown);
  const textOf = el => (el.innerText || el.textContent || '').replace(/\s+/g, ' ').trim();
  const directText = el => [...el.childNodes].some(node => {
    if (node.nodeType !== Node.TEXT_NODE || !node.textContent.trim()) return false;
    const range = document.createRange();
    range.selectNodeContents(node);
    return [...range.getClientRects()].some(rect => rect.width > 0 && rect.height > 0);
  });
  // Colour helpers: computed colours are rgb()/rgba() or color(srgb ...) in Chromium.
  const parse = value => {
    let m = /^rgba?\(([^)]+)\)$/.exec(value);
    if (m) { const p = m[1].split(/[\s,/]+/).filter(Boolean).map(Number); return [p[0], p[1], p[2], p.length > 3 ? p[3] : 1]; }
    m = /^color\(srgb ([^)]+)\)$/.exec(value);
    if (m) { const p = m[1].split(/[\s/]+/).filter(Boolean).map(Number); return [p[0] * 255, p[1] * 255, p[2] * 255, p.length > 3 ? p[3] : 1]; }
    return null;
  };
  const over = (top, under) => top[3] >= 1 ? top : [0, 1, 2].map(i => top[i] * top[3] + under[i] * (1 - top[3])).concat([1]);
  const luminance = c => { const f = v => { v /= 255; return v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4; };
    return 0.2126 * f(c[0]) + 0.7152 * f(c[1]) + 0.0722 * f(c[2]); };
  const ratio = (a, b) => { const [hi, lo] = [luminance(a), luminance(b)].sort((x, y) => y - x); return (hi + 0.05) / (lo + 0.05); };
  const backdrop = el => {
    const layers = [];
    for (let node = el; node; node = node.parentElement) {
      const c = parse(getComputedStyle(node).backgroundColor);
      if (c && c[3] > 0) { layers.push(c); if (c[3] >= 1) break; }
    }
    return layers.reverse().reduce((under, top) => over(top, under), [255, 255, 255, 1]);
  };
  const behind = el => el.parentElement ? backdrop(el.parentElement) : [255, 255, 255, 1];
  // dom: attribute and structure rules
  qs('input[placeholder]:not([placeholder=""]),textarea[placeholder]:not([placeholder=""])').forEach(el => add('dom:placeholder', el));
  qs('input[readonly],textarea[readonly]').forEach(el => add('dom:readonly-input', el));
  qs('input[disabled],textarea[disabled],select[disabled]').forEach(el => add('dom:disabled-input', el));
  qs('input[maxlength],textarea[maxlength]').forEach(el => add('dom:maxlength', el));
  qs('input[pattern]').forEach(el => add('dom:input-pattern', el));
  qs('ol').forEach(el => add('dom:ol', el));
  q('[tabindex]').filter(el => parseInt(el.getAttribute('tabindex'), 10) > 0).forEach(el => add('dom:positive-tabindex', el));
  qs('input[type="date"]').forEach(el => add('dom:date-input', el));
  qs('button[disabled],input[type="submit"][disabled],input[type="button"][disabled]').forEach(el => add('dom:disabled-button', el));
  q('td[colspan],td[rowspan],th[colspan],th[rowspan]').filter(el =>
    parseInt(el.getAttribute('colspan') || '1', 10) > 1 || parseInt(el.getAttribute('rowspan') || '1', 10) > 1)
    .forEach(el => add('dom:table-span', el));
  qs('table').filter(t => !t.querySelector('th')).forEach(el => add('dom:table-headers', el));
  qs('td').forEach(td => {
    const s = getComputedStyle(td);
    if (!['top', 'text-top'].includes(s.verticalAlign)) add('dom:table-align', td, 'vertical-align ' + s.verticalAlign);
    else if (/^[-+]?[¥$€]?[\d,]+(\.\d+)?%?$/.test(textOf(td)) && !['right', 'end'].includes(s.textAlign))
      add('dom:table-align', td, 'number aligned ' + s.textAlign);
  });
  q('[aria-expanded]').forEach(el => {
    if (el.tagName === 'SUMMARY') return;
    const role = el.getAttribute('role');
    if (el.tagName !== 'BUTTON' && role !== 'button') add('dom:disclosure-semantics', el, 'aria-expanded on non-button');
    else if (el.tagName !== 'BUTTON' && !el.hasAttribute('tabindex')) add('dom:disclosure-semantics', el, 'role=button not focusable');
  });
  const vague = /^(ここ|こちら|こちらから|ここをクリック|こちらをクリック|詳しくはこちら|click here|here)$/i;
  qs('a[href]').filter(el => vague.test(textOf(el))).forEach(el => add('dom:vague-link-text', el));
  qs('a[href]').filter(el => getComputedStyle(el).cursor !== 'pointer').forEach(el => add('dom:link-cursor', el, getComputedStyle(el).cursor));
  qs('button,[role="button"]').filter(el => el.scrollWidth > el.clientWidth + 1).forEach(el => add('dom:text-overflow', el));
  const textInputs = 'input:not([type]),input[type="text"],input[type="email"],input[type="tel"],input[type="url"],input[type="number"],input[type="password"],input[type="search"]';
  qs(textInputs + ',select').filter(el => el.getBoundingClientRect().height < 36).forEach(el =>
    add('dom:control-height', el, Math.round(el.getBoundingClientRect().height) + 'px'));
  qs(textInputs + ',textarea,select').forEach(el => {
    const label = (el.id && document.querySelector('label[for="' + CSS.escape(el.id) + '"]')) || el.closest('label');
    if (!label) return;
    const lr = label.getBoundingClientRect(), ir = el.getBoundingClientRect();
    if (label.contains(el)) { const text = [...label.childNodes].find(n => n.nodeType === Node.TEXT_NODE && n.textContent.trim());
      if (!text) return; const r = document.createRange(); r.selectNodeContents(text); const tr = r.getBoundingClientRect();
      if (tr.bottom > ir.top + 2) add('dom:label-above', el); return; }
    if (lr.bottom > ir.top + 2) add('dom:label-above', el);
  });
  qs(textInputs + ',textarea').forEach(el => {
    for (const type of ['paste', 'copy']) {
      const event = new ClipboardEvent(type, {bubbles:true, cancelable:true});
      el.dispatchEvent(event);
      if (event.defaultPrevented) { add('dom:paste-blocked', el, type + ' prevented'); break; }
    }
  });
  q('form [aria-live]:not([aria-live="off"]),form [role="alert"]').forEach(el => add('dom:form-live-region', el));
  const bullet = /^\s*([・•●■◆]|\d+[.．)）]|[①-⑳])/;
  qs('p,div,td,span').filter(el => !el.closest('li') && directText(el)).forEach(el => {
    const lines = (el.innerText || '').split('\n').filter(line => line.trim());
    if (lines.filter(line => bullet.test(line)).length >= 2) add('dom:fake-list', el);
  });
  const h1 = qs('h1');
  if (want.has('dom:h1-title') && (h1.length !== 1 || !document.title.includes(textOf(h1[0]))))
    add('dom:h1-title', h1[0] || document.documentElement, h1.length !== 1 ? h1.length + ' h1' : 'h1 not in title');
  qs('a[href],button,[role="button"]').forEach(el => {
    const graphic = el.querySelector('svg,img');
    if (!graphic) return;
    if (!textOf(el)) add('dom:icon-only', el);
    else el.querySelectorAll('svg,img').forEach(g => {
      const hidden = g.getAttribute('aria-hidden') === 'true' || g.getAttribute('role') === 'presentation'
        || (g.tagName === 'IMG' && g.getAttribute('alt') === '');
      if (!hidden) add('dom:icon-hidden', g);
    });
  });
  const newTab = /新規タブ|新しいタブ|別ウィンドウ|新しいウィンドウ|別タブ|new tab|new window/i;
  qs('a[target="_blank"]').forEach(el => {
    const names = [textOf(el), el.getAttribute('aria-label') || '', el.getAttribute('title') || '',
      ...[...el.querySelectorAll('img[alt],svg[aria-label],svg title')].map(n => n.getAttribute('alt') || n.getAttribute('aria-label') || n.textContent)];
    if (!names.some(n => newTab.test(n || ''))) add('dom:new-tab-indicator', el);
  });
  qs('p a[href],li a[href],td a[href],dd a[href]').forEach(el => {
    const block = el.parentElement;
    if (textOf(block) === textOf(el)) return;
    const s = getComputedStyle(el);
    if (s.textDecorationLine.includes('underline')) return;
    const a = parse(s.color), b = parse(getComputedStyle(block).color);
    if (!a || !b || ratio(a, b) < 3) add('dom:link-distinct', el);
  });
  const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    const text = node.textContent.trim();
    if (text && text.length <= 12 && /\S　+\S/.test(text) && node.parentElement && shown(node.parentElement))
      add('dom:space-alignment', node.parentElement, JSON.stringify(text));
  }
  qs('p,div,span,li,td,blockquote').filter(el => directText(el) && getComputedStyle(el).fontStyle === 'italic'
    && textOf(el).length > 100).forEach(el => add('dom:long-italic', el));
  if (want.has('dom:explicit-bg')) {
    const rootBg = [document.documentElement, document.body].map(el => parse(getComputedStyle(el).backgroundColor));
    if (!rootBg.some(c => c && c[3] > 0)) add('dom:explicit-bg', document.body, 'page background not set');
  }
  qs('dialog,[role="dialog"],[aria-modal="true"]').forEach(el => {
    const s = getComputedStyle(el);
    const bordered = ['Top', 'Right', 'Bottom', 'Left'].some(side => parseFloat(s['border' + side + 'Width']) > 0
      && !['none', 'hidden'].includes(s['border' + side + 'Style']) && (parse(s['border' + side + 'Color']) || [0, 0, 0, 0])[3] > 0);
    if (!bordered) add('dom:dialog-border', el);
  });
  q('input[type="file"]').forEach(el => {
    if (shown(el)) return;
    const label = el.id && document.querySelector('label[for="' + CSS.escape(el.id) + '"]');
    if (!(label && shown(label)) && !(el.closest('label') && shown(el.closest('label')))) add('dom:file-button', el);
  });
  qs('[aria-roledescription="carousel"]').forEach(el => {
    const slides = el.querySelectorAll('[aria-roledescription="slide"]').length;
    if (slides > 6) add('dom:carousel-slides', el, slides + ' slides');
  });
  const focusables = qs('a[href],button:not([disabled]),input:not([disabled]):not([type="hidden"]),select:not([disabled]),textarea:not([disabled]),[tabindex]:not([tabindex="-1"])');
  for (let i = 1; i < focusables.length; i++) {
    const a = focusables[i - 1].getBoundingClientRect(), b = focusables[i].getBoundingClientRect();
    if (getComputedStyle(focusables[i]).position === 'fixed' || getComputedStyle(focusables[i - 1]).position === 'fixed') continue;
    const overlap = Math.min(a.right, b.right) - Math.max(a.left, b.left);
    // The next stop sits entirely above the previous one in the same column: focus jumps back up.
    if (b.bottom <= a.top + 1 && b.top < a.top && overlap > 0) add('dom:focus-order', focusables[i], 'focus moves up after ' + describe(focusables[i - 1]));
  }
  // contrast of component boundaries (WCAG 1.4.11) and of icons
  qs('button,[role="button"],input:not([type="hidden"]),select,textarea').forEach(el => {
    const s = getComputedStyle(el);
    const border = ['Top', 'Right', 'Bottom', 'Left'].map(side => parseFloat(s['border' + side + 'Width']) >= 1
      && !['none', 'hidden'].includes(s['border' + side + 'Style']) ? parse(s['border' + side + 'Color']) : null).find(c => c && c[3] > 0);
    const fill = parse(s.backgroundColor);
    const around = behind(el);
    const edge = border ? over(border, around) : fill && fill[3] > 0 ? over(fill, around) : null;
    if (!edge) { if (el.tagName !== 'BUTTON' && el.getAttribute('role') !== 'button') add('contrast-nontext', el, 'no visible boundary'); return; }
    const r = ratio(edge, around);
    if (r < 3) add('contrast-nontext', el, r.toFixed(2) + ':1');
  });
  qs('svg').forEach(el => {
    if (el.getAttribute('aria-hidden') === 'true' && !el.closest('a,button,[role="button"]')) return;
    const s = getComputedStyle(el);
    const paint = s.fill && s.fill !== 'none' && !s.fill.startsWith('url') ? s.fill : s.stroke;
    const c = parse(paint || '');
    if (!c || c[3] === 0) return;
    const r = ratio(over(c, backdrop(el)), backdrop(el));
    if (r < 4.5) add('contrast-icon', el, r.toFixed(2) + ':1');
  });
  // values outside the design tokens
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
  // Split a computed box-shadow into its comma-separated layers, respecting parens in color functions.
  const shadowLayers = (value) => {
    const layers = [];
    let depth = 0, start = 0;
    for (let i = 0; i < value.length; i++) {
      if (value[i] === '(') depth++;
      else if (value[i] === ')') depth--;
      else if (value[i] === ',' && depth === 0) { layers.push(value.slice(start, i).trim()); start = i + 1; }
    }
    layers.push(value.slice(start).trim());
    return layers;
  };
  // A layer whose color has zero alpha paints nothing, whatever its offsets or size (e.g. the ring
  // layers Tailwind's multi-layer box-shadow carries alongside a real, token-matching shadow).
  const visibleShadow = (value) => {
    if (value === 'none') return value;
    const visible = shadowLayers(value).filter(layer => !/rgba\([^)]*,\s*0\s*\)/.test(layer));
    return visible.length ? visible.join(', ') : 'none';
  };
  for (const el of [...document.querySelectorAll('body *')]) {
    if (['SCRIPT', 'STYLE', 'TEMPLATE', 'LINK', 'META', 'NOSCRIPT'].includes(el.tagName) || !shown(el)) continue;
    const style = getComputedStyle(el);
    if (directText(el)) {
      if (!sizes.has(style.fontSize)) add('palette:fontSize', el, 'font-size ' + style.fontSize);
      if (!weights.has(style.fontWeight)) add('palette:fontWeight', el, 'font-weight ' + style.fontWeight);
      if (style.textAlign === 'justify') add('dom:text-justify', el);
    }
    for (const side of ['Top', 'Right', 'Bottom', 'Left'])
      if (!spaces.has(style['padding' + side])) add('palette:spacing', el, 'padding-' + side.toLowerCase() + ' ' + style['padding' + side]);
    if (/(flex|grid)/.test(style.display))
      for (const property of ['rowGap', 'columnGap'])
        if (style[property] !== 'normal' && !spaces.has(style[property])) add('palette:spacing', el, property + ' ' + style[property]);
    if (!shadows.has(style.boxShadow) && !shadows.has(visibleShadow(style.boxShadow))) add('palette:shadow', el, 'box-shadow');
  }
  // rules the repository declared for its own design decisions (design/checks), run on every rendered page
  const declared = Object.fromEntries((manifest.rules || []).map(rule => [rule.id, rule]));
  for (const id of want) {
    if (!id.startsWith('rule:')) continue;
    const rule = declared[id.slice(5)];
    const row = results[id];
    if (!rule) { row.checked = false; continue; }
    const within = (root, selectors) => selectors.flatMap(selector => [...root.querySelectorAll(selector)]).filter(shown);
    try {
      const scopes = rule.scope ? within(document, rule.scope) : [document];
      const outside = (n) => (rule.minimum !== undefined && n < rule.minimum) || n > (rule.maximum ?? Infinity);
      let applicable = 0;
      if (rule.kind === 'count') {
        applicable = rule.scope ? scopes.length : 1;
        for (const scope of scopes) {
          const n = within(scope, rule.targets).length;
          if (outside(n)) add(id, scope === document ? document.body : scope, 'count ' + n);
        }
      } else if (rule.kind === 'text') {
        const pattern = new RegExp(rule.pattern, 'gu');
        const targets = within(document, rule.targets);
        applicable = targets.length;
        const bound = rule.maximum ?? (rule.minimum === undefined ? 0 : Infinity);
        for (const el of targets) {
          const n = (textOf(el).match(pattern) || []).length;
          if ((rule.minimum !== undefined && n < rule.minimum) || n > bound) add(id, el, 'matches ' + n);
        }
      } else if (rule.kind === 'style') {
        const property = rule.property;
        const probeEl = document.createElement('div');
        probeEl.style.cssText = 'position:fixed;visibility:hidden;pointer-events:none';
        // Page-wide reduced-motion CSS shortens transitions instead of removing them; a transitioning probe would
        // report the previous allowlist value, so only the first value would ever be recognised.
        probeEl.style.setProperty('transition', 'none', 'important');
        probeEl.style.setProperty('animation', 'none', 'important');
        document.body.append(probeEl);
        const allowed = new Set((rule.resolved_allowlist || rule.allowlist).map(value => {
          probeEl.style.removeProperty(property);
          probeEl.style.setProperty(property, value);
          return probeEl.style.getPropertyValue(property) ? getComputedStyle(probeEl).getPropertyValue(property) : null;
        }).filter(Boolean));
        probeEl.remove();
        const targets = scopes.flatMap(scope => within(scope, rule.targets));
        applicable = targets.length;
        for (const el of targets) {
          if (rule.state === 'focus') el.focus({focusVisible: true});
          const value = getComputedStyle(el).getPropertyValue(property);
          if (rule.state === 'focus') el.blur();
          if (!allowed.has(value)) add(id, el, property + ' ' + value);
        }
      } else if (rule.kind === 'judgment') {
        applicable = within(document, rule.targets).length;
      }
      presence[id] = applicable;
    } catch (error) {
      // An invalid selector or pattern must fail the item, so the rule counts as applicable but unmeasured.
      presence[id] = 1;
      row.checked = false;
      row.sample = [String(error.message || error).slice(0, 120)];
    }
  }
  return {presence, results, uncovered:['margin (auto centering is indistinguishable after layout)', 'pseudo-elements', 'shadow-dom',
    'keyboard behaviour and state changes (no interaction is simulated)']};
}
