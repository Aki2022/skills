/** Fixed Chromium environment. Uses existing runtime and repo fonts; installs nothing. */
import {readFileSync, realpathSync} from 'node:fs';
import {resolve, relative, sep} from 'node:path';
import {createHash} from 'node:crypto';
import os from 'node:os';

export const sha = data => createHash('sha256').update(data).digest('hex');
const canonical = value => Array.isArray(value) ? value.map(canonical) :
  value && typeof value === 'object' ? Object.fromEntries(Object.keys(value).sort().map(k=>[k,canonical(value[k])])) : value;
export const digest = value => sha(JSON.stringify(canonical(value)));
const exact = (value,keys) => value && typeof value==='object' && !Array.isArray(value) &&
  Object.keys(value).sort().join(',')===keys.split(',').sort().join(',');

function fingerprint(chromium,playwrightVersion,browser) {
  return {browser_version:browser.version(),browser_sha256:sha(readFileSync(chromium.executablePath())),
    playwright_version:playwrightVersion,node_version:process.version,
    platform:process.platform,arch:process.arch,os_release:os.release()};
}

export async function probe(chromium, playwrightVersion) {
  const browser=await chromium.launch({executablePath:chromium.executablePath()});
  try { return fingerprint(chromium,playwrightVersion,browser); }
  finally { await browser.close(); }
}

export function loadEnvironment(path,root,widths,schemes) {
  path=resolve(path);root=realpathSync(root);
  if(path!==resolve(root,'design/render-environment.json') || realpathSync(path)!==path)throw Error('invalid environment location');
  const bytes=readFileSync(path);const config=JSON.parse(bytes);
  if(!exact(config,'schema_version,runtime,clock,locale,timezone,device_scale_factor,reduced_motion,viewports,schemes,fonts') ||
    config.schema_version!==1 || !exact(config.runtime,'browser_version,browser_sha256,playwright_version,node_version,platform,arch,os_release') ||
    Object.values(config.runtime).some(v=>typeof v!=='string'||!v) || !/^[a-f0-9]{64}$/.test(config.runtime.browser_sha256) ||
    typeof config.clock!=='string' || new Date(config.clock).toISOString()!==config.clock ||
    typeof config.locale!=='string'||!config.locale || typeof config.timezone!=='string'||!config.timezone ||
    !Number.isFinite(config.device_scale_factor)||config.device_scale_factor<=0||config.device_scale_factor>4 ||
    config.reduced_motion!=='reduce'|| !Array.isArray(config.viewports)||!config.viewports.length ||
    config.viewports.some(v=>!exact(v,'width,height') || !Number.isInteger(v.width)||v.width<=0||!Number.isInteger(v.height)||v.height<=0)||
    config.viewports.map(v=>v.width).join(',')!==widths.join(',') || new Set(widths).size!==widths.length ||
    !Array.isArray(config.schemes)||config.schemes.join(',')!==schemes.join(',') ||
    !schemes.length||schemes.some(s=>!['light','dark'].includes(s))||new Set(schemes).size!==schemes.length ||
    !Array.isArray(config.fonts)||!config.fonts.length)throw Error('invalid fixed environment');
  const fonts=config.fonts.map(font=>{
    if(!exact(font,'path,url_path,family,sha256') || typeof font.path!=='string'||!font.path ||
      font.path.startsWith('/')||font.path.split('/').some(p=>!p||p==='.'||p==='..') ||
      typeof font.url_path!=='string'||!/^\/[A-Za-z0-9_./-]+$/.test(font.url_path) ||
      typeof font.family!=='string'||!font.family || !/^[a-f0-9]{64}$/.test(font.sha256))throw Error('invalid fixed font');
    const file=resolve(root,font.path),rel=relative(root,file);
    if(rel.startsWith('..'+sep)||rel==='..'||realpathSync(file)!==file || !/\.(woff2?|ttf|otf)$/.test(file))throw Error('invalid font location');
    const bytes=readFileSync(file);
    if(!bytes.length||sha(bytes)!==font.sha256)throw Error('fixed font hash mismatch');
    return {...font,bytes};
  });
  if(new Set(fonts.map(f=>f.url_path)).size!==fonts.length)throw Error('duplicate fixed font URL');
  return {config,fonts,config_sha256:sha(bytes),root,path};
}

export async function verifyRuntime(environment,chromium,playwrightVersion,browser) {
  const actual=fingerprint(chromium,playwrightVersion,browser);
  if(digest(actual)!==digest(environment.config.runtime))throw Error('fixed runtime mismatch');
  return {runtime:actual,config_sha256:environment.config_sha256,
    fonts:environment.config.fonts,clock:environment.config.clock,
    viewports:environment.config.viewports,schemes:environment.config.schemes,
    device_scale_factor:environment.config.device_scale_factor,reduced_motion:'reduce',
    locale:environment.config.locale,timezone:environment.config.timezone,repeat_equal:null};
}

export async function configureFonts(context,environment,url) {
  const origin=new URL(url).origin;const rejected=[];const used=new Set();
  await context.route('**/*',async route=>{
    const request=route.request();const target=new URL(request.url());
    const font=environment.fonts.find(f=>target.origin===origin&&target.pathname===f.url_path);
    if(font) {used.add(font.url_path);await route.fulfill({status:200,body:font.bytes,contentType:'font/'+(font.path.endsWith('woff2')?'woff2':'woff')});}
    else if(request.resourceType()==='font'){rejected.push('unlisted font request');await route.abort();}
    else await route.continue();
  });
  return {rejected,used,origin,loaded_fonts:[]};
}

export async function trackFonts(page,context,requests) {
  const client=await context.newCDPSession(page);requests.client=client;
  client.on('CSS.fontsUpdated',({font})=>{if(font)requests.loaded_fonts.push(font);});
  await client.send('DOM.enable');await client.send('CSS.enable');
}

export async function renderEvidence(page,context,environment,requests) {
  const client=requests.client;
  let glyphs=0;const fallback=new Set();const used=new Set();
  try {
    const {root}=await client.send('DOM.getDocument');
    const {nodeIds}=await client.send('DOM.querySelectorAll',{nodeId:root.nodeId,selector:'body,body *'});
    for(const nodeId of nodeIds){
      const {computedStyle}=await client.send('CSS.getComputedStyleForNode',{nodeId});
      const style=Object.fromEntries(computedStyle.map(p=>[p.name,p.value]));
      if(style.display==='none'||style.visibility==='hidden'||style.opacity==='0')continue;
      try{const {model}=await client.send('DOM.getBoxModel',{nodeId});if(model.width<=0||model.height<=0)continue;}catch{continue;}
      const {fonts}=await client.send('CSS.getPlatformFontsForNode',{nodeId});
      for(const font of fonts){
        if(font.glyphCount<=0)continue;
        glyphs+=font.glyphCount;used.add(font.familyName);
        if(!font.isCustomFont || !environment.fonts.some(f=>f.family===font.familyName))fallback.add(font.familyName);
      }
    }
  } finally {await client.detach();}
  const actual=await page.evaluate(()=>({clock:new Date().toISOString(),dpr:devicePixelRatio,
    reduced_motion:matchMedia('(prefers-reduced-motion: reduce)').matches,
    viewport:{width:innerWidth,height:innerHeight},locale:navigator.language,
    timezone:Intl.DateTimeFormat().resolvedOptions().timeZone}));
  const sources=[];
  for(const family of used){
    const loaded=requests.loaded_fonts.filter(f=>f.platformFontFamily===family);
    let bound=false;
    for(const font of loaded){
      const src=typeof font.src==='string'?font.src:'';
      const urls=/^https?:\/\/[^\s]+$/.test(src)?[src]:[...src.matchAll(/url\(\s*(['"]?)(.*?)\1\s*\)/g)].map(m=>m[2]);
      const registered=urls.map(u=>{try{
        const url=new URL(u,requests.origin);
        return environment.fonts.find(f=>url.origin===requests.origin&&url.pathname===f.url_path);
      }catch{return null;}});
      if(urls.length&&!/local\(/i.test(font.src)&&registered.every(Boolean)){
        bound=true;sources.push({family,sha256s:[...new Set(registered.map(f=>f.sha256))].sort()});
      }else fallback.add('unbound custom font: '+family);
    }
    if(!bound)fallback.add('unbound custom font: '+family);
  }
  if(!glyphs||!requests.used.size)fallback.add('no pinned font glyphs');
  return {font_glyphs:glyphs,fonts:[...used].sort(),font_sources:[...new Map(sources.map(s=>[digest(s),s])).values()].sort((a,b)=>digest(a).localeCompare(digest(b))),fallbacks:[...fallback].sort(),
    rejected_font_requests:requests.rejected.length,...actual,
    screenshot_sha256:sha(await page.screenshot({fullPage:true,animations:'disabled',caret:'hide'}))};
}

export function compareRepeats(report) {
  const first=report.runs.filter(r=>r.repeat===0),second=report.runs.filter(r=>r.repeat===1);
  const normalized = row => {const {repeat,...rest}=row;return rest;};
  const equal=first.length>0&&first.length===second.length&&first.every((row,i)=>digest(normalized(row))===digest(normalized(second[i])));
  report.environment.repeat_equal=equal;
  if(!equal){report.uncovered.push('render-repeat-mismatch');report.failures.push({id:'non-deterministic-render'});}
  return equal;
}
