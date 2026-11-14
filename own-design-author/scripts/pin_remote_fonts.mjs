#!/usr/bin/env node
/**
 * Record the pin for fonts a CDN serves (e.g. Google Fonts), for render-environment.json's `remote_fonts`.
 *
 *   node pin_remote_fonts.mjs <url>... --environment design/render-environment.json --origins https://fonts.googleapis.com,https://fonts.gstatic.com
 *
 * Renders each URL in the pinned runtime at every viewport and scheme of the environment, fetches everything the
 * page loads from the declared origins, and prints {origins, families, files} where files maps each exact URL to
 * the SHA-256 of its bytes and families are the platform font families those files drew. It contacts the CDN
 * (run it only when you mean to), writes nothing, and the printed pin is what the owner reviews and commits.
 * Pages that draw other characters can request other subsets later: those are rejected until pinned again.
 */
import {readFileSync} from 'node:fs';
import {createRequire} from 'node:module';
import {sha,fixedLaunchOptions} from './fixed_render.mjs';

const args=process.argv.slice(2);
const flag=name=>{const i=args.indexOf('--'+name);return i===-1?null:args[i+1];};
const urls=args.filter((a,i)=>!a.startsWith('--')&&!(i>0&&args[i-1].startsWith('--')));
const origins=(flag('origins')||'').split(',').filter(Boolean);
if(!urls.length||!origins.length||!flag('environment')){
  console.error('usage: pin_remote_fonts.mjs <url>... --environment <render-environment.json> --origins <origin,...>');process.exit(2);
}
if(origins.some(o=>!/^https?:\/\/[^/?#\s]+$/.test(o))){console.error('origins are scheme://host[:port] without a path');process.exit(2);}
const config=JSON.parse(readFileSync(flag('environment'),'utf8'));
const require=createRequire(process.cwd()+'/');
const {chromium}=require('playwright');
const browser=await chromium.launch(fixedLaunchOptions(chromium));
const files={};const families=new Set();
try{
  for(const scheme of config.schemes)for(const viewport of config.viewports)for(const url of urls){
    const context=await browser.newContext({viewport,colorScheme:scheme,deviceScaleFactor:config.device_scale_factor,
      reducedMotion:config.reduced_motion,locale:config.locale,timezoneId:config.timezone,serviceWorkers:'block'});
    await context.route('**/*',async route=>{
      const request=route.request();
      if(!origins.includes(new URL(request.url()).origin))return route.continue();
      const response=await route.fetch();const body=await response.body();
      if(!response.ok()){console.error('the CDN answered '+response.status()+' for '+request.url()+'; not pinned');process.exitCode=2;return route.abort();}
      files[request.url()]=sha(body);
      await route.fulfill({response,body});
    });
    const page=await context.newPage();
    const client=await context.newCDPSession(page);const loaded=[];
    client.on('CSS.fontsUpdated',({font})=>{if(font)loaded.push(font);});
    await client.send('DOM.enable');await client.send('CSS.enable');
    await page.clock.setFixedTime(new Date(config.clock));
    await page.goto(url,{waitUntil:'domcontentloaded',timeout:60000});
    await page.evaluate(()=>document.fonts.ready);
    await page.waitForTimeout(700);
    for(const font of loaded){
      const src=typeof font.src==='string'?font.src:'';
      const sources=[...src.matchAll(/url\(\s*(['"]?)(.*?)\1\s*\)/g)].map(m=>m[2]).concat(/^https?:\/\//.test(src)?[src]:[]);
      if(sources.some(u=>{try{return origins.includes(new URL(u).origin);}catch{return false;}}))families.add(font.platformFontFamily);
    }
    await context.close();
  }
} finally {await browser.close();}
if(process.exitCode===2)process.exit(2);
if(!Object.keys(files).length||!families.size){console.error('the pages loaded no font from the declared origins');process.exit(2);}
const sorted=Object.fromEntries(Object.keys(files).sort().map(k=>[k,files[k]]));
console.log(JSON.stringify({origins,families:[...families].sort(),files:sorted},null,2));
