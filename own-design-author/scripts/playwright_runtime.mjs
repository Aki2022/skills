// Finds the Playwright package for a rendering run, the same way on every machine and in every repo.
//
// Order:
//   1. OWN_PLAYWRIGHT_DIR: a directory holding node_modules/playwright. Authoritative: when it is set
//      and missing or of another version than the pin, the run stops instead of looking elsewhere.
//   2. The project: node_modules/playwright in the project root or an ancestor. The project root is
//      INFOARCH_PLAYWRIGHT_DIR when set (a snapshot checkout has no node_modules), else cwd.
//      With a pinned version, a project copy of another version is skipped.
//   3. The shared per-machine runtime, only when a version is pinned:
//      ${OWN_PLAYWRIGHT_HOME:-~/.cache/own-playwright}/<version>/node_modules/playwright.
//      One install per version serves every repo; browsers already live in Playwright's own shared cache.
//
// NODE_PATH is never consulted: it silently pointed rendering at another repo's checkout.
// Nothing is installed automatically; when nothing matches, the error names the install command.
import {existsSync, readFileSync} from 'node:fs';
import {createRequire} from 'node:module';
import {dirname, join, resolve} from 'node:path';
import os from 'node:os';
import {layerRel} from './layer_paths.mjs';

export function sharedHome(env = process.env) {
  return env.OWN_PLAYWRIGHT_HOME || join(os.homedir(), '.cache', 'own-playwright');
}

export function installCommand(version, env = process.env) {
  const dir = join(sharedHome(env), version);
  return `npm install --prefix "${dir}" --save-exact playwright@${version} axe-core && "${dir}/node_modules/.bin/playwright" install chromium`;
}

function packageAt(dir) {
  const manifest = join(dir, 'node_modules', 'playwright', 'package.json');
  if (!existsSync(manifest)) return null;
  return {dir, version: JSON.parse(readFileSync(manifest, 'utf8')).version};
}

function projectPackage(root) {
  for (let dir = resolve(root); ; dir = dirname(dir)) {
    const found = packageAt(dir);
    if (found) return found;
    if (dirname(dir) === dir) return null;
  }
}

function load(found, source) {
  const require = createRequire(join(found.dir, 'node_modules', '/'));
  return {playwright: require(join(found.dir, 'node_modules', 'playwright')), version: found.version, source, dir: found.dir};
}

export function locatePlaywright({cwd = process.cwd(), version = null, env = process.env} = {}) {
  const tried = [];
  if (env.OWN_PLAYWRIGHT_DIR) {
    const found = packageAt(resolve(env.OWN_PLAYWRIGHT_DIR));
    if (!found) throw new Error(`OWN_PLAYWRIGHT_DIR に playwright が無い: ${env.OWN_PLAYWRIGHT_DIR}`);
    if (version && found.version !== version)
      throw new Error(`OWN_PLAYWRIGHT_DIR の playwright は ${found.version}、固定の版は ${version}`);
    return {...found, source: 'explicit'};
  }
  const root = env.INFOARCH_PLAYWRIGHT_DIR || cwd;
  const local = projectPackage(root);
  if (local && (!version || local.version === version)) return {...local, source: 'project'};
  tried.push(local ? `project ${local.dir}（版 ${local.version}）` : `project ${resolve(root)}（無し）`);
  if (version) {
    const dir = join(sharedHome(env), version);
    const shared = packageAt(dir);
    if (shared && shared.version === version) return {...shared, source: 'shared'};
    tried.push(`shared ${dir}（${shared ? '版 ' + shared.version : '無し'}）`);
    throw new Error(`playwright ${version} が見つからない。探した場所: ${tried.join(' / ')}。` +
      `端末に一度だけ入れる: ${installCommand(version, env)}`);
  }
  throw new Error(`playwright が見つからない。探した場所: ${tried.join(' / ')}。` +
    `${layerRel('render_environment')} で版を固定するか、project に playwright を入れる。`);
}

// Loads Playwright or exits 2 with the reason; scripts call this once at start.
export function loadPlaywright(options = {}) {
  let found;
  try { found = locatePlaywright(options); }
  catch (error) { console.error(error.message); process.exit(2); }
  return load(found, found.source);
}

// A file another package ships (e.g. "axe-core/axe.min.js"): from the project or an ancestor, else from the
// directory the Playwright runtime was loaded from. NODE_PATH is not consulted. Returns null when absent.
export function locateCompanion(spec, {cwd = process.cwd(), runtimeDir = null, env = process.env} = {}) {
  for (let dir = resolve(env.INFOARCH_PLAYWRIGHT_DIR || cwd); ; dir = dirname(dir)) {
    const path = join(dir, 'node_modules', spec);
    if (existsSync(path)) return path;
    if (dirname(dir) === dir) break;
  }
  if (runtimeDir) {
    const path = join(runtimeDir, 'node_modules', spec);
    if (existsSync(path)) return path;
  }
  return null;
}

// The pinned Playwright version from a render-environment.json, or null when the file is absent.
export function pinnedVersion(environmentPath) {
  if (!environmentPath || !existsSync(environmentPath)) return null;
  return JSON.parse(readFileSync(environmentPath, 'utf8'))?.runtime?.playwright_version || null;
}
