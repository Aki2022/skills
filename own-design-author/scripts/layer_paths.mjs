// Where a repository keeps its layers: the Node reading of references/layer-paths.json (twin of layer_paths.py).
// Every road goes through guard(): a repository that already moved to a layout this version does not read stops
// with LAYER_LAYOUT_FOREIGN_ROOT instead of being read around.
import { readFileSync, statSync, realpathSync, existsSync } from 'node:fs';
import { join, resolve, dirname, basename, isAbsolute, posix, sep } from 'node:path';
import { fileURLToPath } from 'node:url';

const LAYOUT_FILE = fileURLToPath(new URL('../references/layer-paths.json', import.meta.url));
const FOREIGN_CODE = 'LAYER_LAYOUT_FOREIGN_ROOT';

function load() {
  const data = JSON.parse(readFileSync(LAYOUT_FILE, 'utf8'));
  if (data.schema_version !== 1 || typeof data.entries !== 'object' || !Array.isArray(data.foreign_roots) || !data.foreign_roots.length)
    throw new Error('layer-paths.json is malformed');
  return data;
}

export const LAYOUT = load();
export const ENTRIES = LAYOUT.entries;
export const FOREIGN_ROOTS = LAYOUT.foreign_roots;
export const LAYERS_ROOT = ENTRIES.layers;
// Entries whose files are build records, not design inputs (the L0 manifest does not hash them), relative to the layers root.
const NOT_DESIGN_INPUTS = (LAYOUT.not_design_inputs || []).map(key => posix.relative(LAYERS_ROOT, ENTRIES[key]).split('/'));

// False for a file (POSIX path relative to the layers root) under an entry that is not a design input (twin of design_input).
export function designInput(relative) {
  const parts = relative.split('/').filter(Boolean);
  return !NOT_DESIGN_INPUTS.some(prefix => prefix.every((name, i) => parts[i] === name));
}

function foreignMessage(root) {
  const text = `${FOREIGN_CODE}: ${root}/ がある。この repo はこの版の skills が読む置き場所（${LAYERS_ROOT}/）と違う並びに移っている。`
    + `ここで止める（${root}/ を読み飛ばして 0 件にしないため）。skills をこの repo の並びに合う版へ揃える。`
    + ` / ${root}/ exists: this repository follows a layout this version of the skills does not read; stopping `
    + 'instead of reading around it. Update the skills to the version that matches it.';
  return `${text} ${LAYOUT.hint || ''}`.trim();
}

function isDir(path) {
  try { return statSync(path).isDirectory(); } catch { return false; }
}

export function guard(repo) {
  for (const name of FOREIGN_ROOTS) if (isDir(join(repo, name))) throw new Error(foreignMessage(name));
}

export function layerRel(key, ...rest) {
  if (!(key in ENTRIES)) throw new Error(`unknown layer entry: ${key}`);
  let base = ENTRIES[key];
  for (const part of rest) {
    if (isAbsolute(part) || part.split('/').includes('..')) throw new Error(`path escapes the layer entry ${key}: ${part}`);
    base = posix.join(base, part);
  }
  return base;
}

export function layerPath(repo, key, ...rest) {
  guard(repo);
  return join(repo, ...layerRel(key, ...rest).split('/'));
}

// The repository root a resolved manifest path belongs to, from the path text alone (the file may be missing).
export function repoFromResolved(manifest) {
  let full = resolve(manifest);
  try { full = join(realpathSync(dirname(full)), basename(full)); } catch { /* a missing parent keeps the text path */ }
  const want = ENTRIES.resolved.split('/');
  const have = dirname(full).split(sep);
  if (have.length > want.length && have.slice(have.length - want.length).join('/') === want.join('/')) {
    const root = have.slice(0, have.length - want.length).join(sep) || sep;
    if (FOREIGN_ROOTS.includes(basename(root)) && !existsSync(join(root, '.git'))) throw new Error(foreignMessage(basename(root)));
    guard(root);
    return root;
  }
  for (const name of FOREIGN_ROOTS) if (full.split(sep).includes(name)) throw new Error(foreignMessage(name));
  throw new Error(`manifest must be ${ENTRIES.resolved}/<media>.json`);
}
