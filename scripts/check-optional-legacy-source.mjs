#!/usr/bin/env node
// Verify one legacy Office model source entry of the published dist:
//
//  - `legacy-<f>.mjs` exports its ModelSource factory and reaches neither the
//    reader runtime nor the WASM glue (creating a source fetches nothing);
//  - it references its WASM asset and its source module asset;
//  - the source module asset is one self-contained ES module (no import of any
//    sibling chunk or package) that exports `openModelSource`;
//  - no OOXML entry, other legacy entry or render worker reaches this reader.
//
// Usage: node scripts/check-optional-legacy-source.mjs <ppt> [dist]
import { readFile, readdir, stat } from 'node:fs/promises';
import { basename, dirname, join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

const FORMATS = {
  ppt: { target: 'pptx', factory: 'legacyPptSource', nativeClass: 'LegacyPptPresentation' },
};

const family = process.argv[2];
const format = FORMATS[family];
if (!format) throw new Error('usage: check-optional-legacy-source.mjs <ppt> [dist]');
const distDir = resolve(process.argv[3] ?? 'dist');
const entry = `legacy-${family}.mjs`;
const wasmAsset = `legacy_${family}_direct_bg.wasm`;
const moduleAssetPattern = new RegExp(`^legacy-${family}-source-module(?:-[\\w-]+)?\\.m?js$`);
const runtimeMarkers = [
  `packages/legacy-converter/src/direct-${family}-engine.ts`,
  'packages/legacy-converter/src/direct-source-runtime.ts',
  `wasm-direct-${family}/legacy_${family}_direct.js`,
];

async function filesUnder(directory) {
  const out = [];
  for (const item of await readdir(directory, { withFileTypes: true })) {
    const path = join(directory, item.name);
    if (item.isDirectory()) {
      if (item.name !== 'types' && !item.name.startsWith('.')) out.push(...await filesUnder(path));
    } else out.push(path);
  }
  return out;
}

async function dependencyClosure(file) {
  const pending = [file];
  const visited = new Set();
  const contents = [];
  while (pending.length > 0) {
    const current = pending.pop();
    if (!current || visited.has(current)) continue;
    visited.add(current);
    const source = await readFile(current, 'utf8');
    contents.push({ file: current, source });
    for (const match of source.matchAll(/(?:from\s*|import\s*\(?\s*)["'](\.\.?\/[^"']+?\.(?:js|mjs))["']/g)) {
      pending.push(resolve(dirname(current), match[1]));
    }
  }
  return contents;
}

function assertAbsent(files, marker, label) {
  const hit = files.find(({ source }) => source.includes(marker));
  if (hit) throw new Error(`${label} unexpectedly reaches ${marker} via ${basename(hit.file)}`);
}

const all = await filesUnder(distDir);
const moduleAssets = all.filter((path) => moduleAssetPattern.test(basename(path)));
if (moduleAssets.length !== 1) {
  throw new Error(`expected one legacy-${family} source module asset, found ${moduleAssets.length}`);
}
const moduleAsset = moduleAssets[0];
const moduleName = basename(moduleAsset);

// 1. Entry: factory only.
const closure = await dependencyClosure(join(distDir, entry));
for (const marker of runtimeMarkers) assertAbsent(closure, marker, entry);
assertAbsent(closure, format.nativeClass, entry);
if (!closure.some(({ source }) => source.includes(wasmAsset))) {
  throw new Error(`${entry} does not reference ${wasmAsset}`);
}
if (!closure.some(({ source }) => source.includes(moduleName))) {
  throw new Error(`${entry} does not reference its source module asset ${moduleName}`);
}
const factoryModule = await import(pathToFileURL(join(distDir, entry)).href);
if (typeof factoryModule[format.factory] !== 'function') {
  throw new Error(`${entry} does not export ${format.factory}`);
}
const source = factoryModule[format.factory]({
  wasmUrl: 'https://example.test/reader.wasm',
  moduleUrl: 'https://example.test/source-module.mjs',
});
const load = source.beginLoad();
if (
  !Object.isFrozen(source) || source.target !== format.target
  || source.claim(new Uint8Array([0x50, 0x4b, 0x03, 0x04])) !== false
  || load.module.protocol !== 'ooxml-model-source-module/v1'
  || load.module.target !== format.target
  || load.module.moduleUrl !== 'https://example.test/source-module.mjs'
  || load.module.config.wasmUrl !== 'https://example.test/reader.wasm'
) {
  throw new Error(`${entry} returned an invalid model source`);
}
const defaults = factoryModule[format.factory]().beginLoad().module;
if (!defaults.moduleUrl.endsWith(`/${moduleName}`) || !defaults.config.wasmUrl.endsWith(`/${wasmAsset}`)) {
  throw new Error(`${entry} default URLs do not name its emitted assets`);
}

// 2. Source module: self-contained, exports openModelSource.
const moduleSource = await readFile(moduleAsset, 'utf8');
const staticImport = /(?:^|[;\n}])\s*import\s*(?:[\w*{][^'"]*from\s*)?["']([^"']+)["']/m.exec(moduleSource);
const dynamicImport = /\bimport\(\s*["']([^"']+)["']\s*\)/.exec(moduleSource);
if (staticImport || dynamicImport) {
  throw new Error(`${moduleName} is not self-contained: imports ${(staticImport ?? dynamicImport)[1]}`);
}
if (!moduleSource.includes(format.nativeClass)) {
  throw new Error(`${moduleName} does not contain the ${format.nativeClass} glue`);
}
const sourceModule = await import(pathToFileURL(moduleAsset).href);
if (typeof sourceModule.openModelSource !== 'function') {
  throw new Error(`${moduleName} does not export openModelSource`);
}
const wasm = await stat(join(distDir, wasmAsset));
if (!wasm.isFile() || wasm.size === 0) throw new Error(`${wasmAsset} is missing or empty`);

// 3. Nothing else reaches this reader.
const otherEntries = ['index.mjs', 'docx.mjs', 'xlsx.mjs', 'pptx.mjs', 'node.mjs',
  ...Object.keys(FORMATS).filter((other) => other !== family).map((other) => `legacy-${other}.mjs`)];
for (const other of otherEntries) {
  const otherClosure = await dependencyClosure(join(distDir, other));
  assertAbsent(otherClosure, wasmAsset, other);
  assertAbsent(otherClosure, moduleName, other);
  assertAbsent(otherClosure, format.nativeClass, other);
}
const workers = all.filter((path) => /render-worker|(?:^|[/\\])worker-[\w-]+\.js$/.test(path) && path.endsWith('.js'));
for (const worker of workers) {
  const text = await readFile(worker, 'utf8');
  for (const marker of [format.nativeClass, wasmAsset, 'legacy-converter']) {
    if (text.includes(marker)) throw new Error(`${basename(worker)} contains ${marker}`);
  }
}

console.log(`legacy ${family.toUpperCase()} model source boundary verified (${moduleName}, ${wasmAsset})`);
