#!/usr/bin/env node
// Built-bundle isolation of the opt-in entries. One table per kind:
//
//  - optional renderers (`./chart-ex`, `./three-d`, `./region-map`, `./tiff`):
//    the entry exports a usable renderer and reaches its implementation, and
//    no default entry reaches that implementation through static imports. The
//    ownership markers are emitted by scripts/optional-renderer-markers.mjs.
//  - legacy Office model sources (`doc`, `xls`, `ppt`): see checkLegacySource.
//
// Usage: node scripts/check-optional-entries.mjs <chart-ex|three-d|region-map|tiff> [dist]
//        node scripts/check-optional-entries.mjs <doc|xls|ppt> [legacy dist]
import { readFile, readdir, stat } from 'node:fs/promises';
import { basename, dirname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const DEFAULT_ENTRIES = ['index.mjs', 'docx.mjs', 'xlsx.mjs', 'pptx.mjs'];

export const OPTIONAL_RENDERERS = {
  'chart-ex': {
    exportName: 'chartEx',
    marker: 'packages/core/src/chart/chart-ex-renderer.ts',
    label: 'ChartEx',
    implementation: 'the ChartEx family implementation',
    // Rolldown may inline a small source module without retaining its source
    // path in emitted JavaScript. Verify that boundary at source level
    // instead: the optional renderer must own the hierarchy dependency, while
    // the default renderer must not regain the hierarchy implementation.
    source: {
      owner: ['packages/core/src/chart/chart-ex-renderer.ts', "from './chart-ex-hierarchy.js'", 'ChartEx renderer does not own the hierarchy module'],
      defaultRenderer: 'packages/core/src/chart/renderer.ts',
      forbiddenInDefaultRenderer: [
        'interface SunburstNode',
        'function hierarchyInputTooLarge',
        'function buildSunburstTree',
        'function layoutSunburstAngles',
        'function sunburstMaxDepth',
      ],
    },
  },
  'three-d': {
    exportName: 'threeD',
    marker: 'packages/core/src/chart/three-d-renderer.ts',
    label: '3-D',
    implementation: 'the mesh/camera implementation',
  },
  'region-map': {
    exportName: 'regionMap',
    marker: 'packages/core/src/chart/region-map-natural-earth.generated.ts',
    label: 'Region Map',
    implementation: 'the fixed offline geometry implementation',
  },
  tiff: {
    exportName: 'tiff',
    marker: 'packages/core/src/image/tiff.ts',
    label: 'TIFF',
    implementation: 'the TIFF implementation',
  },
};

export const LEGACY_SOURCES = {
  doc: { target: 'docx', constant: 'legacyDoc', factory: 'createLegacyDoc', nativeClass: 'LegacyDocDocument' },
  xls: { target: 'xlsx', constant: 'legacyXls', factory: 'createLegacyXls', nativeClass: 'LegacyXlsWorkbook' },
  ppt: { target: 'pptx', constant: 'legacyPpt', factory: 'createLegacyPpt', nativeClass: 'LegacyPptPresentation' },
};

// The renderer checks follow static `./` edges only; the legacy checks also
// follow `import()` and `../` edges. The two spellings of the static pattern
// differ on apostrophes inside a specifier. An implementation must be reached
// under each spelling; it must be absent from the edges either one reads.
// Every edge read must name a file.
const STATIC_SIBLING_EDGE = /(?:from\s*|import\s*)["'](\.\/[^"']+?\.(?:js|mjs))["']/g;
const STATIC_SIBLING_EDGE_WIDE = /(?:from\s*|import\s*)["'](\.\/[^"]+?\.(?:js|mjs))["']/g;
const ANY_RELATIVE_EDGE = /(?:from\s*|import\s*\(?\s*)["'](\.\.?\/[^"']+?\.(?:js|mjs))["']/g;

/** Files reachable from `file` through any of `edgePatterns`, with their text.
 * Every edge must name a file. */
export async function dependencyClosure(file, ...edgePatterns) {
  const pending = [file];
  const visited = new Set();
  const contents = [];
  while (pending.length > 0) {
    const current = pending.pop();
    if (!current || visited.has(current)) continue;
    visited.add(current);
    const source = await readFile(current, 'utf8');
    contents.push({ file: current, source });
    for (const pattern of edgePatterns) {
      for (const match of source.matchAll(pattern)) pending.push(resolve(dirname(current), match[1]));
    }
  }
  return contents;
}

function assertAbsent(files, marker, label) {
  const hit = files.find(({ source }) => source.includes(marker));
  if (hit) throw new Error(`${label} unexpectedly reaches ${marker} via ${basename(hit.file)}`);
}

export async function checkOptionalRenderer(name, { distDir = 'dist', sourceRoot = '.' } = {}) {
  const renderer = OPTIONAL_RENDERERS[name];
  const entry = `${name}.mjs`;
  const dist = resolve(distDir);
  // Read before the entry is run. The implementation must be reached under
  // each spelling of the pattern on its own.
  const closures = [];
  for (const pattern of [STATIC_SIBLING_EDGE, STATIC_SIBLING_EDGE_WIDE]) {
    closures.push(await dependencyClosure(join(dist, entry), pattern));
  }
  const module = await import(pathToFileURL(join(dist, entry)).href);
  if (typeof module[renderer.exportName]?.render !== 'function') {
    throw new Error(`${entry} does not export a usable ${renderer.label} renderer`);
  }
  for (const closure of closures) {
    if (!closure.some(({ source }) => source.includes(renderer.marker))) {
      throw new Error(`${entry} does not reach ${renderer.implementation}`);
    }
  }
  for (const other of DEFAULT_ENTRIES) {
    assertAbsent(await dependencyClosure(join(dist, other), STATIC_SIBLING_EDGE, STATIC_SIBLING_EDGE_WIDE), renderer.marker, other);
  }
  if (renderer.source) {
    const [ownerPath, ownerImport, ownerMessage] = renderer.source.owner;
    if (!(await readFile(resolve(sourceRoot, ownerPath), 'utf8')).includes(ownerImport)) {
      throw new Error(ownerMessage);
    }
    const defaultRenderer = await readFile(resolve(sourceRoot, renderer.source.defaultRenderer), 'utf8');
    for (const symbol of renderer.source.forbiddenInDefaultRenderer) {
      if (defaultRenderer.includes(symbol)) throw new Error(`default renderer unexpectedly owns ${symbol}`);
    }
  }
  return `optional ${renderer.label} bundle boundary verified`;
}

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

/**
 * Verify one legacy Office model source entry of the published dist:
 *
 *  - `<f>.mjs` exports its default ModelSource constant and its factory and
 *    reaches neither the reader runtime nor the WASM glue (importing the entry
 *    or creating a source fetches nothing);
 *  - it references its WASM asset and its source module asset;
 *  - the source module asset is one self-contained ES module (no import of any
 *    sibling chunk or package) that exports `openModelSource`;
 *  - no OOXML entry, other legacy entry or render worker reaches this reader.
 */
export async function checkLegacySource(family, { distDir = 'packages/legacy-office/dist', rootDistDir = 'dist' } = {}) {
  const format = LEGACY_SOURCES[family];
  const dist = resolve(distDir);
  const rootDist = resolve(rootDistDir);
  const entry = `${family}.mjs`;
  const wasmAsset = `legacy_${family}_direct_bg.wasm`;
  const moduleAssetPattern = new RegExp(`^legacy-${family}-source-module(?:-[\\w-]+)?\\.m?js$`);
  const runtimeMarkers = [
    `packages/legacy-office/src/direct-${family}-engine.ts`,
    'packages/core/src/source/model-source-runtime.ts',
    `wasm-direct-${family}/legacy_${family}_direct.js`,
  ];

  const moduleAssets = (await filesUnder(dist)).filter((path) => moduleAssetPattern.test(basename(path)));
  if (moduleAssets.length !== 1) {
    throw new Error(`expected one legacy-${family} source module asset, found ${moduleAssets.length}`);
  }
  const moduleAsset = moduleAssets[0];
  const moduleName = basename(moduleAsset);

  // 1. Entry: constant and factory only.
  const closure = await dependencyClosure(join(dist, entry), ANY_RELATIVE_EDGE);
  for (const marker of runtimeMarkers) assertAbsent(closure, marker, entry);
  assertAbsent(closure, format.nativeClass, entry);
  if (!closure.some(({ source }) => source.includes(wasmAsset))) {
    throw new Error(`${entry} does not reference ${wasmAsset}`);
  }
  if (!closure.some(({ source }) => source.includes(moduleName))) {
    throw new Error(`${entry} does not reference its source module asset ${moduleName}`);
  }
  const factoryModule = await import(pathToFileURL(join(dist, entry)).href);
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
    || load.module.protocol !== 'ooxml-model-source-module/v2'
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
  const constant = factoryModule[format.constant];
  if (
    typeof constant !== 'object' || constant === null || !Object.isFrozen(constant)
    || constant.target !== format.target
    || constant.claim(new Uint8Array([0x50, 0x4b, 0x03, 0x04])) !== false
    || JSON.stringify(constant.beginLoad().module) !== JSON.stringify(defaults)
  ) {
    throw new Error(`${entry} does not export ${format.constant} as the default model source`);
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
  const wasm = await stat(join(dist, wasmAsset));
  if (!wasm.isFile() || wasm.size === 0) throw new Error(`${wasmAsset} is missing or empty`);

  // 3. Nothing else reaches this reader.
  const otherEntries = [
    ...['index', 'docx', 'xlsx', 'pptx', 'node', 'internal-companion'].map((other) => join(rootDist, `${other}.mjs`)),
    ...Object.keys(LEGACY_SOURCES).filter((other) => other !== family).map((other) => join(dist, `${other}.mjs`)),
  ];
  for (const other of otherEntries) {
    const otherClosure = await dependencyClosure(other, ANY_RELATIVE_EDGE);
    assertAbsent(otherClosure, wasmAsset, other);
    assertAbsent(otherClosure, moduleName, other);
    assertAbsent(otherClosure, format.nativeClass, other);
  }
  const workers = (await filesUnder(rootDist)).filter((path) => /render-worker|(?:^|[/\\])worker-[\w-]+\.js$/.test(path) && path.endsWith('.js'));
  for (const worker of workers) {
    const text = await readFile(worker, 'utf8');
    for (const marker of [format.nativeClass, wasmAsset, 'legacy-office']) {
      if (text.includes(marker)) throw new Error(`${basename(worker)} contains ${marker}`);
    }
  }
  return `legacy ${family.toUpperCase()} model source boundary verified (${moduleName}, ${wasmAsset})`;
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const [name, distDir] = process.argv.slice(2);
  if (Object.hasOwn(OPTIONAL_RENDERERS, name)) {
    console.log(await checkOptionalRenderer(name, { distDir }));
  } else if (Object.hasOwn(LEGACY_SOURCES, name)) {
    console.log(await checkLegacySource(name, { distDir }));
  } else {
    throw new Error(`usage: check-optional-entries.mjs <${[...Object.keys(OPTIONAL_RENDERERS), ...Object.keys(LEGACY_SOURCES)].join('|')}> [dist]`);
  }
}
