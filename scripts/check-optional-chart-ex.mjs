import { readFile } from 'node:fs/promises';
import { basename, dirname, join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

const distDir = resolve(process.argv[2] ?? 'dist');
const implementationMarker = 'packages/core/src/chart/chart-ex-renderer.ts';

async function dependencyClosure(entry) {
  const pending = [join(distDir, entry)];
  const visited = new Set();
  const contents = [];
  while (pending.length > 0) {
    const file = pending.pop();
    if (!file || visited.has(file)) continue;
    visited.add(file);
    const source = await readFile(file, 'utf8');
    contents.push({ file, source });
    for (const match of source.matchAll(/(?:from\s*|import\s*)["'](\.\/[^"']+?\.(?:js|mjs))["']/g)) {
      pending.push(resolve(dirname(file), match[1]));
    }
  }
  return contents;
}

function assertAbsent(files, marker, entry) {
  const hit = files.find(({ source }) => source.includes(marker));
  if (hit) throw new Error(`${entry} unexpectedly reaches ${marker} via ${basename(hit.file)}`);
}

const rendererClosure = await dependencyClosure('chart-ex.mjs');
const rendererModule = await import(pathToFileURL(join(distDir, 'chart-ex.mjs')).href);
if (typeof rendererModule.chartEx?.render !== 'function') {
  throw new Error('chart-ex.mjs does not export a usable ChartEx renderer');
}
if (!rendererClosure.some(({ source }) => source.includes(implementationMarker))) {
  throw new Error('chart-ex.mjs does not reach the ChartEx family implementation');
}

for (const entry of ['index.mjs', 'docx.mjs', 'xlsx.mjs', 'pptx.mjs']) {
  const files = await dependencyClosure(entry);
  assertAbsent(files, implementationMarker, entry);
}

// Rolldown may inline a small source module without retaining its source path
// in emitted JavaScript. Verify that boundary at source level instead: the
// optional renderer must own the hierarchy dependency, while the default
// renderer must not regain the hierarchy implementation itself.
const chartExRendererSource = await readFile(
  resolve('packages/core/src/chart/chart-ex-renderer.ts'),
  'utf8',
);
if (!chartExRendererSource.includes("from './chart-ex-hierarchy.js'")) {
  throw new Error('ChartEx renderer does not own the hierarchy module');
}
const defaultRendererSource = await readFile(
  resolve('packages/core/src/chart/renderer.ts'),
  'utf8',
);
for (const symbol of [
  'interface SunburstNode',
  'function hierarchyInputTooLarge',
  'function buildSunburstTree',
  'function layoutSunburstAngles',
  'function sunburstMaxDepth',
]) {
  if (defaultRendererSource.includes(symbol)) {
    throw new Error(`default renderer unexpectedly owns ${symbol}`);
  }
}

console.log('optional ChartEx bundle boundary verified');
