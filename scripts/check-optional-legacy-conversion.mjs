import { readFile, stat } from 'node:fs/promises';
import { basename, dirname, join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

const distDir = resolve(process.argv[2] ?? 'dist');
const implementationMarker = 'packages/core/src/conversion/legacy-office.ts';

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
    // Static imports only. A dynamic import is the intended opt-in boundary.
    for (const match of source.matchAll(/(?:from\s*|import\s*)["'](\.\/[^"']+?\.(?:js|mjs))["']/g)) {
      pending.push(resolve(dirname(file), match[1]));
    }
  }
  return contents;
}

function assertAbsent(files, marker, entry) {
  const hit = files.find(({ source }) => source.includes(marker));
  if (hit) throw new Error(entry + ' unexpectedly reaches ' + marker + ' via ' + basename(hit.file));
}

const converterClosure = await dependencyClosure('legacy-conversion.mjs');
const converterModule = await import(pathToFileURL(join(distDir, 'legacy-conversion.mjs')).href);
for (const name of [
  'createDisposableWorkerLegacyOfficeConverter',
  'createLegacyOfficeWasmConverter',
  'createLegacyOfficeWasmWorkerConverter',
  'installLegacyOfficeConversionWorkerHandler',
  'validateConvertedOoxml',
]) {
  if (typeof converterModule[name] !== 'function') {
    throw new Error('legacy-conversion.mjs does not export ' + name);
  }
}
if (!converterClosure.some(({ source }) => source.includes(implementationMarker))) {
  throw new Error('legacy-conversion.mjs does not reach the converter boundary implementation');
}
if (!converterClosure.some(({ source }) => source.includes('legacy_office_converter_bg.wasm'))) {
  throw new Error('legacy-conversion.mjs does not reference the opt-in converter WASM asset');
}
const converterWasm = await stat(join(distDir, 'legacy_office_converter_bg.wasm'));
if (!converterWasm.isFile() || converterWasm.size === 0) {
  throw new Error('legacy Office converter WASM asset is missing or empty');
}

for (const entry of ['index.mjs', 'docx.mjs', 'xlsx.mjs', 'pptx.mjs', 'node.mjs']) {
  const closure = await dependencyClosure(entry);
  assertAbsent(closure, implementationMarker, entry);
  assertAbsent(closure, 'legacy_office_converter_bg.wasm', entry);
}

console.log('optional legacy Office conversion bundle boundary verified');
