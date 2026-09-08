import { readFile, stat } from 'node:fs/promises';
import { basename, dirname, join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

const distDir = resolve(process.argv[2] ?? 'dist');
const factoryMarker = 'packages/legacy-converter/src/direct-xls.ts';
const wasmAsset = 'legacy_xls_direct_bg.wasm';

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
  if (hit) throw new Error(entry + ' unexpectedly reaches ' + marker + ' via ' + basename(hit.file));
}

const directClosure = await dependencyClosure('legacy-xls.mjs');
const directModule = await import(pathToFileURL(join(distDir, 'legacy-xls.mjs')).href);
if (typeof directModule.createLegacyXlsSource !== 'function') {
  throw new Error('legacy-xls.mjs does not export createLegacyXlsSource');
}
if (!directClosure.some(({ source }) => source.includes(factoryMarker))) {
  throw new Error('legacy-xls.mjs does not reach the direct XLS descriptor factory');
}
if (!directClosure.some(({ source }) => source.includes(wasmAsset))) {
  throw new Error('legacy-xls.mjs does not reference the direct XLS WASM asset');
}
const directWasm = await stat(join(distDir, wasmAsset));
if (!directWasm.isFile() || directWasm.size === 0) {
  throw new Error('direct XLS WASM asset is missing or empty');
}
// Distinct filenames alone do not prove that the correct native artifact was
// emitted. Compile without instantiating and inspect its direct-session ABI.
const nativeModule = await WebAssembly.compile(await readFile(join(distDir, wasmAsset)));
const nativeExports = new Set(WebAssembly.Module.exports(nativeModule).map(value => value.name));
for (const name of ['legacyxlsworkbook_measurement_request', 'legacyxlsworkbook_pull_sheet_cursor', 'legacyxlsworkbook_close_workbook_session']) {
  if (!nativeExports.has(name)) throw new Error('direct XLS WASM is missing ' + name);
}

const descriptor = directModule.createLegacyXlsSource({
  wasmUrl: 'https://example.test/direct-xls.wasm',
});
if (!Object.isFrozen(descriptor) || descriptor.protocol !== 'ooxml-legacy-xls-source/v1' ||
    descriptor.builtin !== 'xls' || descriptor.wasmUrl !== 'https://example.test/direct-xls.wasm') {
  throw new Error('legacy-xls.mjs returned an invalid descriptor');
}

for (const entry of [
  'index.mjs',
  'docx.mjs',
  'xlsx.mjs',
  'pptx.mjs',
  'node.mjs',
  'legacy-conversion.mjs',
  'legacy-doc.mjs',
  'legacy-ppt.mjs',
]) {
  const closure = await dependencyClosure(entry);
  assertAbsent(closure, factoryMarker, entry);
  assertAbsent(closure, wasmAsset, entry);
}

console.log('optional direct XLS static bundle boundary verified');
