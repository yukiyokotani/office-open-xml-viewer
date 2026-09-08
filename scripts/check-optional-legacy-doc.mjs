import { readFile, stat } from 'node:fs/promises';
import { basename, dirname, join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

const distDir = resolve(process.argv[2] ?? 'dist');
const factoryMarker = 'packages/legacy-converter/src/direct-doc.ts';
const engineMarker = 'packages/legacy-converter/src/direct-doc-engine.ts';
const wasmAsset = 'legacy_doc_direct_bg.wasm';

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

const directClosure = await dependencyClosure('legacy-doc.mjs');
const directModule = await import(pathToFileURL(join(distDir, 'legacy-doc.mjs')).href);
if (typeof directModule.createLegacyDocSource !== 'function') {
  throw new Error('legacy-doc.mjs does not export createLegacyDocSource');
}
if (!directClosure.some(({ source }) => source.includes(factoryMarker))) {
  throw new Error('legacy-doc.mjs does not reach the direct DOC descriptor factory');
}
if (!directClosure.some(({ source }) => source.includes(wasmAsset))) {
  throw new Error('legacy-doc.mjs does not reference the direct DOC WASM asset');
}
assertAbsent(directClosure, engineMarker, 'legacy-doc.mjs');
const directWasm = await stat(join(distDir, wasmAsset));
if (!directWasm.isFile() || directWasm.size === 0) {
  throw new Error('direct DOC WASM asset is missing or empty');
}

const descriptor = directModule.createLegacyDocSource({
  wasmUrl: 'https://example.test/direct-doc.wasm',
});
if (!Object.isFrozen(descriptor) || descriptor.protocol !== 'ooxml-legacy-doc-source/v1' ||
    descriptor.builtin !== 'doc' || descriptor.wasmUrl !== 'https://example.test/direct-doc.wasm') {
  throw new Error('legacy-doc.mjs returned an invalid descriptor');
}

for (const entry of [
  'index.mjs',
  'docx.mjs',
  'xlsx.mjs',
  'pptx.mjs',
  'node.mjs',
  'legacy-conversion.mjs',
  'legacy-ppt.mjs',
  'legacy-xls.mjs',
]) {
  const closure = await dependencyClosure(entry);
  assertAbsent(closure, factoryMarker, entry);
  assertAbsent(closure, wasmAsset, entry);
}

console.log('optional direct DOC static bundle boundary verified');
