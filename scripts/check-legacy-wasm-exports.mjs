#!/usr/bin/env node
// The legacy readers build one WASM binary per format from one crate. Verify
// that each published binary is a direct reader only: no OOXML-generation
// export survives, it exports only its own class, and its wasm-bindgen glue
// declares exactly the reader API.
//
// Usage: node scripts/check-legacy-wasm-exports.mjs [dist]
import { readFile } from 'node:fs/promises';
import { join, resolve } from 'node:path';

const distDir = resolve(process.argv[2] ?? 'dist');
const glueDir = resolve('packages/legacy-converter/src');
const FORBIDDEN = /convert|prepare_legacy|conversionoutput|preparedlegacy/i;
const CLASSES = {
  doc: 'LegacyDocDocument',
  xls: 'LegacyXlsWorkbook',
  ppt: 'LegacyPptPresentation',
};
const GLUE_RUNTIME_EXPORTS = ['default', 'initSync', 'reinit', 'InitInput', 'InitOutput', 'SyncInitInput'];

function declaredExports(dts) {
  const names = new Set();
  for (const match of dts.matchAll(/^export\s+(?:declare\s+)?(?:class|function|type|interface|const|let|var)\s+([A-Za-z_$][\w$]*)/gm)) {
    names.add(match[1]);
  }
  if (/^export\s+default\s/m.test(dts)) names.add('default');
  for (const match of dts.matchAll(/^export\s*\{([^}]*)\}/gm)) {
    for (const part of match[1].split(',')) {
      const name = part.trim().split(/\s+as\s+/).pop();
      if (name) names.add(name);
    }
  }
  return names;
}

const failures = [];
for (const [family, className] of Object.entries(CLASSES)) {
  const binary = `legacy_${family}_direct_bg.wasm`;
  const module = await WebAssembly.compile(await readFile(join(distDir, binary)));
  const exports = WebAssembly.Module.exports(module).map((entry) => entry.name);
  const own = className.toLowerCase();
  for (const name of exports) {
    if (FORBIDDEN.test(name)) failures.push(`${binary} exports ${name}`);
    for (const other of Object.values(CLASSES)) {
      const prefix = other.toLowerCase();
      if (prefix !== own && name.toLowerCase().includes(prefix)) {
        failures.push(`${binary} exports another reader's ${name}`);
      }
    }
  }
  if (!exports.some((name) => name.startsWith(`${own}_`))) {
    failures.push(`${binary} exports no ${className} method`);
  }
  const dts = await readFile(join(glueDir, `wasm-direct-${family}`, `legacy_${family}_direct.d.ts`), 'utf8');
  const declared = [...declaredExports(dts)].sort();
  const allowed = [className, ...GLUE_RUNTIME_EXPORTS].sort();
  if (JSON.stringify(declared) !== JSON.stringify(allowed)) {
    failures.push(`legacy_${family}_direct.d.ts declares ${declared.join(', ')}; expected ${allowed.join(', ')}`);
  }
}
if (failures.length > 0) {
  for (const failure of failures) console.error(failure);
  process.exit(1);
}
console.log('legacy reader WASM binaries export only their own direct reader API');
