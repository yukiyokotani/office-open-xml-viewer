#!/usr/bin/env node
// Build one WASM reader per published @silurus/ooxml-opendocument entry. Each
// `./<format>` export needs a matching `wasm:<format>` package script (one
// cargo feature each), so adding an entry such as `./ods` only needs its export
// and its script; a missing script fails instead of publishing an entry without
// its parser. Builds run sequentially: every feature writes the same cargo
// artifact before wasm-bindgen copies it out.
import { execFileSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

const directory = new URL('../packages/opendocument/', import.meta.url);
const { exports, scripts } = JSON.parse(readFileSync(new URL('package.json', directory), 'utf8'));
const formats = Object.keys(exports).map(entry => entry.replace(/^\.\//, ''));
for (const format of formats) {
  if (!scripts[`wasm:${format}`]) throw new Error(`packages/opendocument exports ./${format} but has no wasm:${format} script`);
}
for (const format of formats) {
  execFileSync('pnpm', ['run', `wasm:${format}`], { cwd: fileURLToPath(directory), stdio: 'inherit' });
}
