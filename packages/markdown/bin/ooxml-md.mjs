#!/usr/bin/env node
import { readFileSync, writeFileSync } from 'node:fs';
import { parseArgs } from 'node:util';
import { resolve, dirname, extname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);

/**
 * Locate a parser package's compiled `_bg.wasm` on disk. Tries the published
 * layout first (the parser packages export `./wasm-binary`, so `require.resolve`
 * finds it under `node_modules` after `npm i`, and via pnpm's symlinks inside
 * this monorepo), then falls back to the monorepo-relative sibling path used
 * when running straight from a raw source checkout (e.g. the GitHub Action,
 * where the parser packages' `dist` may be absent but `src/wasm` is present).
 */
function resolveWasm(pkg, relFallback) {
  try {
    return require.resolve(`${pkg}/wasm-binary`);
  } catch {
    return resolve(here, relFallback);
  }
}

const { values, positionals } = parseArgs({
  allowPositionals: true,
  options: {
    out: { type: 'string', short: 'o' },
    help: { type: 'boolean', short: 'h' },
  },
});

if (values.help || positionals.length === 0) {
  console.log(`ooxml-md — convert .pptx / .docx / .xlsx to GitHub-flavoured markdown

Usage:
  ooxml-md <file>              # writes to stdout
  ooxml-md <file> -o out.md    # writes to file
`);
  process.exit(values.help ? 0 : 1);
}

const filePath = resolve(positionals[0]);
const ext = extname(filePath).toLowerCase();
const here = dirname(fileURLToPath(import.meta.url));

const {
  pptxToMarkdown,
  docxToMarkdown,
  xlsxToMarkdown,
  initPptxFromBytes,
  initDocxFromBytes,
  initXlsxFromBytes,
} = await import('../src/index.ts').catch(() => import('../dist/index.js'));
// Dev (monorepo) runs the TS source directly via Node's type stripping; a
// published install has no `src/` (it ships `dist/`), so the first import
// rejects with a not-found / unknown-extension error and we fall back to the
// compiled `dist/index.js`. Node refuses to strip types from `.ts` files under
// node_modules, so the standalone package MUST expose compiled JS here.

const buf = readFileSync(filePath);
let md;
if (ext === '.pptx') {
  const wasm = readFileSync(resolveWasm('@silurus/ooxml-pptx', '../../pptx/src/wasm/pptx_parser_bg.wasm'));
  initPptxFromBytes(wasm);
  md = pptxToMarkdown(buf);
} else if (ext === '.docx') {
  const wasm = readFileSync(resolveWasm('@silurus/ooxml-docx', '../../docx/src/wasm/docx_parser_bg.wasm'));
  initDocxFromBytes(wasm);
  md = docxToMarkdown(buf);
} else if (ext === '.xlsx') {
  const wasm = readFileSync(resolveWasm('@silurus/ooxml-xlsx', '../../xlsx/src/wasm/xlsx_parser_bg.wasm'));
  initXlsxFromBytes(wasm);
  md = xlsxToMarkdown(buf);
} else {
  console.error(`Unsupported extension: ${ext}. Expected .pptx / .docx / .xlsx`);
  process.exit(2);
}

if (values.out) {
  writeFileSync(resolve(values.out), md);
  console.error(`Wrote ${md.length} bytes to ${values.out}`);
} else {
  process.stdout.write(md);
}
