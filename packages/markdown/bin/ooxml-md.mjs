#!/usr/bin/env node
import { existsSync, readFileSync, writeFileSync } from 'node:fs';
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

Exit codes: 1 usage, 2 unsupported extension, 3 not an OOXML document,
4 OOXML resource limit exceeded.
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
} = await loadAdapter();

// Dev (monorepo) runs the TS source through Vite's module runner: the source
// imports the shared typed errors from `@silurus/ooxml-core`, which ships
// TypeScript that Node's strip-only mode cannot execute (parameter properties,
// bundler-style `.js` specifiers). A published install has no `src/`, so it
// loads the compiled `dist/index.js`, which inlines those helpers. Node refuses
// to strip types from `.ts` files under node_modules, so the standalone package
// MUST expose compiled JS here. When the source exists, a failure to load it
// propagates rather than silently running a possibly stale `dist/`.
async function loadAdapter() {
  const source = fileURLToPath(new URL('../src/index.ts', import.meta.url));
  if (!existsSync(source)) return import('../dist/index.js');
  const { runnerImport } = await import('vite');
  const { module } = await runnerImport(source, { configFile: false, logLevel: 'silent' });
  return module;
}

// The projections throw the shared typed errors (`OoxmlError('not-ooxml')`,
// `OoxmlResourceLimitError`); report each as a distinct exit code. Match on the
// stable `code` rather than the class, because the source and dist entry points
// carry different copies of those classes.
const EXIT_NOT_OOXML = 3;
const EXIT_RESOURCE_LIMIT = 4;

function failOnTypedError(error) {
  if (error?.code === 'not-ooxml') {
    console.error(`ooxml-md: ${positionals[0]}: ${error.message}`);
    process.exit(EXIT_NOT_OOXML);
  }
  if (error?.code === 'ooxml-resource-limit') {
    console.error(`ooxml-md: ${positionals[0]}: ${error.message}`);
    process.exit(EXIT_RESOURCE_LIMIT);
  }
  throw error;
}

function convert(run) {
  try {
    return run();
  } catch (error) {
    return failOnTypedError(error);
  }
}

const buf = readFileSync(filePath);
let md;
if (ext === '.pptx') {
  const wasm = readFileSync(resolveWasm('@silurus/ooxml-pptx', '../../pptx/src/wasm/pptx_parser_bg.wasm'));
  initPptxFromBytes(wasm);
  md = convert(() => pptxToMarkdown(buf));
} else if (ext === '.docx') {
  const wasm = readFileSync(resolveWasm('@silurus/ooxml-docx', '../../docx/src/wasm/docx_parser_bg.wasm'));
  initDocxFromBytes(wasm);
  md = convert(() => docxToMarkdown(buf));
} else if (ext === '.xlsx') {
  const wasm = readFileSync(resolveWasm('@silurus/ooxml-xlsx', '../../xlsx/src/wasm/xlsx_parser_bg.wasm'));
  initXlsxFromBytes(wasm);
  md = convert(() => xlsxToMarkdown(buf));
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
