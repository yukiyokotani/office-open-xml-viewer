#!/usr/bin/env node
import { readFileSync, writeFileSync } from 'node:fs';
import { parseArgs } from 'node:util';
import { resolve, dirname, extname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);

/**
 * Locate one format-specific Markdown WASM binary. Published packages resolve
 * through the subpath export; a source checkout falls back to generated output.
 */
function resolveWasm(specifier, relFallback) {
  try {
    return require.resolve(specifier);
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

const buf = readFileSync(filePath);
const format = ext.slice(1);
if (!['pptx', 'docx', 'xlsx'].includes(format)) {
  console.error(`Unsupported extension: ${ext}. Expected .pptx / .docx / .xlsx`);
  process.exit(2);
}

// A source checkout can run the TypeScript module directly on current Node;
// published installs contain only compiled JavaScript.
const api = await import(`../src/${format}.ts`).catch(() => import(`../dist/${format}.js`));
const wasmPath = resolveWasm(
  `@silurus/ooxml-markdown/${format}/wasm-binary`,
  `../wasm/${format}/ooxml_markdown_${format}_bg.wasm`,
);
api.initFromBytes(readFileSync(wasmPath));
const md = api.toMarkdown(buf);

if (values.out) {
  writeFileSync(resolve(values.out), md);
  console.error(`Wrote ${md.length} bytes to ${values.out}`);
} else {
  process.stdout.write(md);
}
