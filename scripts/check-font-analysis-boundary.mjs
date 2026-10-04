#!/usr/bin/env node
/** Check actual shipped graphs after `pnpm build`, including inline parser
 * source and self-contained worker assets. Metric/preload consumers must not
 * acquire PPTX's resource catalogue or normalization/admission machinery.
 * A source facade or a small source file cannot establish this build contract.
 * Future format support must deliberately revise its consumer boundary. */
import { readFileSync, readdirSync } from 'node:fs';
import path from 'node:path';
import ts from 'typescript-compiler-api';
const dist = path.resolve(process.argv[2] ?? 'dist');
const markers = ['ot-definedness-1', 'canonical-static-v1', 'cjkCoverages',
  'packages/core/src/internal/canonical-font-data.ts', 'packages/core/src/internal/font-shaping-profile-data.ts'];
const texts = new Map();
function text(file) {
  if (!texts.has(file)) texts.set(file, readFileSync(file, 'utf8'));
  return texts.get(file);
}
function assertExcluded(file, role, extra = []) {
  const found = [...markers, ...extra].filter(marker => text(file).includes(marker));
  if (found.length) throw Error(`${role} contains unused font analysis: ${found.join(', ')}`);
}
function graph(entry) {
  const seen = new Set();
  function visit(file) {
    if (seen.has(file)) return;
    seen.add(file);
    const ast = ts.createSourceFile(file, text(file), ts.ScriptTarget.Latest, false, ts.ScriptKind.JS);
    for (const statement of ast.statements) {
      if ((ts.isImportDeclaration(statement) || ts.isExportDeclaration(statement))
        && statement.moduleSpecifier && ts.isStringLiteral(statement.moduleSpecifier)
        && statement.moduleSpecifier.text.startsWith('.')) {
        visit(path.resolve(path.dirname(file), statement.moduleSpecifier.text));
      }
    }
  }
  visit(path.join(dist, entry)); return seen;
}
for (const format of ['docx', 'xlsx']) for (const file of graph(`${format}.mjs`)) assertExcluded(file, format);
const inline = new Set(), workers = new Set();
for (const name of readdirSync(dist)) {
  const file = path.join(dist, name);
  if (/^worker-source-.*\.js$/.test(name)) {
    const format = text(file).match(/packages\/(docx|xlsx|pptx)\/src\/worker-source\.ts/)?.[1];
    if (!format) throw Error(`Unidentified inline parser worker: ${name}`);
    assertExcluded(file, `${format} parser/preflight`, ['packages/core/src/fonts/reference-font-metrics-data.json']); inline.add(format);
  }
  if (/^render-worker(?:-source)?-host-.*\.js$/.test(name)) {
    const format = text(file).match(/packages\/(docx|xlsx|pptx)\/src\/render-worker/)?.[1];
    const asset = text(file).match(/new URL\("([^"]+)"/)?.[1];
    if (!format || !asset) throw Error(`Unidentified render worker: ${name}`);
    const worker = path.resolve(dist, asset), source = name.startsWith('render-worker-source-');
    if (format !== 'pptx') assertExcluded(worker, `${format} render worker`);
    else if (!text(worker).includes('canonical-static-v1')) throw Error('PPTX render worker lost support profile');
    workers.add(`${format}:${source ? 'source' : 'default'}`);
  }
}
if (inline.size !== 3 || workers.size !== 6) throw Error(`Incomplete worker inventory: ${inline.size} parsers, ${workers.size} renderers`);
console.log('Font-analysis ownership matches DOCX/XLSX, all three parser/preflight workers and all six render workers.');
