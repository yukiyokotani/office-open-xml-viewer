#!/usr/bin/env node
// Measures the production asset sizes listed on site/src/pages/bundle-size.astro.
//
// Usage (after `pnpm build:wasm && pnpm build`):
//   node scripts/measure-bundle-size.mjs [--dist dist] [--legacy-dist packages/legacy-office/dist]
//     [--opendocument-dist packages/opendocument/dist] [--json]
//
// Definitions (the same ones the page states):
// - "static JavaScript" for a format is every JavaScript file reachable from the
//   published entry (`docx.mjs`, `xlsx.mjs`, `pptx.mjs`) through static
//   `import`/`export … from` edges in the production build. Dynamic `import()`,
//   `new URL(…, import.meta.url)` assets (workers, WASM) are excluded.
// - an optional module row is the static graph of its entry minus the files
//   already present in the union of the three format graphs.
// - raw is the byte length; gzip is measured per emitted asset at level 9 and
//   summed. Values are KiB (1,024 bytes).
// - add-on entry rows (legacy Office, OpenDocument) exclude the package's shared
//   source helper (counted once in its own row) and the `@silurus/ooxml` viewer
//   peer; a reader module row excludes what the entries already load.
import { readFileSync, existsSync, readdirSync } from 'node:fs';
import { dirname, join, relative, resolve } from 'node:path';
import { gzipSync } from 'node:zlib';

const args = process.argv.slice(2);
const option = (name, fallback) => {
  const index = args.indexOf(name);
  return index === -1 ? fallback : args[index + 1];
};
const dist = resolve(option('--dist', 'dist'));
const legacyDist = resolve(option('--legacy-dist', 'packages/legacy-office/dist'));
const openDocumentDist = resolve(option('--opendocument-dist', 'packages/opendocument/dist'));
const asJson = args.includes('--json');

const STATIC_EDGE = /(?:\bfrom\s*|\bimport\s*)(["'])(\.{1,2}\/[^"']+)\1/g;

function staticGraph(entry) {
  const seen = new Set();
  const stack = [resolve(entry)];
  while (stack.length > 0) {
    const file = stack.pop();
    if (seen.has(file)) continue;
    seen.add(file);
    const source = readFileSync(file, 'utf8');
    for (const match of source.matchAll(STATIC_EDGE)) {
      const target = resolve(dirname(file), match[2]);
      if (!/\.m?js$/.test(target)) continue;
      if (!existsSync(target)) throw new Error(`${relative(process.cwd(), file)} imports missing ${match[2]}`);
      stack.push(target);
    }
  }
  return seen;
}

function size(files) {
  let raw = 0;
  let gzip = 0;
  for (const file of files) {
    const bytes = readFileSync(file);
    raw += bytes.length;
    gzip += gzipSync(bytes, { level: 9 }).length;
  }
  return { raw: raw / 1024, gzip: gzip / 1024 };
}

const without = (files, excluded) => [...files].filter((file) => !excluded.has(file));
const rows = [];
const row = (section, name, measured, digits) => rows.push({ section, name, ...measured, digits });

const formatGraphs = Object.fromEntries(
  ['docx', 'xlsx', 'pptx'].map((format) => [format, staticGraph(join(dist, `${format}.mjs`))]),
);
const shared = new Set(Object.values(formatGraphs).flatMap((graph) => [...graph]));
for (const format of ['docx', 'xlsx', 'pptx']) {
  row('Static JavaScript', `${format.toUpperCase()} static JavaScript`, size(formatGraphs[format]), 0);
}

for (const [name, entry] of [
  ['ChartEx', 'chart-ex.mjs'],
  ['3-D charts', 'three-d.mjs'],
  ['Region Map', 'region-map.mjs'],
  ['Math adapter', 'math.mjs'],
  ['TIFF image codec', 'tiff.mjs'],
]) {
  row('Optional modules', name, size(without(staticGraph(join(dist, entry)), shared)), 1);
}

// An add-on package's rows: one per entry (without the shared helper), the
// shared helper (every add-on file that all entries import), one per reader
// module (without anything an entry already loads) and one per reader WASM.
function addOnRows(section, addOnDist, formats, moduleFile, wasmFile) {
  const files = readdirSync(addOnDist);
  const entryGraphs = Object.fromEntries(formats.map((format) => [format, staticGraph(join(addOnDist, `${format}.mjs`))]));
  const entryFile = new RegExp(`/(${formats.join('|')})\\.mjs$`);
  const helper = new Set(
    [...entryGraphs[formats[0]]].filter(
      (file) => formats.every((format) => entryGraphs[format].has(file)) && !entryFile.test(file),
    ),
  );
  for (const format of formats) {
    row(section, `${format.toUpperCase()} entry (/${format})`, size(without(entryGraphs[format], helper)), 1);
  }
  row(section, 'Shared source helper', size(helper), 1);
  const excluded = new Set([...helper, ...Object.values(entryGraphs).flatMap((graph) => [...graph])]);
  for (const format of formats) {
    const file = files.find((name) => moduleFile(format).test(name));
    if (!file) throw new Error(`${format} reader module not found in ${addOnDist}`);
    row(section, `${format.toUpperCase()} reader module`, size(without(staticGraph(join(addOnDist, file)), excluded)), 1);
  }
  for (const format of formats) {
    const file = files.find((name) => wasmFile(format).test(name));
    if (!file) throw new Error(`${format} WASM not found in ${addOnDist}`);
    row(section, `${format.toUpperCase()} reader WASM`, size([join(addOnDist, file)]), 1);
  }
}

if (existsSync(legacyDist)) {
  addOnRows('Legacy Office add-on', legacyDist, ['doc', 'xls', 'ppt'],
    (format) => new RegExp(`^legacy-${format}-source-module.*\\.m?js$`),
    (format) => new RegExp(`^legacy_${format}_direct.*\\.wasm$`));
}

if (existsSync(openDocumentDist)) {
  addOnRows('OpenDocument add-on', openDocumentDist, ['odt', 'odp', 'ods'],
    (format) => new RegExp(`^${format}-source-module.*\\.m?js$`),
    (format) => new RegExp(`^opendocument_${format}.*\\.wasm$`));
}

for (const format of ['docx', 'xlsx', 'pptx']) {
  row('Runtime assets', `${format.toUpperCase()} parser WASM`, size([join(dist, `${format}_parser_bg.wasm`)]), 0);
}
row('Runtime assets', 'MathJax runtime', size([join(dist, 'mathjax-stix2.js')]), 0);

if (asJson) {
  console.log(JSON.stringify(rows, null, 2));
} else {
  const format = (value, digits) => value.toLocaleString('en-US', { minimumFractionDigits: digits, maximumFractionDigits: digits });
  let section = '';
  for (const { section: current, name, raw, gzip, digits } of rows) {
    if (current !== section) console.log(`\n${(section = current)}`);
    console.log(`  ${name.padEnd(28)} ${format(raw, digits).padStart(10)} KiB  ${format(gzip, digits).padStart(9)} KiB gzip`);
  }
}
