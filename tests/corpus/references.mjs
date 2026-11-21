#!/usr/bin/env node
// Office reference PDFs (see lib/references.mjs).
//
//   node tests/corpus/references.mjs            list missing and stale PDFs
//   node tests/corpus/references.mjs --record   record every present PDF whose
//                                               entry is missing or stale
//   node tests/corpus/references.mjs --record docx-033 pptx-010   only these cases
//
// Record a PDF only after exporting it from the specimen as committed now.

import { existsSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { readZip } from '../visual/corpus-features.mjs';
import { fileName, loadCases, outputDir, outputPath } from './lib/cases.mjs';
import { VARIANTS, needsReference, pdfName, readManifest, sha256, writeManifest } from './lib/references.mjs';

const args = process.argv.slice(2);
const record = args.includes('--record');
const only = args.filter((a) => !a.startsWith('--'));
const manifest = readManifest();
const missing = [];
const stale = [];
let recorded = 0;

for (const c of await loadCases()) {
  if (only.length && !only.some((id) => c.id === id)) continue;
  const file = fileName(c);
  const source = readFileSync(outputPath(c));
  if (!needsReference(c.format, source, readZip)) continue;
  const sourceSha = sha256(source);
  for (const variant of VARIANTS) {
    const pdf = pdfName(file, variant);
    const path = join(outputDir(c.format), pdf);
    const entry = manifest[pdf];
    if (!existsSync(path)) { missing.push(pdf); continue; }
    const pdfSha = sha256(readFileSync(path));
    if (entry && entry.source === sourceSha && entry.pdf === pdfSha) continue;
    if (record) {
      manifest[pdf] = { source: sourceSha, pdf: pdfSha };
      recorded += 1;
    } else stale.push(pdf);
  }
}

if (record) {
  writeManifest(manifest);
  console.log(`recorded ${recorded} PDF(s)`);
}
if (missing.length) console.log(`missing (${missing.length}):\n  ${missing.join('\n  ')}`);
if (stale.length) console.log(`stale or unrecorded (${stale.length}):\n  ${stale.join('\n  ')}`);
if (!missing.length && !stale.length) console.log('every reference PDF is present and current');
