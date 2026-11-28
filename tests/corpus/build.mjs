#!/usr/bin/env node
// Rebuild the public specimen corpus.
//   node tests/corpus/build.mjs           write every case to packages/<fmt>/public/corpus/
//   node tests/corpus/build.mjs --check   exit 1 if a committed file differs from its rebuild

import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { relative } from 'node:path';
import { REPO_ROOT, buildCase, loadCases, outputDir, outputPath } from './lib/cases.mjs';

const check = process.argv.includes('--check');
let stale = 0;
for (const c of await loadCases()) {
  const bytes = buildCase(c);
  const path = outputPath(c);
  const same = existsSync(path) && readFileSync(path).equals(bytes);
  const name = relative(REPO_ROOT, path);
  if (check) {
    if (!same) { stale += 1; console.error(`stale: ${name}`); }
    continue;
  }
  mkdirSync(outputDir(c.format), { recursive: true });
  if (!same) writeFileSync(path, bytes);
  console.log(`${same ? 'unchanged' : 'written  '} ${name} (${bytes.length} bytes)`);
}
if (stale) process.exit(1);
