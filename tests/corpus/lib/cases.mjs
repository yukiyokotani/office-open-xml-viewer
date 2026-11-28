// Case registry and output locations.

import { readdirSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { zipStored } from './zip.mjs';

export const CORPUS_DIR = resolve(dirname(fileURLToPath(import.meta.url)), '..');
export const REPO_ROOT = resolve(CORPUS_DIR, '../..');
export const FORMATS = ['docx', 'xlsx', 'pptx'];
const CASE_FILE = /^(docx|xlsx|pptx)-\d{3}-[a-z0-9]+(?:-[a-z0-9]+)*\.mjs$/;

/** Committed output directory of a format (served by Storybook/VRT as /<fmt>/corpus/). */
export const outputDir = (format) => join(REPO_ROOT, 'packages', format, 'public', 'corpus');
export const fileName = (c) => `${c.id}-${c.slug}.${c.format}`;
export const outputPath = (c) => join(outputDir(c.format), fileName(c));

export async function loadCases() {
  const dir = join(CORPUS_DIR, 'cases');
  const files = readdirSync(dir).filter((f) => f.endsWith('.mjs')).sort();
  const cases = [];
  for (const file of files) {
    if (!CASE_FILE.test(file)) throw new Error(`case file name must be <fmt>-NNN-<slug>.mjs: ${file}`);
    const { default: c } = await import(pathToFileURL(join(dir, file)).href);
    if (`${c.id}-${c.slug}.mjs` !== file) throw new Error(`${file}: id/slug do not match the file name`);
    if (!c.id.startsWith(`${c.format}-`)) throw new Error(`${file}: id must start with its format`);
    cases.push(c);
  }
  return cases;
}

export const buildCase = (c) => zipStored(c.parts());
