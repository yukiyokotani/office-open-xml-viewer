// Copy the public specimens that a candidate adds or rewrites into a baseline
// worktree, so both renderers draw identical bytes (tests/visual/README.md,
// "Added or changed public specimens"). Built-ins only. Run from the candidate
// checkout:
//   node tests/visual/copy-corpus-inputs.mjs <baselineCheckout> <baseRevision>
import { execFileSync } from 'node:child_process';
import { lstatSync, mkdirSync, realpathSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { isCorpusInput, readRegularFile, writeRegularFile } from './checked-read.mjs';

const SPECIMEN = /^packages\/(docx|xlsx|pptx)\/public\/corpus\/([^/]+)$/;

/** Added or modified specimens of `source` relative to `base`, as repository
 * paths. Only a `.docx`/`.xlsx`/`.pptx` directly in its own package's corpus
 * folder qualifies. */
export function changedSpecimens({ source, base }) {
  const output = execFileSync('git', [
    '-C', source, 'diff', '--name-only', '-z', '--diff-filter=AM', base, '--',
    'packages/*/public/corpus/*.docx', 'packages/*/public/corpus/*.xlsx', 'packages/*/public/corpus/*.pptx',
  ], {
    encoding: 'utf8',
    env: Object.fromEntries(Object.entries(process.env).filter(([key]) => !/^GIT_/i.test(key))),
  });
  return output.split('\0').filter(Boolean).filter((path) => {
    const match = SPECIMEN.exec(path);
    return match !== null && isCorpusInput(match[2], match[1]);
  });
}

function plainDirectory(path, label) {
  let stat;
  try {
    stat = lstatSync(path);
  } catch {
    return false;
  }
  if (stat.isSymbolicLink() || !stat.isDirectory()) {
    throw new Error(`${label} must be a real directory, not a symlink or file: ${path}`);
  }
  return true;
}

/** Copy the changed specimens into `destination`. Nothing is written through a
 * symlink: the destination root and every folder below it must be real
 * directories (missing corpus folders are created one level at a time), and
 * each file is written by the checked writer. Returns the copied paths. */
export function copyCorpusInputs({ source, destination, base }) {
  if (!plainDirectory(destination, 'baseline checkout')) {
    throw new Error(`baseline checkout does not exist: ${destination}`);
  }
  const root = realpathSync(destination);
  const copied = [];
  for (const path of changedSpecimens({ source, base })) {
    const data = readRegularFile(join(source, path));
    let current = root;
    for (const component of dirname(path).split('/')) {
      current = join(current, component);
      if (!plainDirectory(current, 'baseline folder')) mkdirSync(current);
    }
    writeRegularFile(join(root, path), data);
    copied.push(path);
  }
  return copied;
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const [destination, base] = process.argv.slice(2);
  if (!destination || !base) {
    console.error('usage: node tests/visual/copy-corpus-inputs.mjs <baselineCheckout> <baseRevision>');
    process.exit(2);
  }
  const copied = copyCorpusInputs({ source: process.cwd(), destination: resolve(destination), base });
  console.log(`copied ${copied.length} specimen(s)`);
  for (const path of copied) console.log(`  ${path}`);
}
