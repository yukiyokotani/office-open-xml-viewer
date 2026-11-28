// Copy the public specimens that a candidate adds or rewrites into the
// baseline worktree, so both renderers draw identical bytes
// (tests/visual/README.md, "Added or changed public specimens"). Built-ins
// only. Run from the candidate checkout:
//   node tests/visual/copy-corpus-inputs.mjs <baselineCheckout> <baseRevision>
//
// The destination must be the root of another Git checkout whose HEAD is the
// base revision. A specimen the candidate deletes or renames cannot be
// compared and stops the copy before anything is written.
import { execFileSync } from 'node:child_process';
import { lstatSync, mkdirSync, realpathSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { isCorpusInputPath, readRegularFile, writeRegularFile } from './checked-read.mjs';

function git(cwd, ...args) {
  return execFileSync('git', ['-C', cwd, ...args], {
    encoding: 'utf8',
    env: Object.fromEntries(Object.entries(process.env).filter(([key]) => !/^GIT_/i.test(key))),
    stdio: ['ignore', 'pipe', 'pipe'],
  });
}

/** Real path of the Git checkout whose root is `directory`. */
function checkoutRoot(directory, label) {
  let root;
  try {
    root = realpathSync(git(directory, 'rev-parse', '--show-toplevel').trim());
  } catch {
    throw new Error(`${label} is not a Git checkout: ${directory}`);
  }
  if (root !== realpathSync(directory)) {
    throw new Error(`${label} must be the root of a Git checkout: ${directory} is inside ${root}`);
  }
  return root;
}

/** Specimens of `source`'s working tree that differ from `revision`, as
 * `[status, repositoryPath]` pairs. Renames are reported as a deletion and an
 * addition. */
function specimenChanges(source, revision) {
  const fields = git(
    source, 'diff', '--name-status', '--no-renames', '-z', revision, '--',
    'packages/docx/public/corpus', 'packages/xlsx/public/corpus', 'packages/pptx/public/corpus',
  ).split('\0').filter(Boolean);
  const changes = [];
  for (let index = 0; index < fields.length; index += 2) {
    if (isCorpusInputPath(fields[index + 1])) changes.push([fields[index], fields[index + 1]]);
  }
  return changes;
}

function plainDirectory(path) {
  let stat;
  try {
    stat = lstatSync(path);
  } catch {
    return false;
  }
  if (!stat.isDirectory()) {
    throw new Error(`baseline folder must be a real directory, not a symlink or file: ${path}`);
  }
  return true;
}

/** Copy the added and modified specimens of the checkout `source` into the
 * checkout `destination`, which must be at `base`. Returns the copied
 * repository paths. Nothing is written through a symlink: missing corpus
 * folders are created one level at a time and each file goes through the
 * checked writer. */
export function copyCorpusInputs({ source, destination, base }) {
  const from = checkoutRoot(source, 'candidate checkout');
  const to = checkoutRoot(destination, 'baseline checkout');
  if (from === to) {
    throw new Error(`baseline checkout must not be the candidate checkout: ${to}`);
  }
  const revision = git(from, 'rev-parse', '--verify', `${base}^{commit}`).trim();
  const head = git(to, 'rev-parse', 'HEAD').trim();
  if (head !== revision) {
    throw new Error(`baseline checkout ${to} is at ${head}, not at the base revision ${revision}`);
  }
  const changes = specimenChanges(from, revision);
  const removed = changes.filter(([status]) => status !== 'A' && status !== 'M');
  if (removed.length > 0) {
    throw new Error(
      'the self-VRT cannot compare specimens the candidate deletes or renames: '
      + removed.map(([status, path]) => `${status} ${path}`).join(', '),
    );
  }
  const copied = [];
  for (const [, path] of changes) {
    const data = readRegularFile(join(from, path));
    let current = to;
    for (const component of dirname(path).split('/')) {
      current = join(current, component);
      if (!plainDirectory(current)) mkdirSync(current);
    }
    writeRegularFile(join(to, path), data);
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
  try {
    const copied = copyCorpusInputs({ source: process.cwd(), destination: resolve(destination), base });
    console.log(`copied ${copied.length} specimen(s)`);
    for (const path of copied) console.log(`  ${path}`);
  } catch (error) {
    console.error(error.message);
    process.exit(1);
  }
}
