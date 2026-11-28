// Copy the public specimens that a candidate adds or rewrites into the
// baseline worktree, so both renderers draw identical bytes
// (tests/visual/README.md, "Added or changed public specimens"). Built-ins
// only. Run from the candidate checkout:
//   node tests/visual/copy-corpus-inputs.mjs [--force] <baselineCheckout> <baseRevision>
//
// The destination must be the root of another Git checkout whose HEAD is the
// base revision. Every destination is checked before the first write: a
// specimen the candidate deletes or renames, a destination that is a symlink,
// a hard link or not a regular file, and a destination file that is neither
// the base revision's nor already the candidate's (someone's edit; `--force`
// overwrites it) each stop the copy with nothing written.
import { execFileSync } from 'node:child_process';
import { existsSync, lstatSync, mkdirSync, realpathSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  assertWritableLeaf, isCorpusInputPath, readRegularFile, writeRegularFile,
} from './checked-read.mjs';

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

/** True when the file at `path` of the checkout `root` has the content that
 * `revision` tracks there. */
function tracksContent(root, revision, path) {
  const [entry] = git(root, 'ls-tree', '-z', revision, '--', path).split('\0');
  const blob = /^\d+ blob ([0-9a-f]+)\t/.exec(entry)?.[1];
  return blob !== undefined
    && git(root, 'hash-object', '--no-filters', '--', join(root, path)).trim() === blob;
}

/** Copy the added and modified specimens of the checkout `source` into the
 * checkout `destination`, which must be at `base`. Returns the copied
 * repository paths. All destinations are checked before the first write, and
 * nothing is written through a symlink: missing corpus folders are created
 * one level at a time and each file goes through the checked writer.
 * `force` overwrites destination files that someone edited. */
export function copyCorpusInputs({ source, destination, base, force = false }) {
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
  const copies = changes.map(([, path]) => ({ path, data: readRegularFile(join(from, path)) }));
  const edited = [];
  for (const { path, data } of copies) {
    let current = to;
    let present = true;
    for (const component of dirname(path).split('/')) {
      current = join(current, component);
      present = present && plainDirectory(current);
    }
    assertWritableLeaf(join(to, path));
    if (present && existsSync(join(to, path))
      && !tracksContent(to, revision, path) && !readRegularFile(join(to, path)).equals(data)) {
      edited.push(path);
    }
  }
  if (edited.length > 0 && !force) {
    throw new Error(
      `baseline checkout ${to} holds specimens that are neither the base revision's nor the candidate's `
      + `(pass --force to overwrite them): ${edited.join(', ')}`,
    );
  }
  for (const { path, data } of copies) {
    let current = to;
    for (const component of dirname(path).split('/')) {
      current = join(current, component);
      if (!plainDirectory(current)) mkdirSync(current);
    }
    writeRegularFile(join(to, path), data);
  }
  return copies.map(({ path }) => path);
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const options = process.argv.slice(2);
  const force = options.includes('--force');
  const [destination, base, ...rest] = options.filter((option) => option !== '--force');
  if (!destination || !base || rest.length > 0) {
    console.error('usage: node tests/visual/copy-corpus-inputs.mjs [--force] <baselineCheckout> <baseRevision>');
    process.exit(2);
  }
  try {
    const copied = copyCorpusInputs({
      source: process.cwd(), destination: resolve(destination), base, force,
    });
    console.log(`copied ${copied.length} specimen(s)`);
    for (const path of copied) console.log(`  ${path}`);
  } catch (error) {
    console.error(error.message);
    process.exit(1);
  }
}
