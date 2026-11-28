// Create the baseline worktree for a self-VRT: a new detached worktree at the
// base revision, plus the public specimens that the candidate adds or
// rewrites, so both renderers draw identical bytes (tests/visual/README.md,
// "Added or changed public specimens"). Built-ins only. Run from the root of
// the candidate checkout:
//   node tests/visual/baseline-worktree.mjs <newWorktreePath> <baseRevision>
//
// The command overwrites nothing an operator could have made: the path must
// not exist, and the specimens go into the tree Git has just checked out. A
// specimen the candidate deletes or renames cannot be compared and stops the
// command before the worktree is created.
import { execFileSync } from 'node:child_process';
import { lstatSync, mkdirSync, realpathSync, rmSync } from 'node:fs';
import { basename, dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { isCorpusInputPath, readRegularFile, writeRegularFile } from './checked-read.mjs';

function git(cwd, ...args) {
  return execFileSync('git', ['-C', cwd, ...args], {
    encoding: 'utf8',
    env: Object.fromEntries(Object.entries(process.env).filter(([key]) => !/^GIT_/i.test(key))),
    stdio: ['ignore', 'pipe', 'pipe'],
  });
}

function exists(path) {
  try {
    lstatSync(path);
    return true;
  } catch {
    return false;
  }
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

/** Create a detached worktree of the checkout `source` at `base` in the new
 * directory `destination` and write into it the specimens that `source`'s
 * working tree adds or modifies relative to `base`. Returns the worktree's
 * real path and the copied repository paths. If the checkout or a write
 * fails, the worktree this call registered is removed again. */
export function createBaselineWorktree({ source, destination, base }) {
  let from;
  try {
    from = realpathSync(git(source, 'rev-parse', '--show-toplevel').trim());
  } catch {
    throw new Error(`candidate checkout is not a Git checkout: ${source}`);
  }
  if (from !== realpathSync(source)) {
    throw new Error(`run from the root of the candidate checkout: ${source} is inside ${from}`);
  }
  if (exists(destination)) {
    throw new Error(`baseline worktree path already exists: ${destination}`);
  }
  const revision = git(from, 'rev-parse', '--verify', `${base}^{commit}`).trim();
  const changes = specimenChanges(from, revision);
  const removed = changes.filter(([status]) => status !== 'A' && status !== 'M');
  if (removed.length > 0) {
    throw new Error(
      'the self-VRT cannot compare specimens the candidate deletes or renames: '
      + removed.map(([status, path]) => `${status} ${path}`).join(', '),
    );
  }
  const copies = changes.map(([, path]) => ({ path, data: readRegularFile(join(from, path)) }));
  const registered = () => git(from, 'worktree', 'list', '--porcelain', '-z').split('\0')
    .filter((field) => field.startsWith('worktree ')).map((field) => field.slice('worktree '.length));
  // Git records the canonical path; the parent may not exist yet.
  const canonical = () => {
    const parent = dirname(destination);
    return join(exists(parent) ? realpathSync(parent) : parent, basename(destination));
  };
  // A moved worktree keeps its registration at the old path; never create
  // (or later roll back) a destination Git already has registered.
  if (registered().includes(canonical())) {
    throw new Error(`baseline worktree path is already registered with Git: ${destination}`);
  }
  let worktree;
  try {
    git(from, 'worktree', 'add', '--quiet', '--detach', destination, revision);
    worktree = realpathSync(destination);
    for (const { path, data } of copies) {
      let current = worktree;
      for (const component of dirname(path).split('/')) {
        current = join(current, component);
        if (!exists(current)) mkdirSync(current);
        else if (!lstatSync(current).isDirectory()) throw new Error(`not a directory at the base revision: ${current}`);
      }
      writeRegularFile(join(worktree, path), data);
    }
  } catch (error) {
    // Undo only the destination this call created (the path did not exist
    // before; a failing checkout hook leaves it registered), so that the
    // command can be run again. Other registrations are never touched: no
    // repository-wide prune.
    if (registered().includes(canonical())) {
      git(from, 'worktree', 'remove', '--force', destination);
    } else if (exists(destination)) {
      rmSync(destination, { recursive: true, force: true });
    }
    throw error;
  }
  return { worktree, copied: copies.map(({ path }) => path) };
}

// Canonical paths, so that the command also runs when invoked through a symlink.
if (process.argv[1] && realpathSync(process.argv[1]) === realpathSync(fileURLToPath(import.meta.url))) {
  const [destination, base, ...rest] = process.argv.slice(2);
  if (!destination || !base || rest.length > 0) {
    console.error('usage: node tests/visual/baseline-worktree.mjs <newWorktreePath> <baseRevision>');
    process.exit(2);
  }
  try {
    const { worktree, copied } = createBaselineWorktree({
      source: process.cwd(), destination: resolve(destination), base,
    });
    console.log(`created ${worktree} with ${copied.length} copied specimen(s)`);
    for (const path of copied) console.log(`  ${path}`);
  } catch (error) {
    console.error(error.message);
    process.exit(1);
  }
}
