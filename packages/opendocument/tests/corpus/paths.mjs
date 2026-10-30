// The private corpora live in the main checkout, not in worktrees: resolve it
// from git's common directory unless ODF_CORPUS_REPO overrides it.
import { execFileSync } from 'node:child_process';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));

/** Root of the checkout whose sources the harness runs (this worktree). */
export const checkoutRoot = resolve(here, '../../../..');

/** Root of the main checkout holding packages/<fmt>/public/private. */
export function corpusRepo() {
  if (process.env.ODF_CORPUS_REPO) return resolve(process.env.ODF_CORPUS_REPO);
  const common = execFileSync('git', ['rev-parse', '--path-format=absolute', '--git-common-dir'], {
    cwd: checkoutRoot, encoding: 'utf8',
  }).trim();
  return dirname(common);
}
