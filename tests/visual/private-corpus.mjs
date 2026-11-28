// Self-VRT harness: exact-pixel comparison of a candidate renderer with the
// images a previous renderer drew from the same inputs.
//
// What a passing comparison establishes for every page, sheet or slide:
//
// 1. Same input. A fixture never loads a corpus input by file name. The specs
//    ask it for a harness URL (`/__selfvrt/<corpus>/<index>.<ext>`,
//    routeSelfVrtInput), which no file name can collide with, and that request
//    is answered by this module's loopback endpoint: it reads the file once
//    with the checked reader, requires that read to have the SHA-256 the
//    corpus manifest holds, and sends those bytes. The digest is written into
//    the per-document manifest beside the image digests, and a document
//    cannot be finished unless its input went through the endpoint. So a
//    baseline image is compared only with a candidate image drawn from bytes
//    of the same digest, whatever names or copies exist on either side.
// 2. Previous renderer. The baseline images are read from the checkout named
//    by VRT_BASELINE_CHECKOUT: a Git checkout root, distinct from the
//    candidate, at VRT_BASELINE_REVISION, whose only differences from that
//    revision are harness files (VRT_ALLOW_HARNESS_CHANGES=1) and public
//    specimen inputs (VRT_ALLOW_CORPUS_INPUT_CHANGES=1). Manifests name the
//    revision and the checkout that wrote them.
// 3. Same images. A document's baseline directory holds exactly the items its
//    manifest lists, each with the recorded SHA-256, reached without a
//    symlink.
//
// Because (1) binds inputs by content, a specimen copied into the baseline
// checkout needs no check of its own: a forgotten, renamed or stale copy
// changes the input list or a digest and fails the manifest comparison,
// whether or not Git reports the file.
//
// Operator mistakes this is meant to catch: the wrong or a modified checkout,
// a forgotten, renamed or stale specimen copy, images left by an earlier run
// or a removed document, symlinked output directories, and non-regular files
// (every read and write goes through checked-read.mjs).
//
// What the harness changes on disk: it writes manifests and images below its
// own checkout, and it removes only item images of the document it is
// rendering (pages a previous run of that document left). Images of documents
// this run does not render (a removed input, a fast-tier input without a
// selected item) are reported, not removed, and so is a symlinked entry. Not covered: a local
// process that replaces files while a run is in progress, and whatever the
// renderer loads besides the input (dev server reuse, WASM provenance).
import {
  existsSync,
  lstatSync,
  mkdirSync,
  readdirSync,
  realpathSync,
  unlinkSync,
} from 'node:fs';
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { createServer } from 'node:http';
import { dirname, isAbsolute, join, parse, relative, resolve, sep } from 'node:path';
import { PNG } from 'pngjs';
import {
  assertWritableLeaf, isCorpusInput, isCorpusInputPath, readRegularFile, writeRegularFile,
} from './checked-read.mjs';

const SCHEMA_VERSION = 2;

// Bootstrap-only exception for running the current VRT harness against an old
// renderer commit that predates the harness. No renderer/parser source is
// allowed in this set. Future baselines should be clean and need no exception.
// The package Vite configs are listed only because they hold the
// `@ooxml-test-*` aliases through which the fixtures load optional renderers
// (for example ChartEx); `harnessBootstrapDiffViolations` enforces that a
// bootstrap diff there adds or removes nothing but such alias lines.
const VRT_HARNESS_PATHS = new Set([
  'package.json',
  'packages/docx/package.json',
  'packages/docx/playwright.config.ts',
  'packages/docx/tests/visual/fixture.html',
  'packages/docx/vite.config.ts',
  'packages/docx/tests/visual/stable-canvas-render.mjs',
  'packages/docx/tests/visual/visual.spec.ts',
  'packages/xlsx/package.json',
  'packages/xlsx/playwright.config.ts',
  'packages/xlsx/tests/visual/fixture.html',
  'packages/xlsx/vite.config.ts',
  'packages/xlsx/tests/visual/visual.spec.ts',
  'packages/pptx/package.json',
  'packages/pptx/playwright.config.ts',
  'packages/pptx/tests/visual/fixture.html',
  'packages/pptx/vite.config.ts',
  'packages/pptx/tests/visual/visual.spec.ts',
  'tests/visual/private-corpus.mjs',
  'tests/visual/checked-read.mjs',
  'tests/visual/corpus-features.mjs',
  // Only for the fixture font packages; see lockfileBootstrapDiffViolations.
  'pnpm-lock.yaml',
]);

const TEST_RENDERER_ALIAS_LINE =
  /^[+-]\s*'@ooxml-test-[a-z0-9-]+': resolve\((?:__)?dirname, '\.\.\/\.\.\/src\/[a-z0-9-]+\.ts'\),$/;

/** Lines of a `git diff -U0` for an allowlisted Vite config that are not a
 * test-renderer alias addition/removal. Any such line makes the harness
 * bootstrap unsafe, because the config also drives the package build. */
export function harnessBootstrapDiffViolations(diff) {
  return diff.split('\n').filter((line) =>
    (line.startsWith('+') || line.startsWith('-'))
    && !line.startsWith('+++')
    && !line.startsWith('---')
    && !TEST_RENDERER_ALIAS_LINE.test(line));
}

// A bootstrap lockfile change may only add or remove the fixtures' font
// packages (`@fontsource/*`): every `-U0` hunk must name such a package, so a
// renderer dependency bump (its own hunk) is rejected.
const FIXTURE_FONT_PACKAGE = /@fontsource\//;
const LOCKFILE_STRUCTURE_LINE =
  /^[+-]\s*(?:'@fontsource\/[a-z0-9-]+(?:@[0-9.]+)?':(?: \{\})?|specifier: \S+|version: \S+|resolution: \{integrity: \S+\}|)$/;

/** Changed lines of a `git diff -U0` for `pnpm-lock.yaml` that are not part of
 * an `@fontsource/*` addition/removal. */
export function lockfileBootstrapDiffViolations(diff) {
  const violations = [];
  for (const hunk of diff.split(/^@@.*$/m).slice(1)) {
    const lines = hunk.split('\n').filter((line) => line.startsWith('+') || line.startsWith('-'));
    if (lines.length === 0) continue;
    if (!lines.some((line) => FIXTURE_FONT_PACKAGE.test(line))) {
      violations.push(...lines);
      continue;
    }
    violations.push(...lines.filter((line) => !LOCKFILE_STRUCTURE_LINE.test(line)));
  }
  return violations;
}

/** Environment for every harness Git command. GIT_DIR, GIT_WORK_TREE,
 * GIT_INDEX_FILE and the other GIT_* variables would otherwise let a caller
 * make a plain directory answer as a checkout, so all of them are removed. */
function gitEnvironment() {
  return Object.fromEntries(
    Object.entries(process.env).filter(([key]) => !/^GIT_/i.test(key)),
  );
}

function git(cwd, args, { trim = true } = {}) {
  const output = execFileSync('git', ['-C', cwd, ...args], {
    encoding: 'utf8',
    env: gitEnvironment(),
    stdio: ['ignore', 'pipe', 'pipe'],
    maxBuffer: 64 * 1024 * 1024,
  });
  return trim ? output.trim() : output;
}

function gitRevision(revision, cwd = process.cwd()) {
  return git(cwd, ['rev-parse', '--verify', `${revision}^{commit}`]);
}

// Facts about the candidate's own checkout do not depend on the baseline, so
// they are resolved once per process; everything about the baseline checkout
// is re-verified on every read.
const candidateFacts = new Map();

function memoizedCandidateFact(key, compute) {
  if (!candidateFacts.has(key)) candidateFacts.set(key, compute());
  return candidateFacts.get(key);
}

function requiredBaselineRevision() {
  const revision = process.env.VRT_BASELINE_REVISION?.trim();
  if (!revision) {
    throw new Error('VRT_BASELINE_REVISION is required for self-VRT');
  }
  return memoizedCandidateFact(
    `revision\0${process.cwd()}\0${revision}`,
    () => gitRevision(revision),
  );
}

/** HEAD and changed paths of a checkout from ONE `git status` call, so the
 * baseline can be re-verified before every item read at low cost. Paths under
 * dependencies and the local private corpus are ignored. */
function checkoutState(root) {
  const output = git(root, [
    'status', '--porcelain=v2', '-z', '--branch', '--untracked-files=all', '--no-renames',
  ], { trim: false });
  let head = null;
  const tracked = [];
  const untracked = [];
  for (const record of output.split('\0').filter(Boolean)) {
    if (record.startsWith('# branch.oid ')) {
      head = record.slice('# branch.oid '.length);
    } else if (record.startsWith('? ')) {
      untracked.push(record.slice(2));
    } else if (record.startsWith('1 ')) {
      tracked.push(record.split(' ').slice(8).join(' '));
    } else if (record.startsWith('u ')) {
      tracked.push(record.split(' ').slice(10).join(' '));
    } else if (!record.startsWith('# ')) {
      throw new Error(`unexpected git status record at ${root}: ${record}`);
    }
  }
  const relevant = (path) => !/(^|\/)node_modules(?:\/|$)/.test(path)
    && !/^packages\/(docx|xlsx|pptx)\/public\/private(?:\/|$)/.test(path);
  return {
    head,
    tracked: tracked.filter(relevant),
    untracked: untracked.filter(relevant),
  };
}

/** Fail unless `root` is a checkout at exactly `revision` whose renderer is
 * unmodified: no tracked change and no untracked file outside dependencies and
 * the local private corpus. With VRT_ALLOW_HARNESS_CHANGES=1, only the
 * allowlisted VRT harness files may differ (the one-time bootstrap), and a Vite
 * config only by test-renderer alias lines. */
function assertRendererCheckout({ root, revision, label }) {
  const { head, tracked, untracked } = checkoutState(root);
  if (head !== revision) {
    throw new Error(`${label} checkout mismatch: expected ${revision}, found ${head} at ${root}`);
  }
  // A public specimen the candidate adds or rewrites is copied into the
  // baseline checkout so that the previous renderer draws it too. Its content
  // is not checked here: the manifests bind every input by SHA-256. A deleted
  // specimen, or one replaced by a symlink, is not admitted.
  const specimenCopy = process.env.VRT_ALLOW_CORPUS_INPUT_CHANGES === '1'
    ? (path) => isCorpusInputPath(path) && isPlainFile(root, path)
    : () => false;
  const changed = [...new Set([...tracked, ...untracked])].filter((path) => !specimenCopy(path));
  if (changed.length === 0) return;
  const harnessBootstrap = process.env.VRT_ALLOW_HARNESS_CHANGES === '1'
    && changed.every((path) => VRT_HARNESS_PATHS.has(path));
  if (!harnessBootstrap) {
    throw new Error(
      `${label} requires a clean renderer checkout; changed paths at ${root}: ${changed.join(', ')}`,
    );
  }
  if (changed.includes('pnpm-lock.yaml')) {
    const violations = lockfileBootstrapDiffViolations(git(root, ['diff', '-U0', 'HEAD', '--', 'pnpm-lock.yaml']));
    if (violations.length > 0) {
      throw new Error(
        `${label} harness bootstrap may only add fixture font packages to pnpm-lock.yaml: `
        + violations.join(' | '),
      );
    }
  }
  for (const path of changed.filter((changedPath) => changedPath.endsWith('vite.config.ts'))) {
    // An untracked config has no HEAD version to diff against, so its
    // whole content would escape the alias-only bound.
    if (untracked.includes(path)) {
      throw new Error(`${label} harness bootstrap cannot add an untracked ${path}`);
    }
    const violations = harnessBootstrapDiffViolations(git(root, ['diff', '-U0', 'HEAD', '--', path]));
    if (violations.length > 0) {
      throw new Error(
        `${label} harness bootstrap may only change test renderer aliases in ${path}: `
        + violations.join(' | '),
      );
    }
  }
}

/** True when `path` below `root` is a regular file and no component of it is
 * a symlink. Only lstat is used, so nothing is followed or opened. */
function isPlainFile(root, path) {
  let current = root;
  let stat;
  for (const component of path.split('/')) {
    current = join(current, component);
    try {
      stat = lstatSync(current);
    } catch {
      return false;
    }
    if (stat.isSymbolicLink()) return false;
  }
  return stat.isFile();
}

function checkoutRoot(path) {
  return git(path, ['rev-parse', '--show-toplevel']);
}

/** Verify that `checkout` is the root of a Git checkout at the full commit
 * `revision` with an unmodified renderer, so its images were produced by that
 * revision's renderer and not copied into an arbitrary directory. Returns the
 * checkout's real path. */
export function verifyBaselineCheckout({ checkout, revision }) {
  if (!/^[0-9a-f]{40}$/.test(revision)) {
    throw new Error(`baseline revision must be a full commit SHA, got ${revision}`);
  }
  const directory = resolve(checkout);
  if (!existsSync(directory)) {
    throw new Error(`VRT_BASELINE_CHECKOUT does not exist: ${directory}`);
  }
  let root;
  try {
    root = checkoutRoot(directory);
  } catch {
    throw new Error(`VRT_BASELINE_CHECKOUT is not a Git checkout: ${directory}`);
  }
  const real = realpathSync(directory);
  if (realpathSync(root) !== real) {
    throw new Error(
      `VRT_BASELINE_CHECKOUT must be the root of a Git checkout: ${directory} is inside ${root}`,
    );
  }
  assertRendererCheckout({ root: real, revision, label: 'self-VRT baseline' });
  return real;
}

/** Real path of the checkout that runs this process. */
function ownCheckout() {
  return memoizedCandidateFact(
    `checkout\0${process.cwd()}`,
    () => realpathSync(checkoutRoot(process.cwd())),
  );
}

function baselineRevision(snapshot = false) {
  const resolved = requiredBaselineRevision();
  if (snapshot) {
    assertRendererCheckout({ root: ownCheckout(), revision: resolved, label: 'self-VRT snapshot' });
  }
  return resolved;
}

// Self-VRT corpora. `demo` is the tracked public demo set under `public/demo/`;
// `corpus` is the tracked public specimen corpus under `public/corpus/`
// (generated by tests/corpus/build.mjs); `private` is the local, gitignored
// corpus under `public/private/<format>/`, and `private-fast` its
// structural-feature subset (tests/visual/corpus-features.mjs). All use the
// same exact-pixel previous-renderer oracle and the same manifest binding; they
// differ only in where the inputs live and where their images go.
// `directory` is the input directory below the public prefix; null means the
// format name.
const CORPORA = {
  demo: {
    publicPrefix: '',
    directory: 'demo',
    outputPrefix: '',
    manifestDirectory: 'demo',
    stemPattern: /^demo\/[^/\\]+$/,
  },
  // The demo decks rendered through PptxScrollViewer: same inputs as `demo`,
  // separate outputs and manifests.
  'demo-scroll': {
    publicPrefix: '',
    directory: 'demo',
    outputPrefix: 'scroll-viewer/',
    manifestDirectory: 'scroll-viewer-demo',
    stemPattern: /^demo\/[^/\\]+$/,
  },
  corpus: {
    publicPrefix: '',
    directory: 'corpus',
    outputPrefix: 'public-corpus/',
    manifestDirectory: 'public-corpus',
    stemPattern: /^corpus\/[^/\\]+$/,
  },
  private: {
    publicPrefix: 'private/',
    directory: null,
    outputPrefix: 'private-corpus/',
    manifestDirectory: 'private-corpus',
    stemPattern: /^(docx|xlsx|pptx)\/[^/\\]+$/,
  },
  // Separate trees, so a subset baseline can never satisfy, or be overwritten
  // by, a full-corpus run.
  'private-fast': {
    publicPrefix: 'private/',
    directory: null,
    outputPrefix: 'private-corpus-fast/',
    manifestDirectory: 'private-corpus-fast',
    stemPattern: /^(docx|xlsx|pptx)\/[^/\\]+$/,
    selection: true,
  },
};

function corpusConfig(corpus) {
  const config = CORPORA[corpus];
  if (!config) throw new Error(`unknown self-VRT corpus: ${corpus}`);
  return config;
}

/** Sorted input files of a corpus, relative to that corpus's public prefix
 * (`demo/sample-1.pptx`, `pptx/deck.pptx`). */
export function selfVrtCorpusFiles({ corpus, format }) {
  const { publicPrefix, directory: configured } = corpusConfig(corpus);
  const directory = configured ?? format;
  return readdirSync(`public/${publicPrefix}${directory}`)
    .filter((file) => isCorpusInput(file, format))
    .map((file) => `${directory}/${file}`)
    .sort((left, right) => left.localeCompare(right, undefined, { numeric: true }));
}

/** Path of a corpus input below the package's `public/` directory. */
export function selfVrtInputPath({ corpus, file }) {
  return `${corpusConfig(corpus).publicPrefix}${file}`;
}

/** Where previous-renderer images live for this run.
 *
 * Snapshots are always written into the checkout that renders them, under the
 * package's `tests/visual/baseline/`, and their manifests record that
 * checkout's real path.
 *
 * A comparison run reads the same package's baseline only from
 * `VRT_BASELINE_CHECKOUT`. The checkout is re-verified (HEAD at
 * VRT_BASELINE_REVISION, clean renderer) on every call, so a baseline that
 * changes mid-run is caught, and every file read below it must resolve to its
 * own path: a symlink anywhere below the checkout root is rejected, so the
 * baseline cannot point at copied images or at the candidate's own output. */
function baselineLocation({ snapshot = false } = {}) {
  const checkout = process.env.VRT_BASELINE_CHECKOUT?.trim();
  if (snapshot) {
    if (checkout) {
      throw new Error(
        'VRT_BASELINE_CHECKOUT is for comparison runs; capture snapshots in the baseline checkout itself',
      );
    }
    return { checkout: ownCheckout(), directory: resolve('tests/visual/baseline'), candidate: null };
  }
  if (!checkout) {
    throw new Error(
      'VRT_BASELINE_CHECKOUT is required for self-VRT comparison: name the previous-renderer '
      + 'checkout at VRT_BASELINE_REVISION',
    );
  }
  const candidate = ownCheckout();
  const packagePath = relative(candidate, realpathSync(process.cwd()));
  if (existsSync(checkout) && realpathSync(checkout) === candidate) {
    throw new Error('VRT_BASELINE_CHECKOUT must name the previous-renderer checkout, not this one');
  }
  const real = verifyBaselineCheckout({ checkout, revision: requiredBaselineRevision() });
  return { checkout: real, directory: join(real, packagePath, 'tests/visual/baseline'), candidate };
}

/** Exported for tests and callers that only need the directory. */
export function selfVrtBaselineRoot({ snapshot = false } = {}) {
  return baselineLocation({ snapshot }).directory;
}

function isWithin(path, directory) {
  const offset = relative(directory, path);
  return offset === '' || (!offset.startsWith('..') && !isAbsolute(offset));
}

/** Resolve a path below a comparison baseline and prove it is a real,
 * symlink-free location inside the verified checkout and outside the
 * candidate. Returns null when the path does not exist. */
function verifiedBaselinePath(location, path) {
  if (!location.candidate) return existsSync(path) ? path : null;
  if (!isWithin(path, location.checkout)) {
    throw new Error(`self-VRT baseline path escapes the baseline checkout: ${path}`);
  }
  let ancestor = location.checkout;
  for (const component of relative(location.checkout, path).split(sep).filter(Boolean)) {
    ancestor = join(ancestor, component);
    if (!existsSync(ancestor) && !isSymlink(ancestor)) return null;
    if (isSymlink(ancestor)) {
      throw new Error(`self-VRT baseline must not contain symlinks: ${ancestor}`);
    }
  }
  const real = realpathSync(path);
  if (real !== path || isWithin(real, location.candidate)) {
    throw new Error(`self-VRT baseline path resolves outside the baseline checkout: ${path}`);
  }
  return real;
}

function isSymlink(path) {
  try {
    return lstatSync(path).isSymbolicLink();
  } catch {
    return false;
  }
}

/** Fail when `path` leaves this checkout or any existing component of it,
 * from the checkout root down to the leaf, is a symlink. Generated
 * directories are sometimes linked to another worktree; nothing is written,
 * listed for removal or removed through such a link. */
function assertNoSymlinkBelowCheckout(path) {
  const root = ownCheckout();
  const offset = relative(root, resolve(realpathSync(process.cwd()), path));
  if (offset.startsWith('..') || isAbsolute(offset)) {
    throw new Error(`self-VRT output path is outside the checkout: ${path}`);
  }
  let current = root;
  for (const component of offset.split(sep).filter(Boolean)) {
    current = join(current, component);
    if (isSymlink(current)) {
      throw new Error(`refusing to use a symlinked self-VRT output path: ${current}`);
    }
  }
}

/** Fail unless the harness may write `path`: below this checkout, reached
 * without a symlink, and absent or a regular file with no other hard link.
 * Operations that remove files before they write call it first, so that a
 * destination they cannot write stops them before anything is removed. */
function assertOutputWritable(path) {
  assertNoSymlinkBelowCheckout(path);
  assertWritableLeaf(path);
}

/** The one way the harness writes a manifest or an image. */
function writeOutput(path, data) {
  assertOutputWritable(path);
  mkdirSync(dirname(path), { recursive: true });
  writeRegularFile(path, data);
}

/** Entries of `directory` whose names match `pattern`, all of them regular
 * files: only those are the harness's to remove, so anything else under such
 * a name fails before a caller removes the first one. */
function generatedFiles(directory, pattern) {
  const matching = readdirSync(directory, { withFileTypes: true }).filter((file) => pattern.test(file.name));
  const foreign = matching.find((file) => !file.isFile());
  if (foreign) {
    throw new Error(
      `refusing to remove a self-VRT output that is not a regular file: ${join(directory, foreign.name)}`,
    );
  }
  return matching.map((file) => file.name);
}

const ITEM_IMAGE = /^(page|sheet|slide)-\d+\.png$/;

/** Fail when `tree`, which holds one directory per document, has images of a
 * document this run does not render (a removed or renamed input, a fast-tier
 * input that no longer has a selected item, or another run's output). The harness deletes nothing it did not write for the
 * document it is rendering, so the operator removes them. A listed document's
 * directory is recognized by file identity, not by its spelling: on a
 * case-insensitive file system `Deck.pptx` writes into an existing `deck/`. */
function assertNoUnlistedDocuments(tree, files, directory) {
  assertNoSymlinkBelowCheckout(tree);
  const identity = (path) => {
    try {
      const { dev, ino } = lstatSync(path);
      return `${dev}:${ino}`;
    } catch {
      return null;
    }
  };
  const listed = new Set(files.map(
    (file) => identity(join(tree, file.slice(directory.length + 1, file.lastIndexOf('.')))),
  ));
  let entries;
  try {
    entries = readdirSync(tree, { withFileTypes: true });
  } catch (error) {
    if (error.code === 'ENOENT') return;
    throw error;
  }
  // A symlink is never followed, so what it holds cannot be judged: any
  // symlinked entry fails.
  const linked = entries.find((entry) => entry.isSymbolicLink());
  if (linked) {
    throw new Error(`refusing to use a symlinked self-VRT output path: ${join(tree, linked.name)}`);
  }
  const unlisted = entries
    .filter((entry) => entry.isDirectory())
    .map((entry) => join(tree, entry.name))
    .filter((path) => !listed.has(identity(path)) && readdirSync(path).some((file) => ITEM_IMAGE.test(file)));
  if (unlisted.length > 0) {
    throw new Error(
      `self-VRT output holds images of documents this run does not render: ${unlisted.join(', ')}. `
      + 'Delete these directories (or the whole output tree) and run again',
    );
  }
}

function sha256(buffer) {
  return createHash('sha256').update(buffer).digest('hex');
}

// Inputs of each prepared corpus, in manifest order. `sha256` is the digest
// the corpus manifest holds; `url` is the only address a fixture is given.
const preparedInputs = new Map();

function corpusFiles(corpus, format, files) {
  const inputs = files.map((name, index) => {
    const path = resolve(`public/${selfVrtInputPath({ corpus, file: name })}`);
    return {
      name, path, sha256: sha256(readRegularFile(path)),
      url: `/__selfvrt/${corpus}/${index}.${format}`, served: 0, failure: null,
    };
  });
  preparedInputs.set(corpus, inputs);
  return inputs.map(({ name, sha256: digest }) => ({ name, sha256: digest }));
}

function preparedInput(corpus, name) {
  const input = preparedInputs.get(corpus)?.find((candidate) => candidate.name === name);
  if (!input) throw new Error(`self-VRT input ${name} was not prepared for the ${corpus} corpus`);
  return input;
}

// One loopback endpoint per process serves the prepared inputs. Each request
// is one checked read: the bytes sent are the bytes whose digest was compared
// with the manifest's, and a mismatch drops the connection, so the fixture
// fails to load. A body goes out as it was read, without the base64 copy that
// a Playwright `route.fulfill` body takes (which caps the input size).
let inputEndpoint;

function selfVrtInputEndpoint() {
  inputEndpoint ??= new Promise((resolveOrigin, reject) => {
    const server = createServer((request, response) => {
      const input = [...preparedInputs.values()].flat().find((candidate) => candidate.url === request.url);
      if (!input) {
        response.writeHead(404, { connection: 'close' }).end();
        return;
      }
      try {
        const body = readRegularFile(input.path);
        if (sha256(body) !== input.sha256) {
          throw new Error(`self-VRT input ${input.path} changed after its manifest recorded it`);
        }
        input.served += 1;
        response.writeHead(200, {
          'content-type': 'application/octet-stream', 'content-length': body.length, connection: 'close',
        }).end(body);
      } catch (error) {
        input.failure = error;
        response.destroy();
      }
    });
    server.unref();
    server.once('error', reject);
    server.listen(0, '127.0.0.1', () => resolveOrigin(`http://127.0.0.1:${server.address().port}`));
  });
  return inputEndpoint;
}

/** Route the page's request for one prepared input to the harness endpoint and
 * return the path to hand the fixture in place of a file name (root-relative,
 * with the format's extension). Call it before the fixture is opened.
 * @param {{ route(matcher: (url: URL) => boolean,
 *   handler: (route: { continue(overrides: { url: string }): Promise<void> }) => unknown): Promise<unknown> }} page
 *   a Playwright page
 * @param {{ corpus: string, file: string }} input */
export async function routeSelfVrtInput(page, { corpus, file }) {
  const input = preparedInput(corpus, file);
  input.served = 0;
  input.failure = null;
  const origin = await selfVrtInputEndpoint();
  await page.route(
    (url) => url.pathname === input.url,
    (route) => route.continue({ url: `${origin}${input.url}` }),
  );
  return input.url.slice(1);
}

/** Name and digest of an input for its document manifest. Fails unless the
 * endpoint served it since it was routed, and served only verified bytes. */
function servedInput(corpus, name) {
  const input = preparedInput(corpus, name);
  if (input.failure) throw input.failure;
  if (input.served === 0) {
    throw new Error(`self-VRT input ${name} was not loaded through the harness endpoint (routeSelfVrtInput)`);
  }
  return { name, sha256: input.sha256 };
}

function readBaselineJson(location, path) {
  const verified = verifiedBaselinePath(location, path);
  if (!verified) throw new Error(`missing previous-renderer manifest: ${path}`);
  return JSON.parse(readRegularFile(verified).toString('utf8'));
}

function assertExactManifest(actual, expected, path) {
  if (JSON.stringify(actual) !== JSON.stringify(expected)) {
    throw new Error(
      `previous-renderer manifest mismatch at ${path}\n`
      + `expected ${JSON.stringify(expected)}\nreceived ${JSON.stringify(actual)}`,
    );
  }
}

/** Call from a suite hook (`beforeAll`), never while tests are being
 * collected: it writes the snapshot manifest and fails on leftover output,
 * which a list-only run must not do.
 *
 * Fail closed on an empty/stale corpus and bind every baseline to an explicit
 * merge-base revision. This prevents a candidate server or unrelated old
 * snapshot from silently becoming its own regression oracle. A subset corpus
 * (`private-fast`) is additionally bound to its selection manifest digest.
 * @param {{ corpus: string, format: string, files: string[], snapshot: boolean,
 *   selectionDigest?: string, selectedFiles?: string[] }} options
 *   `selectedFiles` (subset corpus only): the files with at least one selected
 *   item. Every file is bound by the manifest, but only these are rendered, so
 *   only these may have images in the output tree. */
export function prepareSelfVrtCorpus({
  corpus, format, files, snapshot, selectionDigest, selectedFiles,
}) {
  if (files.length === 0) {
    throw new Error(`${format} ${corpus} corpus is empty; zero-test self-VRT is not coverage`);
  }
  const { selection } = corpusConfig(corpus);
  if (Boolean(selection) !== Boolean(selectionDigest)) {
    throw new Error(selection
      ? `${corpus} self-VRT requires a selection digest`
      : `${corpus} self-VRT does not take a selection digest`);
  }
  if (Boolean(selection) !== Boolean(selectedFiles)
    || selectedFiles?.some((file) => !files.includes(file))) {
    throw new Error(selection
      ? `${corpus} self-VRT requires the corpus files that have selected items`
      : `${corpus} self-VRT renders every file; it does not take selected files`);
  }
  const rendered = selectedFiles ?? files;
  const location = baselineLocation({ snapshot });
  const root = join(location.directory, corpusConfig(corpus).manifestDirectory);
  const path = join(root, 'manifest.json');
  const manifest = {
    schemaVersion: SCHEMA_VERSION,
    format,
    baselineRevision: baselineRevision(snapshot),
    checkout: location.checkout,
    ...(selection ? { selectionDigest } : {}),
    files: corpusFiles(corpus, format, files),
  };
  const { outputPrefix, directory: configured } = corpusConfig(corpus);
  const directory = configured ?? format;
  if (snapshot) {
    assertNoUnlistedDocuments(join(location.directory, `${outputPrefix}${directory}`), rendered, directory);
    writeOutput(path, `${JSON.stringify(manifest, null, 2)}\n`);
  } else {
    assertExactManifest(readBaselineJson(location, path), manifest, path);
    assertNoUnlistedDocuments(resolve(`tests/visual/screenshots/${outputPrefix}${directory}`), rendered, directory);
  }
}

/** A prior candidate run may have rendered more pages than this one. Remove
 * only this input's generated screenshots before capture, so downstream
 * reviews cannot mistake stale page-N files for the current renderer output.
 * Baselines and non-image evidence are never touched. */
export function clearSelfVrtCandidateOutput({
  corpus,
  stem,
  itemKind,
  outputRoot = `tests/visual/screenshots/${corpusConfig(corpus).outputPrefix}`,
}) {
  if (!corpusConfig(corpus).stemPattern.test(stem)
    || ['.', '..'].includes(stem.split('/')[1])
    || !/^(page|sheet|slide)$/.test(itemKind)) {
    throw new Error('invalid self-VRT output identity');
  }
  const directory = resolve(outputRoot, stem);
  // Generated-output directories can be local symlinks. Never follow one when
  // removing stale files: it may point into another worktree or dependencies.
  let ancestor = parse(directory).root;
  for (const component of directory.slice(ancestor.length).split(sep).filter(Boolean)) {
    ancestor = join(ancestor, component);
    if (existsSync(ancestor) && lstatSync(ancestor).isSymbolicLink()) {
      throw new Error(`refusing to clear symlinked self-VRT output: ${ancestor}`);
    }
  }
  if (!existsSync(directory)) return;
  for (const file of generatedFiles(directory, itemPattern(itemKind))) unlinkSync(join(directory, file));
}

function itemPattern(itemKind) {
  return new RegExp(`^${itemKind}-\\d+\\.png$`);
}

function sortedItems(files, itemKind) {
  return files
    .filter((file) => itemPattern(itemKind).test(file))
    .sort((left, right) => left.localeCompare(right, undefined, { numeric: true }));
}

/** Verify that the baseline contains exactly the complete page/sheet/slide set
 * reported by the previous renderer (for a subset corpus, exactly the selected
 * `itemIndices`), each image unchanged since that checkout's snapshot run
 * recorded its SHA-256. An item-count reduction, a stale extra PNG or a
 * replaced image is therefore a hard failure.
 * @param {{ corpus: string, format: string, stem: string, itemKind: string,
 *   itemCount: number, snapshot: boolean, itemIndices?: number[] }} options */
export function verifySelfVrtItemManifest({
  corpus,
  format,
  stem,
  itemKind,
  itemCount,
  snapshot,
  itemIndices,
}) {
  const { selection, outputPrefix } = corpusConfig(corpus);
  if (Boolean(selection) !== Boolean(itemIndices)) {
    throw new Error(selection
      ? `${corpus} self-VRT requires the selected item indices`
      : `${corpus} self-VRT renders every item; it does not take item indices`);
  }
  const indices = itemIndices ?? Array.from({ length: itemCount }, (_, index) => index);
  if (indices.some((index, position) => !Number.isInteger(index) || index < 0 || index >= itemCount
    || (position > 0 && index <= indices[position - 1]))) {
    throw new Error(`invalid ${itemKind} selection for ${stem}: ${JSON.stringify(indices)} of ${itemCount}`);
  }
  const location = baselineLocation({ snapshot });
  const directory = join(location.directory, `${outputPrefix}${stem}`);
  const path = join(directory, 'manifest.json');
  const names = indices.map((index) => `${itemKind}-${index + 1}.png`);
  const header = {
    schemaVersion: SCHEMA_VERSION,
    format,
    baselineRevision: baselineRevision(snapshot),
    checkout: location.checkout,
    input: servedInput(corpus, `${stem}.${format}`),
    itemKind,
    itemCount,
  };
  if (snapshot) {
    // Everything is checked before the first removal: the manifest
    // destination, the type of every item image, and that this run captured
    // every expected item.
    assertOutputWritable(path);
    const present = sortedItems(generatedFiles(directory, itemPattern(itemKind)), itemKind);
    const missing = names.filter((name) => !present.includes(name));
    if (missing.length > 0) {
      throw new Error(
        `snapshot item set mismatch at ${directory}\n`
        + `expected ${JSON.stringify(names)}\nreceived ${JSON.stringify(present)}`,
      );
    }
    // A previous snapshot in this checkout may have rendered more items.
    for (const stale of present) {
      if (!names.includes(stale)) unlinkSync(join(directory, stale));
    }
    const items = names.map((name) => ({ name, sha256: sha256(readRegularFile(join(directory, name))) }));
    writeOutput(path, `${JSON.stringify({ ...header, items }, null, 2)}\n`);
    return;
  }
  const recorded = readBaselineJson(location, path);
  const { items: recordedItems, ...recordedHeader } = recorded;
  if (recordedHeader.input === undefined) {
    throw new Error(
      `${path} has no "input" record: the baseline was captured by a harness without input binding. `
      + 'Bootstrap the baseline worktree with this harness (docs/workflow/vrt.md, "Harness bootstrap"), '
      + 'set VRT_ALLOW_HARNESS_CHANGES=1 for both runs and capture the snapshot again',
    );
  }
  assertExactManifest(recordedHeader, header, path);
  if (!Array.isArray(recordedItems)
    || JSON.stringify(recordedItems.map((item) => item?.name)) !== JSON.stringify(names)) {
    throw new Error(
      `previous-renderer item set mismatch at ${path}\n`
      + `expected ${JSON.stringify(names)}\nreceived ${JSON.stringify(recordedItems)}`,
    );
  }
  const verifiedDirectory = verifiedBaselinePath(location, directory);
  const actualItems = sortedItems(verifiedDirectory ? readdirSync(verifiedDirectory) : [], itemKind);
  if (JSON.stringify(actualItems) !== JSON.stringify(names)) {
    throw new Error(
      `previous-renderer item set mismatch at ${directory}\n`
      + `expected ${JSON.stringify(names)}\nreceived ${JSON.stringify(actualItems)}`,
    );
  }
  for (const { name, sha256: expected } of recordedItems) {
    const verified = verifiedBaselinePath(location, join(directory, name));
    if (!verified || sha256(readRegularFile(verified)) !== expected) {
      throw new Error(`previous-renderer image changed since its snapshot: ${join(directory, name)}`);
    }
  }
}

/** Persist the candidate artifact and return a diagnostic instead of throwing,
 * so one changed item cannot prevent later items in the same document from
 * being rendered and compared. The baseline image must match the SHA-256 its
 * snapshot run recorded. */
export function captureOrCompareSelfVrtItem({
  corpus,
  stem,
  itemKind,
  itemIndex,
  actual,
  snapshot,
}) {
  const key = `${itemKind}-${itemIndex + 1}.png`;
  const { outputPrefix } = corpusConfig(corpus);
  const outputDirectory = `tests/visual/${snapshot ? 'baseline' : 'screenshots'}/${outputPrefix}${stem}`;
  writeOutput(`${outputDirectory}/${key}`, actual);
  if (snapshot) return null;
  const location = baselineLocation();
  const directory = join(location.directory, `${outputPrefix}${stem}`);
  const manifestPath = join(directory, 'manifest.json');
  const recorded = readBaselineJson(location, manifestPath);
  const expected = Array.isArray(recorded.items)
    ? recorded.items.find((item) => item?.name === key)?.sha256
    : undefined;
  const baselinePath = join(directory, key);
  const verified = verifiedBaselinePath(location, baselinePath);
  if (!verified || !expected) return `missing previous-renderer baseline: ${baselinePath}`;
  const baseline = readRegularFile(verified);
  if (sha256(baseline) !== expected) {
    throw new Error(`previous-renderer image changed since its snapshot: ${baselinePath}`);
  }
  return pngPixelsEqual(actual, baseline)
    ? null
    : `${stem} ${key} differs from the previous renderer`;
}

/** PNG byte streams may differ in encoder metadata/compression while decoding
 * to the same canvas. Self-VRT compares the actual rendered RGBA pixels; width
 * or height changes remain regressions. */
export function pngPixelsEqual(leftBuffer, rightBuffer) {
  const left = PNG.sync.read(leftBuffer);
  const right = PNG.sync.read(rightBuffer);
  return left.width === right.width
    && left.height === right.height
    && left.data.equals(right.data);
}
