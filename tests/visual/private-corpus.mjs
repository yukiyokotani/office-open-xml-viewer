import {
  existsSync,
  lstatSync,
  mkdirSync,
  readdirSync,
  realpathSync,
  statSync,
  unlinkSync,
  writeFileSync,
} from 'node:fs';
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { isAbsolute, join, parse, relative, resolve, sep } from 'node:path';
import { PNG } from 'pngjs';
import { readRegularFile } from './checked-read.mjs';

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
    untracked: [...new Set([...untracked, ...ignoredSpecimens(root, head, [...tracked, ...untracked])])]
      .filter(relevant),
  };
}

// `git status` omits ignored files, and a baseline revision that predates the
// public corpus may ignore the specimen folders. So the specimen inputs are
// also listed from disk and compared with the files tracked at the baseline
// revision; any other one counts as an untracked input and goes through the
// same allowance and copy-record checks. The tracked list depends only on the
// revision, so it is read once per checkout and revision.
const trackedSpecimens = new Map();

function ignoredSpecimens(root, head, reported) {
  const key = `${root}\0${head}`;
  if (!trackedSpecimens.has(key)) {
    const output = git(root, [
      'ls-tree', '-r', '-z', '--name-only', head, '--',
      'packages/docx/public/corpus', 'packages/xlsx/public/corpus', 'packages/pptx/public/corpus',
    ], { trim: false });
    trackedSpecimens.set(key, new Set(output.split('\0').filter(Boolean)));
  }
  const tracked = trackedSpecimens.get(key);
  const known = new Set(reported);
  const found = [];
  for (const format of ['docx', 'xlsx', 'pptx']) {
    const directory = `packages/${format}/public/corpus`;
    let names;
    try {
      names = readdirSync(join(root, directory));
    } catch {
      continue;
    }
    for (const name of names) {
      const path = `${directory}/${name}`;
      if (CORPUS_INPUT_PATH.test(path) && !tracked.has(path) && !known.has(path)) found.push(path);
    }
  }
  return found;
}

// A public specimen that the PR under test adds or rewrites has no image from
// the previous renderer. With VRT_ALLOW_CORPUS_INPUT_CHANGES=1 the
// candidate's specimen files may be copied into the baseline checkout (added
// or replacing a tracked one), so both renderers draw identical bytes. Only a
// `.docx`/`.xlsx`/`.pptx` directly in its own package's corpus folder
// qualifies; it must be a regular file reached without a symlink, and when the
// baseline is read from another checkout each copy must still equal the
// candidate's file at that moment. The snapshot manifest binds the inputs by
// SHA-256 as well.
const CORPUS_INPUT_PATH = /^packages\/(docx|xlsx|pptx)\/public\/corpus\/[^/]+\.\1$/;

// The baseline is re-verified before every item it reads. A copy is hashed
// again only when its identity, size or modification time changed, so the
// check costs a stat per input and item rather than a read of every copied
// specimen.
const copyDigests = new Map();

function copyDigest(path) {
  const { dev, ino, size, mtimeMs, ctimeMs } = statSync(path);
  const key = `${dev}:${ino}:${size}:${mtimeMs}:${ctimeMs}`;
  const cached = copyDigests.get(path);
  if (cached?.key === key) return cached.digest;
  const digest = sha256(readRegularFile(path));
  copyDigests.set(path, { key, digest });
  return digest;
}

/** True when `path` below `root` is a regular file and no component of it,
 * from `root` down, is a symlink. Only lstat is used, so a named pipe or other
 * special file is never opened. */
function isPlainFile(root, path) {
  let current = root;
  for (const component of path.split('/').filter(Boolean)) {
    current = join(current, component);
    let stat;
    try {
      stat = lstatSync(current);
    } catch {
      return false;
    }
    if (stat.isSymbolicLink()) return false;
  }
  return lstatSync(current).isFile();
}

function assertCandidateCorpusInputs({ root, candidate, paths, label }) {
  for (const path of paths) {
    const file = join(root, path);
    if (!existsSync(file) && !isSymlink(file)) {
      throw new Error(`${label} corpus input ${path} is missing at ${root}`);
    }
    if (!isPlainFile(root, path) || realpathSync(file) !== file) {
      throw new Error(`${label} corpus input ${path} must be a regular file without symlinks at ${root}`);
    }
    if (candidate === root) continue;
    const own = join(candidate, path);
    if (!isPlainFile(candidate, path)) {
      throw new Error(`${label} corpus input ${path} must be a regular file without symlinks at ${candidate}`);
    }
    if (copyDigest(own) !== copyDigest(file)) {
      throw new Error(`${label} corpus input ${path} differs from the candidate's ${own}`);
    }
  }
}

/** Changed public specimens that VRT_ALLOW_CORPUS_INPUT_CHANGES=1 admits. */
function allowedCorpusInputs({ tracked, untracked }) {
  return process.env.VRT_ALLOW_CORPUS_INPUT_CHANGES === '1'
    ? [...new Set([...tracked, ...untracked])].filter((path) => CORPUS_INPUT_PATH.test(path)).sort()
    : [];
}

/** Specimen copies a snapshot run admitted under the allowance, with their
 * SHA-256, for the corpus manifest. */
function snapshotCorpusInputs(root) {
  const paths = allowedCorpusInputs(checkoutState(root));
  assertCandidateCorpusInputs({ root, candidate: root, paths, label: 'self-VRT snapshot' });
  return paths.map((path) => ({ path, sha256: copyDigest(join(root, path)) }));
}

// Copies the comparison's corpus manifest recorded, per baseline checkout. The
// baseline is re-checked against them before every item it reads, so a copy
// deleted or edited after the snapshot cannot pass.
const recordedCopies = new Map();

/** `present` is the admitted-copy list from the status call that verified the
 * checkout, so no further git process runs here. */
function assertRecordedCopies({ root, candidate, label, present }) {
  const recorded = recordedCopies.get(root);
  if (!recorded || (recorded.length === 0 && present.length === 0)) return;
  if (JSON.stringify(present) !== JSON.stringify(recorded.map((copy) => copy.path))) {
    throw new Error(
      `${label} corpus input copies differ from the snapshot's record: `
      + `expected ${JSON.stringify(recorded.map((copy) => copy.path))}, found ${JSON.stringify(present)}`,
    );
  }
  assertCandidateCorpusInputs({ root, candidate, paths: present, label });
  for (const { path, sha256: expected } of recorded) {
    if (copyDigest(join(root, path)) !== expected) {
      throw new Error(`${label} corpus input ${path} changed since its snapshot recorded it`);
    }
  }
}

/** Fail unless `root` is a checkout at exactly `revision` whose renderer is
 * unmodified: no tracked change and no untracked file outside dependencies and
 * the local private corpus. With VRT_ALLOW_HARNESS_CHANGES=1, only the
 * allowlisted VRT harness files may differ (the one-time bootstrap), and a Vite
 * config only by test-renderer alias lines. */
function assertRendererCheckout({ root, revision, label, candidate = root }) {
  const { head, tracked, untracked } = checkoutState(root);
  if (head !== revision) {
    throw new Error(`${label} checkout mismatch: expected ${revision}, found ${head} at ${root}`);
  }
  const corpusInputs = allowedCorpusInputs({ tracked, untracked });
  assertCandidateCorpusInputs({ root, candidate, paths: corpusInputs, label });
  const changed = [...new Set([...tracked, ...untracked])].filter((path) => !corpusInputs.includes(path));
  if (changed.length === 0) return corpusInputs;
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
  return corpusInputs;
}

function checkoutRoot(path) {
  return git(path, ['rev-parse', '--show-toplevel']);
}

/** Verify that `checkout` is the root of a Git checkout at the full commit
 * `revision` with an unmodified renderer, so its images were produced by that
 * revision's renderer and not copied into an arbitrary directory. Returns the
 * checkout's real path. */
export function verifyBaselineCheckout(options) {
  return verifyBaselineCheckoutState(options).real;
}

function verifyBaselineCheckoutState({ checkout, revision, candidate = ownCheckout() }) {
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
  const corpusInputs = assertRendererCheckout({ root: real, revision, label: 'self-VRT baseline', candidate });
  return { real, corpusInputs };
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
    .filter((file) => file.endsWith(`.${format}`) && !file.startsWith('~$'))
    .map((file) => `${directory}/${file}`)
    .sort((left, right) => left.localeCompare(right, undefined, { numeric: true }));
}

/** URL path of a corpus input, relative to the package's public root. */
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
  const { real, corpusInputs } = verifyBaselineCheckoutState({
    checkout, revision: requiredBaselineRevision(), candidate,
  });
  assertRecordedCopies({ root: real, candidate, label: 'self-VRT baseline', present: corpusInputs });
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

/** Fail when any existing component of `path`, from the checkout root down, is
 * a symlink, so a write or a stale-image removal cannot be redirected out of
 * the checkout. Call it before every snapshot or candidate-output write. */
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
      throw new Error(`refusing to write below a symlinked self-VRT path: ${current}`);
    }
  }
}

function isSymlink(path) {
  try {
    return lstatSync(path).isSymbolicLink();
  } catch {
    return false;
  }
}

function sha256(buffer) {
  return createHash('sha256').update(buffer).digest('hex');
}

// Each input is read once by prepareSelfVrtCorpus, and that buffer is what the
// manifest's SHA-256 describes. Public corpora (demo, demo-scroll, corpus) are
// small, so their buffer is kept and later served to the renderer as is. A
// private corpus can be large, so only the digest is kept there and the route
// re-reads the file through the checked reader and serves that buffer only when
// its digest matches. Either way the images captured or compared in this
// process belong to the manifest's input even when the file on disk is
// replaced during the run.
const preparedInputs = new Map();
const BODY_CACHED_CORPORA = new Set(['demo', 'demo-scroll', 'corpus']);

function corpusFiles(corpus, files, selectionFiles) {
  return files.map((name) => {
    const bytes = readRegularFile(`public/${selfVrtInputPath({ corpus, file: name })}`);
    const digest = sha256(bytes);
    if (selectionFiles) {
      const selected = selectionFiles.get(name)?.sha256;
      if (selected !== digest) {
        throw new Error(
          `${name} changed between the fast-tier selection and corpus preparation `
          + `(selection ${selected}, prepared ${digest}); regenerate the selection`,
        );
      }
    }
    preparedInputs.set(`${corpus}\0${name}`, {
      digest,
      bytes: BODY_CACHED_CORPORA.has(corpus) ? bytes : null,
    });
    return { name, sha256: digest };
  });
}

/** The bytes of a prepared self-VRT input, for the spec to serve to the
 * fixture (page.route) instead of letting the dev server read the file again.
 * Throws when the file no longer has the digest its manifest recorded.
 * @param {{ corpus: string, file: string }} options */
export function selfVrtInputBody({ corpus, file }) {
  const prepared = preparedInputs.get(`${corpus}\0${file}`);
  if (!prepared) throw new Error(`self-VRT input ${file} was not prepared for the ${corpus} corpus`);
  if (prepared.bytes) return prepared.bytes;
  const path = `public/${selfVrtInputPath({ corpus, file })}`;
  const body = readRegularFile(path);
  if (sha256(body) !== prepared.digest) {
    throw new Error(`self-VRT input ${path} changed after its manifest recorded it`);
  }
  return body;
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

/** Fail closed on an empty/stale corpus and bind every baseline to an explicit
 * merge-base revision. This prevents a candidate server or unrelated old
 * snapshot from silently becoming its own regression oracle. A subset corpus
 * (`private-fast`) is additionally bound to its selection manifest digest.
 * `selectionFiles` (private-fast) carries the digest the selection loader
 * validated for each file; every prepared input, including a file without any
 * selected item, must hash to it.
 * @param {{ corpus: string, format: string, files: string[], snapshot: boolean,
 *   selectionDigest?: string, selectionFiles?: Map<string, { sha256: string }> }} options */
export function prepareSelfVrtCorpus({
  corpus, format, files, snapshot, selectionDigest, selectionFiles,
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
  if (Boolean(selection) !== Boolean(selectionFiles)) {
    throw new Error(selection
      ? `${corpus} self-VRT requires the selection's per-file digests`
      : `${corpus} self-VRT does not take per-file selection digests`);
  }
  // Each corpus manifest is read afresh; a record from an earlier prepare must
  // not stand in for it.
  recordedCopies.clear();
  const location = baselineLocation({ snapshot });
  const root = join(location.directory, corpusConfig(corpus).manifestDirectory);
  const path = join(root, 'manifest.json');
  const manifest = {
    schemaVersion: SCHEMA_VERSION,
    format,
    baselineRevision: baselineRevision(snapshot),
    checkout: location.checkout,
    ...(selection ? { selectionDigest } : {}),
    files: corpusFiles(corpus, files, selectionFiles),
  };
  if (snapshot) {
    const copies = snapshotCorpusInputs(location.checkout);
    assertNoSymlinkBelowCheckout(path);
    mkdirSync(root, { recursive: true });
    writeFileSync(
      path,
      `${JSON.stringify(copies.length > 0 ? { ...manifest, corpusInputCopies: copies } : manifest, null, 2)}\n`,
    );
  } else {
    const { corpusInputCopies = [], ...header } = readBaselineJson(location, path);
    assertExactManifest(header, manifest, path);
    if (!Array.isArray(corpusInputCopies)
      || corpusInputCopies.some((copy) => typeof copy?.path !== 'string' || typeof copy?.sha256 !== 'string')) {
      throw new Error(`invalid corpus input copy record in ${path}`);
    }
    if (corpusInputCopies.length > 0 && process.env.VRT_ALLOW_CORPUS_INPUT_CHANGES !== '1') {
      throw new Error(
        `${path} records specimen copies admitted by VRT_ALLOW_CORPUS_INPUT_CHANGES=1; `
        + 'set it for the comparison run too',
      );
    }
    recordedCopies.set(location.checkout, corpusInputCopies);
    // Verify the baseline against the record just loaded.
    baselineLocation();
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
  const item = new RegExp(`^${itemKind}-\\d+\\.png$`);
  for (const file of readdirSync(directory)) {
    if (item.test(file)) unlinkSync(`${directory}/${file}`);
  }
}

function itemPattern(itemKind) {
  return new RegExp(`^${itemKind}-\\d+\\.png$`);
}

function sortedItems(files, itemKind) {
  return files
    .filter((file) => itemPattern(itemKind).test(file))
    .sort((left, right) => left.localeCompare(right, undefined, { numeric: true }));
}

// Per document and run: the plan validated by beginSelfVrtItems, kept so that
// capture, comparison and finalization all use that one object (a snapshot's
// digests of the captured buffers; a comparison's single read of the baseline
// manifest) instead of reading the disk again.
const itemStates = new Map();

function itemPlan({ corpus, format, stem, itemKind, itemCount, snapshot, itemIndices }) {
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
  const inputName = `${stem}.${format}`;
  const prepared = preparedInputs.get(`${corpus}\0${inputName}`);
  if (!prepared) throw new Error(`self-VRT input ${inputName} was not prepared for the ${corpus} corpus`);
  const location = baselineLocation({ snapshot });
  const directory = join(location.directory, `${outputPrefix}${stem}`);
  return {
    location,
    directory,
    path: join(directory, 'manifest.json'),
    names: indices.map((index) => `${itemKind}-${index + 1}.png`),
    header: {
      schemaVersion: SCHEMA_VERSION,
      format,
      baselineRevision: baselineRevision(snapshot),
      checkout: location.checkout,
      input: { name: inputName, sha256: prepared.digest },
      itemKind,
      itemCount,
    },
  };
}

/** Call once per document, after its item count is known and before the first
 * item is captured or compared.
 *
 * Snapshot: prepares the output directory (removing stale images from an
 * earlier snapshot) and starts recording the digest of every captured buffer.
 * Comparison: reads the baseline's per-document manifest exactly once, checks
 * its header (input digest, checkout, revision), its item set and the
 * baseline directory's contents, and keeps the validated hashes; every image
 * comparison then derives from this object. The baseline contains exactly the
 * complete page/sheet/slide set the previous renderer reported (for a subset
 * corpus, exactly the selected `itemIndices`), so an item-count reduction, a
 * stale extra PNG or a replaced image is a hard failure.
 * @param {{ corpus: string, format: string, stem: string, itemKind: string,
 *   itemCount: number, snapshot: boolean, itemIndices?: number[] }} options */
export function beginSelfVrtItems(options) {
  const { corpus, stem, itemKind, snapshot } = options;
  const plan = itemPlan(options);
  const { location, directory, path, names, header } = plan;
  const key = `${corpus}\0${stem}`;
  if (snapshot) {
    assertNoSymlinkBelowCheckout(directory);
    mkdirSync(directory, { recursive: true });
    // A previous snapshot in this checkout may have rendered more items.
    for (const stale of sortedItems(readdirSync(directory), itemKind)) {
      if (!names.includes(stale)) unlinkSync(join(directory, stale));
    }
    itemStates.set(key, { snapshot, names, header, digests: new Map() });
    return;
  }
  const recorded = readBaselineJson(location, path);
  const { items: recordedItems, ...recordedHeader } = recorded;
  assertExactManifest(recordedHeader, header, path);
  if (!Array.isArray(recordedItems)
    || JSON.stringify(recordedItems.map((item) => item?.name)) !== JSON.stringify(names)
    || recordedItems.some((item) => typeof item?.sha256 !== 'string')) {
    throw new Error(
      `previous-renderer item set mismatch at ${path}\n`
      + `expected ${JSON.stringify(names)}\nreceived ${JSON.stringify(recordedItems)}`,
    );
  }
  assertBaselineItemSet({ location, directory, itemKind, names });
  itemStates.set(key, {
    snapshot,
    names,
    header,
    expected: new Map(recordedItems.map((item) => [item.name, item.sha256])),
    compared: new Set(),
  });
}

function assertBaselineItemSet({ location, directory, itemKind, names }) {
  const verifiedDirectory = verifiedBaselinePath(location, directory);
  const actualItems = sortedItems(verifiedDirectory ? readdirSync(verifiedDirectory) : [], itemKind);
  if (JSON.stringify(actualItems) !== JSON.stringify(names)) {
    throw new Error(
      `previous-renderer item set mismatch at ${directory}\n`
      + `expected ${JSON.stringify(names)}\nreceived ${JSON.stringify(actualItems)}`,
    );
  }
}

/** Finish a document's items. Same options as beginSelfVrtItems, which must
 * have run first. Snapshot: every item must have been captured, the file
 * written for it must still be the buffer that was captured, and the manifest
 * records the captured buffers' digests. Comparison: every item must have been
 * compared against the manifest beginSelfVrtItems validated; the baseline's
 * item set is checked again, and no image content is read a second time. */
export function verifySelfVrtItemManifest(options) {
  const { corpus, stem, itemKind, snapshot } = options;
  const key = `${corpus}\0${stem}`;
  const state = itemStates.get(key);
  if (!state || state.snapshot !== snapshot) {
    throw new Error(`self-VRT items of ${stem} were not begun for this run`);
  }
  const plan = itemPlan(options);
  const { location, directory, path, names, header } = plan;
  if (JSON.stringify(header) !== JSON.stringify(state.header)
    || JSON.stringify(names) !== JSON.stringify(state.names)) {
    throw new Error(`self-VRT items of ${stem} changed since they were begun`);
  }
  itemStates.delete(key);
  if (snapshot) {
    assertNoSymlinkBelowCheckout(directory);
    const present = sortedItems(readdirSync(directory), itemKind);
    if (JSON.stringify(present) !== JSON.stringify(names)) {
      throw new Error(
        `snapshot item set mismatch at ${directory}\n`
        + `expected ${JSON.stringify(names)}\nreceived ${JSON.stringify(present)}`,
      );
    }
    const items = names.map((name) => {
      const digest = state.digests.get(name);
      if (!digest) throw new Error(`snapshot item ${name} of ${stem} was never captured`);
      if (sha256(readRegularFile(join(directory, name))) !== digest) {
        throw new Error(`snapshot image changed after capture: ${join(directory, name)}`);
      }
      return { name, sha256: digest };
    });
    writeFileSync(path, `${JSON.stringify({ ...header, items }, null, 2)}\n`);
    return;
  }
  for (const name of names) {
    if (!state.compared.has(name)) throw new Error(`${stem} ${name} was never compared`);
  }
  assertBaselineItemSet({ location, directory, itemKind, names });
}

/** Persist the candidate artifact and return a diagnostic instead of throwing,
 * so one changed item cannot prevent later items in the same document from
 * being rendered and compared. The baseline image is read once; its SHA-256
 * must equal the one in the manifest beginSelfVrtItems validated, and the
 * pixel comparison uses that same buffer. Snapshot runs record the digest of
 * the captured buffer. */
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
  const state = itemStates.get(`${corpus}\0${stem}`);
  if (!state || state.snapshot !== snapshot) {
    throw new Error(`self-VRT items of ${stem} were not begun before capture`);
  }
  const outputDirectory = `tests/visual/${snapshot ? 'baseline' : 'screenshots'}/${outputPrefix}${stem}`;
  assertNoSymlinkBelowCheckout(`${outputDirectory}/${key}`);
  mkdirSync(outputDirectory, { recursive: true });
  writeFileSync(`${outputDirectory}/${key}`, actual);
  if (snapshot) {
    state.digests.set(key, sha256(actual));
    return null;
  }
  const location = baselineLocation();
  const baselinePath = join(location.directory, `${outputPrefix}${stem}`, key);
  const expected = state.expected.get(key);
  const verified = verifiedBaselinePath(location, baselinePath);
  if (!verified || !expected) return `missing previous-renderer baseline: ${baselinePath}`;
  const baseline = readRegularFile(verified);
  if (sha256(baseline) !== expected) {
    throw new Error(`previous-renderer image changed since its snapshot: ${baselinePath}`);
  }
  state.compared.add(key);
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
