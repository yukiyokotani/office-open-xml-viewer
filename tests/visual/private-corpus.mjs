import {
  existsSync,
  lstatSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  realpathSync,
  unlinkSync,
  writeFileSync,
} from 'node:fs';
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { isAbsolute, join, parse, relative, resolve, sep } from 'node:path';
import { PNG } from 'pngjs';

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
  const changed = [...new Set([...tracked, ...untracked])];
  if (changed.length === 0) return;
  const harnessBootstrap = process.env.VRT_ALLOW_HARNESS_CHANGES === '1'
    && changed.every((path) => VRT_HARNESS_PATHS.has(path));
  if (!harnessBootstrap) {
    throw new Error(
      `${label} requires a clean renderer checkout; changed paths at ${root}: ${changed.join(', ')}`,
    );
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
// `private` is the local, gitignored corpus under `public/private/<format>/`.
// Both use the same exact-pixel previous-renderer oracle and the same manifest
// binding; they differ only in where the inputs live and where their images go.
const CORPORA = {
  demo: {
    publicPrefix: '',
    outputPrefix: '',
    manifestDirectory: 'demo',
    stemPattern: /^demo\/[^/\\]+$/,
  },
  // The demo decks rendered through PptxScrollViewer: same inputs as `demo`,
  // separate outputs and manifests.
  'demo-scroll': {
    publicPrefix: '',
    outputPrefix: 'scroll-viewer/',
    manifestDirectory: 'scroll-viewer-demo',
    stemPattern: /^demo\/[^/\\]+$/,
  },
  private: {
    publicPrefix: 'private/',
    outputPrefix: 'private-corpus/',
    manifestDirectory: 'private-corpus',
    stemPattern: /^(docx|xlsx|pptx)\/[^/\\]+$/,
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
  const directory = corpus === 'demo' ? 'demo' : format;
  const { publicPrefix } = corpusConfig(corpus);
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

function sha256(buffer) {
  return createHash('sha256').update(buffer).digest('hex');
}

function corpusFiles(corpus, files) {
  return files.map((name) => ({
    name,
    sha256: createHash('sha256')
      .update(readFileSync(`public/${selfVrtInputPath({ corpus, file: name })}`))
      .digest('hex'),
  }));
}

function readBaselineJson(location, path) {
  const verified = verifiedBaselinePath(location, path);
  if (!verified) throw new Error(`missing previous-renderer manifest: ${path}`);
  return JSON.parse(readFileSync(verified, 'utf8'));
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
 * snapshot from silently becoming its own regression oracle. */
export function prepareSelfVrtCorpus({ corpus, format, files, snapshot }) {
  if (files.length === 0) {
    throw new Error(`${format} ${corpus} corpus is empty; zero-test self-VRT is not coverage`);
  }
  const location = baselineLocation({ snapshot });
  const root = join(location.directory, corpusConfig(corpus).manifestDirectory);
  const path = join(root, 'manifest.json');
  const manifest = {
    schemaVersion: SCHEMA_VERSION,
    format,
    baselineRevision: baselineRevision(snapshot),
    checkout: location.checkout,
    files: corpusFiles(corpus, files),
  };
  if (snapshot) {
    mkdirSync(root, { recursive: true });
    writeFileSync(path, `${JSON.stringify(manifest, null, 2)}\n`);
  } else {
    assertExactManifest(readBaselineJson(location, path), manifest, path);
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

/** Verify that the baseline contains exactly the complete page/sheet/slide set
 * reported by the previous renderer, each image unchanged since that
 * checkout's snapshot run recorded its SHA-256. An item-count reduction, a
 * stale extra PNG or a replaced image is therefore a hard failure. */
export function verifySelfVrtItemManifest({
  corpus,
  format,
  stem,
  itemKind,
  itemCount,
  snapshot,
}) {
  const location = baselineLocation({ snapshot });
  const directory = join(location.directory, `${corpusConfig(corpus).outputPrefix}${stem}`);
  const path = join(directory, 'manifest.json');
  const names = Array.from({ length: itemCount }, (_, index) => `${itemKind}-${index + 1}.png`);
  const header = {
    schemaVersion: SCHEMA_VERSION,
    format,
    baselineRevision: baselineRevision(snapshot),
    checkout: location.checkout,
    itemKind,
    itemCount,
  };
  if (snapshot) {
    // A previous snapshot in this checkout may have rendered more items.
    for (const stale of sortedItems(readdirSync(directory), itemKind)) {
      if (!names.includes(stale)) unlinkSync(join(directory, stale));
    }
    const present = sortedItems(readdirSync(directory), itemKind);
    if (JSON.stringify(present) !== JSON.stringify(names)) {
      throw new Error(
        `snapshot item set mismatch at ${directory}\n`
        + `expected ${JSON.stringify(names)}\nreceived ${JSON.stringify(present)}`,
      );
    }
    const items = names.map((name) => ({ name, sha256: sha256(readFileSync(join(directory, name))) }));
    writeFileSync(path, `${JSON.stringify({ ...header, items }, null, 2)}\n`);
    return;
  }
  const recorded = readBaselineJson(location, path);
  const { items: recordedItems, ...recordedHeader } = recorded;
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
    if (!verified || sha256(readFileSync(verified)) !== expected) {
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
  mkdirSync(outputDirectory, { recursive: true });
  writeFileSync(`${outputDirectory}/${key}`, actual);
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
  const baseline = readFileSync(verified);
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
