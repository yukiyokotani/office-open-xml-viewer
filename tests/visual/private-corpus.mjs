import { existsSync, mkdirSync, readFileSync, readdirSync, writeFileSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { PNG } from 'pngjs';

const SCHEMA_VERSION = 1;

// Bootstrap-only exception for running the current VRT harness against an old
// renderer commit that predates the harness. No renderer/parser source is
// allowed in this set. Future baselines should be clean and need no exception.
const VRT_HARNESS_PATHS = new Set([
  'package.json',
  'packages/docx/package.json',
  'packages/docx/playwright.config.ts',
  'packages/docx/tests/visual/fixture.html',
  'packages/docx/tests/visual/stable-canvas-render.mjs',
  'packages/docx/tests/visual/visual.spec.ts',
  'packages/xlsx/package.json',
  'packages/xlsx/playwright.config.ts',
  'packages/xlsx/tests/visual/fixture.html',
  'packages/xlsx/tests/visual/visual.spec.ts',
  'packages/pptx/package.json',
  'packages/pptx/playwright.config.ts',
  'packages/pptx/tests/visual/fixture.html',
  'packages/pptx/tests/visual/visual.spec.ts',
  'tests/visual/private-corpus.mjs',
]);

function gitRevision(revision) {
  return execFileSync('git', ['rev-parse', `${revision}^{commit}`], {
    encoding: 'utf8',
  }).trim();
}

function baselineRevision(snapshot = false) {
  const revision = process.env.VRT_BASELINE_REVISION?.trim();
  if (!revision) {
    throw new Error('VRT_BASELINE_REVISION is required for private corpus self-VRT');
  }
  const resolved = gitRevision(revision);
  if (snapshot) {
    const checkout = gitRevision('HEAD');
    if (checkout !== resolved) {
      throw new Error(
        `private self-VRT snapshot checkout mismatch: expected ${resolved}, running ${checkout}`,
      );
    }
    const root = execFileSync('git', ['rev-parse', '--show-toplevel'], { encoding: 'utf8' }).trim();
    const tracked = execFileSync(
      'git',
      ['diff', '--name-only', 'HEAD'],
      { cwd: root, encoding: 'utf8' },
    ).trim().split('\n').filter(Boolean);
    const untracked = execFileSync(
      'git',
      ['ls-files', '--others', '--exclude-standard'],
      { cwd: root, encoding: 'utf8' },
    ).trim().split('\n').filter(Boolean).filter((path) =>
      !/(^|\/)node_modules(?:\/|$)/.test(path)
      && !/^packages\/(docx|xlsx|pptx)\/public\/private(?:\/|$)/.test(path));
    const changed = [...new Set([...tracked, ...untracked])];
    if (changed.length > 0) {
      const harnessBootstrap = process.env.VRT_ALLOW_HARNESS_CHANGES === '1'
        && changed.every((path) => VRT_HARNESS_PATHS.has(path));
      if (!harnessBootstrap) {
        throw new Error(
          'private self-VRT snapshot requires a clean renderer checkout; changed paths: '
          + changed.join(', '),
        );
      }
    }
  }
  return resolved;
}

function corpusFiles(files) {
  return files.map((name) => ({
    name,
    sha256: createHash('sha256').update(readFileSync(`public/private/${name}`)).digest('hex'),
  }));
}

function readJson(path) {
  if (!existsSync(path)) throw new Error(`missing previous-renderer manifest: ${path}`);
  return JSON.parse(readFileSync(path, 'utf8'));
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
export function preparePrivateCorpus({ format, files, snapshot }) {
  if (files.length === 0) {
    throw new Error(`${format} private corpus is empty; zero-test self-VRT is not coverage`);
  }
  const root = 'tests/visual/baseline/private-corpus';
  const path = `${root}/manifest.json`;
  const manifest = {
    schemaVersion: SCHEMA_VERSION,
    format,
    baselineRevision: baselineRevision(snapshot),
    files: corpusFiles(files),
  };
  if (snapshot) {
    mkdirSync(root, { recursive: true });
    writeFileSync(path, `${JSON.stringify(manifest, null, 2)}\n`);
  } else {
    assertExactManifest(readJson(path), manifest, path);
  }
}

/** Verify that the baseline contains exactly the complete page/sheet/slide set
 * reported by the previous renderer. An item-count reduction or stale extra PNG
 * is therefore a hard failure instead of an ignored file. */
export function verifyPrivateItemManifest({
  format,
  stem,
  itemKind,
  itemCount,
  snapshot,
}) {
  const directory = `tests/visual/baseline/private-corpus/${stem}`;
  const path = `${directory}/manifest.json`;
  const items = Array.from({ length: itemCount }, (_, index) => `${itemKind}-${index + 1}.png`);
  const manifest = {
    schemaVersion: SCHEMA_VERSION,
    format,
    baselineRevision: baselineRevision(snapshot),
    itemKind,
    itemCount,
    items,
  };
  if (snapshot) {
    mkdirSync(directory, { recursive: true });
    writeFileSync(path, `${JSON.stringify(manifest, null, 2)}\n`);
    return;
  }
  assertExactManifest(readJson(path), manifest, path);
  const actualItems = readdirSync(directory)
    .filter((file) => new RegExp(`^${itemKind}-\\d+\\.png$`).test(file))
    .sort((left, right) => left.localeCompare(right, undefined, { numeric: true }));
  if (JSON.stringify(actualItems) !== JSON.stringify(items)) {
    throw new Error(
      `previous-renderer item set mismatch at ${directory}\n`
      + `expected ${JSON.stringify(items)}\nreceived ${JSON.stringify(actualItems)}`,
    );
  }
}

/** Persist the candidate artifact and return a diagnostic instead of throwing,
 * so one changed item cannot prevent later items in the same document from
 * being rendered and compared. */
export function captureOrComparePrivateItem({ stem, itemKind, itemIndex, actual, snapshot }) {
  const key = `${itemKind}-${itemIndex + 1}.png`;
  const outputDirectory = `tests/visual/${snapshot ? 'baseline' : 'screenshots'}/private-corpus/${stem}`;
  mkdirSync(outputDirectory, { recursive: true });
  writeFileSync(`${outputDirectory}/${key}`, actual);
  if (snapshot) return null;
  const baselinePath = `tests/visual/baseline/private-corpus/${stem}/${key}`;
  if (!existsSync(baselinePath)) return `missing previous-renderer baseline: ${baselinePath}`;
  return pngPixelsEqual(actual, readFileSync(baselinePath))
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
