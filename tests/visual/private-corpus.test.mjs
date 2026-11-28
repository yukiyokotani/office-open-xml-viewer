import assert from 'node:assert/strict';
import { execFileSync, spawnSync } from 'node:child_process';
import {
  cpSync,
  existsSync,
  linkSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  readdirSync,
  realpathSync,
  rmSync,
  symlinkSync,
  unlinkSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import test from 'node:test';
import { PNG } from 'pngjs';
import { copyCorpusInputs } from './copy-corpus-inputs.mjs';
import {
  captureOrCompareSelfVrtItem,
  clearSelfVrtCandidateOutput,
  prepareSelfVrtCorpus,
  harnessBootstrapDiffViolations,
  lockfileBootstrapDiffViolations,
  pngPixelsEqual,
  routeSelfVrtInput,
  selfVrtBaselineRoot,
  selfVrtCorpusFiles,
  selfVrtInputPath,
  verifyBaselineCheckout,
  verifySelfVrtItemManifest,
} from './private-corpus.mjs';

test('private corpus self-VRT compares decoded pixels, not encoder bytes', () => {
  const image = new PNG({ width: 2, height: 1 });
  image.data.set([255, 0, 0, 255, 0, 0, 255, 255]);
  const fast = PNG.sync.write(image, { deflateLevel: 0 });
  const compact = PNG.sync.write(image, { deflateLevel: 9 });

  assert.equal(fast.equals(compact), false);
  assert.equal(pngPixelsEqual(fast, compact), true);
});

test('private corpus self-VRT rejects a one-channel pixel change', () => {
  const left = new PNG({ width: 1, height: 1 });
  left.data.set([1, 2, 3, 255]);
  const right = new PNG({ width: 1, height: 1 });
  right.data.set([1, 2, 4, 255]);

  assert.equal(pngPixelsEqual(PNG.sync.write(left), PNG.sync.write(right)), false);
});

test('candidate capture discards stale pages without following local evidence symlinks', () => {
  const root = mkdtempSync(join(realpathSync(tmpdir()), 'ooxml-private-vrt-output-'));
  try {
    const directory = join(root, 'docx', 'case');
    mkdirSync(directory, { recursive: true });
    writeFileSync(join(directory, 'page-29.png'), 'stale');
    writeFileSync(join(directory, 'notes.json'), 'retain');
    clearSelfVrtCandidateOutput({
      corpus: 'private', stem: 'docx/case', itemKind: 'page', outputRoot: root,
    });
    assert.equal(existsSync(join(directory, 'page-29.png')), false);
    assert.equal(readFileSync(join(directory, 'notes.json'), 'utf8'), 'retain');

    symlinkSync(directory, join(root, 'docx', 'linked'), 'dir');
    assert.throws(() => clearSelfVrtCandidateOutput({
      corpus: 'private', stem: 'docx/linked', itemKind: 'page', outputRoot: root,
    }), /symlinked self-VRT output/);
    assert.equal(readFileSync(join(directory, 'notes.json'), 'utf8'), 'retain');
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
});

test('harness bootstrap allows only test-renderer alias lines in Vite configs', () => {
  const header = [
    'diff --git a/packages/pptx/vite.config.ts b/packages/pptx/vite.config.ts',
    '--- a/packages/pptx/vite.config.ts',
    '+++ b/packages/pptx/vite.config.ts',
    '@@ -20,0 +21 @@',
  ];
  const alias = "+      '@ooxml-test-chart-ex-renderer': resolve(dirname, '../../src/chart-ex.ts'),";
  assert.deepEqual(harnessBootstrapDiffViolations([...header, alias].join('\n')), []);
  assert.deepEqual(
    harnessBootstrapDiffViolations([
      ...header,
      alias,
      "+  define: { __OOXML_MODEL_SOURCES__: 'false' },",
    ].join('\n')),
    ["+  define: { __OOXML_MODEL_SOURCES__: 'false' },"],
  );
});

test('harness bootstrap allows only @fontsource additions in pnpm-lock.yaml', () => {
  const addition = [
    'diff --git a/pnpm-lock.yaml b/pnpm-lock.yaml',
    '@@ -262,0 +263,3 @@ importers:',
    "+      '@fontsource/roboto':",
    '+        specifier: ^5.3.0',
    '+        version: 5.3.0',
    '@@ -1080,0 +1084,3 @@ packages:',
    "+  '@fontsource/roboto@5.3.0':",
    '+    resolution: {integrity: sha512-AAAA==}',
    '+',
    '@@ -5320,0 +5327,2 @@ snapshots:',
    "+  '@fontsource/roboto@5.3.0': {}",
    '+',
  ];
  assert.deepEqual(lockfileBootstrapDiffViolations(addition.join('\n')), []);
  // A renderer dependency bump is its own hunk, without a font package.
  const bump = ['@@ -300 +300 @@ importers:', '-        version: 0.18.5', '+        version: 0.18.6'];
  assert.deepEqual(lockfileBootstrapDiffViolations([...addition, ...bump].join('\n')), bump.slice(1));
  // Anything but the package's own structure inside a font hunk is rejected.
  assert.deepEqual(
    lockfileBootstrapDiffViolations(['@@ -1 +1 @@', "+  '@fontsource/roboto@5.3.0':", '+    dependencies:'].join('\n')),
    ['+    dependencies:'],
  );
});

test('self-VRT corpora keep public specimen and fast-tier stems in their own trees', () => {
  assert.equal(selfVrtInputPath({ corpus: 'corpus', file: 'corpus/docx-001.docx' }), 'corpus/docx-001.docx');
  assert.equal(selfVrtInputPath({ corpus: 'private-fast', file: 'pptx/deck.pptx' }), 'private/pptx/deck.pptx');
  assert.throws(() => clearSelfVrtCandidateOutput({
    corpus: 'corpus', stem: 'demo/sample-1', itemKind: 'page',
  }), /invalid self-VRT output identity/);
  assert.throws(() => clearSelfVrtCandidateOutput({
    corpus: 'private-fast', stem: 'corpus/docx-001', itemKind: 'page',
  }), /invalid self-VRT output identity/);
});

test('self-VRT corpora keep demo and private stems in their own namespaces', () => {
  assert.equal(selfVrtInputPath({ corpus: 'demo', file: 'demo/sample-1.docx' }), 'demo/sample-1.docx');
  assert.equal(selfVrtInputPath({ corpus: 'private', file: 'docx/case.docx' }), 'private/docx/case.docx');
  assert.throws(() => clearSelfVrtCandidateOutput({
    corpus: 'demo', stem: 'docx/case', itemKind: 'page',
  }), /invalid self-VRT output identity/);
  assert.throws(() => clearSelfVrtCandidateOutput({
    corpus: 'private', stem: 'demo/sample-1', itemKind: 'page',
  }), /invalid self-VRT output identity/);
  assert.throws(() => selfVrtInputPath({ corpus: 'public', file: 'x' }), /unknown self-VRT corpus/);
});

function git(cwd, ...args) {
  return execFileSync('git', args, { cwd, encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] }).trim();
}

/** A throwaway repository with two commits; HEAD is left at the first. */
function baselineRepository() {
  const root = mkdtempSync(join(realpathSync(tmpdir()), 'ooxml-vrt-baseline-checkout-'));
  git(root, 'init', '-q');
  git(root, 'config', 'user.email', 'vrt@example.invalid');
  git(root, 'config', 'user.name', 'vrt');
  writeFileSync(join(root, '.gitignore'), 'tests/visual/baseline/\n');
  writeFileSync(join(root, 'renderer.js'), 'export const version = 1;\n');
  git(root, 'add', '.');
  git(root, 'commit', '-q', '-m', 'first');
  const first = git(root, 'rev-parse', 'HEAD');
  writeFileSync(join(root, 'renderer.js'), 'export const version = 2;\n');
  git(root, 'commit', '-q', '-am', 'second');
  const second = git(root, 'rev-parse', 'HEAD');
  git(root, 'checkout', '-q', '--detach', first);
  return { root, first, second };
}

function withEnv(values, run) {
  const previous = Object.fromEntries(Object.keys(values).map((key) => [key, process.env[key]]));
  try {
    for (const [key, value] of Object.entries(values)) {
      if (value === undefined) delete process.env[key];
      else process.env[key] = value;
    }
    return run();
  } finally {
    for (const [key, value] of Object.entries(previous)) {
      if (value === undefined) delete process.env[key];
      else process.env[key] = value;
    }
  }
}

test('a baseline checkout must be a clean Git checkout at the baseline revision', () => {
  const { root, first, second } = baselineRepository();
  const plain = mkdtempSync(join(realpathSync(tmpdir()), 'ooxml-vrt-baseline-plain-'));
  try {
    withEnv({ VRT_ALLOW_HARNESS_CHANGES: undefined }, () => {
      // The previous renderer's own checkout, with ignored baseline images.
      mkdirSync(join(root, 'tests/visual/baseline'), { recursive: true });
      writeFileSync(join(root, 'tests/visual/baseline/manifest.json'), '{}');
      verifyBaselineCheckout({ checkout: root, revision: first });

      // A plain directory holding copied images is not a renderer checkout.
      mkdirSync(join(plain, 'tests/visual/baseline'), { recursive: true });
      assert.throws(
        () => verifyBaselineCheckout({ checkout: plain, revision: first }),
        /not a Git checkout/,
      );
      // Nor is a directory inside a checkout.
      assert.throws(
        () => verifyBaselineCheckout({ checkout: join(root, 'tests'), revision: first }),
        /must be the root of a Git checkout/,
      );
      // The checkout must be at the baseline revision.
      assert.throws(
        () => verifyBaselineCheckout({ checkout: root, revision: second }),
        /checkout mismatch/,
      );
      assert.throws(
        () => verifyBaselineCheckout({ checkout: root, revision: 'HEAD' }),
        /full commit SHA/,
      );

      // A tracked modification means the images need not come from that revision.
      writeFileSync(join(root, 'renderer.js'), 'export const version = 3;\n');
      assert.throws(
        () => verifyBaselineCheckout({ checkout: root, revision: first }),
        /clean renderer checkout; changed paths at .*: renderer\.js/,
      );
      git(root, 'checkout', '-q', '--', 'renderer.js');
      // So does an untracked source file.
      writeFileSync(join(root, 'patch.js'), 'export {};\n');
      assert.throws(
        () => verifyBaselineCheckout({ checkout: root, revision: first }),
        /clean renderer checkout; changed paths at .*: patch\.js/,
      );
      rmSync(join(root, 'patch.js'));
      verifyBaselineCheckout({ checkout: root, revision: first });
    });

    // The harness bootstrap never excuses a renderer change.
    withEnv({ VRT_ALLOW_HARNESS_CHANGES: '1' }, () => {
      writeFileSync(join(root, 'renderer.js'), 'export const version = 3;\n');
      assert.throws(
        () => verifyBaselineCheckout({ checkout: root, revision: first }),
        /clean renderer checkout/,
      );
    });
  } finally {
    rmSync(root, { recursive: true, force: true });
    rmSync(plain, { recursive: true, force: true });
  }
});

test('comparison runs read only a verified previous-renderer checkout', () => {
  const { root, first } = baselineRepository();
  const plain = mkdtempSync(join(realpathSync(tmpdir()), 'ooxml-vrt-baseline-plain-'));
  // This test runs from the repository root, whose HEAD is a real commit.
  const head = git(process.cwd(), 'rev-parse', 'HEAD');
  try {
    withEnv({ VRT_BASELINE_CHECKOUT: undefined, VRT_BASELINE_REVISION: head }, () => {
      assert.equal(selfVrtBaselineRoot({ snapshot: true }), resolve('tests/visual/baseline'));
      assert.throws(() => selfVrtBaselineRoot(), /VRT_BASELINE_CHECKOUT is required/);
    });
    withEnv({ VRT_BASELINE_CHECKOUT: plain, VRT_BASELINE_REVISION: head }, () => {
      assert.throws(() => selfVrtBaselineRoot({ snapshot: true }), /capture snapshots in the baseline checkout/);
      assert.throws(() => selfVrtBaselineRoot(), /not a Git checkout/);
    });
    withEnv({ VRT_BASELINE_CHECKOUT: root, VRT_BASELINE_REVISION: head }, () => {
      // A real checkout, but not at the revision the manifests are bound to.
      assert.notEqual(first, head);
      assert.throws(() => selfVrtBaselineRoot(), /checkout mismatch/);
    });
    withEnv({ VRT_BASELINE_CHECKOUT: process.cwd(), VRT_BASELINE_REVISION: head }, () => {
      assert.throws(() => selfVrtBaselineRoot(), /not this one/);
    });
  } finally {
    rmSync(root, { recursive: true, force: true });
    rmSync(plain, { recursive: true, force: true });
  }
});

// ── End-to-end baseline binding ──────────────────────────────────────────────
// A previous-renderer checkout `base` snapshots one demo slide; a candidate
// clone at the same commit compares against it. Each bypass below must fail.

function png(value) {
  const image = new PNG({ width: 1, height: 1 });
  image.data.set([value, 0, 0, 255]);
  return PNG.sync.write(image);
}

function selfVrtFixture() {
  const scratch = mkdtempSync(join(realpathSync(tmpdir()), 'ooxml-vrt-e2e-'));
  const base = join(scratch, 'base');
  mkdirSync(join(base, 'packages/pptx/public/demo'), { recursive: true });
  git(base, 'init', '-q');
  git(base, 'config', 'user.email', 'vrt@example.invalid');
  git(base, 'config', 'user.name', 'vrt');
  writeFileSync(
    join(base, '.gitignore'),
    'packages/*/tests/visual/baseline/\npackages/*/tests/visual/screenshots/\n',
  );
  writeFileSync(join(base, 'packages/pptx/public/demo/deck.pptx'), 'deck');
  mkdirSync(join(base, 'packages/pptx/public/corpus'));
  writeFileSync(join(base, 'packages/pptx/public/corpus/old.pptx'), 'v1');
  writeFileSync(join(base, 'renderer.js'), 'export const version = 1;\n');
  git(base, 'add', '.');
  git(base, 'commit', '-q', '-m', 'renderer');
  const revision = git(base, 'rev-parse', 'HEAD');
  const candidate = join(scratch, 'candidate');
  git(scratch, 'clone', '-q', base, candidate);
  return { scratch, base, candidate, revision };
}

const DEMO = { corpus: 'demo', format: 'pptx', stem: 'demo/deck', itemKind: 'slide' };

async function inPackage(checkout, env, run) {
  const previous = process.cwd();
  const names = ['VRT_ALLOW_HARNESS_CHANGES', 'VRT_ALLOW_CORPUS_INPUT_CHANGES', ...Object.keys(env)];
  const saved = Object.fromEntries(names.map((key) => [key, process.env[key]]));
  const assign = (values) => {
    for (const [key, value] of Object.entries(values)) {
      if (value === undefined) delete process.env[key];
      else process.env[key] = value;
    }
  };
  process.chdir(join(checkout, 'packages/pptx'));
  assign({ VRT_ALLOW_HARNESS_CHANGES: undefined, VRT_ALLOW_CORPUS_INPUT_CHANGES: undefined, ...env });
  try {
    return await run();
  } finally {
    assign(saved);
    process.chdir(previous);
  }
}

/** Load a prepared input the way a fixture does: route it on a stand-in for a
 * Playwright page, request the returned path, and follow the route to the
 * harness endpoint. Rejects when the endpoint refuses the input. */
async function loadInput(corpus, file) {
  let matches;
  let handle;
  const path = await routeSelfVrtInput({
    route: async (matcher, handler) => {
      matches = matcher;
      handle = handler;
    },
  }, { corpus, file });
  assert.equal(matches(new URL(`http://fixture.invalid/${path}`)), true);
  let body;
  await handle({
    continue: async ({ url }) => {
      const response = await fetch(url);
      assert.equal(response.status, 200);
      body = Buffer.from(await response.arrayBuffer());
    },
  });
  return { path, body, matches };
}

function snapshot(checkout, revision, value = 7) {
  return inPackage(checkout, { VRT_BASELINE_CHECKOUT: undefined, VRT_BASELINE_REVISION: revision }, async () => {
    prepareSelfVrtCorpus({ ...DEMO, files: ['demo/deck.pptx'], snapshot: true });
    await loadInput('demo', 'demo/deck.pptx');
    assert.equal(captureOrCompareSelfVrtItem({ ...DEMO, itemIndex: 0, actual: png(value), snapshot: true }), null);
    verifySelfVrtItemManifest({ ...DEMO, itemCount: 1, snapshot: true });
  });
}

function compare({ candidate, base, revision }, value = 7) {
  return inPackage(candidate, { VRT_BASELINE_CHECKOUT: base, VRT_BASELINE_REVISION: revision }, async () => {
    prepareSelfVrtCorpus({ ...DEMO, files: ['demo/deck.pptx'], snapshot: false });
    await loadInput('demo', 'demo/deck.pptx');
    const difference = captureOrCompareSelfVrtItem({
      ...DEMO, itemIndex: 0, actual: png(value), snapshot: false,
    });
    verifySelfVrtItemManifest({ ...DEMO, itemCount: 1, snapshot: false });
    return difference;
  });
}

test('a verified baseline checkout detects a changed slide and accepts an unchanged one', async () => {
  const fixture = selfVrtFixture();
  try {
    await snapshot(fixture.base, fixture.revision);
    assert.equal(await compare(fixture), null);
    assert.match(await compare(fixture, 8), /differs from the previous renderer/);
  } finally {
    rmSync(fixture.scratch, { recursive: true, force: true });
  }
});

test('bypass: a symlink in the baseline cannot redirect reads to other images', async () => {
  const fixture = selfVrtFixture();
  try {
    await snapshot(fixture.base, fixture.revision);
    const baselineDemo = join(fixture.base, 'packages/pptx/tests/visual/baseline/demo');
    const screenshots = join(fixture.candidate, 'packages/pptx/tests/visual/screenshots/demo');
    mkdirSync(screenshots, { recursive: true });
    cpSync(join(baselineDemo, 'deck'), join(screenshots, 'deck'), { recursive: true });
    cpSync(join(baselineDemo, 'manifest.json'), join(screenshots, 'manifest.json'));

    // The whole corpus directory linked to the candidate's own output.
    rmSync(baselineDemo, { recursive: true });
    symlinkSync(screenshots, baselineDemo, 'dir');
    await assert.rejects(compare(fixture, 8), /must not contain symlinks/);

    // A single image linked to the candidate's output.
    unlinkSync(baselineDemo);
    cpSync(screenshots, baselineDemo, { recursive: true });
    rmSync(join(baselineDemo, 'deck/slide-1.png'));
    symlinkSync(join(screenshots, 'deck/slide-1.png'), join(baselineDemo, 'deck/slide-1.png'));
    await assert.rejects(compare(fixture, 8), /must not contain symlinks/);
  } finally {
    rmSync(fixture.scratch, { recursive: true, force: true });
  }
});

test('bypass: images copied from another checkout or replaced after the snapshot are rejected', async () => {
  const fixture = selfVrtFixture();
  const other = join(fixture.scratch, 'other');
  try {
    git(fixture.scratch, 'clone', '-q', fixture.base, other);
    await snapshot(other, fixture.revision, 8);
    // Real files, but written by another checkout's snapshot run.
    cpSync(
      join(other, 'packages/pptx/tests/visual/baseline'),
      join(fixture.base, 'packages/pptx/tests/visual/baseline'),
      { recursive: true },
    );
    await assert.rejects(compare(fixture, 8), /manifest mismatch[\s\S]*"checkout"/);

    // The genuine snapshot, with its image replaced afterwards.
    await snapshot(fixture.base, fixture.revision);
    writeFileSync(join(fixture.base, 'packages/pptx/tests/visual/baseline/demo/deck/slide-1.png'), png(8));
    await assert.rejects(compare(fixture, 8), /changed since its snapshot/);
  } finally {
    rmSync(fixture.scratch, { recursive: true, force: true });
  }
});

test('bypass: GIT_* variables cannot make a plain copy answer as a checkout', () => {
  const fixture = selfVrtFixture();
  const plain = join(fixture.scratch, 'plain');
  try {
    cpSync(fixture.base, plain, { recursive: true, filter: (path) => !path.includes(`${join(fixture.base, '.git')}`) });
    assert.equal(existsSync(join(plain, '.git')), false);
    withEnv({
      GIT_DIR: join(fixture.base, '.git'),
      GIT_WORK_TREE: plain,
      GIT_INDEX_FILE: join(fixture.base, '.git/index'),
      GIT_CEILING_DIRECTORIES: fixture.scratch,
      VRT_ALLOW_HARNESS_CHANGES: undefined,
    }, () => {
      assert.throws(
        () => verifyBaselineCheckout({ checkout: plain, revision: fixture.revision }),
        /not a Git checkout/,
      );
    });
  } finally {
    rmSync(fixture.scratch, { recursive: true, force: true });
  }
});

test('bypass: a baseline checkout changed mid-run is caught at the next item', async () => {
  const fixture = selfVrtFixture();
  try {
    await snapshot(fixture.base, fixture.revision);
    await inPackage(
      fixture.candidate,
      { VRT_BASELINE_CHECKOUT: fixture.base, VRT_BASELINE_REVISION: fixture.revision },
      () => {
        prepareSelfVrtCorpus({ ...DEMO, files: ['demo/deck.pptx'], snapshot: false });
        writeFileSync(join(fixture.base, 'renderer.js'), 'export const version = 2;\n');
        assert.throws(
          () => captureOrCompareSelfVrtItem({ ...DEMO, itemIndex: 0, actual: png(7), snapshot: false }),
          /clean renderer checkout; changed paths at .*: renderer\.js/,
        );
        git(fixture.base, 'checkout', '-q', '--', 'renderer.js');
        git(fixture.base, 'commit', '-q', '--allow-empty', '-m', 'moved');
        assert.throws(
          () => captureOrCompareSelfVrtItem({ ...DEMO, itemIndex: 0, actual: png(7), snapshot: false }),
          /checkout mismatch/,
        );
      },
    );
  } finally {
    rmSync(fixture.scratch, { recursive: true, force: true });
  }
});

// ── Fast tier ────────────────────────────────────────────────────────────────
// `private-fast` renders a selected subset; its baseline is bound to the
// selection digest and holds exactly the selected items.

const FAST = { corpus: 'private-fast', format: 'pptx', stem: 'pptx/deck', itemKind: 'slide' };

function fastFixture() {
  const fixture = selfVrtFixture();
  for (const checkout of [fixture.base, fixture.candidate]) {
    mkdirSync(join(checkout, 'packages/pptx/public/private/pptx'), { recursive: true });
    writeFileSync(join(checkout, 'packages/pptx/public/private/pptx/deck.pptx'), 'private deck');
  }
  return fixture;
}

function fastRun(checkout, env, { snapshot, digest = 'selection-a', indices = [1, 3], value = 7 }) {
  return inPackage(checkout, env, async () => {
    prepareSelfVrtCorpus({ ...FAST, files: ['pptx/deck.pptx'], snapshot, selectionDigest: digest });
    await loadInput('private-fast', 'pptx/deck.pptx');
    const differences = indices.map((itemIndex) => captureOrCompareSelfVrtItem({
      ...FAST, itemIndex, actual: png(value), snapshot,
    })).filter(Boolean);
    verifySelfVrtItemManifest({ ...FAST, itemCount: 4, snapshot, itemIndices: indices });
    return differences;
  });
}

test('fast tier compares exactly the selected items under the same selection', async () => {
  const fixture = fastFixture();
  const { base, candidate, revision } = fixture;
  const snapshotEnv = { VRT_BASELINE_CHECKOUT: undefined, VRT_BASELINE_REVISION: revision };
  const compareEnv = { VRT_BASELINE_CHECKOUT: base, VRT_BASELINE_REVISION: revision };
  try {
    assert.deepEqual(await fastRun(base, snapshotEnv, { snapshot: true }), []);
    assert.ok(existsSync(join(base, 'packages/pptx/tests/visual/baseline/private-corpus-fast/pptx/deck/slide-4.png')));
    assert.ok(!existsSync(join(base, 'packages/pptx/tests/visual/baseline/private-corpus-fast/pptx/deck/slide-1.png')));
    assert.deepEqual(await fastRun(candidate, compareEnv, { snapshot: false }), []);
    assert.equal((await fastRun(candidate, compareEnv, { snapshot: false, value: 8 })).length, 2);
    // A different selection is a different oracle.
    await assert.rejects(fastRun(candidate, compareEnv, { snapshot: false, digest: 'selection-b' }), /manifest mismatch/);
    await assert.rejects(fastRun(candidate, compareEnv, { snapshot: false, indices: [1] }), /item set mismatch/);
  } finally {
    rmSync(fixture.scratch, { recursive: true, force: true });
  }
});

test('fast tier and full corpora refuse each other\'s selection arguments', async () => {
  const fixture = fastFixture();
  const env = { VRT_BASELINE_CHECKOUT: undefined, VRT_BASELINE_REVISION: fixture.revision };
  try {
    await inPackage(fixture.base, env, () => {
      assert.throws(() => prepareSelfVrtCorpus({ ...FAST, files: ['pptx/deck.pptx'], snapshot: true }),
        /requires a selection digest/);
      assert.throws(() => prepareSelfVrtCorpus({
        ...FAST, corpus: 'private', files: ['pptx/deck.pptx'], snapshot: true, selectionDigest: 'x',
      }), /does not take a selection digest/);
      assert.throws(() => verifySelfVrtItemManifest({ ...FAST, itemCount: 4, snapshot: true }),
        /requires the selected item indices/);
      assert.throws(() => verifySelfVrtItemManifest({
        ...FAST, itemCount: 4, snapshot: true, itemIndices: [3, 1],
      }), /invalid slide selection/);
      assert.throws(() => verifySelfVrtItemManifest({
        ...FAST, corpus: 'private', itemCount: 4, snapshot: true, itemIndices: [1],
      }), /does not take item indices/);
    });
  } finally {
    rmSync(fixture.scratch, { recursive: true, force: true });
  }
});

// ── Input binding ────────────────────────────────────────────────────────────
// A fixture gets its document from the harness endpoint, never by file name.

const PRIVATE = { corpus: 'private', format: 'pptx', stem: 'pptx/deck', itemKind: 'slide' };

test('inputs are addressed by manifest position, so file names cannot alias one another', async () => {
  const fixture = selfVrtFixture();
  const directory = join(fixture.base, 'packages/pptx/public/private/pptx');
  try {
    mkdirSync(directory, { recursive: true });
    // A dev server that decodes the request path would serve one for the other.
    writeFileSync(join(directory, 'A&B.pptx'), 'ampersand');
    writeFileSync(join(directory, 'A%26B.pptx'), 'percent');
    await inPackage(
      fixture.base,
      { VRT_BASELINE_CHECKOUT: undefined, VRT_BASELINE_REVISION: fixture.revision },
      async () => {
        const files = selfVrtCorpusFiles({ corpus: 'private', format: 'pptx' });
        assert.deepEqual([...files].sort(), ['pptx/A%26B.pptx', 'pptx/A&B.pptx']);
        prepareSelfVrtCorpus({ corpus: 'private', format: 'pptx', files, snapshot: true });
        const ampersand = await loadInput('private', 'pptx/A&B.pptx');
        const percent = await loadInput('private', 'pptx/A%26B.pptx');
        assert.match(ampersand.path, /^__selfvrt\/private\/\d+\.pptx$/);
        assert.match(percent.path, /^__selfvrt\/private\/\d+\.pptx$/);
        assert.notEqual(ampersand.path, percent.path);
        assert.equal(ampersand.body.toString(), 'ampersand');
        assert.equal(percent.body.toString(), 'percent');
        // Only the harness URL is routed; a request by name is left alone.
        assert.equal(ampersand.matches(new URL('http://fixture.invalid/private/pptx/A%26B.pptx')), false);
        assert.equal(ampersand.matches(new URL(`http://fixture.invalid/${percent.path}`)), false);
        await assert.rejects(
          routeSelfVrtInput({ route: async () => {} }, { corpus: 'private', file: 'pptx/other.pptx' }),
          /was not prepared for the private corpus/,
        );
      },
    );
  } finally {
    rmSync(fixture.scratch, { recursive: true, force: true });
  }
});

test('a document is finished only if the endpoint served its input with the manifest digest', async () => {
  const fixture = fastFixture();
  const input = join(fixture.base, 'packages/pptx/public/private/pptx/deck.pptx');
  try {
    await inPackage(
      fixture.base,
      { VRT_BASELINE_CHECKOUT: undefined, VRT_BASELINE_REVISION: fixture.revision },
      async () => {
        const finish = () => verifySelfVrtItemManifest({ ...PRIVATE, itemCount: 1, snapshot: true });
        prepareSelfVrtCorpus({ ...PRIVATE, files: ['pptx/deck.pptx'], snapshot: true });
        captureOrCompareSelfVrtItem({ ...PRIVATE, itemIndex: 0, actual: png(7), snapshot: true });
        // The fixture loaded something, but not through the harness.
        assert.throws(finish, /was not loaded through the harness endpoint/);
        // The file no longer is what the corpus manifest recorded.
        writeFileSync(input, 'replaced');
        await assert.rejects(loadInput('private', 'pptx/deck.pptx'), /fetch failed/);
        assert.throws(finish, /deck\.pptx changed after its manifest recorded it/);
        writeFileSync(input, 'private deck');
        assert.equal((await loadInput('private', 'pptx/deck.pptx')).body.toString(), 'private deck');
        finish();
      },
    );
  } finally {
    rmSync(fixture.scratch, { recursive: true, force: true });
  }
});

test('images kept from a snapshot of other input bytes are refused', async () => {
  const fixture = fastFixture();
  const { base, candidate, revision } = fixture;
  const run = (checkout, env, snapshot, render = true) => inPackage(checkout, env, async () => {
    prepareSelfVrtCorpus({ ...PRIVATE, files: ['pptx/deck.pptx'], snapshot });
    if (!render) return null;
    await loadInput('private', 'pptx/deck.pptx');
    const difference = captureOrCompareSelfVrtItem({ ...PRIVATE, itemIndex: 0, actual: png(7), snapshot });
    verifySelfVrtItemManifest({ ...PRIVATE, itemCount: 1, snapshot });
    return difference;
  });
  const snapshotEnv = { VRT_BASELINE_CHECKOUT: undefined, VRT_BASELINE_REVISION: revision };
  const compareEnv = { VRT_BASELINE_CHECKOUT: base, VRT_BASELINE_REVISION: revision };
  try {
    await run(base, snapshotEnv, true);
    assert.equal(await run(candidate, compareEnv, false), null);
    // The input is replaced in both checkouts and the snapshot is started
    // again but does not render this document (a filtered or aborted run):
    // the corpus manifest is current, the document's images are not.
    for (const checkout of [base, candidate]) {
      writeFileSync(join(checkout, 'packages/pptx/public/private/pptx/deck.pptx'), 'private deck, edited');
    }
    await run(base, snapshotEnv, true, false);
    await assert.rejects(run(candidate, compareEnv, false), /manifest mismatch[\s\S]*"input"/);
  } finally {
    rmSync(fixture.scratch, { recursive: true, force: true });
  }
});

// ── Added or rewritten public specimens ──────────────────────────────────────

const SPECIMEN = { corpus: 'corpus', format: 'pptx', itemKind: 'slide' };

/** Snapshot or compare every specimen (one slide each); returns the differences. */
function specimenRun(checkout, env, snapshot, value = 7) {
  return inPackage(checkout, env, async () => {
    const files = selfVrtCorpusFiles({ corpus: 'corpus', format: 'pptx' });
    prepareSelfVrtCorpus({ ...SPECIMEN, files, snapshot });
    const differences = [];
    for (const file of files) {
      const stem = file.slice(0, -'.pptx'.length);
      await loadInput('corpus', file);
      const difference = captureOrCompareSelfVrtItem({
        ...SPECIMEN, stem, itemIndex: 0, actual: png(value), snapshot,
      });
      if (difference) differences.push(difference);
      verifySelfVrtItemManifest({ ...SPECIMEN, stem, itemCount: 1, snapshot });
    }
    return differences;
  });
}

test('specimen copies are admitted only under the flag, and only as present specimen files', () => {
  const fixture = selfVrtFixture();
  const { base, revision } = fixture;
  const corpus = join(base, 'packages/pptx/public/corpus');
  const verify = () => verifyBaselineCheckout({ checkout: base, revision });
  const cases = (flag, run) => withEnv(
    { VRT_ALLOW_CORPUS_INPUT_CHANGES: flag, VRT_ALLOW_HARNESS_CHANGES: undefined }, run,
  );
  try {
    writeFileSync(join(corpus, 'new.pptx'), 'new');
    writeFileSync(join(corpus, 'old.pptx'), 'v2');
    cases(undefined, () => assert.throws(verify, /clean renderer checkout; changed paths at .*old\.pptx.*new\.pptx/));
    cases('1', () => {
      verify();
      // Nothing but a specimen of the folder's own format, directly in it.
      for (const other of ['new.win.pdf', '~$new.pptx', 'new.docx', 'nested/new.pptx', '../demo/new.pptx']) {
        const path = join(corpus, other);
        mkdirSync(join(path, '..'), { recursive: true });
        writeFileSync(path, 'x');
        assert.throws(verify, /clean renderer checkout/, other);
        rmSync(path);
      }
      rmSync(join(corpus, 'nested'), { recursive: true });
      // The allowance never covers the renderer, or a specimen removed from the baseline.
      writeFileSync(join(base, 'renderer.js'), 'export const version = 3;\n');
      assert.throws(verify, /clean renderer checkout; changed paths at .*: renderer\.js$/);
      git(base, 'checkout', '-q', '--', 'renderer.js');
      rmSync(join(corpus, 'old.pptx'));
      assert.throws(verify, /clean renderer checkout; changed paths at .*: packages\/pptx\/public\/corpus\/old\.pptx$/);
      // Nor one replaced by a symlink (a Git type change), even to the same bytes.
      writeFileSync(join(fixture.scratch, 'elsewhere.pptx'), 'v1');
      symlinkSync(join(fixture.scratch, 'elsewhere.pptx'), join(corpus, 'old.pptx'));
      assert.throws(verify, /clean renderer checkout; changed paths at .*: packages\/pptx\/public\/corpus\/old\.pptx$/);
      // Nor one reached through a symlinked corpus folder.
      rmSync(corpus, { recursive: true });
      mkdirSync(join(fixture.scratch, 'elsewhere'));
      writeFileSync(join(fixture.scratch, 'elsewhere/old.pptx'), 'v1');
      symlinkSync(join(fixture.scratch, 'elsewhere'), corpus);
      assert.throws(verify, /clean renderer checkout; changed paths at /);
    });
  } finally {
    rmSync(fixture.scratch, { recursive: true, force: true });
  }
});

test('a specimen the candidate adds or rewrites is compared only against a copy of the same bytes', async () => {
  const fixture = selfVrtFixture();
  const { base, candidate, revision } = fixture;
  const corpus = 'packages/pptx/public/corpus';
  const snapshotEnv = { VRT_BASELINE_CHECKOUT: undefined, VRT_BASELINE_REVISION: revision };
  const compareEnv = { VRT_BASELINE_CHECKOUT: base, VRT_BASELINE_REVISION: revision };
  const allow = { VRT_ALLOW_CORPUS_INPUT_CHANGES: '1' };
  try {
    writeFileSync(join(candidate, corpus, 'old.pptx'), 'v2');
    writeFileSync(join(candidate, corpus, 'new.pptx'), 'new');
    git(candidate, 'add', '.');
    git(candidate, '-c', 'user.email=vrt@example.invalid', '-c', 'user.name=vrt', 'commit', '-q', '-m', 'specimens');

    // Forgotten copies: the baseline rendered another input set.
    await specimenRun(base, snapshotEnv, true);
    await assert.rejects(specimenRun(candidate, compareEnv, false), /manifest mismatch/);

    assert.deepEqual(
      copyCorpusInputs({ source: candidate, destination: base, base: revision }),
      [`${corpus}/new.pptx`, `${corpus}/old.pptx`],
    );
    await assert.rejects(specimenRun(base, snapshotEnv, true), /clean renderer checkout; changed paths/);
    await specimenRun(base, { ...snapshotEnv, ...allow }, true);
    await assert.rejects(specimenRun(candidate, compareEnv, false), /clean renderer checkout; changed paths/);
    assert.deepEqual(await specimenRun(candidate, { ...compareEnv, ...allow }, false), []);
    assert.equal((await specimenRun(candidate, { ...compareEnv, ...allow }, false, 8)).length, 2);

    // A stale copy: the snapshot drew other bytes than the candidate will.
    writeFileSync(join(base, corpus, 'old.pptx'), 'v1.5');
    await specimenRun(base, { ...snapshotEnv, ...allow }, true);
    await assert.rejects(
      specimenRun(candidate, { ...compareEnv, ...allow }, false),
      /manifest mismatch[\s\S]*corpus\/old\.pptx/,
    );
  } finally {
    rmSync(fixture.scratch, { recursive: true, force: true });
  }
});

test('the specimen copy command checks its destination and refuses deletions and renames', () => {
  const fixture = selfVrtFixture();
  const { scratch, base, candidate, revision } = fixture;
  const corpus = 'packages/pptx/public/corpus';
  const outside = join(scratch, 'outside');
  const copy = (destination = base) => copyCorpusInputs({ source: candidate, destination, base: revision });
  const put = (path, bytes) => {
    mkdirSync(join(candidate, path, '..'), { recursive: true });
    writeFileSync(join(candidate, path), bytes);
  };
  try {
    put(`${corpus}/old.pptx`, 'v2');
    put('packages/xlsx/public/corpus/new.xlsx', 'new');
    put('packages/xlsx/public/corpus/new.win.pdf', 'reference');
    put('packages/xlsx/public/corpus/wrong.pptx', 'wrong format');
    put('renderer.js', 'export const version = 2;\n');
    git(candidate, 'add', '-A');
    // The baseline has no xlsx corpus folder yet: it is created.
    assert.deepEqual(copy(), ['packages/pptx/public/corpus/old.pptx', 'packages/xlsx/public/corpus/new.xlsx']);
    assert.equal(readFileSync(join(base, corpus, 'old.pptx'), 'utf8'), 'v2');
    assert.deepEqual(readdirSync(join(base, 'packages/xlsx/public/corpus')), ['new.xlsx']);
    assert.equal(readFileSync(join(base, 'renderer.js'), 'utf8'), 'export const version = 1;\n');

    // The destination must be the root of another checkout at the base revision.
    mkdirSync(outside);
    assert.throws(() => copy(outside), /baseline checkout is not a Git checkout/);
    assert.throws(() => copy(join(base, 'packages')), /must be the root of a Git checkout/);
    assert.throws(() => copy(candidate), /must not be the candidate checkout/);
    const moved = join(scratch, 'moved');
    git(scratch, 'clone', '-q', base, moved);
    git(moved, '-c', 'user.email=vrt@example.invalid', '-c', 'user.name=vrt', 'commit', '-q', '--allow-empty', '-m', 'moved');
    assert.throws(() => copy(moved), /not at the base revision/);
    assert.equal(readFileSync(join(moved, corpus, 'old.pptx'), 'utf8'), 'v1');

    // Nothing is written through a symlink in the destination.
    writeFileSync(join(outside, 'target.xlsx'), 'keep');
    rmSync(join(base, 'packages/xlsx/public/corpus/new.xlsx'));
    symlinkSync(join(outside, 'target.xlsx'), join(base, 'packages/xlsx/public/corpus/new.xlsx'));
    assert.throws(copy, /symlink or non-regular file/);
    assert.equal(readFileSync(join(outside, 'target.xlsx'), 'utf8'), 'keep');
    // Nor over a hard link, which shares its inode with another file.
    rmSync(join(base, 'packages/xlsx/public/corpus/new.xlsx'));
    linkSync(join(outside, 'target.xlsx'), join(base, 'packages/xlsx/public/corpus/new.xlsx'));
    assert.throws(copy, /hard-linked file/);
    assert.equal(readFileSync(join(outside, 'target.xlsx'), 'utf8'), 'keep');
    rmSync(join(base, 'packages/xlsx'), { recursive: true });
    mkdirSync(join(base, 'packages/xlsx/public'), { recursive: true });
    symlinkSync(outside, join(base, 'packages/xlsx/public/corpus'));
    assert.throws(copy, /must be a real directory/);
    assert.deepEqual(readdirSync(outside), ['target.xlsx']);
    rmSync(join(base, 'packages/xlsx'), { recursive: true });

    // A rename is a deletion and an addition; neither side is copied.
    git(candidate, 'mv', `${corpus}/old.pptx`, `${corpus}/renamed.pptx`);
    git(base, 'checkout', '-q', '--', `${corpus}/old.pptx`);
    assert.throws(copy, /deletes or renames: D packages\/pptx\/public\/corpus\/old\.pptx$/);
    assert.deepEqual(readdirSync(join(base, corpus)), ['old.pptx']);
    assert.equal(existsSync(join(base, 'packages/xlsx')), false);
  } finally {
    rmSync(scratch, { recursive: true, force: true });
  }
});

// ── Generated output ─────────────────────────────────────────────────────────

test('documents the corpus no longer lists keep no generated file after a snapshot or a comparison', async () => {
  const fixture = selfVrtFixture();
  const baselineDemo = join(fixture.base, 'packages/pptx/tests/visual/baseline/demo');
  const screenshots = join(fixture.candidate, 'packages/pptx/tests/visual/screenshots/demo');
  const leftover = (root, name, extra = []) => {
    mkdirSync(join(root, name), { recursive: true });
    writeFileSync(join(root, name, 'slide-1.png'), 'old');
    writeFileSync(join(root, name, 'manifest.json'), '{}');
    for (const file of extra) writeFileSync(join(root, name, file), 'mine');
  };
  try {
    leftover(baselineDemo, 'gone');
    leftover(baselineDemo, 'kept', ['notes.txt']);
    await snapshot(fixture.base, fixture.revision);
    assert.deepEqual(readdirSync(baselineDemo).sort(), ['deck', 'kept', 'manifest.json']);
    assert.deepEqual(readdirSync(join(baselineDemo, 'kept')), ['notes.txt']);
    // Candidate screenshots hold images only; a manifest there is not the harness's.
    leftover(screenshots, 'gone');
    leftover(screenshots, 'kept', ['notes.txt']);
    assert.equal(await compare(fixture), null);
    assert.deepEqual(readdirSync(screenshots).sort(), ['deck', 'gone', 'kept']);
    assert.deepEqual(readdirSync(join(screenshots, 'gone')), ['manifest.json']);
    assert.deepEqual(readdirSync(join(screenshots, 'kept')).sort(), ['manifest.json', 'notes.txt']);
    // A generated name that is not a regular file is not the harness's to
    // remove: the run fails before any removal, so this stale document and
    // another one listed before it keep every entry.
    leftover(baselineDemo, 'a-early');
    leftover(baselineDemo, 'gone');
    try {
      execFileSync('mkfifo', [join(baselineDemo, 'gone/slide-2.png')]);
    } catch {
      mkdirSync(join(baselineDemo, 'gone/slide-2.png'));
    }
    await assert.rejects(snapshot(fixture.base, fixture.revision), /not a regular file: .*gone\/slide-2\.png/);
    assert.deepEqual(readdirSync(join(baselineDemo, 'gone')).sort(), ['manifest.json', 'slide-1.png', 'slide-2.png']);
    assert.deepEqual(readdirSync(join(baselineDemo, 'a-early')).sort(), ['manifest.json', 'slide-1.png']);
  } finally {
    rmSync(fixture.scratch, { recursive: true, force: true });
  }
});

test('nothing is written or removed through a symlinked output path', async () => {
  const fixture = selfVrtFixture();
  const outside = join(fixture.scratch, 'outside');
  const baseline = join(fixture.base, 'packages/pptx/tests/visual/baseline');
  const screenshots = join(fixture.candidate, 'packages/pptx/tests/visual/screenshots');
  const untouched = () => {
    assert.deepEqual(readdirSync(outside).sort(), ['gone', 'manifest.json']);
    assert.deepEqual(readdirSync(join(outside, 'gone')), ['slide-1.png']);
    assert.equal(readFileSync(join(outside, 'manifest.json'), 'utf8'), 'keep');
  };
  try {
    mkdirSync(join(outside, 'gone'), { recursive: true });
    writeFileSync(join(outside, 'gone/slide-1.png'), 'old');
    writeFileSync(join(outside, 'manifest.json'), 'keep');
    // Ignore the output folders themselves (the tracked pattern has a trailing
    // slash), so that a symlink there reaches the guard under test and not
    // the clean-checkout check.
    writeFileSync(join(fixture.base, '.git/info/exclude'), 'packages/*/tests/visual/baseline\n');
    await snapshot(fixture.base, fixture.revision);

    // Snapshot: the baseline folder, a document folder and a manifest leaf.
    const links = [
      [baseline, outside],
      [join(baseline, 'demo'), outside],
      [join(baseline, 'demo/gone'), join(outside, 'gone')],
      [join(baseline, 'demo/deck'), join(outside, 'gone')],
      [join(baseline, 'demo/manifest.json'), join(outside, 'manifest.json')],
    ];
    for (const [link, target] of links) {
      const saved = join(fixture.scratch, 'saved');
      if (existsSync(link)) cpSync(link, saved, { recursive: true });
      rmSync(link, { recursive: true, force: true });
      symlinkSync(target, link);
      await assert.rejects(snapshot(fixture.base, fixture.revision), /symlinked self-VRT output path/, link);
      untouched();
      unlinkSync(link);
      if (existsSync(saved)) {
        cpSync(saved, link, { recursive: true });
        rmSync(saved, { recursive: true });
      }
    }
    await snapshot(fixture.base, fixture.revision);

    // Comparison: the candidate's screenshot tree, and a document folder in it.
    mkdirSync(screenshots, { recursive: true });
    symlinkSync(outside, join(screenshots, 'demo'));
    await assert.rejects(compare(fixture), /symlinked self-VRT output path/);
    untouched();
    unlinkSync(join(screenshots, 'demo'));
    mkdirSync(join(screenshots, 'demo'));
    symlinkSync(join(outside, 'gone'), join(screenshots, 'demo/gone'));
    await assert.rejects(compare(fixture), /symlinked self-VRT output path/);
    untouched();
    unlinkSync(join(screenshots, 'demo/gone'));
    assert.equal(await compare(fixture), null);
  } finally {
    rmSync(fixture.scratch, { recursive: true, force: true });
  }
});

test('a named pipe or symlink in place of an input or an output is refused without blocking', (t) => {
  const fixture = fastFixture();
  const module = new URL('./private-corpus.mjs', import.meta.url).href;
  const input = join(fixture.base, 'packages/pptx/public/private/pptx/deck.pptx');
  try {
    try {
      execFileSync('mkfifo', [join(fixture.scratch, 'probe')]);
    } catch {
      t.skip('mkfifo unavailable');
      return;
    }
    // A separate process with a timeout: a blocking open of a pipe never returns.
    const child = spawnSync(process.execPath, ['--input-type=module', '-e', `
      import { execFileSync } from 'node:child_process';
      import { mkdirSync, renameSync, symlinkSync, unlinkSync } from 'node:fs';
      import { captureOrCompareSelfVrtItem, prepareSelfVrtCorpus } from ${JSON.stringify(module)};
      process.chdir(${JSON.stringify(join(fixture.base, 'packages/pptx'))});
      const PRIVATE = { corpus: 'private', format: 'pptx', stem: 'pptx/deck', itemKind: 'slide' };
      const prepare = () => prepareSelfVrtCorpus({ ...PRIVATE, files: ['pptx/deck.pptx'], snapshot: true });
      const attempt = (label, run) => {
        try { run(); console.log(label + ': accepted'); } catch (error) { console.log(label + ': ' + error.message); }
      };
      const input = ${JSON.stringify(input)};
      renameSync(input, input + '.saved');
      execFileSync('mkfifo', [input]);
      attempt('input pipe', prepare);
      unlinkSync(input);
      symlinkSync(input + '.saved', input);
      attempt('input symlink', prepare);
      unlinkSync(input);
      renameSync(input + '.saved', input);
      const output = 'tests/visual/baseline/private-corpus';
      mkdirSync(output + '/pptx/deck', { recursive: true });
      execFileSync('mkfifo', [output + '/manifest.json']);
      attempt('manifest pipe', prepare);
      unlinkSync(output + '/manifest.json');
      prepare();
      execFileSync('mkfifo', [output + '/pptx/deck/slide-1.png']);
      attempt('image pipe', () => captureOrCompareSelfVrtItem({
        ...PRIVATE, itemIndex: 0, actual: Buffer.from('x'), snapshot: true,
      }));
    `], {
      encoding: 'utf8',
      timeout: 15000,
      env: {
        ...process.env, VRT_BASELINE_CHECKOUT: '', VRT_BASELINE_REVISION: fixture.revision,
        VRT_ALLOW_HARNESS_CHANGES: '', VRT_ALLOW_CORPUS_INPUT_CHANGES: '',
      },
    });
    assert.equal(child.error, undefined, `must not block on a pipe: ${child.stderr}`);
    assert.match(child.stdout, /input pipe: .*not a regular file/, child.stderr);
    assert.match(child.stdout, /input symlink: ELOOP/, child.stderr);
    assert.match(child.stdout, /manifest pipe: .*non-regular file/, child.stderr);
    assert.match(child.stdout, /image pipe: .*non-regular file/, child.stderr);
  } finally {
    rmSync(fixture.scratch, { recursive: true, force: true });
  }
});

test('listing the tests of a spec prepares nothing', (t) => {
  // Preparation writes the snapshot manifest and removes stale documents; a
  // list-only run collects the specs and must do neither. The real pptx spec
  // is collected by the package's Playwright with a disposable checkout as
  // the working directory, which is where the harness reads and writes.
  const packageDirectory = fileURLToPath(new URL('../../packages/pptx/', import.meta.url));
  const cli = join(packageDirectory, 'node_modules/@playwright/test/cli.js');
  if (!existsSync(cli)) {
    t.skip('Playwright is not installed for packages/pptx');
    return;
  }
  const fixture = fastFixture();
  const output = join(fixture.base, 'packages/pptx/tests/visual/baseline/private-corpus');
  try {
    mkdirSync(join(output, 'pptx/gone'), { recursive: true });
    writeFileSync(join(output, 'pptx/gone/slide-1.png'), 'old');
    const child = spawnSync(process.execPath, [
      cli, 'test', '--config', join(packageDirectory, 'playwright.config.ts'),
      '--list', '--grep', 'private corpus self regression',
    ], {
      cwd: join(fixture.base, 'packages/pptx'),
      encoding: 'utf8',
      timeout: 120000,
      env: {
        ...process.env, VRT_PRIVATE_CORPUS: '1', VRT_PRIVATE_TIER: '', VRT_SNAPSHOT: '1', VRT_MODE: '',
        VRT_BASELINE_REVISION: fixture.revision, VRT_BASELINE_CHECKOUT: '',
        VRT_ALLOW_HARNESS_CHANGES: '', VRT_ALLOW_CORPUS_INPUT_CHANGES: '',
      },
    });
    assert.equal(child.status, 0, `${child.stdout}\n${child.stderr}`);
    assert.match(child.stdout, /private corpus self regression › pptx\/deck\.pptx/);
    assert.deepEqual(readdirSync(output), ['pptx']);
    assert.deepEqual(readdirSync(join(output, 'pptx')), ['gone']);
    assert.equal(readFileSync(join(output, 'pptx/gone/slide-1.png'), 'utf8'), 'old');
  } finally {
    rmSync(fixture.scratch, { recursive: true, force: true });
  }
});
