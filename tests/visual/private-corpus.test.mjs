import assert from 'node:assert/strict';
import { execFileSync, spawnSync } from 'node:child_process';
import {
  cpSync,
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  realpathSync,
  rmSync,
  symlinkSync,
  unlinkSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import test from 'node:test';
import { PNG } from 'pngjs';
import { createHash } from 'node:crypto';
import {
  beginSelfVrtItems,
  captureOrCompareSelfVrtItem,
  clearSelfVrtCandidateOutput,
  prepareSelfVrtCorpus,
  harnessBootstrapDiffViolations,
  lockfileBootstrapDiffViolations,
  pngPixelsEqual,
  selfVrtBaselineRoot,
  selfVrtInputBody,
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

test('candidate corpus inputs may be copied into the baseline checkout only when allowed', () => {
  const { root, first } = baselineRepository();
  const candidate = mkdtempSync(join(realpathSync(tmpdir()), 'ooxml-vrt-candidate-'));
  const outside = mkdtempSync(join(realpathSync(tmpdir()), 'ooxml-vrt-outside-'));
  const rel = 'packages/xlsx/public/corpus';
  const place = (dir, name, bytes) => {
    mkdirSync(join(dir, rel), { recursive: true });
    writeFileSync(join(dir, rel, name), bytes);
  };
  const verify = (revision = first) => verifyBaselineCheckout({ checkout: root, revision, candidate });
  try {
    // The baseline revision already tracks one specimen.
    place(root, 'xlsx-001-old.xlsx', 'v1');
    git(root, 'add', '-f', `${rel}/xlsx-001-old.xlsx`);
    git(root, 'commit', '-q', '-m', 'track specimen');
    const tracked = git(root, 'rev-parse', 'HEAD');
    // The candidate rewrites it and adds another; both are copied over.
    place(candidate, 'xlsx-001-old.xlsx', 'v2');
    place(candidate, 'xlsx-999-new.xlsx', 'new');
    place(root, 'xlsx-001-old.xlsx', 'v2');
    place(root, 'xlsx-999-new.xlsx', 'new');
    withEnv({ VRT_ALLOW_CORPUS_INPUT_CHANGES: undefined, VRT_ALLOW_HARNESS_CHANGES: undefined }, () => {
      assert.throws(() => verify(tracked), /clean renderer checkout; changed paths at .*xlsx-001-old\.xlsx/);
    });
    withEnv({ VRT_ALLOW_CORPUS_INPUT_CHANGES: '1', VRT_ALLOW_HARNESS_CHANGES: undefined }, () => {
      verify(tracked);
      // A copy that differs from the candidate's file is rejected, also after the snapshot.
      place(root, 'xlsx-999-new.xlsx', 'changed');
      assert.throws(() => verify(tracked), /corpus input .*xlsx-999-new\.xlsx differs from the candidate/);
      place(root, 'xlsx-999-new.xlsx', 'new');
      // So is a symlink, even to identical bytes.
      writeFileSync(join(outside, 'x.xlsx'), 'new');
      rmSync(join(root, rel, 'xlsx-999-new.xlsx'));
      symlinkSync(join(outside, 'x.xlsx'), join(root, rel, 'xlsx-999-new.xlsx'));
      assert.throws(() => verify(tracked), /must be a regular file without symlinks/);
      rmSync(join(root, rel, 'xlsx-999-new.xlsx'));
      place(root, 'xlsx-999-new.xlsx', 'new');
      // A deleted tracked input is not a candidate copy.
      rmSync(join(root, rel, 'xlsx-001-old.xlsx'));
      assert.throws(() => verify(tracked), /corpus input .*xlsx-001-old\.xlsx is missing/);
      place(root, 'xlsx-001-old.xlsx', 'v2');
      // Only a specimen of the folder's own format qualifies.
      writeFileSync(join(root, rel, 'patch.js'), 'export {};\n');
      writeFileSync(join(root, rel, 'other.docx'), 'x');
      assert.throws(() => verify(tracked), /changed paths at .*: .*other\.docx/);
      rmSync(join(root, rel, 'patch.js'));
      rmSync(join(root, rel, 'other.docx'));
      verify(tracked);
      // It never excuses a renderer change.
      writeFileSync(join(root, 'renderer.js'), 'export const version = 3;\n');
      assert.throws(() => verify(tracked), /changed paths at .*: renderer\.js/);
    });
  } finally {
    rmSync(root, { recursive: true, force: true });
    rmSync(candidate, { recursive: true, force: true });
    rmSync(outside, { recursive: true, force: true });
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
  writeFileSync(join(base, 'renderer.js'), 'export const version = 1;\n');
  git(base, 'add', '.');
  git(base, 'commit', '-q', '-m', 'renderer');
  const revision = git(base, 'rev-parse', 'HEAD');
  const candidate = join(scratch, 'candidate');
  git(scratch, 'clone', '-q', base, candidate);
  return { scratch, base, candidate, revision };
}

const DEMO = { corpus: 'demo', format: 'pptx', stem: 'demo/deck', itemKind: 'slide' };

function inPackage(checkout, env, run) {
  const previous = process.cwd();
  process.chdir(join(checkout, 'packages/pptx'));
  try {
    return withEnv({ VRT_ALLOW_HARNESS_CHANGES: undefined, ...env }, run);
  } finally {
    process.chdir(previous);
  }
}

function snapshot(checkout, revision, value = 7) {
  inPackage(checkout, { VRT_BASELINE_CHECKOUT: undefined, VRT_BASELINE_REVISION: revision }, () => {
    prepareSelfVrtCorpus({ ...DEMO, files: ['demo/deck.pptx'], snapshot: true });
    beginSelfVrtItems({ ...DEMO, itemCount: 1, snapshot: true });
    assert.equal(captureOrCompareSelfVrtItem({ ...DEMO, itemIndex: 0, actual: png(value), snapshot: true }), null);
    verifySelfVrtItemManifest({ ...DEMO, itemCount: 1, snapshot: true });
  });
}

function compare({ candidate, base, revision }, value = 7) {
  return inPackage(candidate, { VRT_BASELINE_CHECKOUT: base, VRT_BASELINE_REVISION: revision }, () => {
    prepareSelfVrtCorpus({ ...DEMO, files: ['demo/deck.pptx'], snapshot: false });
    beginSelfVrtItems({ ...DEMO, itemCount: 1, snapshot: false });
    const difference = captureOrCompareSelfVrtItem({
      ...DEMO, itemIndex: 0, actual: png(value), snapshot: false,
    });
    verifySelfVrtItemManifest({ ...DEMO, itemCount: 1, snapshot: false });
    return difference;
  });
}

test('the renderer is served only the input bytes its manifest records', () => {
  const fixture = selfVrtFixture();
  try {
    inPackage(fixture.base, { VRT_BASELINE_CHECKOUT: undefined, VRT_BASELINE_REVISION: fixture.revision }, () => {
      assert.throws(() => selfVrtInputBody({ corpus: 'demo', file: 'demo/unprepared.pptx' }), /was not prepared/);
      prepareSelfVrtCorpus({ ...DEMO, files: ['demo/deck.pptx'], snapshot: true });
      assert.equal(selfVrtInputBody({ corpus: 'demo', file: 'demo/deck.pptx' }).toString(), 'deck');
      // A public corpus serves the buffer preparation hashed, so a specimen
      // replaced during capture cannot reach the renderer.
      writeFileSync('public/demo/deck.pptx', 'other');
      assert.equal(selfVrtInputBody({ corpus: 'demo', file: 'demo/deck.pptx' }).toString(), 'deck');
    });
  } finally {
    rmSync(fixture.scratch, { recursive: true, force: true });
  }
});

test('a private corpus input is re-read and refused when it no longer has the prepared digest', () => {
  const fixture = fastFixture();
  try {
    inPackage(fixture.base, { VRT_BASELINE_CHECKOUT: undefined, VRT_BASELINE_REVISION: fixture.revision }, () => {
      prepareSelfVrtCorpus({
        ...FAST, files: ['pptx/deck.pptx'], snapshot: true, selectionDigest: 'selection-a',
        selectionFiles: selectionFiles('private deck'),
      });
      const route = () => selfVrtInputBody({ corpus: 'private-fast', file: 'pptx/deck.pptx' });
      assert.equal(route().toString(), 'private deck');
      const path = 'public/private/pptx/deck.pptx';
      writeFileSync(path, 'other');
      assert.throws(route, /changed after its manifest recorded it/);
      writeFileSync(path, 'private deck');
      assert.equal(route().toString(), 'private deck');
    });
  } finally {
    rmSync(fixture.scratch, { recursive: true, force: true });
  }
});

test('an item manifest requires its input to be prepared in this process', () => {
  const fixture = selfVrtFixture();
  try {
    inPackage(fixture.base, { VRT_BASELINE_CHECKOUT: undefined, VRT_BASELINE_REVISION: fixture.revision }, () => {
      assert.throws(
        () => beginSelfVrtItems({ ...DEMO, stem: 'demo/unprepared', itemCount: 1, snapshot: true }),
        /was not prepared/,
      );
    });
  } finally {
    rmSync(fixture.scratch, { recursive: true, force: true });
  }
});

test('a verified baseline checkout detects a changed slide and accepts an unchanged one', () => {
  const fixture = selfVrtFixture();
  try {
    snapshot(fixture.base, fixture.revision);
    assert.equal(compare(fixture), null);
    assert.match(compare(fixture, 8), /differs from the previous renderer/);
  } finally {
    rmSync(fixture.scratch, { recursive: true, force: true });
  }
});

test('bypass: a symlink in the baseline cannot redirect reads to other images', () => {
  const fixture = selfVrtFixture();
  try {
    snapshot(fixture.base, fixture.revision);
    const baselineDemo = join(fixture.base, 'packages/pptx/tests/visual/baseline/demo');
    const screenshots = join(fixture.candidate, 'packages/pptx/tests/visual/screenshots/demo');
    mkdirSync(screenshots, { recursive: true });
    cpSync(join(baselineDemo, 'deck'), join(screenshots, 'deck'), { recursive: true });
    cpSync(join(baselineDemo, 'manifest.json'), join(screenshots, 'manifest.json'));

    // The whole corpus directory linked to the candidate's own output.
    rmSync(baselineDemo, { recursive: true });
    symlinkSync(screenshots, baselineDemo, 'dir');
    assert.throws(() => compare(fixture, 8), /must not contain symlinks/);

    // A single image linked to the candidate's output.
    unlinkSync(baselineDemo);
    cpSync(screenshots, baselineDemo, { recursive: true });
    rmSync(join(baselineDemo, 'deck/slide-1.png'));
    symlinkSync(join(screenshots, 'deck/slide-1.png'), join(baselineDemo, 'deck/slide-1.png'));
    assert.throws(() => compare(fixture, 8), /must not contain symlinks/);
  } finally {
    rmSync(fixture.scratch, { recursive: true, force: true });
  }
});

test('bypass: images copied from another checkout or replaced after the snapshot are rejected', () => {
  const fixture = selfVrtFixture();
  const other = join(fixture.scratch, 'other');
  try {
    git(fixture.scratch, 'clone', '-q', fixture.base, other);
    snapshot(other, fixture.revision, 8);
    // Real files, but written by another checkout's snapshot run.
    cpSync(
      join(other, 'packages/pptx/tests/visual/baseline'),
      join(fixture.base, 'packages/pptx/tests/visual/baseline'),
      { recursive: true },
    );
    assert.throws(() => compare(fixture, 8), /manifest mismatch[\s\S]*"checkout"/);

    // The genuine snapshot, with its image replaced afterwards.
    snapshot(fixture.base, fixture.revision);
    writeFileSync(join(fixture.base, 'packages/pptx/tests/visual/baseline/demo/deck/slide-1.png'), png(8));
    assert.throws(() => compare(fixture, 8), /changed since its snapshot/);
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

test('bypass: a baseline checkout changed mid-run is caught at the next item', () => {
  const fixture = selfVrtFixture();
  try {
    snapshot(fixture.base, fixture.revision);
    inPackage(
      fixture.candidate,
      { VRT_BASELINE_CHECKOUT: fixture.base, VRT_BASELINE_REVISION: fixture.revision },
      () => {
        prepareSelfVrtCorpus({ ...DEMO, files: ['demo/deck.pptx'], snapshot: false });
        beginSelfVrtItems({ ...DEMO, itemCount: 1, snapshot: false });
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

function selectionFiles(content) {
  return new Map([['pptx/deck.pptx', { unitCount: 4, units: [1, 3], sha256: createHash('sha256').update(content).digest('hex') }]]);
}

function fastRun(checkout, env, { snapshot, digest = 'selection-a', indices = [1, 3], value = 7, content = 'private deck' }) {
  return inPackage(checkout, env, () => {
    prepareSelfVrtCorpus({
      ...FAST, files: ['pptx/deck.pptx'], snapshot, selectionDigest: digest,
      selectionFiles: selectionFiles(content),
    });
    beginSelfVrtItems({ ...FAST, itemCount: 4, snapshot, itemIndices: indices });
    const differences = indices.map((itemIndex) => captureOrCompareSelfVrtItem({
      ...FAST, itemIndex, actual: png(value), snapshot,
    })).filter(Boolean);
    verifySelfVrtItemManifest({ ...FAST, itemCount: 4, snapshot, itemIndices: indices });
    return differences;
  });
}

test('images captured for an input are refused after the input is replaced', () => {
  const fixture = fastFixture();
  try {
    const baseEnv = { VRT_BASELINE_CHECKOUT: undefined, VRT_BASELINE_REVISION: fixture.revision };
    const compareEnv = { VRT_BASELINE_CHECKOUT: fixture.base, VRT_BASELINE_REVISION: fixture.revision };
    assert.deepEqual(fastRun(fixture.base, baseEnv, { snapshot: true }), []);
    assert.deepEqual(fastRun(fixture.candidate, compareEnv, { snapshot: false }), []);
    for (const checkout of [fixture.base, fixture.candidate]) {
      writeFileSync(join(checkout, 'packages/pptx/public/private/pptx/deck.pptx'), 'replaced deck');
    }
    // Snapshot preparation runs again for the replaced input, but its images are
    // not recaptured or re-verified; the comparison must not accept them.
    inPackage(fixture.base, baseEnv, () => {
      prepareSelfVrtCorpus({
        ...FAST, files: ['pptx/deck.pptx'], snapshot: true, selectionDigest: 'selection-a',
        selectionFiles: selectionFiles('replaced deck'),
      });
    });
    assert.throws(
      () => fastRun(fixture.candidate, compareEnv, { snapshot: false, content: 'replaced deck' }),
      /previous-renderer manifest mismatch/,
    );
  } finally {
    rmSync(fixture.scratch, { recursive: true, force: true });
  }
});

test('fast tier compares exactly the selected items under the same selection', () => {
  const fixture = fastFixture();
  const { base, candidate, revision } = fixture;
  const snapshotEnv = { VRT_BASELINE_CHECKOUT: undefined, VRT_BASELINE_REVISION: revision };
  const compareEnv = { VRT_BASELINE_CHECKOUT: base, VRT_BASELINE_REVISION: revision };
  try {
    assert.deepEqual(fastRun(base, snapshotEnv, { snapshot: true }), []);
    assert.ok(existsSync(join(base, 'packages/pptx/tests/visual/baseline/private-corpus-fast/pptx/deck/slide-4.png')));
    assert.ok(!existsSync(join(base, 'packages/pptx/tests/visual/baseline/private-corpus-fast/pptx/deck/slide-1.png')));
    assert.deepEqual(fastRun(candidate, compareEnv, { snapshot: false }), []);
    assert.equal(fastRun(candidate, compareEnv, { snapshot: false, value: 8 }).length, 2);
    // A different selection is a different oracle.
    assert.throws(() => fastRun(candidate, compareEnv, { snapshot: false, digest: 'selection-b' }), /manifest mismatch/);
    assert.throws(() => fastRun(candidate, compareEnv, { snapshot: false, indices: [1] }), /item set mismatch/);
  } finally {
    rmSync(fixture.scratch, { recursive: true, force: true });
  }
});

test('fast tier and full corpora refuse each other\'s selection arguments', () => {
  const fixture = fastFixture();
  const env = { VRT_BASELINE_CHECKOUT: undefined, VRT_BASELINE_REVISION: fixture.revision };
  try {
    inPackage(fixture.base, env, () => {
      assert.throws(() => prepareSelfVrtCorpus({ ...FAST, files: ['pptx/deck.pptx'], snapshot: true }),
        /requires a selection digest/);
      assert.throws(() => prepareSelfVrtCorpus({
        ...FAST, files: ['pptx/deck.pptx'], snapshot: true, selectionDigest: 'x',
      }), /requires the selection's per-file digests/);
      assert.throws(() => prepareSelfVrtCorpus({
        ...FAST, corpus: 'private', files: ['pptx/deck.pptx'], snapshot: true, selectionDigest: 'x',
      }), /does not take a selection digest/);
      assert.throws(() => beginSelfVrtItems({ ...FAST, itemCount: 4, snapshot: true }),
        /requires the selected item indices/);
      assert.throws(() => beginSelfVrtItems({
        ...FAST, itemCount: 4, snapshot: true, itemIndices: [3, 1],
      }), /invalid slide selection/);
      assert.throws(() => beginSelfVrtItems({
        ...FAST, corpus: 'private', itemCount: 4, snapshot: true, itemIndices: [1],
      }), /does not take item indices/);
    });
  } finally {
    rmSync(fixture.scratch, { recursive: true, force: true });
  }
});

test('allowance copies recorded at snapshot time must stay present and identical', () => {
  const scratch = mkdtempSync(join(realpathSync(tmpdir()), 'ooxml-vrt-allow-'));
  const base = join(scratch, 'base');
  const candidate = join(scratch, 'candidate');
  const rel = 'packages/xlsx/public/corpus';
  const specimen = 'xlsx-999-new.xlsx';
  try {
    mkdirSync(join(base, rel), { recursive: true });
    git(base, 'init', '-q');
    git(base, 'config', 'user.email', 'vrt@example.invalid');
    git(base, 'config', 'user.name', 'vrt');
    writeFileSync(join(base, rel, '.gitkeep'), '');
    writeFileSync(join(base, '.gitignore'), 'packages/*/tests/visual/baseline/\n');
    git(base, 'add', '-f', '.');
    git(base, 'commit', '-q', '-m', 'renderer');
    const revision = git(base, 'rev-parse', 'HEAD');
    git(scratch, 'clone', '-q', base, candidate);
    writeFileSync(join(candidate, rel, specimen), 'new');
    writeFileSync(join(base, rel, specimen), 'new');
    const files = [`corpus/${specimen}`];
    const spec = { corpus: 'corpus', format: 'xlsx', files };
    const inXlsx = (checkout, env, run) => {
      const previous = process.cwd();
      process.chdir(join(checkout, 'packages/xlsx'));
      try {
        return withEnv({ VRT_ALLOW_HARNESS_CHANGES: undefined, ...env }, run);
      } finally {
        process.chdir(previous);
      }
    };
    const compare = (flag) => inXlsx(candidate, {
      VRT_ALLOW_CORPUS_INPUT_CHANGES: flag,
      VRT_BASELINE_CHECKOUT: base,
      VRT_BASELINE_REVISION: revision,
    }, () => {
      prepareSelfVrtCorpus({ ...spec, snapshot: false });
      selfVrtBaselineRoot();
    });
    inXlsx(base, {
      VRT_ALLOW_CORPUS_INPUT_CHANGES: '1',
      VRT_BASELINE_CHECKOUT: undefined,
      VRT_BASELINE_REVISION: revision,
    }, () => prepareSelfVrtCorpus({ ...spec, snapshot: true }));
    const record = JSON.parse(readFileSync(
      join(base, 'packages/xlsx/tests/visual/baseline/public-corpus/manifest.json'),
      'utf8',
    ));
    assert.deepEqual(record.corpusInputCopies.map((copy) => copy.path), [`${rel}/${specimen}`]);

    compare('1');
    // The comparison needs the allowance the snapshot used.
    assert.throws(() => compare(undefined), /clean renderer checkout/);

    // A copy deleted after the snapshot is caught, with the flag set or unset.
    rmSync(join(base, rel, specimen));
    assert.throws(() => compare('1'), /copies differ from the snapshot's record/);
    assert.throws(() => compare(undefined), /set it for the comparison run too/);

    // A modified copy is caught.
    writeFileSync(join(base, rel, specimen), 'edited');
    assert.throws(() => compare('1'), /differs from the candidate/);
    // Even when the candidate's file was edited along with it.
    writeFileSync(join(candidate, rel, specimen), 'edited');
    assert.throws(() => compare('1'), /changed since its snapshot recorded it|SHA-256|manifest mismatch/);

    // Edited mid-run, in both checkouts to identical new bytes: only the
    // snapshot's digest still rejects it.
    writeFileSync(join(candidate, rel, specimen), 'new');
    writeFileSync(join(base, rel, specimen), 'new');
    inXlsx(candidate, {
      VRT_ALLOW_CORPUS_INPUT_CHANGES: '1',
      VRT_BASELINE_CHECKOUT: base,
      VRT_BASELINE_REVISION: revision,
    }, () => {
      prepareSelfVrtCorpus({ ...spec, snapshot: false });
      writeFileSync(join(base, rel, specimen), 'rewritten');
      writeFileSync(join(candidate, rel, specimen), 'rewritten');
      assert.throws(() => selfVrtBaselineRoot(), /corpus input .*xlsx-999-new\.xlsx changed since its snapshot recorded it/);
    });

    // Deleted mid-run, after the corpus manifest was verified.
    writeFileSync(join(candidate, rel, specimen), 'new');
    writeFileSync(join(base, rel, specimen), 'new');
    inXlsx(candidate, {
      VRT_ALLOW_CORPUS_INPUT_CHANGES: '1',
      VRT_BASELINE_CHECKOUT: base,
      VRT_BASELINE_REVISION: revision,
    }, () => {
      prepareSelfVrtCorpus({ ...spec, snapshot: false });
      rmSync(join(base, rel, specimen));
      assert.throws(() => selfVrtBaselineRoot(), /copies differ from the snapshot's record/);
    });
  } finally {
    rmSync(scratch, { recursive: true, force: true });
  }
});

test('a non-regular or symlinked candidate input is rejected without being read', (t) => {
  const { root, first } = baselineRepository();
  const candidate = mkdtempSync(join(realpathSync(tmpdir()), 'ooxml-vrt-candidate-'));
  const outside = mkdtempSync(join(realpathSync(tmpdir()), 'ooxml-vrt-outside-'));
  const rel = 'packages/pptx/public/corpus';
  const module = new URL('./private-corpus.mjs', import.meta.url).href;
  // A separate process with a timeout: reading a named pipe would block forever.
  const verify = () => spawnSync(process.execPath, ['--input-type=module', '-e', `
    import { verifyBaselineCheckout } from ${JSON.stringify(module)};
    try {
      verifyBaselineCheckout({ checkout: ${JSON.stringify(root)}, revision: ${JSON.stringify(first)}, candidate: ${JSON.stringify(candidate)} });
      console.log('accepted');
    } catch (error) { console.log(error.message); }
  `], {
    encoding: 'utf8',
    timeout: 15000,
    env: { ...process.env, VRT_ALLOW_CORPUS_INPUT_CHANGES: '1', VRT_ALLOW_HARNESS_CHANGES: '' },
  });
  try {
    mkdirSync(join(root, rel), { recursive: true });
    mkdirSync(join(candidate, rel), { recursive: true });
    writeFileSync(join(root, rel, 'deck.pptx'), 'x');
    const fifo = join(candidate, rel, 'deck.pptx');
    try {
      execFileSync('mkfifo', [fifo]);
    } catch {
      t.diagnostic('mkfifo unavailable; skipping the named-pipe case');
    }
    if (existsSync(fifo)) {
      const result = verify();
      assert.equal(result.error, undefined, 'must not block on the pipe');
      assert.match(result.stdout, /must be a regular file without symlinks at .*candidate/);
      unlinkSync(fifo);
    }
    // A symlinked candidate input, even to identical bytes.
    writeFileSync(join(outside, 'x.pptx'), 'x');
    symlinkSync(join(outside, 'x.pptx'), join(candidate, rel, 'deck.pptx'));
    assert.match(verify().stdout, /must be a regular file without symlinks at .*candidate/);
    unlinkSync(join(candidate, rel, 'deck.pptx'));
    // A symlinked directory component.
    rmSync(join(candidate, rel), { recursive: true });
    mkdirSync(join(outside, 'dir'));
    writeFileSync(join(outside, 'dir/deck.pptx'), 'x');
    symlinkSync(join(outside, 'dir'), join(candidate, rel));
    assert.match(verify().stdout, /must be a regular file without symlinks at .*candidate/);
  } finally {
    rmSync(root, { recursive: true, force: true });
    rmSync(candidate, { recursive: true, force: true });
    rmSync(outside, { recursive: true, force: true });
  }
});

test('a prepared input replaced by a named pipe is refused without blocking', (t) => {
  const { root, first } = baselineRepository();
  const rel = 'packages/pptx/public/private/pptx';
  const module = new URL('./private-corpus.mjs', import.meta.url).href;
  try {
    mkdirSync(join(root, rel), { recursive: true });
    writeFileSync(join(root, rel, 'deck.pptx'), 'x');
    try {
      execFileSync('mkfifo', [join(root, 'probe')]);
      unlinkSync(join(root, 'probe'));
    } catch {
      t.skip('mkfifo unavailable');
      return;
    }
    // Separate process with a timeout: a blocking read of the pipe never returns.
    const result = spawnSync(process.execPath, ['--input-type=module', '-e', `
      import { execFileSync } from 'node:child_process';
      import { unlinkSync } from 'node:fs';
      import { prepareSelfVrtCorpus, selfVrtInputBody } from ${JSON.stringify(module)};
      process.chdir(${JSON.stringify(join(root, 'packages/pptx'))});
      prepareSelfVrtCorpus({ corpus: 'private', format: 'pptx', files: ['pptx/deck.pptx'], snapshot: true });
      unlinkSync('public/private/pptx/deck.pptx');
      execFileSync('mkfifo', ['public/private/pptx/deck.pptx']);
      try {
        selfVrtInputBody({ corpus: 'private', file: 'pptx/deck.pptx' });
        console.log('accepted');
      } catch (error) { console.log(error.message); }
    `], {
      encoding: 'utf8',
      timeout: 15000,
      env: {
        ...process.env,
        VRT_ALLOW_CORPUS_INPUT_CHANGES: '1',
        VRT_ALLOW_HARNESS_CHANGES: '',
        VRT_BASELINE_CHECKOUT: '',
        VRT_BASELINE_REVISION: first,
      },
    });
    assert.equal(result.error, undefined, 'must not block on the pipe');
    assert.match(result.stdout, /not a regular file/);
  } finally {
    rmSync(root, { recursive: true, force: true });
  }
});

test('a snapshot image swapped between capture and finalization is refused', () => {
  const fixture = selfVrtFixture();
  try {
    inPackage(fixture.base, { VRT_BASELINE_CHECKOUT: undefined, VRT_BASELINE_REVISION: fixture.revision }, () => {
      prepareSelfVrtCorpus({ ...DEMO, files: ['demo/deck.pptx'], snapshot: true });
      beginSelfVrtItems({ ...DEMO, itemCount: 1, snapshot: true });
      captureOrCompareSelfVrtItem({ ...DEMO, itemIndex: 0, actual: png(7), snapshot: true });
      writeFileSync('tests/visual/baseline/demo/deck/slide-1.png', png(9));
      assert.throws(
        () => verifySelfVrtItemManifest({ ...DEMO, itemCount: 1, snapshot: true }),
        /snapshot image changed after capture/,
      );
      assert.equal(existsSync('tests/visual/baseline/demo/deck/manifest.json'), false);
    });
    // The failed run left no manifest, so no comparison can accept the swap.
    assert.throws(() => compare(fixture, 9), /missing previous-renderer manifest/);
  } finally {
    rmSync(fixture.scratch, { recursive: true, force: true });
  }
});

test('a baseline manifest and image substituted mid-comparison are not trusted, restored or not', () => {
  const fixture = selfVrtFixture();
  try {
    snapshot(fixture.base, fixture.revision);
    const directory = join(fixture.base, 'packages/pptx/tests/visual/baseline/demo/deck');
    const manifest = readFileSync(join(directory, 'manifest.json'), 'utf8');
    const image = readFileSync(join(directory, 'slide-1.png'));
    inPackage(
      fixture.candidate,
      { VRT_BASELINE_CHECKOUT: fixture.base, VRT_BASELINE_REVISION: fixture.revision },
      () => {
        prepareSelfVrtCorpus({ ...DEMO, files: ['demo/deck.pptx'], snapshot: false });
        beginSelfVrtItems({ ...DEMO, itemCount: 1, snapshot: false });
        // A consistent forgery: a different image and a manifest that records it.
        const forged = png(8);
        const record = JSON.parse(manifest);
        record.items[0].sha256 = createHash('sha256').update(forged).digest('hex');
        writeFileSync(join(directory, 'slide-1.png'), forged);
        writeFileSync(join(directory, 'manifest.json'), JSON.stringify(record));
        assert.throws(
          () => captureOrCompareSelfVrtItem({ ...DEMO, itemIndex: 0, actual: forged, snapshot: false }),
          /changed since its snapshot/,
        );
        // Restoring the files afterwards does not erase the observed tampering
        // from this run's validated record, and the genuine image still matches it.
        writeFileSync(join(directory, 'slide-1.png'), image);
        writeFileSync(join(directory, 'manifest.json'), manifest);
        assert.equal(
          captureOrCompareSelfVrtItem({ ...DEMO, itemIndex: 0, actual: png(7), snapshot: false }),
          null,
        );
        verifySelfVrtItemManifest({ ...DEMO, itemCount: 1, snapshot: false });
      },
    );
  } finally {
    rmSync(fixture.scratch, { recursive: true, force: true });
  }
});

test('a specimen replaced between selection load and preparation is refused, with or without selected items', () => {
  const fixture = fastFixture();
  try {
    inPackage(fixture.base, { VRT_BASELINE_CHECKOUT: undefined, VRT_BASELINE_REVISION: fixture.revision }, () => {
      for (const units of [[1, 3], []]) {
        const loaded = selectionFiles('private deck');
        loaded.get('pptx/deck.pptx').units = units;
        // The loader validated 'private deck'; the file is replaced before preparation.
        writeFileSync('public/private/pptx/deck.pptx', 'replaced deck');
        assert.throws(
          () => prepareSelfVrtCorpus({
            ...FAST, files: ['pptx/deck.pptx'], snapshot: true, selectionDigest: 'selection-a',
            selectionFiles: loaded,
          }),
          /changed between the fast-tier selection and corpus preparation/,
        );
        writeFileSync('public/private/pptx/deck.pptx', 'private deck');
        prepareSelfVrtCorpus({
          ...FAST, files: ['pptx/deck.pptx'], snapshot: true, selectionDigest: 'selection-a',
          selectionFiles: loaded,
        });
      }
    });
  } finally {
    rmSync(fixture.scratch, { recursive: true, force: true });
  }
});
