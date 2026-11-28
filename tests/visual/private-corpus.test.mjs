import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
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
import {
  captureOrCompareSelfVrtItem,
  clearSelfVrtCandidateOutput,
  prepareSelfVrtCorpus,
  harnessBootstrapDiffViolations,
  pngPixelsEqual,
  selfVrtBaselineRoot,
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
    assert.equal(captureOrCompareSelfVrtItem({ ...DEMO, itemIndex: 0, actual: png(value), snapshot: true }), null);
    verifySelfVrtItemManifest({ ...DEMO, itemCount: 1, snapshot: true });
  });
}

function compare({ candidate, base, revision }, value = 7) {
  return inPackage(candidate, { VRT_BASELINE_CHECKOUT: base, VRT_BASELINE_REVISION: revision }, () => {
    prepareSelfVrtCorpus({ ...DEMO, files: ['demo/deck.pptx'], snapshot: false });
    const difference = captureOrCompareSelfVrtItem({
      ...DEMO, itemIndex: 0, actual: png(value), snapshot: false,
    });
    verifySelfVrtItemManifest({ ...DEMO, itemCount: 1, snapshot: false });
    return difference;
  });
}

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

test('a baseline rendered under another browser locale is rejected', () => {
  const fixture = selfVrtFixture();
  try {
    withEnv({ VRT_LOCALE: 'fr-FR' }, () => snapshot(fixture.base, fixture.revision));
    withEnv({ VRT_LOCALE: undefined }, () => {
      assert.throws(() => compare(fixture), /previous-renderer manifest mismatch/);
    });
    withEnv({ VRT_LOCALE: 'fr-FR' }, () => assert.equal(compare(fixture), null));
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
