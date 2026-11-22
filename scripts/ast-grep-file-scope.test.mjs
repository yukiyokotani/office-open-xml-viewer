// The ast-grep rule tests (rule-tests/*.yml) check snippets without a path.
// These tests run the real rules over a fixture tree, for what depends on
// where a file lives: which modules a rule inspects and which it skips.
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { cpSync, mkdirSync, mkdtempSync, readFileSync, rmSync, symlinkSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import test from 'node:test';

const repo = resolve(import.meta.dirname, '..');
const astGrep = join(repo, 'node_modules/.bin/ast-grep');
// The exact scan `pnpm lint:docx-import-zones` runs: ignore files and hidden
// names must not take a production module out of a zone.
const lintScript = JSON.parse(readFileSync(join(repo, 'package.json'), 'utf8')).scripts['lint:docx-import-zones'];
const lintArguments = lintScript.replace(/^ast-grep /, '').replaceAll("'", '').split(' ');

function scan(t, files) {
  const root = mkdtempSync(join(tmpdir(), 'ooxml-ast-grep-scope-'));
  t.after(() => rmSync(root, { recursive: true, force: true }));
  for (const name of ['sgconfig.yml', 'rules', 'rule-utils']) {
    cpSync(join(repo, name), join(root, name), { recursive: true });
  }
  for (const [path, contents] of Object.entries(files)) {
    mkdirSync(dirname(join(root, 'packages/docx/src', path)), { recursive: true });
    writeFileSync(join(root, 'packages/docx/src', path), contents);
  }
  const result = spawnSync(astGrep, [...lintArguments, '--json=compact'], { cwd: root, encoding: 'utf8' });
  return JSON.parse(result.stdout)
    // The import-zone rules; older rules overlap on some fixtures.
    .filter((match) => match.ruleId.startsWith('docx-'))
    .map((match) => `${match.ruleId} ${match.file.replace('packages/docx/src/', '')}`)
    .sort();
}

const facade = "export { composeAffine, inverseMapAffinePoint, inverseMapAffineVector, mapAffinePoint,\n"
  + "  quarterTurnAffine, scaleAffine, translationAffine } from '../layout/affine.js';\n";
const layoutEdge = "import { paginateBody } from '../layout/body-paginator.js';\n";
const rendererEdge = "import type { Hidden } from '../renderer.js';\n";

test('the affine facade is accepted in paint/affine.ts only', (t) => {
  assert.deepEqual(scan(t, {
    'paint/affine.ts': `import type { Matrix2DData } from '../layout/types.js';\n${facade}`,
    'paint/helper.ts': facade,
    'paint/nested/affine.ts': facade,
  }), [
    'docx-paint-import-zone paint/helper.ts',
    'docx-paint-import-zone paint/nested/affine.ts',
  ]);
  assert.deepEqual(scan(t, { 'paint/affine.ts': `${facade}${layoutEdge}` }), [
    'docx-paint-affine-facade paint/affine.ts',
  ]);
});

test('every production paint module is inspected, so a chain of paint modules cannot leave paint', (t) => {
  // The retired graph walk reported graph.ts -> helper.ts -> paragraph-measure.ts.
  assert.deepEqual(scan(t, {
    'paint/graph.ts': "import { helper } from './helper.js';\nexport const paint = helper;\n",
    'paint/helper.ts': "import { measure } from '../paragraph-measure.js';\nexport const helper = measure;\n",
    'paint/nested/deep.ts': "import { measure } from '../canvas-page.js';\nimport { measureText } from '@silurus/ooxml-core';\n",
    'paragraph-measure.ts': 'export const measure = 1;\n',
    'paint/.hidden.ts': layoutEdge,
    'paint/dist/ignored.ts': layoutEdge,
    '../../../.gitignore': 'dist/\n',
    '../../../.ignore': 'dist/\n',
  }), [
    'docx-paint-import-zone paint/.hidden.ts',
    'docx-paint-import-zone paint/dist/ignored.ts',
    'docx-paint-import-zone paint/helper.ts',
    'docx-paint-import-zone paint/nested/deep.ts',
    'docx-paint-package-imports paint/nested/deep.ts',
  ]);
});

test('paint tests, test support, benches, stories and declarations are not production paint', (t) => {
  const files = {};
  for (const name of ['a.test.ts', 'a.spec.ts', 'a.stories.ts', 'a.test-support.ts', 'a.bench.ts', 'a.d.ts', 'wasm/glue.ts']) {
    files[`paint/${name}`] = `${layoutEdge}import { measureText } from '@silurus/ooxml-core';\n`;
  }
  assert.deepEqual(scan(t, files), []);
});

test('TSX modules cannot enter an import zone unread', (t) => {
  assert.deepEqual(scan(t, {
    'paint/view.tsx': layoutEdge,
    'line-breaker/view.tsx': rendererEdge,
    'layout/view.tsx': 'export const view = 1;\n',
  }), [
    'docx-import-zones-are-ts-only line-breaker/view.tsx',
    'docx-import-zones-are-ts-only paint/view.tsx',
  ]);
});

test('acquisition surfaces, their consumers and every line-breaker module are renderer-free', (t) => {
  const inspected = [
    'layout/acquisition-context.ts', 'layout/acquisition-state.ts', 'layout/acquisition-input-projections.ts',
    'layout/anchor-classification.ts', 'layout/measurement-environment.ts', 'layout/measurement-capabilities.ts',
    'layout/section-orientation.ts', 'layout/body-table-measurement.ts',
    'anchor-geometry.ts', 'float-table-geometry.ts', 'frame-geometry.ts', 'line-layout.ts', 'paragraph-measure.ts',
    'line-breaker/added-module.ts', 'line-breaker/nested/added-module.ts',
  ];
  const edge = (path) => (path.includes('/') ? rendererEdge : rendererEdge.replace('../', './'));
  const skipped = [
    'line-breaker/added-module.test.ts', 'line-breaker/added-module.bench.ts',
    'line-breaker/added-module.test-support.ts', 'line-breaker/added-module.d.ts', 'layout/document.ts',
  ];
  const files = Object.fromEntries([...inspected, ...skipped].map((path) => [path, edge(path)]));
  assert.deepEqual(
    scan(t, files),
    inspected.map((path) => `docx-acquisition-no-renderer-import ${path}`).sort(),
  );
});

test('only the renderer adapter is denied the legacy module names', (t) => {
  const edge = "import { layoutLines } from './layout/line-layout.js';\n";
  assert.deepEqual(scan(t, { 'renderer.ts': edge, 'document.ts': edge }), [
    'docx-renderer-legacy-imports renderer.ts',
  ]);
});
