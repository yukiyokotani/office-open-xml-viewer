import assert from 'node:assert/strict';
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import test from 'node:test';
import { tombstoneViolations } from './check-tombstones.mjs';
import { DELETED_IDENTIFIERS, DELETED_PATHS } from './tombstones.mjs';

function violations(t, files) {
  const root = mkdtempSync(join(tmpdir(), 'ooxml-tombstones-'));
  t.after(() => rmSync(root, { recursive: true, force: true }));
  for (const [path, contents] of Object.entries(files)) {
    mkdirSync(dirname(join(root, path)), { recursive: true });
    writeFileSync(join(root, path), contents);
  }
  return tombstoneViolations(root);
}

const docx = DELETED_IDENTIFIERS.find((entry) => entry.root === 'packages/docx/src');

test('a tree without resurrected paths or identifiers passes', (t) => {
  assert.deepEqual(violations(t, {
    'packages/docx/src/layout/document.ts': 'export function layoutDocument() { const colIndex = 0; return colIndex + 1; }\n',
    'packages/docx/src/options.ts': 'export interface Options { legacyDoc?: unknown }\n',
    'packages/docx/src/table-geometry.test.ts': 'export {};\n',
  }), []);
});

test('every deleted path is rejected', (t) => {
  for (const entry of DELETED_PATHS) {
    const path = entry.path ?? `${entry.under}/nested/${entry.basename}`;
    assert.deepEqual(violations(t, { [path]: '{}\n' }), [`DELETED_PATH: ${path}`]);
  }
  // The pinned entries, so shrinking the list fails here.
  for (const path of [
    'scripts/docx-layout-boundary-baseline.json',
    'scripts/docx-compatibility-observation-baseline.json',
    'packages/docx/src/paragraph-anchor-frame-adapter.ts',
    'packages/docx/src/table-geometry.ts',
    'packages/docx/src/layout/table-fragments.ts',
    'packages/docx/src/paint/layout-fragments.ts',
  ]) {
    assert.deepEqual(violations(t, { [path]: '{}\n' }), [`DELETED_PATH: ${path}`]);
  }
});

test('every deleted identifier is rejected wherever it appears in production DOCX source', (t) => {
  for (const name of docx.names) {
    assert.deepEqual(
      violations(t, { 'packages/docx/src/layout/deleted-path.ts': `export const ${name} = true;\n` }),
      [`DELETED_IDENTIFIER: packages/docx/src/layout/deleted-path.ts#${name}`],
    );
  }
  // The pinned identifiers of the retired per-list guards.
  const pinned = [
    'computePages', 'paginateDocument', 'paginateWithHeaderFooterReserve', 'physicalPageSizeForPage',
    'PaginatedBodyElement', 'prebuiltPages', 'retainedLayout', 'bodyFragmentFor', 'bodyLayoutFallback',
    'sectionBreakSpacer', 'collapsedSpacer', 'leadsCollapsedRun', 'hiddenCollapsed', 'computeTableRowHeights',
    'estimateTableHeight', 'measureRetainedCellContentHeightPt', 'measureCellParagraphWindow',
    'measureCellElementHeight', 'trimTrailingStructuralMarker', 'rescaleLayoutLines',
    'paragraphSegsStateSensitive', 'RenderState', 'computeTableLayout', 'renderHeaderFooter',
  ];
  for (const name of pinned) assert.ok(docx.names.includes(name), name);
  assert.deepEqual(violations(t, {
    'packages/docx/src/renderer.ts': 'export function paginateDocument() {}\n',
    'packages/docx/src/view.tsx': 'export const view = (props) => props.prebuiltPages;\n',
  }), [
    'DELETED_IDENTIFIER: packages/docx/src/renderer.ts#paginateDocument',
    'DELETED_IDENTIFIER: packages/docx/src/view.tsx#prebuiltPages',
  ]);
});

test('migration switches and alternate-engine fallbacks are rejected by pattern', (t) => {
  for (const name of ['useOldLayoutEngine', 'preferAlternateLayoutPath', 'dryRun', 'tableRequiresLegacy', 'paintEnabled', 'legacyDocument']) {
    assert.deepEqual(
      violations(t, { 'packages/docx/src/layout/path.ts': `export const ${name} = true;\n` }),
      [`DELETED_IDENTIFIER: packages/docx/src/layout/path.ts#${name}`],
    );
  }
  assert.deepEqual(violations(t, {
    'packages/docx/src/layout/obsolete-state.ts': 'export interface RenderState { readonly dryRun: boolean }\n',
  }), [
    'DELETED_IDENTIFIER: packages/docx/src/layout/obsolete-state.ts#RenderState',
    'DELETED_IDENTIFIER: packages/docx/src/layout/obsolete-state.ts#dryRun',
  ]);
});

test('legacy stamp fields are rejected in every property position', (t) => {
  for (const name of docx.propertyNames) {
    for (const source of [
      `export const stamp = { ${name}: true };\n`,
      `export const read = (stamp) => stamp.${name};\n`,
      `export const read = (stamp) => stamp['${name}'];\n`,
      `export const read = ({ ${name} }) => ${name};\n`,
      `export const read = ({ ${name}: local }) => local;\n`,
      `export const stamp = (${name}) => ({ ${name} });\n`,
      `export interface Stamp { ${name}: number }\n`,
      `export class Stamp { ${name} = 0; }\n`,
      `export class Stamp { ${name}() {} }\n`,
      `export class Stamp { get ${name}() { return 0; } }\n`,
    ]) {
      assert.deepEqual(
        violations(t, { 'packages/docx/src/deleted-path.ts': source }),
        [`DELETED_PROPERTY: packages/docx/src/deleted-path.ts#${name}`],
        source,
      );
    }
  }
  assert.deepEqual(docx.propertyNames, [
    'colIndex', 'colGeom', 'colTopPt', 'sectionHF', 'sectionGeom', 'sectionPageNumType', 'sectionTextDirection',
  ]);
});

test('tests, benches, test support, declarations and generated WASM glue are not production source', (t) => {
  const files = {};
  for (const name of ['a.test.ts', 'a.spec.ts', 'a.stories.ts', 'a.test-support.ts', 'a.bench.ts', 'a.d.ts', 'wasm/glue.ts']) {
    files[`packages/docx/src/${name}`] = 'export const paginateDocument = { colIndex: true };\n';
  }
  assert.deepEqual(violations(t, files), []);
});
