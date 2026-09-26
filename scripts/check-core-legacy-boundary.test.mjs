import assert from 'node:assert/strict';
import test from 'node:test';
import { findViolations, isGuardedPath } from './check-core-legacy-boundary.mjs';

const rules = (text) => findViolations([{ path: 'packages/docx/src/x.ts', text }]).map((v) => v.rule);

test('flags legacy imports, options, branches and names in the OOXML packages', () => {
  assert.deepEqual(rules("import { x } from '@silurus/ooxml-legacy-converter/internal/direct-doc-engine';"),
    ['legacy-package', 'direct-format-name']);
  assert.deepEqual(rules('await DocxDocument.load(bytes, { legacyConversion: {} });'), ['legacy-office-api']);
  assert.deepEqual(rules("if (kind === 'legacy-xls') return;"), ['legacy-format-name']);
  assert.deepEqual(rules('const nativeDoc = await openNative();'), ['native-legacy-source']);
  assert.deepEqual(rules('archive.revision_markup_in_print?.()'), ['legacy-revision-view']);
  assert.deepEqual(rules('measureLegacyXlsNormalFont?: unknown;'), ['legacy-xls-measurement']);
  assert.deepEqual(rules('langDefault?: string;'), ['docx-lang-default']);
});

test('allows the pre-existing generic CFB rejection and unrelated "legacy" wording', () => {
  assert.deepEqual(rules("throw new OoxmlError('legacy-binary-format', 'a legacy binary .doc file');"), []);
  assert.deepEqual(rules('// legacy VML shapes and the legacy chart family fallback'), []);
  assert.deepEqual(rules('export function selectModelSource(sources, target, bytes) {}'), []);
});

test('guards source, parser and manifest files but not generated or private output', () => {
  assert.equal(isGuardedPath('packages/core/src/index.ts'), true);
  assert.equal(isGuardedPath('packages/docx/parser/src/parser.rs'), true);
  assert.equal(isGuardedPath('packages/node/package.json'), true);
  assert.equal(isGuardedPath('packages/xlsx/src/wasm/xlsx_parser.js'), false);
  assert.equal(isGuardedPath('packages/pptx/public/demo/readme.md'), false);
  assert.equal(isGuardedPath('packages/legacy-converter/src/legacy-doc.ts'), false);
});
