import assert from 'node:assert/strict';
import test from 'node:test';
import { auditAwaitCase, hasCombinedXlsxHostRun } from './check-model-source-ooxml-path.mjs';

test('an OOXML await added outside a selected-source branch fails the AST audit', () => {
  const previous = 'async function load() { await open(); await parse(); }';
  const changed = 'async function load() { await open(); await helper(); await parse(); }';
  assert.throws(() => auditAwaitCase('sample.ts', 'load', previous, changed), /OOXML awaits changed/);
  const sourceOnly = 'async function load() { if (opts.modelSources !== undefined) await helper(); await open(); await parse(); }';
  assert.equal(auditAwaitCase('sample.ts', 'load', previous, sourceOnly), 2);
  const gatedSourceOnly = 'async function load() { if (__OOXML_MODEL_SOURCES__ && opts.modelSources !== undefined) await helper(); await open(); await parse(); }';
  assert.equal(auditAwaitCase('sample.ts', 'load', previous, gatedSourceOnly), 2);
  const inverted = 'async function load() { if (__OOXML_MODEL_SOURCES__ && opts.modelSources === undefined) await helper(); await open(); await parse(); }';
  assert.throws(() => auditAwaitCase('sample.ts', 'load', previous, inverted), /OOXML awaits changed/);
});

test('only the explicit bundled-font opt-in may add a DOCX load await', () => {
  const previous = "async function load() { await parse(); }";
  const optional = "async function load() { const doc = { _mode: 'main' }; const opts = { useBundledOfficeFonts: true }; const fonts = doc._mode === 'main' && opts.useBundledOfficeFonts ? await loadBundledCalibri() : []; await parse(); }";
  assert.equal(auditAwaitCase('sample.ts', 'load', previous, optional), 1);
  const hostUrls = "async function load() { const urls = opts.useBundledOfficeFonts ? (await import('./urls.js')).CARLITO_URLS : undefined; await parse(); }";
  assert.equal(auditAwaitCase('sample.ts', 'load', previous, hostUrls), 1);
  const ungated = "async function load() { await loadBundledCalibri(); await parse(); }";
  assert.throws(() => auditAwaitCase('sample.ts', 'load', previous, ungated), /OOXML awaits changed/);
  const inverted = "async function load() { const fonts = !opts.useBundledOfficeFonts ? await loadBundledCalibri() : []; await parse(); }";
  assert.throws(() => auditAwaitCase('sample.ts', 'load', previous, inverted), /OOXML awaits changed/);
});

test('XLSX construction and parse must share one host.run', () => {
  assert.equal(hasCombinedXlsxHostRun('host.run(() => { const archive = new XlsxArchive(bytes); return archive.parse(); });'), true);
  assert.equal(hasCombinedXlsxHostRun('host.run(() => new XlsxArchive(bytes)); host.run(() => archive.parse());'), false);
});
