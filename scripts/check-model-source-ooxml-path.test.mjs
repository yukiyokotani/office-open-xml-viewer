import assert from 'node:assert/strict';
import test from 'node:test';
import { auditAwaitCase, hasCombinedXlsxHostRun } from './check-model-source-ooxml-path.mjs';

test('an OOXML await added outside a selected-source branch fails the AST audit', () => {
  const previous = 'async function load() { await open(); await parse(); }';
  const changed = 'async function load() { await open(); await helper(); await parse(); }';
  assert.throws(() => auditAwaitCase('sample.ts', 'load', previous, changed), /OOXML awaits changed/);
  const sourceOnly = 'async function load() { if (opts.legacyDoc !== undefined) await helper(); await open(); await parse(); }';
  assert.equal(auditAwaitCase('sample.ts', 'load', previous, sourceOnly), 2);
  const gatedSourceOnly = 'async function load() { if (__OOXML_MODEL_SOURCES__ && opts.legacyDoc !== undefined) await helper(); await open(); await parse(); }';
  assert.equal(auditAwaitCase('sample.ts', 'load', previous, gatedSourceOnly), 2);
  const inverted = 'async function load() { if (__OOXML_MODEL_SOURCES__ && opts.legacyDoc === undefined) await helper(); await open(); await parse(); }';
  assert.throws(() => auditAwaitCase('sample.ts', 'load', previous, inverted), /OOXML awaits changed/);
  const listGate = 'async function load() { if (__OOXML_MODEL_SOURCES__ && (opts.legacyDoc !== undefined || opts.odt !== undefined)) await helper(); await open(); await parse(); }';
  assert.equal(auditAwaitCase('sample.ts', 'load', previous, listGate), 2);
  // Only alternatives of source presence form a source-only gate.
  const widened = 'async function load() { if (__OOXML_MODEL_SOURCES__ && (opts.legacyDoc !== undefined || opts.other)) await helper(); await open(); await parse(); }';
  assert.throws(() => auditAwaitCase('sample.ts', 'load', previous, widened), /OOXML awaits changed/);
});

test('XLSX construction and parse must share one host.run', () => {
  assert.equal(hasCombinedXlsxHostRun('host.run(() => { const archive = new XlsxArchive(bytes); return archive.parse(); });'), true);
  assert.equal(hasCombinedXlsxHostRun('host.run(() => new XlsxArchive(bytes)); host.run(() => archive.parse());'), false);
});

test('PPTX retains the same opt-in font barrier when its owner becomes a lease', () => {
  const file = 'packages/pptx/src/presentation.ts';
  const previous = 'async function load() { await parse(); if (mode === "main" && opts.useGoogleFonts && ready) await preloadGoogleFonts(names); }';
  const changed = 'async function load() { await parse(); if (mode === "main" && opts.useGoogleFonts && ready) await pres._ensureGoogleFonts(names); }';
  assert.equal(auditAwaitCase(file, 'load', previous, changed), 2);
  const unguarded = 'async function load() { await parse(); await pres._ensureGoogleFonts(names); }';
  assert.throws(() => auditAwaitCase(file, 'load', previous, unguarded), /OOXML awaits changed/);
  const extra = 'async function load() { await parse(); if (opts.useGoogleFonts) { await pres._ensureGoogleFonts(names); await helper(); } }';
  assert.throws(() => auditAwaitCase(file, 'load', previous, extra), /OOXML awaits changed/);
  const optional = 'async function load() { await parse(); if (mode === "main" || opts.useGoogleFonts) await pres._ensureGoogleFonts(names); }';
  assert.throws(() => auditAwaitCase(file, 'load', previous, optional), /OOXML awaits changed/);
});

test('PPTX admits only the font-source barrier in place of the Google lease barrier', () => {
  const file = 'packages/pptx/src/presentation.ts';
  const previous = 'async function load() { await parse(); if (mode === "main" && opts.useGoogleFonts && ready) await pres._ensureGoogleFonts(names); }';
  const changed = 'async function load() { await parse(); if (mode === "main" && opts.fonts && ready) await pres._ensureFontSources(names); }';
  assert.equal(auditAwaitCase(file, 'load', previous, changed), 2);
  const moved = 'async function load() { if (opts.fonts) await pres._ensureFontSources(names); await parse(); }';
  assert.throws(() => auditAwaitCase(file, 'load', previous, moved), /OOXML awaits changed/);
  const twice = 'async function load() { await parse(); await pres._ensureFontSources(names); await pres._ensureFontSources(names); }';
  assert.throws(() => auditAwaitCase(file, 'load', previous, twice), /OOXML awaits changed/);
  assert.throws(() => auditAwaitCase('sample.ts', 'load', previous, changed), /OOXML awaits changed/);
});

test('DOCX admits only the native route barrier after math preparation', () => {
  const file = 'packages/docx/src/document.ts';
  const previous = 'async function load() { await parse(); await prepareMathRuns(); await publish(); }';
  const changed = 'async function load() { await parse(); await prepareMathRuns(); await nativeRoutesPending(); await publish(); }';
  assert.equal(auditAwaitCase(file, 'load', previous, changed), 3);
  const moved = 'async function load() { await nativeRoutesPending(); await parse(); await prepareMathRuns(); await publish(); }';
  assert.throws(() => auditAwaitCase(file, 'load', previous, moved), /OOXML awaits changed/);
  const twice = 'async function load() { await parse(); await prepareMathRuns(); await nativeRoutesPending(); await nativeRoutesPending(); await publish(); }';
  assert.throws(() => auditAwaitCase(file, 'load', previous, twice), /OOXML awaits changed/);
  assert.throws(() => auditAwaitCase('sample.ts', 'load', previous, changed), /OOXML awaits changed/);
});
