#!/usr/bin/env node
// Exercise only installed tarballs in an isolated consumer: no workspace aliases.
// Authored synthetic inputs keep this PR/release gate redistributable.
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { mkdtemp, mkdir, realpath, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { pathToFileURL, fileURLToPath } from 'node:url';
import { build, preview } from 'vite';
import { rolldown } from 'rolldown';
import { chromium } from '@playwright/test';

const root = fileURLToPath(new URL('../', import.meta.url));
const scratch = await realpath(await mkdtemp(join(tmpdir(), 'ooxml-packed-companions-')));
let server, browser;
const run = (command, args, cwd = scratch) => execFileSync(command, args, { cwd, stdio: 'pipe', encoding: 'utf8' });
try {
  run(process.execPath, ['scripts/sync-companion-version.mjs'], root);
  const tarballs = [];
  for (const directory of [root, join(root, 'packages/legacy-office'), join(root, 'packages/opendocument')]) {
    const packed = JSON.parse(run('npm', ['pack', '--json', '--pack-destination', scratch], directory));
    const files = packed[0].files.map(file => file.path).sort();
    const wasm = files.filter(path => path.endsWith('.wasm'));
    if (directory.endsWith('opendocument')) {
      // Exactly the reader assets: never the local private corpus.
      assert.deepEqual(files, [
        'LICENSE', 'README.md', 'THIRD_PARTY_NOTICES.md', 'dist/odt-source-module.mjs', 'dist/odt.mjs',
        'dist/opendocument_odt_bg.wasm', 'dist/types/odt.d.ts', 'package.json',
      ]);
    } else if (directory !== root) {
      assert.deepEqual(wasm, ['doc', 'ppt', 'xls'].map(f => `dist/legacy_${f}_direct_bg.wasm`));
    } else {
      assert.ok(files.every(path => !/^dist\/(?:legacy[-_]|opendocument|odt[-.])/.test(path)), 'root tarball contains a companion reader asset');
    }
    tarballs.push(join(scratch, packed[0].filename));
  }
  await writeFile(join(scratch, 'package.json'), JSON.stringify({ private: true, type: 'module' }));
  run('npm', ['install', '--ignore-scripts', '--no-audit', '--no-fund', ...tarballs]);
  // Fixture bundling is test setup only; the consumer itself resolves exclusively
  // from its installed packages. Extension aliases support the TS fixture imports.
  const fixtureEntry = join(scratch, 'fixtures.ts');
  await writeFile(fixtureEntry, `export { buildDocFixture, buildXlsFixture, buildPptFixture } from ${JSON.stringify(join(root, 'packages/legacy-office/src/test-fixtures.ts'))};\nexport { buildPptShapeImageFillFixture } from ${JSON.stringify(join(root, 'packages/legacy-office/src/ppt-shape-image-fill-fixture.ts'))};\nexport { buildCfbFixture } from ${JSON.stringify(join(root, 'packages/core/src/testing/cfb-fixture.ts'))};`);
  const fixturesBuild = await rolldown({ input: fixtureEntry, platform: 'node', resolve: { extensionAlias: { '.js': ['.ts', '.js'] } } });
  try { await fixturesBuild.write({ file: join(scratch, 'fixtures.mjs'), format: 'es', codeSplitting: false }); }
  finally { await fixturesBuild.close(); }
  const fixtures = await import(pathToFileURL(join(scratch, 'fixtures.mjs')).href);
  await mkdir(join(scratch, 'public'));
  for (const [family, bytes] of Object.entries({ doc: fixtures.buildDocFixture(), xls: fixtures.buildXlsFixture(), ppt: fixtures.buildPptShapeImageFillFixture() })) {
    await writeFile(join(scratch, 'public', `sample.${family}`), bytes);
  }
  for (const [family, stream] of [['doc', 'WordDocument'], ['xls', 'Workbook'], ['ppt', 'PowerPoint Document']]) {
    await writeFile(join(scratch, 'public', `unsupported.${family}`), new Uint8Array(fixtures.buildCfbFixture(['Root Entry', stream])));
  }
  await writeFile(join(scratch, 'public', 'deferred.ppt'), fixtures.buildPptFixture());
  await writeFile(join(scratch, 'node-smoke.mjs'), `
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { DocxViewer } from '@silurus/ooxml/docx';
import { XlsxViewer } from '@silurus/ooxml/xlsx';
import { PptxViewer } from '@silurus/ooxml/pptx';
import { legacyDoc, createLegacyDoc } from '@silurus/ooxml-legacy-office/doc';
import { legacyXls, createLegacyXls } from '@silurus/ooxml-legacy-office/xls';
import { legacyPpt, createLegacyPpt } from '@silurus/ooxml-legacy-office/ppt';
import { materializeDocxDocument, materializeXlsxWorkbook, materializePptxPresentation, OoxmlError } from '@silurus/ooxml/node';
import { assertCompanionRuntime, CompanionRuntimeMismatchError } from '@silurus/ooxml/internal/companion';
for (const value of [DocxViewer, XlsxViewer, PptxViewer, createLegacyDoc, createLegacyXls, createLegacyPpt]) assert.equal(typeof value, 'function');
for (const [value, target] of [[legacyDoc, 'docx'], [legacyXls, 'xlsx'], [legacyPpt, 'pptx']]) assert.equal(value.target, target);
for (const [entry, key] of [['chart-ex','chartEx'], ['tiff','tiff']]) assert.equal(typeof (await import('@silurus/ooxml/' + entry))[key].render, 'function');
assert.equal(typeof globalThis.document, 'undefined');
assert.throws(() => assertCompanionRuntime('0.0.0', 'stale'), error => error instanceof CompanionRuntimeMismatchError && error.code === 'companion-runtime-mismatch');
for (const entry of ['legacy-doc','legacy-xls','legacy-ppt']) await assert.rejects(import('@silurus/ooxml/' + entry), { code: 'ERR_PACKAGE_PATH_NOT_EXPORTED' });
const doc = await materializeDocxDocument(await readFile('public/sample.doc'), { legacyDoc });
assert.match(JSON.stringify(doc.body), /Hello/);
const xls = await materializeXlsxWorkbook(await readFile('public/sample.xls'), { legacyXls });
assert.equal(xls.worksheets.length, 1);
const ppt = await materializePptxPresentation(await readFile('public/sample.ppt'), { legacyPpt });
assert.match(JSON.stringify(ppt), /Picture fill text/);
for (const [family, open, options] of [
  ['doc', materializeDocxDocument, { legacyDoc: createLegacyDoc() }],
  ['xls', materializeXlsxWorkbook, { legacyXls: createLegacyXls() }],
  ['ppt', materializePptxPresentation, { legacyPpt: createLegacyPpt() }],
]) {
  await assert.rejects(async () => open(await readFile('public/unsupported.' + family), options),
    error => error instanceof OoxmlError && error.code === 'unsupported-content' && error.details.format === family && error.details.reason === 'source-subset');
}
await assert.rejects(async () => materializePptxPresentation(await readFile('public/deferred.ppt'), { legacyPpt }),
  error => error instanceof OoxmlError && error.code === 'unsupported-content' && error.details.format === 'ppt');
console.log('Packed Node imports and all three legacy readers passed');
`);
  process.stdout.write(run(process.execPath, ['node-smoke.mjs']));
  await writeFile(join(scratch, 'index.html'), '<div id="result">loading</div><script type="module" src="/main.js"></script>');
  await writeFile(join(scratch, 'main.js'), `
import { DocxDocument, DocxViewer, OoxmlError } from '@silurus/ooxml/docx';
import { XlsxWorkbook, XlsxViewer } from '@silurus/ooxml/xlsx';
import { PptxPresentation, PptxViewer } from '@silurus/ooxml/pptx';
import { legacyDoc, createLegacyDoc } from '@silurus/ooxml-legacy-office/doc';
import { legacyXls, createLegacyXls } from '@silurus/ooxml-legacy-office/xls';
import { legacyPpt, createLegacyPpt } from '@silurus/ooxml-legacy-office/ppt';
try {
  for (const [family, Loader, Viewer, sourceOptions] of [
    ['doc', DocxDocument, DocxViewer, { legacyDoc }],
    ['xls', XlsxWorkbook, XlsxViewer, { legacyXls }],
    ['ppt', PptxPresentation, PptxViewer, { legacyPpt }],
  ]) {
    for (const mode of ['main', 'worker']) {
      const bytes = await (await fetch('/sample.' + family)).arrayBuffer();
      const options = { mode, ...sourceOptions };
      const loaded = await Loader.load(bytes.slice(0), options);
      try { if (loaded.mode !== mode) throw new Error(family + ': expected ' + mode + ', got ' + loaded.mode); }
      finally { loaded.destroy(); }
      const unsupported = await (await fetch('/unsupported.' + family)).arrayBuffer();
      try { await Loader.load(unsupported, options); throw new Error('unsupported input was accepted'); }
      catch (error) {
        if (!(error instanceof OoxmlError) || error.code !== 'unsupported-content'
          || error.details?.format !== family || error.details?.reason !== 'source-subset') throw error;
      }
      if (family === 'ppt') {
        const deferred = await (await fetch('/deferred.ppt')).arrayBuffer();
        let deck;
        try {
          deck = await Loader.load(deferred, options);
          await deck.collectSlideRuns(0);
          throw new Error('unsupported slide was accepted');
        } catch (error) {
          if (!(error instanceof OoxmlError) || error.code !== 'unsupported-content' || error.details?.format !== family) throw error;
        } finally { deck?.destroy(); }
      }
      const element = document.createElement(family === 'xls' ? 'div' : 'canvas');
      element.style.cssText = 'width:600px;height:400px';
      document.body.append(element);
      const viewer = new Viewer(element, options);
      try { await viewer.load(bytes); }
      finally { viewer.destroy(); element.remove(); }
    }
  }
  window.packedResult = 'passed';
} catch (error) { window.packedResult = String(error?.stack || error); }
`);
  await build({ configFile: false, root: scratch, base: './', logLevel: 'warn', build: { target: 'esnext' } });
  server = await preview({ configFile: false, root: scratch, preview: { host: '127.0.0.1', port: 0, open: false } });
  browser = await chromium.launch({ channel: 'chrome', headless: true });
  const page = await browser.newPage();
  const errors = [];
  page.on('pageerror', error => errors.push(error.message));
  await page.goto(server.resolvedUrls.local[0]);
  await page.waitForFunction(() => window.packedResult !== undefined, undefined, { timeout: 90_000 });
  assert.equal(await page.evaluate(() => window.packedResult), 'passed');
  assert.deepEqual(errors, []);
  console.log('Packed Vite viewers/loaders: DOC/XLS/PPT main + worker passed');
} finally {
  await browser?.close();
  await server?.close();
  await rm(scratch, { recursive: true, force: true });
}
