import { readFileSync } from 'node:fs';
import { expect, test, type Browser } from '@playwright/test';
// @ts-ignore - shared plain-JS test fixture builder.
import { storedZip } from '../fixtures/chart-ex-packages.mjs';
import { deck, para, textBox } from '../../packages/pptx/src/test-support/synthetic-deck.js';

// `fontResolver` and `fontFiles` data: URLs through the published dist in
// main and worker mode, for every format (font-provider design §3.4, §4.3,
// §12.1). In worker mode the application function stays on main: the worker's
// session asks main, which resolves through its realm resolution cache and
// sends a copy of the declared bytes. Main and worker mode must paint the same pixels, and the
// fixture face must actually paint (it differs visibly from host fallback).
// The font is the committed synthetic fixture from
// scripts/generate-fixture-font.py (no third-party copyright).

const FONT = readFileSync(new URL('../fixtures/fonts/ooxml-fixture.ttf', import.meta.url));
const FAMILY = 'Ooxml Fixture';
const TEXT = 'Fixture text ABC 123';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OFFICE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

function fixtureDocx(): Uint8Array {
  const run = `<w:r><w:rPr><w:rFonts w:ascii="${FAMILY}" w:hAnsi="${FAMILY}" w:eastAsia="${FAMILY}" w:cs="${FAMILY}"/>`
    + `<w:sz w:val="48"/></w:rPr><w:t xml:space="preserve">${TEXT}</w:t></w:r>`;
  return storedZip([
    ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + '</Types>'],
    ['_rels/.rels', `<Relationships xmlns="${PKG_REL}"><Relationship Id="rId1" Type="${OFFICE_REL}/officeDocument" Target="word/document.xml"/></Relationships>`],
    ['word/document.xml', `<w:document xmlns:w="${W}"><w:body><w:p>${run}</w:p>`
      + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"'
      + ' w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>'],
  ]) as Uint8Array;
}

function fixtureXlsx(): Uint8Array {
  return storedZip([
    ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
      + '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
      + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
      + '</Types>'],
    ['_rels/.rels', `<Relationships xmlns="${PKG_REL}"><Relationship Id="rId1" Type="${OFFICE_REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`],
    ['xl/workbook.xml', `<workbook xmlns="${S}" xmlns:r="${OFFICE_REL}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`],
    ['xl/_rels/workbook.xml.rels', `<Relationships xmlns="${PKG_REL}">`
      + `<Relationship Id="rId1" Type="${OFFICE_REL}/worksheet" Target="worksheets/sheet1.xml"/>`
      + `<Relationship Id="rId2" Type="${OFFICE_REL}/styles" Target="styles.xml"/></Relationships>`],
    ['xl/styles.xml', `<styleSheet xmlns="${S}"><fonts count="1"><font><sz val="20"/><name val="${FAMILY}"/></font></fonts>`
      + '<fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders>'
      + '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs></styleSheet>'],
    ['xl/worksheets/sheet1.xml', `<worksheet xmlns="${S}"><cols><col min="1" max="1" width="40" customWidth="1"/></cols><sheetData>`
      + `<row r="1" ht="30" customHeight="1"><c r="A1" t="inlineStr"><is><t>${TEXT}</t></is></c></row>`
      + '</sheetData></worksheet>'],
  ]) as Uint8Array;
}

const fixturePptx = () => deck(textBox(para(`<a:r><a:rPr lang="en-US" sz="4000"><a:latin typeface="${FAMILY}"/>`
  + `<a:ea typeface="${FAMILY}"/><a:cs typeface="${FAMILY}"/></a:rPr><a:t>${TEXT}</a:t></a:r>`)));

const FIXTURES = { docx: fixtureDocx(), xlsx: fixtureXlsx(), pptx: fixturePptx() } as const;

type Format = keyof typeof FIXTURES;
type Source = 'resolver' | 'data-url' | 'none';

/** RGBA pixels of one rendered bitmap. */
interface Pixels { readonly width: number; readonly height: number; readonly data: Buffer }

interface Rendered {
  readonly png: Pixels;
  /** Resolver calls, the demand each call received (its field names and
   * families), and whether the answered buffer is still attached on main
   * after the load. */
  readonly calls: number;
  readonly demands: readonly { readonly keys: readonly string[]; readonly families: readonly string[] }[];
  readonly answeredBytes: readonly number[];
}

/** `modelSource`: load through an application model source (the repository's
 * test-only `fake-model-source.mjs`, which opens the OOXML package itself),
 * the path model-source loaders take instead of the ordinary loader. */
async function render(browser: Browser, format: Format, mode: 'main' | 'worker', source: Source, modelSource = false): Promise<Rendered> {
  // A fresh context per load: no shared FontFaceSet, worker or realm state.
  const context = await browser.newContext();
  try {
    const page = await context.newPage();
    await page.route('**/font-resolver-bridge', (route) => route.fulfill({
      contentType: 'text/html', body: '<!doctype html><body></body>',
    }));
    await page.goto('/font-resolver-bridge');
    const result = await page.evaluate(async ({ format, mode, source, modelSource, documentBytes, fontBytes, family }) => {
      const entry = `/dist/${format}.mjs`;
      const api = await import(entry);
      const font = new Uint8Array(fontBytes);
      const calls: { keys: string[]; families: string[] }[] = [];
      const answered: ArrayBuffer[] = [];
      let fonts: unknown[] | undefined;
      if (source === 'resolver') {
        fonts = [api.fontResolver((demand: { families: string[]; fallbackFamilies: string[] }, context: { signal: AbortSignal }) => {
          if (!(context.signal instanceof AbortSignal)) throw new Error('resolver received no AbortSignal');
          calls.push({ keys: Object.keys(demand).sort(), families: [...demand.families] });
          if (!demand.families.includes(family)) return [];
          const data = font.slice().buffer;
          answered.push(data);
          return [{ family, data }];
        })];
      } else if (source === 'data-url') {
        let binary = '';
        for (const byte of font) binary += String.fromCharCode(byte);
        fonts = [api.fontFiles({ faces: [{ family, url: `data:font/ttf;base64,${btoa(binary)}` }] })];
      }
      const sourceOption = { docx: 'legacyDoc', pptx: 'legacyPpt', xlsx: 'legacyXls' }[format];
      const modelSourceOptions = modelSource ? {
        [sourceOption]: {
          target: format,
          claim: () => true,
          beginLoad: () => ({
            module: {
              protocol: 'ooxml-model-source-module/v2',
              target: format,
              moduleUrl: new URL('/tests/worker-dist/fake-model-source.mjs', location.href).href,
              config: { format },
            },
            release() {},
          }),
        },
      } : {};
      const options = { mode, ...(fonts ? { fonts } : {}), ...modelSourceOptions };
      const bytes = new Uint8Array(documentBytes).buffer;
      let bitmap: ImageBitmap;
      if (format === 'docx') {
        const loaded = await api.DocxDocument.load(bytes, options);
        bitmap = await loaded.renderPageToBitmap(0, { width: 816, dpr: 1 });
        loaded.destroy();
      } else if (format === 'pptx') {
        const loaded = await api.PptxPresentation.load(bytes, options);
        bitmap = await loaded.renderSlideToBitmap(0, { width: 960, dpr: 1 });
        loaded.destroy();
      } else {
        const loaded = await api.XlsxWorkbook.load(bytes, options);
        bitmap = await loaded.renderViewportToBitmap(0, { row: 1, col: 1, rows: 4, cols: 4 }, { width: 640, height: 160, dpr: 1 });
        loaded.destroy();
      }
      const canvas = document.createElement('canvas');
      canvas.width = bitmap.width;
      canvas.height = bitmap.height;
      canvas.getContext('2d')!.drawImage(bitmap, 0, 0);
      bitmap.close();
      const rgba = canvas.getContext('2d')!.getImageData(0, 0, canvas.width, canvas.height).data;
      let binary = '';
      for (let offset = 0; offset < rgba.length; offset += 0x8000) {
        binary += String.fromCharCode(...rgba.subarray(offset, offset + 0x8000));
      }
      return {
        width: canvas.width,
        height: canvas.height,
        rgba: btoa(binary),
        calls: calls.length,
        demands: calls,
        answeredBytes: answered.map((buffer) => buffer.byteLength),
      };
    }, { format, mode, source, modelSource, documentBytes: [...FIXTURES[format]], fontBytes: [...FONT], family: FAMILY });
    return {
      png: { width: result.width, height: result.height, data: Buffer.from(result.rgba, 'base64') },
      calls: result.calls,
      demands: result.demands,
      answeredBytes: result.answeredBytes,
    };
  } finally {
    await context.close();
  }
}

/** Percentage of pixels whose channels differ by more than ~10%. */
function diffPct(a: Pixels, b: Pixels): number {
  let differing = 0;
  for (let offset = 0; offset < a.data.length; offset += 4) {
    let delta = 0;
    for (let channel = 0; channel < 4; channel++) {
      delta = Math.max(delta, Math.abs(a.data[offset + channel] - b.data[offset + channel]));
    }
    if (delta > 25) differing++;
  }
  return (differing / (a.width * a.height)) * 100;
}

for (const format of ['docx', 'pptx', 'xlsx'] as const) {
  test(`${format}: fontResolver bytes and fontFiles data: URLs paint alike in main and worker mode`, async ({ browser }) => {
    test.setTimeout(180_000);
    const none = await render(browser, format, 'main', 'none');
    for (const source of ['resolver', 'data-url'] as const) {
      const main = await render(browser, format, 'main', source);
      const worker = await render(browser, format, 'worker', source);
      expect(worker.png.width).toBe(main.png.width);
      expect(worker.png.height).toBe(main.png.height);
      const mainVsWorker = diffPct(main.png, worker.png);
      const mainVsNone = diffPct(main.png, none.png);
      console.log(`  ${format} ${source}: worker-vs-main ${mainVsWorker.toFixed(3)}%, fixture-vs-fallback ${mainVsNone.toFixed(3)}%`);
      expect(mainVsWorker).toBeLessThanOrEqual(0.2);
      expect(mainVsNone).toBeGreaterThan(0.5);
      if (source !== 'resolver') continue;
      // The application function ran on main in both modes and saw the same
      // demand: family names only (FontDemand), the fixture family among them.
      for (const demand of [...main.demands, ...worker.demands]) {
        expect(demand.keys).toEqual(['fallbackFamilies', 'families']);
        expect(demand.families).toContain(FAMILY);
      }
      expect(worker.demands).toEqual(main.demands);
      expect(main.calls).toBe(1);
      expect(main.answeredBytes).toEqual([FONT.byteLength]);
      // Once per family in both modes: the worker's request resolves through
      // main's resolution cache (shared with the XLSX viewer realm), which
      // keeps the original; the worker receives a copy, nothing is detached.
      expect(worker.calls).toBe(1);
      expect(worker.answeredBytes).toEqual([FONT.byteLength]);
    }
  });
}

// Model-source loaders construct the document and parse through their own
// path; the worker's resolver calls must still reach this document's bridge.
for (const format of ['docx', 'pptx', 'xlsx'] as const) {
  test(`${format}: a model-source load in worker mode runs the fontResolver on main`, async ({ browser }) => {
    test.setTimeout(120_000);
    const main = await render(browser, format, 'main', 'resolver');
    const worker = await render(browser, format, 'worker', 'resolver', true);
    expect(worker.calls).toBe(1);
    expect(worker.answeredBytes).toEqual([FONT.byteLength]);
    expect(diffPct(main.png, worker.png)).toBeLessThanOrEqual(0.2);
  });
}
