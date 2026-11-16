import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { chromium, type Browser, type Page } from '@playwright/test';
import { createServer, type ViteDevServer } from 'vite';
import { existsSync, readFileSync } from 'node:fs';
import { storedZip, wmfFixture, wmfRecords } from '@silurus/ooxml-core/testing';

// Metafile text demands its own scalars when the picture plays (core
// bindMetafileFontDemand), through the production load and render paths in
// main and worker mode: a served family split into two `unicode-range`
// subsets, cell text in the first only, a WMF picture painting text of the
// second. The second subset loads when, and only when, the picture plays.

type Api = typeof import('./index');

const FONT = readFileSync(new URL('../../../tests/fixtures/fonts/ooxml-fixture-gap.ttf', import.meta.url));
const FAMILY = 'Ooxml Meta Gap';
const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';

function workbookBytes(withPicture: boolean): number[] {
  // Greek capital omega (CP1253 0xD9) in the served family.
  const wmf = wmfFixture([wmfRecords.setWindowExt(1000, 1000), wmfRecords.createFont(FAMILY, { charset: 161 }),
    wmfRecords.selectObject(0), wmfRecords.textOut(new Uint8Array([0x41, 0xd9]))]);
  const anchor = '<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
    + '<xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>8</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
    + '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="P"/><xdr:cNvPicPr/></xdr:nvPicPr>'
    + '<xdr:blipFill><a:blip r:embed="rIdM0"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>'
    + '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3000000" cy="1200000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>';
  return Array.from(storedZip([
    ['[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      + '<Default Extension="wmf" ContentType="image/x-wmf"/>'
      + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
      + '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
      + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
      + '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>'],
    ['_rels/.rels', `<Relationships xmlns="${PKG}"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`],
    ['xl/workbook.xml', `<workbook xmlns="${MAIN}" xmlns:r="${REL}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`],
    ['xl/_rels/workbook.xml.rels', `<Relationships xmlns="${PKG}"><Relationship Id="rId1" Type="${REL}/worksheet" Target="worksheets/sheet1.xml"/>`
      + `<Relationship Id="rId2" Type="${REL}/styles" Target="styles.xml"/></Relationships>`],
    ['xl/styles.xml', `<styleSheet xmlns="${MAIN}"><fonts count="1"><font><sz val="11"/><name val="${FAMILY}"/></font></fonts>`
      + '<fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders>'
      + '<cellXfs count="1"><xf numFmtId="0" fontId="0"/></cellXfs></styleSheet>'],
    ['xl/worksheets/sheet1.xml', `<worksheet xmlns="${MAIN}" xmlns:r="${REL}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Body</t></is></c></row></sheetData>`
      + '<drawing r:id="rIdD"/></worksheet>'],
    ['xl/worksheets/_rels/sheet1.xml.rels', `<Relationships xmlns="${PKG}"><Relationship Id="rIdD" Type="${REL}/drawing" Target="../drawings/drawing1.xml"/></Relationships>`],
    ['xl/drawings/drawing1.xml', '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
      + ` xmlns:r="${REL}">${withPicture ? anchor : ''}</xdr:wsDr>`],
    ['xl/drawings/_rels/drawing1.xml.rels', `<Relationships xmlns="${PKG}"><Relationship Id="rIdM0" Type="${REL}/image" Target="../media/image1.wmf"/></Relationships>`],
    ['xl/media/image1.wmf', wmf],
  ]));
}

async function launch(): Promise<Browser | null> {
  try {
    return existsSync(chromium.executablePath()) ? await chromium.launch() : await chromium.launch({ channel: 'chrome' });
  } catch {
    return null;
  }
}

describe('XLSX metafile text demand in the browser', () => {
  let server: ViteDevServer;
  let browser: Browser | null;
  let page: Page;
  const requests: string[] = [];

  beforeAll(async () => {
    browser = await launch();
    if (!browser) return;
    server = await createServer({
      configFile: new URL('../vite.config.ts', import.meta.url).pathname,
      root: new URL('..', import.meta.url).pathname,
      server: { host: '127.0.0.1', port: 6689, strictPort: true },
      logLevel: 'error',
    });
    await server.listen();
    const context = await browser.newContext();
    page = await context.newPage();
    await page.addInitScript(() => { (window as unknown as { __OOXML_MODEL_SOURCES__: boolean }).__OOXML_MODEL_SOURCES__ = true; });
    await page.route('**/metafile-demand.html', (route) => route.fulfill({ contentType: 'text/html', body: '<!doctype html><body></body>' }));
    await context.route('**/meta-fonts/**', (route) => {
      // /meta-fonts/<load>/<file>: a fresh path per load, so no realm cache
      // of an earlier load answers for this one.
      const [, , load, name] = new URL(route.request().url()).pathname.split('/');
      requests.push(name!);
      if (name!.endsWith('.css')) {
        return route.fulfill({ contentType: 'text/css', body: `@font-face { font-family: '${FAMILY}'; src: url(/meta-fonts/${load}/latin.ttf); unicode-range: U+0000-00FF; }`
          + `@font-face { font-family: '${FAMILY}'; src: url(/meta-fonts/${load}/greek.ttf); unicode-range: U+0370-03FF; }` });
      }
      return route.fulfill({ contentType: 'font/ttf', body: FONT });
    });
    await page.goto('http://127.0.0.1:6689/metafile-demand.html');
    await page.addScriptTag({ type: 'module', content: `
      import * as api from '/src/index.ts';
      window.metafileDemandTest = api;
    ` });
    await page.waitForFunction(() => 'metafileDemandTest' in window, null, { timeout: 90_000 });
  }, 120_000);
  afterAll(async () => { await browser?.close(); await server?.close(); }, 60_000);

  /** The distinct face files requested while loading and rendering the first sheet's viewport. */
  async function faceRequests(bytes: number[], mode: 'main' | 'worker'): Promise<string[]> {
    requests.length = 0;
    await page.evaluate(async ({ bytes, mode }) => {
      const api = (window as unknown as { metafileDemandTest: Api }).metafileDemandTest;
      // A fresh stylesheet URL per load: no source cache answers for another.
      const url = `${location.origin}/meta-fonts/${Math.random().toString(36).slice(2)}/faces.css`;
      const workbook = await api.XlsxWorkbook.load(new Uint8Array(bytes).buffer, { mode, fonts: [api.fontCss({ url })] });
      const viewport = { row: 1, col: 1, rows: 20, cols: 10 };
      try {
        if (mode === 'worker') (await workbook.renderViewportToBitmap(0, viewport, { width: 800, height: 400 })).close();
        else {
          const canvas = document.createElement('canvas');
          canvas.width = 800; canvas.height = 400;
          await workbook.renderViewport(canvas, 0, viewport);
        }
      } finally { workbook.destroy(); }
    }, { bytes, mode });
    // A realm fetches each face once; worker mode may load in two realms.
    return [...new Set(requests.filter((name) => name.endsWith('.ttf')))].sort();
  }

  it.for(['main', 'worker'] as const)('loads the subset a WMF picture paints when it plays (%s mode)', async (mode, { skip }) => {
    // Reported as skipped, never as passing, where no browser can launch.
    if (!browser) skip();
    // Without the picture only the body text's subset loads (finite demand).
    expect(await faceRequests(workbookBytes(false), mode)).toEqual(['latin.ttf']);
    expect(await faceRequests(workbookBytes(true), mode)).toEqual(['greek.ttf', 'latin.ttf']);
  }, 60_000);
});
