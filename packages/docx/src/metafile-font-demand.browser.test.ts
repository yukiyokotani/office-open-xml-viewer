import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { chromium, type Browser, type Page } from '@playwright/test';
import { createServer, type ViteDevServer } from 'vite';
import { existsSync, readFileSync } from 'node:fs';
import { wmfFixture, wmfRecords } from '@silurus/ooxml-core/testing/metafile-paint';
import { storeZip } from './conformance/generate.js';

// Metafile text demands its own scalars when the picture plays (core
// bindMetafileFontDemand), through the production load and render paths in
// main and worker mode: a served family split into two `unicode-range`
// subsets, body text in the first only, a WMF picture painting text of the
// second. The second subset loads when, and only when, the picture plays.

type Api = typeof import('./index');

const FONT = readFileSync(new URL('../../../tests/fixtures/fonts/ooxml-fixture-gap.ttf', import.meta.url));
const FAMILY = 'Ooxml Meta Gap';
const encoder = new TextEncoder();

function docx(withPicture: boolean): number[] {
  const parts = new Map<string, Uint8Array>();
  parts.set('[Content_Types].xml', encoder.encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="wmf" ContentType="image/x-wmf"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'));
  parts.set('_rels/.rels', encoder.encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'));
  parts.set('word/_rels/document.xml.rels', encoder.encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdM0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.wmf"/></Relationships>'));
  // Greek capital omega (CP1253 0xD9) in the served family.
  parts.set('word/media/image1.wmf', wmfFixture([wmfRecords.setWindowExt(1000, 1000), wmfRecords.createFont(FAMILY, { charset: 161 }),
    wmfRecords.selectObject(0), wmfRecords.textOut(new Uint8Array([0x41, 0xd9]))]));
  const picture = '<w:r><w:drawing><wp:inline><wp:extent cx="1905000" cy="952500"/><wp:docPr id="1" name="p"/>'
    + '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>'
    + '<pic:nvPicPr><pic:cNvPr id="1" name="p"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdM0"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
    + '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
    + '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
  const body = `<w:p><w:r><w:rPr><w:rFonts w:ascii="${FAMILY}" w:hAnsi="${FAMILY}"/></w:rPr><w:t>Body</w:t></w:r>${withPicture ? picture : ''}</w:p>`;
  parts.set('word/document.xml', encoder.encode('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    + ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
    + ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
    + ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
    + `<w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`));
  return Array.from(storeZip(parts));
}

async function launch(): Promise<Browser | null> {
  try {
    return existsSync(chromium.executablePath()) ? await chromium.launch() : await chromium.launch({ channel: 'chrome' });
  } catch {
    return null;
  }
}

describe('DOCX metafile text demand in the browser', () => {
  let server: ViteDevServer;
  let browser: Browser | null;
  let page: Page;
  const requests: string[] = [];

  beforeAll(async () => {
    browser = await launch();
    if (!browser) return;
    server = await createServer({
      configFile: new URL('../vite.config.ts', import.meta.url).pathname,
      server: { host: '127.0.0.1', port: 6687, strictPort: true },
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
    await page.goto('http://127.0.0.1:6687/metafile-demand.html');
    await page.addScriptTag({ type: 'module', content: `
      import * as api from '/src/index.ts';
      window.metafileDemandTest = api;
    ` });
    await page.waitForFunction(() => 'metafileDemandTest' in window, null, { timeout: 90_000 });
  }, 120_000);
  afterAll(async () => { await browser?.close(); await server?.close(); }, 60_000);

  /** The distinct face files requested while loading and rendering page 1. */
  async function faceRequests(bytes: number[], mode: 'main' | 'worker'): Promise<string[]> {
    requests.length = 0;
    await page.evaluate(async ({ bytes, mode }) => {
      const api = (window as unknown as { metafileDemandTest: Api }).metafileDemandTest;
      // A fresh stylesheet URL per load: no source cache answers for another.
      const url = `${location.origin}/meta-fonts/${Math.random().toString(36).slice(2)}/faces.css`;
      const doc = await api.DocxDocument.load(new Uint8Array(bytes).buffer, { mode, fonts: [api.fontCss({ url })] });
      try {
        if (mode === 'worker') (await doc.renderPageToBitmap(0, { width: 600 })).close();
        else await doc.renderPage(document.createElement('canvas'), 0, { width: 600 });
      } finally { doc.destroy(); }
    }, { bytes, mode });
    // A realm fetches each face once; worker mode may load in two realms.
    return [...new Set(requests.filter((name) => name.endsWith('.ttf')))].sort();
  }

  for (const mode of ['main', 'worker'] as const) it(`loads the subset a WMF picture paints when it plays (${mode} mode)`, async ({ skip }) => {
    // Reported as skipped, never as passing, where no browser can launch.
    if (!browser) skip();
    // Without the picture only the body text's subset loads (finite demand).
    expect(await faceRequests(docx(false), mode)).toEqual(['latin.ttf']);
    expect(await faceRequests(docx(true), mode)).toEqual(['greek.ttf', 'latin.ttf']);
  }, 60_000);
});
