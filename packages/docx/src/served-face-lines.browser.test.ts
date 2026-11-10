import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { chromium, type Browser, type Page } from '@playwright/test';
import { createServer, type ViteDevServer } from 'vite';
import { existsSync, readFileSync } from 'node:fs';
import { storeZip } from './conformance/generate.js';

// Served faces in a real engine: the line pitch of 11pt text in a face with
// hhea ascender 800, descender 200 and lineGap 300 (per 1000) must be Word's
// 1.3 em whichever way the face reaches the page, in main and worker mode.
// Synthetic fixtures: scripts/generate-fixture-font.py.

type Api = typeof import('./index');

const FIXTURES = new URL('../../../tests/fixtures/fonts/', import.meta.url);
const font = (name: string) => readFileSync(new URL(name, FIXTURES));
const EM_PX = 11 * 4 / 3;

function docx(family: string, bold = false): number[] {
  const encoder = new TextEncoder();
  const parts = new Map<string, Uint8Array>();
  parts.set('[Content_Types].xml', encoder.encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'));
  parts.set('_rels/.rels', encoder.encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'));
  const paragraph = `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="${family}" w:hAnsi="${family}"/>${bold ? '<w:b/>' : ''}<w:sz w:val="22"/></w:rPr><w:t>Served face line</w:t></w:r></w:p>`;
  parts.set('word/document.xml', encoder.encode(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${paragraph.repeat(3)}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`));
  return Array.from(storeZip(parts));
}

async function launch(): Promise<Browser | null> {
  try {
    return existsSync(chromium.executablePath()) ? await chromium.launch() : await chromium.launch({ channel: 'chrome' });
  } catch {
    return null;
  }
}

describe('served face line metrics in the browser', () => {
  let server: ViteDevServer;
  let browser: Browser | null;
  let page: Page;

  beforeAll(async () => {
    browser = await launch();
    if (!browser) return;
    server = await createServer({
      configFile: new URL('../vite.config.ts', import.meta.url).pathname,
      server: { host: '127.0.0.1', port: 6686, strictPort: true },
      logLevel: 'error',
    });
    await server.listen();
    page = await browser.newPage();
    await page.addInitScript(() => { (window as unknown as { __OOXML_MODEL_SOURCES__: boolean }).__OOXML_MODEL_SOURCES__ = true; });
    await page.route('**/served-face.html', (route) => route.fulfill({ contentType: 'text/html', body: '<!doctype html><body></body>' }));
    await page.route('**/fixture-fonts/**', (route) => {
      const name = new URL(route.request().url()).pathname.split('/').pop()!;
      if (name.endsWith('.css')) {
        return route.fulfill({ contentType: 'text/css', body: `@font-face { font-family: 'Ooxml Css Gap'; src: url(/fixture-fonts/ooxml-fixture-gap.woff2) format('woff2'); }` });
      }
      return route.fulfill({ contentType: 'font/ttf', body: font(name) });
    });
    await page.goto('http://127.0.0.1:6686/served-face.html');
    await page.addScriptTag({ type: 'module', content: `
      import * as api from '/src/index.ts';
      window.servedFaceTest = api;
    ` });
    await page.waitForFunction(() => 'servedFaceTest' in window);
  }, 60_000);
  afterAll(async () => { await browser?.close(); await server?.close(); }, 60_000);

  /** Line pitch in em of the document's first two lines, per mode. */
  async function pitches(
    bytes: number[],
    fonts: 'files-url' | 'files-woff' | 'files-woff2' | 'files-var' | 'css' | 'none',
    appFace?: string,
  ): Promise<number[]> {
    return page.evaluate(async ({ bytes, fonts, appFace, em }) => {
      const api = (window as unknown as { servedFaceTest: Api }).servedFaceTest;
      const origin = location.origin;
      const file = { 'files-url': 'ttf', 'files-woff': 'woff', 'files-woff2': 'woff2' }[fonts as string];
      const sources = file
        ? [api.fontFiles({ faces: [{ family: 'Ooxml Fixture Gap', url: `${origin}/fixture-fonts/ooxml-fixture-gap.${file}` }] })]
        : fonts === 'files-var'
          ? [api.fontFiles({ faces: [{ family: 'Ooxml Var Gap', url: `${origin}/fixture-fonts/ooxml-fixture-gap-var.ttf`, weight: '100 900' }] })]
          : fonts === 'css' ? [api.fontCss({ url: `${origin}/fixture-fonts/gap.css` })] : undefined;
      let registered: FontFace | undefined;
      if (appFace) {
        const data = await (await fetch(`${origin}/fixture-fonts/ooxml-fixture-gap.ttf`)).arrayBuffer();
        registered = new FontFace(appFace, data);
        document.fonts.add(registered);
        await registered.load();
      }
      const values: number[] = [];
      try {
        for (const mode of appFace ? ['main'] as const : ['main', 'worker'] as const) {
          const doc = await api.DocxDocument.load(new Uint8Array(bytes).buffer, { mode, ...(sources ? { fonts: sources } : {}) });
          try {
            const runs = await doc.collectPageRuns(0);
            const ys = [...new Set(runs.map((run) => run.y))].sort((a, b) => a - b);
            values.push((ys[1]! - ys[0]!) / em);
          } finally { doc.destroy(); }
        }
      } finally {
        if (registered) document.fonts.delete(registered);
      }
      return values;
    }, { bytes, fonts, appFace, em: EM_PX });
  }

  it('sizes a fontFiles URL face from its own tables (library fetch)', async (test) => {
    if (!browser) test.skip();
    for (const pitch of await pitches(docx('Ooxml Fixture Gap'), 'files-url')) expect(pitch).toBeCloseTo(1.3, 2);
  }, 60_000);

  it.each(['files-woff', 'files-woff2'] as const)('sizes a %s face from its tables, or its measured line without native Brotli', async (kind) => {
    if (!browser) return;
    for (const pitch of await pitches(docx('Ooxml Fixture Gap'), kind)) expect(pitch).toBeCloseTo(1.3, 2);
  }, 60_000);

  it('sizes a stylesheet face from its own tables (library fetch)', async (test) => {
    if (!browser) test.skip();
    for (const pitch of await pitches(docx('Ooxml Css Gap'), 'css')) expect(pitch).toBeCloseTo(1.3, 2);
  }, 60_000);

  it('sizes an application-registered FontFace from its measured normal line', async (test) => {
    if (!browser) test.skip();
    for (const pitch of await pitches(docx('Ooxml App Gap'), 'none', 'Ooxml App Gap')) expect(pitch).toBeCloseTo(1.3, 2);
  }, 60_000);

  it('sizes a variable face at the instance the run selects', async (test) => {
    if (!browser) test.skip();
    // wght 400: (800 + 200 + 300) / 1000; wght 700: (860 + 200 + 420) / 1000.
    for (const pitch of await pitches(docx('Ooxml Var Gap'), 'files-var')) expect(pitch).toBeCloseTo(1.3, 2);
    for (const pitch of await pitches(docx('Ooxml Var Gap', true), 'files-var')) expect(pitch).toBeCloseTo(1.48, 2);
  }, 60_000);
});
