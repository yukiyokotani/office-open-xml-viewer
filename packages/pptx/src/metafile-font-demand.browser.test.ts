import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { chromium, type Browser, type Page } from '@playwright/test';
import { createServer, type ViteDevServer } from 'vite';
import { existsSync, readFileSync } from 'node:fs';
import { wmfFixture, wmfRecords } from '@silurus/ooxml-core/testing';
import { deck, para, picture, textBox } from './test-support/synthetic-deck.js';

// Metafile text demands its own scalars when the picture plays (core
// bindMetafileFontDemand), through the production load and render paths in
// main and worker mode: a served family split into two `unicode-range`
// subsets, body text in the first only, a WMF picture painting text of the
// second. The second subset loads when, and only when, the picture plays.

type Api = typeof import('./index');

const FONT = readFileSync(new URL('../../../tests/fixtures/fonts/ooxml-fixture-gap.ttf', import.meta.url));
const FAMILY = 'Ooxml Meta Gap';

function slide(withPicture: boolean): number[] {
  // Greek capital omega (CP1253 0xD9) in the served family.
  const wmf = wmfFixture([wmfRecords.setWindowExt(1000, 1000), wmfRecords.createFont(FAMILY, { charset: 161 }),
    wmfRecords.selectObject(0), wmfRecords.textOut(new Uint8Array([0x41, 0xd9]))]);
  const shapes = textBox(para(`<a:r><a:rPr lang="en-US" sz="2400"><a:latin typeface="${FAMILY}"/></a:rPr><a:t>Body</a:t></a:r>`))
    + (withPicture ? picture(0) : '');
  return Array.from(deck(shapes, withPicture ? new Map([['image1.wmf', wmf]]) : new Map()));
}

async function launch(): Promise<Browser | null> {
  try {
    return existsSync(chromium.executablePath()) ? await chromium.launch() : await chromium.launch({ channel: 'chrome' });
  } catch {
    return null;
  }
}

describe('PPTX metafile text demand in the browser', () => {
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
      server: { host: '127.0.0.1', port: 6688, strictPort: true },
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
    await page.goto('http://127.0.0.1:6688/metafile-demand.html');
    await page.addScriptTag({ type: 'module', content: `
      import * as api from '/src/index.ts';
      window.metafileDemandTest = api;
    ` });
    await page.waitForFunction(() => 'metafileDemandTest' in window, null, { timeout: 90_000 });
  }, 120_000);
  afterAll(async () => { await browser?.close(); await server?.close(); }, 60_000);

  /** The distinct face files requested while loading and rendering slide 1. */
  async function faceRequests(bytes: number[], mode: 'main' | 'worker'): Promise<string[]> {
    requests.length = 0;
    await page.evaluate(async ({ bytes, mode }) => {
      const api = (window as unknown as { metafileDemandTest: Api }).metafileDemandTest;
      // A fresh stylesheet URL per load: no source cache answers for another.
      const url = `${location.origin}/meta-fonts/${Math.random().toString(36).slice(2)}/faces.css`;
      const deck = await api.PptxPresentation.load(new Uint8Array(bytes).buffer, { mode, fonts: [api.fontCss({ url })] });
      try {
        if (mode === 'worker') (await deck.renderSlideToBitmap(0, { width: 600 })).close();
        else await deck.renderSlide(document.createElement('canvas'), 0, { width: 600 });
      } finally { deck.destroy(); }
    }, { bytes, mode });
    // A realm fetches each face once; worker mode may load in two realms.
    return [...new Set(requests.filter((name) => name.endsWith('.ttf')))].sort();
  }

  it.for(['main', 'worker'] as const)('loads the subset a WMF picture paints when it plays (%s mode)', async (mode, { skip }) => {
    // Reported as skipped, never as passing, where no browser can launch.
    if (!browser) skip();
    // Without the picture only the body text's subset loads (finite demand).
    expect(await faceRequests(slide(false), mode)).toEqual(['latin.ttf']);
    expect(await faceRequests(slide(true), mode)).toEqual(['greek.ttf', 'latin.ttf']);
  }, 60_000);
});
