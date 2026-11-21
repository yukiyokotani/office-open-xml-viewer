import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { chromium, type Browser, type Page } from '@playwright/test';
import { createServer, type ViteDevServer } from 'vite';
import { existsSync } from 'node:fs';

type Api = typeof import('@silurus/ooxml-core') & typeof import('./layout-runtime');

// core reference-advances.ts in a real engine, with the DOCX probe clusters.
// A size-dependent design (macOS system-ui has an opsz axis and tracking:
// its 11px advances are wider than 256px advances scaled down, by fractional
// amounts) is font data and must keep measuring at its own size; a colour
// bitmap face whose small-size advances are whole pixels (Apple Color Emoji
// on macOS) is grid fitting and moves to the reference em.
describe.skipIf(!existsSync(chromium.executablePath()))('reference-em planning in Chromium', () => {
  let server: ViteDevServer;
  let browser: Browser;
  let page: Page;

  beforeAll(async () => {
    server = await createServer({
      configFile: new URL('../vite.config.ts', import.meta.url).pathname,
      server: { host: '127.0.0.1', port: 6687, strictPort: true },
      logLevel: 'error',
    });
    await server.listen();
    browser = await chromium.launch();
    page = await browser.newPage();
    await page.addInitScript(() => { (window as unknown as { __OOXML_MODEL_SOURCES__: boolean }).__OOXML_MODEL_SOURCES__ = true; });
    await page.route('**/reference-advances.html', route => route.fulfill({
      contentType: 'text/html', body: '<!doctype html><body></body>',
    }));
    await page.goto('http://127.0.0.1:6687/reference-advances.html');
    await page.addScriptTag({ type: 'module', content: `
      import * as runtime from '/src/layout-runtime.ts';
      import * as core from ${JSON.stringify('/@fs' + new URL('../../core/src/index.ts', import.meta.url).pathname)};
      window.referenceAdvanceTest = { ...core, ...runtime };
    ` });
    await page.waitForFunction(() => 'referenceAdvanceTest' in window);
  }, 60_000);
  afterAll(async () => { await browser?.close(); await server?.close(); });

  it('keeps optical-size designs at their own size and measures them natively', async () => {
    const result = await page.evaluate(() => {
      const api = (window as unknown as { referenceAdvanceTest: Api }).referenceAdvanceTest;
      const ctx = document.createElement('canvas').getContext('2d') as CanvasRenderingContext2D;
      const planner = api.createTextAdvancePlanner(ctx, { probeClusters: api.TEXT_ADVANCE_PROBE_CLUSTERS });
      const view = api.createReferenceAdvanceContext(ctx, planner);
      const text = 'AV To ffi Abcdefg';
      return ['normal 400 11px system-ui', 'normal 400 6.825px system-ui', 'normal 400 3.25px system-ui']
        .map((font) => {
          ctx.fontKerning = 'none';
          ctx.font = font;
          const native = ctx.measureText(text).width;
          const scale = Number(/([\d.]+)px/.exec(font)![1]) / 256;
          ctx.font = font.replace(/[\d.]+px/, '256px');
          const scaled = ctx.measureText(text).width * scale;
          view.font = font;
          return { font, planned: planner.plan(font), native, scaled, measured: view.measureText(text).width };
        });
    });
    for (const entry of result) {
      expect(entry.planned).toBeNull();
      expect(entry.measured).toBe(entry.native);
    }
  });

  it('moves whole-pixel colour bitmap advances to the reference em', async (test) => {
    const result = await page.evaluate(() => {
      const api = (window as unknown as { referenceAdvanceTest: Api }).referenceAdvanceTest;
      const ctx = document.createElement('canvas').getContext('2d') as CanvasRenderingContext2D;
      const font = 'normal 400 10px "Apple Color Emoji", serif';
      ctx.font = font;
      const digit = ctx.measureText('0').width;
      if (!(Number.isInteger(digit) && digit !== 10)) return null;
      const planner = api.createTextAdvancePlanner(ctx, { probeClusters: api.TEXT_ADVANCE_PROBE_CLUSTERS });
      const view = api.createReferenceAdvanceContext(ctx, planner);
      view.font = font;
      return { planned: planner.plan(font), digit, measured: view.measureText('0').width };
    });
    if (!result) test.skip();
    expect(result!.planned).not.toBeNull();
    expect(result!.measured).toBeCloseTo(10, 6);
  });
});
