import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { chromium, type Browser, type Page } from '@playwright/test';
import { createServer, type ViteDevServer } from 'vite';
import { existsSync } from 'node:fs';
import type { DocxTextRunInfo } from './types.js';

type Api = typeof import('./text-layer') & typeof import('./find-highlight-layer');

/** WORD_CELL_CONTENT_CLIP in Chromium: the overlay clip containers clip hit
 * testing and highlights after a responsive resize and inside turned clips. */
describe.skipIf(!existsSync(chromium.executablePath()))('DOCX overlay clip containers', () => {
  let server: ViteDevServer;
  let browser: Browser;
  let page: Page;

  beforeAll(async () => {
    server = await createServer({
      configFile: new URL('../vite.config.ts', import.meta.url).pathname,
      server: { host: '127.0.0.1', port: 0 },
      logLevel: 'error',
    });
    await server.listen();
    const address = server.httpServer?.address();
    const port = typeof address === 'object' && address ? address.port : 0;
    browser = await chromium.launch();
    page = await browser.newPage({ viewport: { width: 1000, height: 800 } });
    await page.route('**/overlay-clip.html', (route) => route.fulfill({
      contentType: 'text/html',
      body: '<!doctype html><body style="margin:0"><div id="layer" style="position:relative;width:800px;height:600px"></div></body>',
    }));
    await page.goto(`http://127.0.0.1:${port}/overlay-clip.html`);
    await page.addScriptTag({ type: 'module', content: `
      import * as text from '/src/text-layer.ts';
      import * as find from '/src/find-highlight-layer.ts';
      window.overlayClipTest = { ...text, ...find };
    ` });
    await page.waitForFunction(() => 'overlayClipTest' in window);
  }, 60_000);
  afterAll(async () => { await browser?.close(); await server?.close(); });

  /** Builds the text layer for `runs` on an 800 x 600 page and returns, for
   * each probe point (in page px at the given layer width), the text of the
   * span hit there ('' for none). */
  const hits = (runs: DocxTextRunInfo[], width: number, points: Array<[number, number]>) => page.evaluate(
    ({ runs, width, points }) => {
      const api = (window as unknown as { overlayClipTest: Api }).overlayClipTest;
      const layer = document.getElementById('layer') as HTMLDivElement;
      layer.style.width = '800px';
      layer.style.height = '600px';
      api.buildDocxTextLayer(layer, runs, 800, 600, () => undefined);
      layer.style.width = `${width}px`;
      layer.style.height = `${width * 0.75}px`;
      return points.map(([x, y]) => {
        const hit = document.elementFromPoint(x, y);
        return hit?.tagName === 'SPAN' ? hit.textContent ?? '' : '';
      });
    },
    { runs, width, points },
  );

  const link = (partial: Partial<DocxTextRunInfo>): DocxTextRunInfo => ({
    text: 'MMMMMMMMMM', x: 100, y: 100, w: 200, h: 20, fontSize: 20, font: '20px monospace',
    hyperlink: { kind: 'external', url: 'https://example.com' } as DocxTextRunInfo['hyperlink'],
    ...partial,
  });

  it('clips a partly clipped link at its cell edge before and after a responsive resize', async () => {
    const cell = { x: 160, y: 90, width: 300, height: 40 };
    const runs = [link({ clips: [cell], clipBounds: { x: 160, y: 100, width: 140, height: 20 } })];
    // At 800px: hidden left of x=160, live right of it.
    expect(await hits(runs, 800, [[150, 110], [170, 110]])).toEqual(['', 'MMMMMMMMMM']);
    // At 400px everything halves: the edge is at x=80.
    expect(await hits(runs, 400, [[75, 55], [85, 55]])).toEqual(['', 'MMMMMMMMMM']);
  });

  it('keeps a link out of reach when a turned clip hides it', async () => {
    // A text box turned 90deg whose clip lies away from the run.
    const box = { x: 600, y: 300, width: 100, height: 50, transform: 'rotate(90deg)' };
    const runs = [link({ x: 100, y: 100, clips: [box] })];
    expect(await hits(runs, 800, [[150, 110], [250, 110]])).toEqual(['', '']);
    // Inside the turned clip, a run placed there is reachable.
    const inside = [link({ x: 590, y: 310, w: 30, text: 'M', transform: 'rotate(90deg)', clips: [box] })];
    expect(await hits(inside, 800, [[580, 320]])).toEqual(['M']);
  });
});
