import { afterAll, beforeAll, describe, expect, it } from 'vitest';
import { chromium, type Browser, type Page } from '@playwright/test';
import { createServer, type ViteDevServer } from 'vite';
import { existsSync } from 'node:fs';
import { readFile } from 'node:fs/promises';
import init, { DocxArchive } from './wasm/docx_parser.js';
import { storeZip } from './conformance/generate.js';
import { measureContext } from './test-support/tab-fitting.test-support.js';
import { normalizeInternalDocumentModel } from './parser-model.js';
import { layoutDocument } from './document-layout.js';
import { createLayoutServices } from './layout-runtime.js';
import { textRunsForPage } from './text-run-projection.js';
import type { DocxTextRunInfo } from './types.js';

const encode = (text: string) => new TextEncoder().encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${text}`);

/** Lays out a one-page document whose body is `body` and returns its runs at
 * scale 1 (page CSS px = points). */
async function runsOf(body: string): Promise<DocxTextRunInfo[]> {
  const ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    + 'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"';
  const pkg = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const bytes = storeZip(new Map([
    ['[Content_Types].xml', encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
    ['_rels/.rels', encode(`<Relationships xmlns="${pkg}"><Relationship Id="rId1" `
      + 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
    ['word/document.xml', encode(`<w:document ${ns}><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>`
      + '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr></w:body></w:document>')],
  ]));
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
  const archive = new DocxArchive(bytes);
  let parsed;
  try {
    parsed = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse())));
  } finally { archive.free(); }
  const model = parsed.document;
  const layout = layoutDocument(model, createLayoutServices(model, { measureContext: measureContext(5) }), { currentDateMs: 0 });
  return textRunsForPage(layout, 0, { scale: 1 });
}

type Api = typeof import('./text-layer') & typeof import('./find-highlight-layer') & typeof import('./comment-margin')
  & typeof import('./selection-context');

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
    page = await browser.newPage({ viewport: { width: 1800, height: 1400 } });
    await page.route('**/overlay-clip.html', (route) => route.fulfill({
      contentType: 'text/html',
      body: '<!doctype html><body style="margin:0"><div id="layer" style="position:relative;width:800px;height:600px"></div></body>',
    }));
    await page.goto(`http://127.0.0.1:${port}/overlay-clip.html`);
    await page.addScriptTag({ type: 'module', content: `
      import * as text from '/src/text-layer.ts';
      import * as find from '/src/find-highlight-layer.ts';
      import * as margin from '/src/comment-margin.ts';
      import * as selection from '/src/selection-context.ts';
      window.overlayClipTest = { ...text, ...find, ...margin, ...selection };
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
      const origin = layer.getBoundingClientRect();
      return points.map(([x, y]) => {
        const hit = document.elementFromPoint(origin.left + x, origin.top + y);
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
    const runs = [link({ clips: [cell] })];
    // At 800px: hidden left of x=160, live right of it.
    expect(await hits(runs, 800, [[150, 110], [170, 110]])).toEqual(['', 'MMMMMMMMMM']);
    // At 400px everything halves: the edge is at x=80.
    expect(await hits(runs, 400, [[75, 55], [85, 55]])).toEqual(['', 'MMMMMMMMMM']);
  });

  it('keeps a VML text box run reachable where its authored text rectangle paints it', async () => {
    // A 100 x 60pt VML text box whose text rectangle is the whole shape, with
    // a 30pt left inset and the paragraph indented 20pt back into it: the run
    // paints between the inset and the shape edge.
    const box = '<w:p><w:r><w:pict><v:group style="width:120pt;height:80pt" coordsize="120,80">'
      + '<v:shapetype id="box" coordsize="1000,1000"/><v:shape type="#box" style="width:100;height:60" filled="f" stroked="f">'
      + '<v:path textboxrect="0,0,1000,1000"/><v:textbox inset="30pt,0,0,0"><w:txbxContent>'
      + '<w:p><w:pPr><w:ind w:left="-400"/></w:pPr><w:r><w:t>Link</w:t></w:r></w:p>'
      + '</w:txbxContent></v:textbox></v:shape></v:group></w:pict></w:r></w:p>';
    const runs = await runsOf(box);
    const run = runs.find((item) => item.text === 'Link');
    expect(run?.clips?.at(-1)?.rects).toBeDefined();
    const point: [number, number] = [run!.x + run!.w / 2, run!.y + run!.h / 2];
    expect(await hits(runs, 800, [point])).toEqual(['Link']);
  });

  it('keeps the visible part of a clipped link reachable after a uniform enlargement', async () => {
    // A 20px monospace link at (100, 100) clipped from x=170.
    const cell = { x: 170, y: 90, width: 300, height: 40 };
    const runs = [link({ text: 'MMMMMMMMMM', w: 120, clips: [cell] })];
    expect(await hits(runs, 800, [[190, 110]])).toEqual(['MMMMMMMMMM']);
    // At 1600px the clip starts at x=340 and the span has grown with the page.
    expect(await hits(runs, 1600, [[380, 220]])).toEqual(['MMMMMMMMMM']);
  });

  it('sizes spans by the page width inside a vertical-rl host', async () => {
    await page.evaluate(() => { document.body.style.writingMode = 'vertical-rl'; });
    try {
      const runs = [link({ text: 'MMMMMMMMMM', w: 120 })];
      // The 20px monospace link covers about x=100-220 as on a horizontal page.
      expect(await hits(runs, 800, [[150, 110], [300, 110]])).toEqual(['MMMMMMMMMM', '']);
    } finally {
      await page.evaluate(() => { document.body.style.writingMode = ''; });
    }
  });

  it('leaves text its cell clip hides wholly out of a drag selection', async () => {
    const result = await page.evaluate(() => {
      const api = (window as unknown as { overlayClipTest: Api }).overlayClipTest;
      const layer = document.getElementById('layer') as HTMLDivElement;
      layer.style.width = '800px';
      layer.style.height = '600px';
      const cell = { x: 160, y: 90, width: 300, height: 40 };
      const run = (text: string, x: number) => ({
        text, x, y: 100, w: 60, h: 20, fontSize: 20, font: '20px monospace', clips: [cell],
      });
      // 'Hidden' lies wholly left of the cell; the others are inside it.
      api.buildDocxTextLayer(layer, [run('Shown1', 170), run('Hidden', 50), run('Shown2', 300)], 800, 600, undefined, undefined, 0);
      const range = document.createRange();
      range.selectNodeContents(layer);
      const selection = window.getSelection()!;
      selection.removeAllRanges();
      selection.addRange(range);
      const context = api.readDocxTextSelectionContext(layer, selection);
      const native = selection.toString();
      selection.removeAllRanges();
      return { text: context?.text ?? '', runs: context?.runs.map((item) => item.runIndex) ?? [], native };
    });
    expect(result.text).not.toContain('Hidden');
    expect(result.runs).toEqual([0, 2]);
    expect(result.native).not.toContain('Hidden');
  });

  it('keeps a link out of reach under an empty clip union', async () => {
    const box = { x: 50, y: 50, width: 400, height: 200, rects: [] };
    expect(await hits([link({ clips: [box] })], 800, [[150, 110]])).toEqual(['']);
  });

  it('places a comment marker after the visible part of a clipped run', async () => {
    const left = await page.evaluate(() => {
      const api = (window as unknown as { overlayClipTest: Api }).overlayClipTest;
      const layer = document.getElementById('layer') as HTMLDivElement;
      layer.style.width = '800px';
      layer.style.height = '600px';
      const tint = document.createElement('div');
      tint.style.cssText = 'position:absolute;inset:0';
      layer.replaceChildren(tint);
      const source = { story: 'body' as const, storyInstance: 'body', path: [1] };
      // A 200px run in a 100px cell that starts at x=72: it is cut at x=172.
      api.buildDocxCommentMargin(tint, null, [{
        source, sourceRunIndex: 0, text: 'commented', x: 72, y: 100, w: 200, h: 20, fontSize: 20, font: '20px serif',
        clips: [{ x: 72, y: 90, width: 100, height: 40 }],
      }], {
        comments: [{ id: 'c', author: 'Ada', text: 'Note' }] as never,
        anchors: [{ commentId: 'c', source, startRunIndex: 0, endRunIndex: 1,
          reference: { source, runIndex: 1, affinity: 'preceding' } }],
      }, 800, 600, null, () => undefined, 1, 260, true);
      const marker = tint.querySelector('[data-ooxml-comment-marker]') as HTMLElement;
      return marker.getBoundingClientRect().left;
    });
    // The marker sits just after x=172, not after the hidden end at x=272.
    expect(left).toBeGreaterThan(172);
    expect(left).toBeLessThan(200);
  });

  it('keeps a link out of reach when a turned clip hides it', async () => {
    // A text box turned 90deg whose clip lies away from the run.
    const box = { x: 600, y: 300, width: 100, height: 50, transform: 'rotate(90deg)' };
    const runs = [link({ x: 100, y: 100, clips: [box] })];
    expect(await hits(runs, 800, [[150, 110], [250, 110]])).toEqual(['', '']);
    // Inside the turned clip, a run placed there is reachable.
    const inside = [link({ x: 590, y: 310, w: 30, text: 'M', transform: 'rotate(90deg)', clips: [box] })];
    expect(await hits(inside, 800, [[580, 320]])).toEqual(['M']);
    // A uniform resize (the page keeps its aspect ratio) keeps the live area.
    expect(await hits(inside, 400, [[290, 160]])).toEqual(['M']);
  });
});
