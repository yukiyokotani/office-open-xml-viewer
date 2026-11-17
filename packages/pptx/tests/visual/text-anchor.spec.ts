import { expect, test } from '@playwright/test';
import type * as Pptx from '../../src/index';

test('native main and render-worker snapshots resolve the same stored anchor after reflow', async ({ page }) => {
  test.setTimeout(120_000);
  await page.route('**/text-anchor-fixture', (route) => route.fulfill({ contentType: 'text/html', body: '<!doctype html><html><body></body></html>' }));
  await page.goto('/text-anchor-fixture');
  const result = await page.evaluate(async () => {
    const entry = '/src/index.ts';
    const engine = await import(entry) as typeof Pptx;
    const bytes = await (await fetch('/demo/sample-1.pptx')).arrayBuffer();
    const main = await engine.PptxPresentation.load(bytes.slice(0));
    const worker = await engine.PptxPresentation.load(bytes.slice(0), { mode: 'worker' });
    try {
      const original = await main.collectSlideRuns(0, 960);
      const index = original.findIndex((run) => run.source && run.text.length > 0);
      if (index < 0) throw new Error('Demo slide has no source-backed text');
      const created = engine.createPptxTextAnchor(original, [{ runIndex: index, start: 0, end: original[index].text.length }]);
      if (!created.isSuccess) throw new Error(created.reason);
      const anchor = JSON.parse(JSON.stringify(created.anchor));
      const mainRuns = await main.collectSlideRuns(0, 480);
      const workerRuns = await worker.collectSlideRuns(0, 480);
      return {
        main: engine.resolvePptxTextAnchor(anchor, mainRuns),
        worker: engine.resolvePptxTextAnchor(anchor, workerRuns),
        mainSources: mainRuns.map((run) => run.source),
        workerSources: workerRuns.map((run) => run.source),
      };
    } finally {
      main.destroy();
      worker.destroy();
    }
  });
  expect(result.main.isSuccess).toBe(true);
  expect(result.worker.isSuccess).toBe(true);
  expect(result.workerSources).toEqual(result.mainSources);
});

test('native selection anchors preserve source context and reject truncation', async ({ page }) => {
  await page.route('**/text-anchor-fixture', (route) => route.fulfill({ contentType: 'text/html', body: '<!doctype html><html><body></body></html>' }));
  await page.goto('/text-anchor-fixture');
  const result = await page.evaluate(async () => {
    const entry = '/src/index.ts';
    const engine = await import(entry) as typeof Pptx;
    const layer = document.createElement('div');
    document.body.append(layer);
    const context = { text: 'A𝄞éZ\nNext' };
    const runs = [
      { text: 'A𝄞', start: 0, end: 3 }, { text: 'éZ', start: 3, end: 6 }, { text: 'Next', start: 7, end: 11 },
    ].map(({ text, start, end }) => ({
      text, inShapeX: 0, inShapeY: start * 10, w: 30, h: 12, fontSize: 12, font: '12px serif',
      shapeX: 0, shapeY: 0, shapeW: 100, shapeH: 50, rotation: 0,
      source: { body: { slideIndex: 0, slidePart: 'ppt/slides/slide1.xml', shapeId: '2' }, context, start, end },
    }));
    engine.buildPptxTextLayer(layer, runs, 100, 50, undefined, 0);
    const spans = layer.querySelectorAll('span');
    const range = document.createRange();
    range.setStart(spans[0].firstChild!, 1);
    range.setEnd(spans[2].firstChild!, 2);
    const selection = window.getSelection()!;
    selection.removeAllRanges();
    selection.addRange(range);
    const created = engine.readPptxTextAnchor(layer, selection);
    const textLimited = engine.readPptxTextAnchor(layer, selection, { maxTextCharacters: 1 });
    const runLimited = engine.readPptxTextAnchor(layer, selection, { maxRunLocators: 1 });
    const restored = created.isSuccess ? engine.resolvePptxTextAnchor(JSON.parse(JSON.stringify(created.anchor)), structuredClone(runs)) : created;
    delete layer.dataset.slideIndex;
    const missingContext = engine.readPptxTextAnchor(layer, selection);
    layer.dataset.slideIndex = '0';
    spans[0].textContent = 'changed';
    const mutated = engine.readPptxTextAnchor(layer, selection);
    return { created, restored, textLimited, runLimited, missingContext, mutated };
  });
  expect(result.created).toMatchObject({ isSuccess: true, anchor: { start: 1, end: 9, exact: '𝄞éZ\nNe' } });
  expect(result.restored).toEqual({ isSuccess: true, slices: [
    { runIndex: 0, start: 1, end: 3 }, { runIndex: 1, start: 0, end: 3 }, { runIndex: 2, start: 0, end: 2 },
  ] });
  expect(result.textLimited).toEqual({ isSuccess: false, reason: 'incomplete' });
  expect(result.runLimited).toEqual({ isSuccess: false, reason: 'incomplete' });
  expect(result.missingContext).toEqual({ isSuccess: false, reason: 'incomplete' });
  expect(result.mutated.isSuccess).toBe(false);
});
