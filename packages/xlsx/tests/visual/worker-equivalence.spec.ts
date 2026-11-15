import { test, expect } from '@playwright/test';
import { PNG } from 'pngjs';
import pixelmatch from 'pixelmatch';
// @ts-ignore - shared plain-JS test fixture builder.
import { chartExXlsxBytes } from '../../../../tests/fixtures/chart-ex-packages.mjs';
// @ts-ignore - shared plain-JS test fixture builder.
import { frozenOverflowXlsxBytes } from '../../../../tests/fixtures/frozen-overflow-xlsx.mjs';
// @ts-ignore - shared plain-JS test fixture builder.
import { symbolFontXlsxBytes } from '../../../../tests/fixtures/symbol-font-xlsx.mjs';

// Worker mode must produce (near-)identical pixels to main mode: same
// renderer, same per-sheet parse, same fonts, different thread. The only
// expected drift is sub-pixel text rasterization between the main-thread
// <canvas> and the worker OffscreenCanvas. Per-sheet tolerances grant slack
// only to a sheet that needs it, so the (near-)bit-identical sheets keep
// enough sensitivity that even a single dropped cell fails the diff.
const SHEETS = [0, 1];
const MAX_DIFF_PCT = [0.5, 0.5];

for (const sheetIndex of SHEETS) {
  test(`worker mode matches main mode › demo/sample-1 sheet ${sheetIndex + 1}`, async ({ page }) => {
    await page.goto(`/tests/visual/worker-fixture.html?xlsx=demo/sample-1&sheet=${sheetIndex}`);
    // Two full loads (main + worker) plus worker spin-up per test — twice the
    // single-render budget visual.spec.ts uses.
    await page.waitForFunction(
      () => document.body.dataset.status === 'ready' || document.body.dataset.status === 'error',
      { timeout: 60_000 },
    );
    const status = await page.evaluate(() => document.body.dataset.status);
    if (status === 'error') {
      throw new Error(await page.evaluate(() => document.body.dataset.errorMessage ?? ''));
    }

    // The fixture also drives a worker-mode getWorksheet() round-trip, which
    // posts a `parseSheet` message the render worker must handle. If that arm
    // regresses the promise never settles, the fixture stalls before 'ready'
    // and the waitForFunction above fails loudly. Assert the probe resolved
    // with a real worksheet so a silently-empty result can't slip through.
    const worksheetOk = await page.evaluate(() => document.body.dataset.worksheetOk);
    expect(worksheetOk).toBe('true');

    const [mainUrl, workerUrl] = await page.evaluate(() => [
      (document.getElementById('main-canvas') as HTMLCanvasElement).toDataURL('image/png'),
      (document.getElementById('worker-canvas') as HTMLCanvasElement).toDataURL('image/png'),
    ]);
    const a = PNG.sync.read(Buffer.from(mainUrl.split(',')[1], 'base64'));
    const b = PNG.sync.read(Buffer.from(workerUrl.split(',')[1], 'base64'));
    // A zero-size canvas means a silently failed render; fail with a readable
    // assertion instead of the NaN the diff percentage would produce.
    expect(a.width).toBeGreaterThan(0);
    expect(a.height).toBeGreaterThan(0);
    expect(b.width).toBe(a.width);
    expect(b.height).toBe(a.height);

    const diff = pixelmatch(a.data, b.data, undefined, a.width, a.height, { threshold: 0.1 });
    const pct = (diff / (a.width * a.height)) * 100;
    console.log(`  sheet ${sheetIndex + 1}: worker-vs-main diff ${pct.toFixed(3)}%`);
    expect(pct).toBeLessThanOrEqual(MAX_DIFF_PCT[sheetIndex]);
  });
}

test('worker mode matches main mode for the sheet viewer CSV preview', async ({ page }) => {
  await page.goto('/tests/visual/delimited-text-worker-fixture.html');
  await page.waitForFunction(
    () => document.body.dataset.status === 'ready' || document.body.dataset.status === 'error',
    { timeout: 60_000 },
  );
  const state = await page.evaluate(() => ({
    status: document.body.dataset.status,
    error: document.body.dataset.errorMessage,
    mainName: document.body.dataset.mainName,
    workerName: document.body.dataset.workerName,
    values: document.body.dataset.values,
  }));
  if (state.status === 'error') throw new Error(state.error ?? 'CSV preview failed');
  expect(state.mainName).toBe('CSV preview');
  expect(state.workerName).toBe('CSV preview');
  expect(JSON.parse(state.values ?? '[]')).toEqual([
    ['text', '00123'],
    ['text', 'first line\nsecond, line'],
    ['text', '=1+1'],
    ['text', '9007199254740993'],
    ['text', 'plain'],
    ['text', '2026-09-04'],
  ]);

  const [mainUrl, workerUrl] = await page.evaluate(() => [
    (document.getElementById('main-canvas') as HTMLCanvasElement).toDataURL('image/png'),
    (document.getElementById('worker-canvas') as HTMLCanvasElement).toDataURL('image/png'),
  ]);
  const main = PNG.sync.read(Buffer.from(mainUrl.split(',')[1], 'base64'));
  const worker = PNG.sync.read(Buffer.from(workerUrl.split(',')[1], 'base64'));
  expect(main.width).toBeGreaterThan(0);
  expect(main.height).toBeGreaterThan(0);
  expect(worker.width).toBe(main.width);
  expect(worker.height).toBe(main.height);
  let ink = 0;
  for (let offset = 0; offset < main.data.length; offset += 4) {
    if (
      main.data[offset] < 250 ||
      main.data[offset + 1] < 250 ||
      main.data[offset + 2] < 250
    ) ink++;
  }
  expect(ink).toBeGreaterThan(100);

  const diff = pixelmatch(
    main.data,
    worker.data,
    undefined,
    main.width,
    main.height,
    { threshold: 0.1 },
  );
  expect((diff / (main.width * main.height)) * 100).toBeLessThanOrEqual(0.5);
});

// The optional ChartEx renderer crosses the worker boundary as a descriptor.
// Public demo files carry no ChartEx part, so serve a self-authored waterfall
// and require identical, coloured (not grayscale-placeholder) output.
test('worker mode matches main mode › ChartEx waterfall', async ({ page }) => {
  const bytes = chartExXlsxBytes();
  await page.route('**/generated/chart-ex.xlsx', (route) => route.fulfill({
    status: 200, contentType: 'application/octet-stream', body: Buffer.from(bytes),
  }));
  await page.goto('/tests/visual/worker-fixture.html?xlsx=generated/chart-ex&sheet=0');
  await page.waitForFunction(
    () => document.body.dataset.status === 'ready' || document.body.dataset.status === 'error',
    { timeout: 60_000 },
  );
  const status = await page.evaluate(() => document.body.dataset.status);
  if (status === 'error') {
    throw new Error(await page.evaluate(() => document.body.dataset.errorMessage ?? ''));
  }
  const [mainUrl, workerUrl] = await page.evaluate(() => [
    (document.getElementById('main-canvas') as HTMLCanvasElement).toDataURL('image/png'),
    (document.getElementById('worker-canvas') as HTMLCanvasElement).toDataURL('image/png'),
  ]);
  const main = PNG.sync.read(Buffer.from(mainUrl.split(',')[1], 'base64'));
  const worker = PNG.sync.read(Buffer.from(workerUrl.split(',')[1], 'base64'));
  expect(worker.width).toBe(main.width);
  expect(worker.height).toBe(main.height);
  const coloured = (png: PNG): number => {
    let count = 0;
    for (let offset = 0; offset < png.data.length; offset += 4) {
      const r = png.data[offset], g = png.data[offset + 1], b = png.data[offset + 2];
      if (Math.max(r, g, b) - Math.min(r, g, b) > 60) count++;
    }
    return count;
  };
  expect(coloured(main)).toBeGreaterThan(500);
  expect(coloured(worker)).toBeGreaterThan(500);
  const diff = pixelmatch(main.data, worker.data, undefined, main.width, main.height, { threshold: 0.1 });
  expect((diff / (main.width * main.height)) * 100).toBeLessThanOrEqual(0.1);
});

// Text overflow crosses the freeze line in both directions (#1973). Both
// threads run the same renderer, so the unscrolled and scrolled frames must
// match: the spill from a frozen column moves with the scrollable pane, and
// leftward spill into the frozen pane stays after its column scrolls away.
async function frozenOverflowFrames(
  page: import('@playwright/test').Page,
  query: string,
): Promise<{ main: PNG; worker: PNG }> {
  const bytes = frozenOverflowXlsxBytes();
  await page.route('**/generated/frozen-overflow.xlsx', (route) => route.fulfill({
    status: 200, contentType: 'application/octet-stream', body: Buffer.from(bytes),
  }));
  await page.goto(`/tests/visual/worker-fixture.html?xlsx=generated/frozen-overflow&sheet=0&${query}`);
  await page.waitForFunction(
    () => document.body.dataset.status === 'ready' || document.body.dataset.status === 'error',
    { timeout: 60_000 },
  );
  const status = await page.evaluate(() => document.body.dataset.status);
  if (status === 'error') {
    throw new Error(await page.evaluate(() => document.body.dataset.errorMessage ?? ''));
  }
  const [mainUrl, workerUrl] = await page.evaluate(() => [
    (document.getElementById('main-canvas') as HTMLCanvasElement).toDataURL('image/png'),
    (document.getElementById('worker-canvas') as HTMLCanvasElement).toDataURL('image/png'),
  ]);
  return {
    main: PNG.sync.read(Buffer.from(mainUrl.split(',')[1], 'base64')),
    worker: PNG.sync.read(Buffer.from(workerUrl.split(',')[1], 'base64')),
  };
}

for (const col of [4, 8]) {
  test(`worker mode matches main mode › frozen-column overflow, scrollable pane at column ${col}`, async ({ page }) => {
    const { main, worker } = await frozenOverflowFrames(
      page, `row=2&col=${col}&freezeRows=1&freezeCols=3`,
    );
    expect(worker.width).toBe(main.width);
    expect(worker.height).toBe(main.height);
    const diff = pixelmatch(main.data, worker.data, undefined, main.width, main.height, { threshold: 0.1 });
    expect((diff / (main.width * main.height)) * 100).toBeLessThanOrEqual(0.1);
  });
}

test('unscrolled frozen panes paint overflow exactly like the unfrozen sheet', async ({ page }) => {
  // Unscrolled, each pane's sheet coordinates coincide with the unfrozen
  // sheet's, so every overflow (rightward from B, leftward from E, both ways
  // from C) must be identical; only the freeze separator lines may differ.
  const frozen = await frozenOverflowFrames(page, 'row=2&col=4&freezeRows=1&freezeCols=3');
  // Same grid extent: 3 frozen + 12 scrollable columns, 1 frozen + 40 rows.
  const unfrozen = await frozenOverflowFrames(page, 'row=1&col=1&cols=15&rows=41');
  for (const [a, b] of [[frozen.main, unfrozen.main], [frozen.worker, unfrozen.worker]]) {
    expect(a.width).toBe(b.width);
    expect(a.height).toBe(b.height);
    const { width, height } = a;
    const differs = (x: number, y: number): boolean => {
      const o = (y * width + x) * 4;
      return a.data[o] !== b.data[o] || a.data[o + 1] !== b.data[o + 1] || a.data[o + 2] !== b.data[o + 2];
    };
    const colCounts = new Array<number>(width).fill(0);
    const rowCounts = new Array<number>(height).fill(0);
    for (let y = 0; y < height; y++) {
      for (let x = 0; x < width; x++) {
        if (differs(x, y)) { colCounts[x]++; rowCounts[y]++; }
      }
    }
    // A separator spans the drawn grid; text overflow differences do not.
    const near = (counts: number[], i: number, extent: number) =>
      [i - 1, i, i + 1].some((j) => (counts[j] ?? 0) > extent / 4);
    let stray = 0;
    for (let y = 0; y < height; y++) {
      for (let x = 0; x < width; x++) {
        if (differs(x, y) && !near(colCounts, x, height) && !near(rowCounts, y, width)) stray++;
      }
    }
    expect(stray).toBe(0);
  }
});

// Symbol-encoded faces are judged in each realm by its own oracle (main
// document, worker realm), each with route sentinels loaded: kept or
// converted codes, measured and wrapped identically in both modes.
test('worker mode matches main mode › symbol-font cells, runs and text box', async ({ page }) => {
  const bytes = symbolFontXlsxBytes();
  await page.route('**/generated/symbol-fonts.xlsx', (route) => route.fulfill({
    status: 200, contentType: 'application/octet-stream', body: Buffer.from(bytes),
  }));
  await page.goto('/tests/visual/worker-fixture.html?xlsx=generated/symbol-fonts&sheet=0');
  await page.waitForFunction(
    () => document.body.dataset.status === 'ready' || document.body.dataset.status === 'error',
    { timeout: 60_000 },
  );
  const status = await page.evaluate(() => document.body.dataset.status);
  if (status === 'error') {
    throw new Error(await page.evaluate(() => document.body.dataset.errorMessage ?? ''));
  }
  const [mainUrl, workerUrl] = await page.evaluate(() => [
    (document.getElementById('main-canvas') as HTMLCanvasElement).toDataURL('image/png'),
    (document.getElementById('worker-canvas') as HTMLCanvasElement).toDataURL('image/png'),
  ]);
  const main = PNG.sync.read(Buffer.from(mainUrl.split(',')[1], 'base64'));
  const worker = PNG.sync.read(Buffer.from(workerUrl.split(',')[1], 'base64'));
  expect(worker.width).toBe(main.width);
  expect(worker.height).toBe(main.height);
  const diff = pixelmatch(main.data, worker.data, undefined, main.width, main.height, { threshold: 0.1 });
  expect((diff / (main.width * main.height)) * 100).toBeLessThanOrEqual(0.1);
});
