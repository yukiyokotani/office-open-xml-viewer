import { test, expect } from '@playwright/test';

/**
 * IX6 functional worker-equivalence (+ IX9 zoom integration): the
 * text-selection overlay, the findText match set, and the find-highlight boxes
 * must be identical in `mode: 'worker'` and `mode: 'main'`, both at the initial
 * scale AND after a `setScale(1.5)` zoom. The run geometry is collected
 * off-thread in worker mode and shipped back beside the bitmap, so:
 *  - the selection spans (text + rounded box) match main mode exactly;
 *  - findText returns the same match set (the pre-IX6 guard returned `[]`);
 *  - after setScale(1.5) the render width is exactly natural × 1.5 and the
 *    overlay/highlight coordinates follow the zoom (worker re-render round-trips
 *    fresh runs at the new width).
 * This is the functional twin of the pixel worker-equivalence spec.
 */

interface Snap {
  overlay: { text: string; left: number; top: number }[];
  highlights: { left: number; top: number; width: number }[];
  canvasW: number;
  scale: number;
  page: number;
}
interface ModeData {
  matches: { text: string; page: number }[];
  before: Snap;
  after: Snap;
  textHighlights: {
    matches: { text: string; page: number; group: string; term: string }[];
    active: { matchIndex: number; group: string } | null;
    boxes: { background: string; ringed: boolean; left: number; top: number; width: number }[];
    findNextIndex: number | undefined;
  };
}

test('worker selection + findText + zoom match main mode › demo/sample-1', async ({ page }) => {
  test.setTimeout(120_000);
  await page.goto('/tests/visual/worker-selection-fixture.html?docx=demo/sample-1&q=the');
  await page.waitForFunction(
    () => document.body.dataset.status === 'ready' || document.body.dataset.status === 'error',
    { timeout: 90_000 },
  );
  const status = await page.evaluate(() => document.body.dataset.status);
  if (status === 'error') {
    throw new Error(await page.evaluate(() => document.body.dataset.errorMessage ?? ''));
  }

  const raw = await page.evaluate(() => ({
    main: document.body.dataset.main ?? '{}',
    worker: document.body.dataset.worker ?? '{}',
  }));
  const main = JSON.parse(raw.main) as ModeData;
  const worker = JSON.parse(raw.worker) as ModeData;

  // --- findText: real matches, identical across modes (the pre-IX6 worker
  // guard returned [] + console.warn — this is the regression tripwire).
  expect(main.matches.length).toBeGreaterThan(0);
  expect(worker.matches).toEqual(main.matches);

  // --- initial scale: overlay + highlight parity (non-empty — findNext
  // navigated to the first match's page, so highlight boxes must exist).
  expect(main.before.overlay.length).toBeGreaterThan(0);
  expect(worker.before.overlay).toEqual(main.before.overlay);
  expect(main.before.highlights.length).toBeGreaterThan(0);
  // Every captured box is a real match highlight, not an empty placeholder.
  for (const h of main.before.highlights) expect(h.width).toBeGreaterThan(0);
  expect(worker.before.highlights).toEqual(main.before.highlights);
  expect(worker.before.page).toBe(main.before.page);

  // --- IX9 zoom in worker mode: scale latched, render width = natural × 1.5.
  expect(worker.after.scale).toBe(1.5);
  expect(main.after.scale).toBe(1.5);
  // naturalWidth = beforeW / beforeScale (getScale() before any zoom is the
  // effective opts.width / natural factor), so the post-zoom canvas must be
  // round(natural × 1.5) — allow ±1 for the two roundings involved.
  const naturalPx = worker.before.canvasW / worker.before.scale;
  expect(Math.abs(worker.after.canvasW - naturalPx * 1.5)).toBeLessThanOrEqual(1);
  expect(worker.after.canvasW).toBe(main.after.canvasW);

  // --- post-zoom parity: overlay + highlights identical across modes.
  expect(main.after.overlay.length).toBeGreaterThan(0);
  expect(worker.after.overlay).toEqual(main.after.overlay);
  expect(main.after.highlights.length).toBeGreaterThan(0);
  expect(worker.after.highlights).toEqual(main.after.highlights);

  // --- the worker overlay/highlight coordinates FOLLOW the zoom: every box
  // scaled by (1.5 / beforeScale) within a small rounding tolerance.
  const ratio = 1.5 / worker.before.scale;
  const follow = (before: number, after: number) =>
    Math.abs(after - before * ratio) <= 3;
  expect(worker.after.overlay.length).toBe(worker.before.overlay.length);
  for (let i = 0; i < worker.before.overlay.length; i++) {
    expect(follow(worker.before.overlay[i].left, worker.after.overlay[i].left)).toBe(true);
    expect(follow(worker.before.overlay[i].top, worker.after.overlay[i].top)).toBe(true);
  }
  expect(worker.after.highlights.length).toBe(worker.before.highlights.length);
  for (let i = 0; i < worker.before.highlights.length; i++) {
    expect(follow(worker.before.highlights[i].left, worker.after.highlights[i].left)).toBe(true);
    expect(follow(worker.before.highlights[i].top, worker.after.highlights[i].top)).toBe(true);
  }

  // --- text highlights (highlightText): the same set, cursor and boxes in both
  // modes, next to the find and without disturbing it.
  const highlights = main.textHighlights;
  expect(worker.textHighlights).toEqual(highlights);
  const any = highlights.matches.filter((match) => match.group === 'any');
  const word = highlights.matches.filter((match) => match.group === 'word');
  // A group without options matches what the find matched.
  expect(any.map(({ text, page }) => ({ text, page }))).toEqual(main.matches);
  // The whole-word group keeps some of those occurrences and drops the ones
  // inside longer words.
  expect(word.length).toBeGreaterThan(0);
  expect(word.length).toBeLessThan(any.length);
  expect(highlights.active?.matchIndex).toBe(0);
  const colors = new Set(highlights.boxes.map((box) => box.background));
  expect(colors.has('rgba(0, 0, 255, 0.3)') || colors.has('rgba(255, 0, 0, 0.3)')).toBe(true);
  // The active highlight keeps its group colour and gains a ring.
  const ringed = highlights.boxes.filter((box) => box.ringed);
  expect(ringed.length).toBeGreaterThan(0);
  for (const box of ringed) expect(box.background).toMatch(/^rgba\((0, 0, 255|255, 0, 0), 0\.3\)$/);
  for (const box of highlights.boxes) expect(box.width).toBeGreaterThan(0);
  expect(highlights.findNextIndex).toBe(main.matches.length > 1 ? 1 : 0);
});
