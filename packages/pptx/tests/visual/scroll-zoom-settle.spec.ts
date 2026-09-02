import { expect, test } from '@playwright/test';

test('PptxScrollViewer keeps its scroll anchor when a worker zoom preview settles', async ({ page }) => {
  test.setTimeout(120_000);
  await page.goto('/tests/visual/scroll-zoom-settle-fixture.html');
  await page.waitForFunction(
    () => document.body.dataset.status === 'ready' || document.body.dataset.status === 'error',
    undefined,
    { timeout: 90_000 },
  );
  if (await page.evaluate(() => document.body.dataset.status) === 'error') {
    throw new Error(await page.evaluate(() => document.body.dataset.errorMessage ?? ''));
  }
  const result = JSON.parse(await page.evaluate(() => document.body.dataset.result ?? '{}')) as {
    preview: { left: number; top: number; width: number; height: number };
    settled: { left: number; top: number; width: number; height: number };
  };

  // The preview and crisp render represent the same scaled document geometry.
  // A text overlay based on the already-expanded wrapper used to make the
  // preview extent larger; clearing its transform then shrank the range and the
  // browser clamped both axes toward the top-left.
  expect(Math.abs(result.settled.left - result.preview.left)).toBeLessThanOrEqual(1);
  expect(Math.abs(result.settled.top - result.preview.top)).toBeLessThanOrEqual(1);
  expect(Math.abs(result.settled.width - result.preview.width)).toBeLessThanOrEqual(1);
  expect(Math.abs(result.settled.height - result.preview.height)).toBeLessThanOrEqual(1);
});

test('PptxScrollViewer keeps its scroll anchor when the Try Yours configuration settles at maximum zoom', async ({ page }) => {
  test.setTimeout(120_000);
  await page.goto('/tests/visual/scroll-zoom-settle-fixture.html?maximum');
  await page.waitForFunction(
    () => document.body.dataset.status === 'ready' || document.body.dataset.status === 'error',
    undefined,
    { timeout: 90_000 },
  );
  if (await page.evaluate(() => document.body.dataset.status) === 'error') {
    throw new Error(await page.evaluate(() => document.body.dataset.errorMessage ?? ''));
  }
  const result = JSON.parse(await page.evaluate(() => document.body.dataset.result ?? '{}')) as {
    preview: { left: number; top: number; width: number; height: number };
    settled: { left: number; top: number; width: number; height: number };
  };

  expect(Math.abs(result.settled.left - result.preview.left)).toBeLessThanOrEqual(1);
  expect(Math.abs(result.settled.top - result.preview.top)).toBeLessThanOrEqual(1);
  expect(Math.abs(result.settled.width - result.preview.width)).toBeLessThanOrEqual(1);
  expect(Math.abs(result.settled.height - result.preview.height)).toBeLessThanOrEqual(1);
});

test.describe('Retina Try Yours maximum zoom', () => {
  test.use({ deviceScaleFactor: 2 });

  test('keeps the scroll anchor after the high-resolution settle', async ({ page }) => {
    test.setTimeout(120_000);
    await page.goto('/tests/visual/scroll-zoom-settle-fixture.html?maximum');
    await page.waitForFunction(
      () => document.body.dataset.status === 'ready' || document.body.dataset.status === 'error',
      undefined,
      { timeout: 90_000 },
    );
    if (await page.evaluate(() => document.body.dataset.status) === 'error') {
      throw new Error(await page.evaluate(() => document.body.dataset.errorMessage ?? ''));
    }
    const result = JSON.parse(await page.evaluate(() => document.body.dataset.result ?? '{}')) as {
      preview: { left: number; top: number; width: number; height: number };
      settled: {
        left: number; top: number; width: number; height: number;
        slideWidth: number; slideHeight: number; canvasWidth: number; canvasHeight: number;
        canvasBufferWidth: number; canvasBufferHeight: number; scale: number;
      };
    };
    expect(Math.abs(result.settled.left - result.preview.left)).toBeLessThanOrEqual(1);
    expect(Math.abs(result.settled.top - result.preview.top)).toBeLessThanOrEqual(1);
    expect(Math.abs(result.settled.width - result.preview.width)).toBeLessThanOrEqual(1);
    expect(Math.abs(result.settled.height - result.preview.height)).toBeLessThanOrEqual(1);
    expect(Math.abs(result.settled.canvasWidth - result.settled.slideWidth)).toBeLessThanOrEqual(1);
    expect(Math.abs(result.settled.canvasHeight - result.settled.slideHeight)).toBeLessThanOrEqual(1);
    expect(result.settled.scale).toBe(4);
    expect(result.settled.canvasBufferWidth).toBeLessThan(result.settled.canvasWidth * 2);
    expect(result.settled.canvasBufferHeight).toBeLessThan(result.settled.canvasHeight * 2);
  });
});
