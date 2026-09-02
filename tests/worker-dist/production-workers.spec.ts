import { expect, test } from '@playwright/test';

async function expectWorkerBitmaps(page: import('@playwright/test').Page, url: string) {
  await page.goto(url);
  await expect(page.locator('body')).toHaveAttribute('data-status', 'ready', { timeout: 60_000 });

  for (const id of ['docx', 'math', 'xlsx', 'pptx', 'pptx-text', 'xlsx-bordered']) {
    const ink = await page.locator(`#${id}`).evaluate((canvas: HTMLCanvasElement) => {
      const context = canvas.getContext('2d');
      if (!context) return 0;
      const pixels = context.getImageData(0, 0, canvas.width, canvas.height).data;
      let count = 0;
      for (let offset = 0; offset < pixels.length; offset += 4) {
        if (pixels[offset] < 250 || pixels[offset + 1] < 250 || pixels[offset + 2] < 250) count++;
      }
      return count;
    });
    expect(ink, `${id} worker bitmap should contain ink`).toBeGreaterThan(100);
  }

  const pptxTextRuns = await page.evaluate(() => (
    window as typeof window & { pptxTextRuns?: Array<Record<string, unknown>> }
  ).pptxTextRuns);
  expect(pptxTextRuns).toHaveLength(1);
  for (const field of ['inShapeX', 'inShapeY', 'w', 'h', 'fontSize']) {
    expect(
      Number.isFinite(pptxTextRuns?.[0]?.[field]),
      `PPTX worker text run ${field} should be finite`,
    ).toBe(true);
  }

  const fontProviderRequests = await page.evaluate(() => (
    window as typeof window & { fontProviderRequests?: string[][] }
  ).fontProviderRequests);
  expect(fontProviderRequests?.length).toBeGreaterThanOrEqual(3);
  expect(fontProviderRequests?.every((families) => families.length > 0)).toBe(true);

  for (const id of ['docx-chart-ex', 'xlsx-chart-ex', 'pptx-chart-ex']) {
    const coloredInk = await page.locator(`#${id}`).evaluate((canvas: HTMLCanvasElement) => {
      const context = canvas.getContext('2d');
      if (!context) return 0;
      const pixels = context.getImageData(0, 0, canvas.width, canvas.height).data;
      let count = 0;
      for (let offset = 0; offset < pixels.length; offset += 4) {
        const red = pixels[offset];
        const green = pixels[offset + 1];
        const blue = pixels[offset + 2];
        if (Math.max(red, green, blue) - Math.min(red, green, blue) > 30) count++;
      }
      return count;
    });
    expect(coloredInk, `${id} worker bitmap should execute the ChartEx painter`)
      .toBeGreaterThan(100);
  }
}

test('published dist starts all three render workers with optional renderers', async ({ page }) => {
  await expectWorkerBitmaps(page, '/');
});

test('Vite consumer bundle preserves all three render workers', async ({ page }) => {
  await expectWorkerBitmaps(page, '/consumer/index.html');
});
