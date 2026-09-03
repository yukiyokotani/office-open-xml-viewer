import { expect, test, type Page } from '@playwright/test';
import { fileURLToPath } from 'node:url';

const docxSample = fileURLToPath(
  new URL('../../packages/docx/public/demo/sample-1.docx', import.meta.url),
);

const dispatchPersistedPagehide = (page: Page) => page.evaluate(() => {
  window.dispatchEvent(new PageTransitionEvent('pagehide', { persisted: true }));
});

for (const format of ['docx', 'xlsx', 'pptx'] as const) {
  test(`${format.toUpperCase()} live and comment demos initialize`, async ({ page }) => {
    const pageErrors: string[] = [];
    page.on('pageerror', (error) => pageErrors.push(error.message));

    await page.goto(`/${format}/?all`);

    await expect(page.locator('[data-built-in-comment-status]')).toBeHidden({ timeout: 60_000 });
    await expect(page.locator('canvas').first()).toBeVisible();
    await expect(page.locator('body')).not.toContainText(/not a constructor|Failed:/i);
    expect(pageErrors).toEqual([]);
  });
}

test('DOCX comment demo survives browser back', async ({ page }) => {
  await page.goto('/docx/?all');
  const status = page.locator('[data-built-in-comment-status]');
  const viewer = page.locator('[data-built-in-comment-viewer]');
  await expect(status).toBeHidden({ timeout: 60_000 });
  await expect(viewer.locator('canvas').first()).toBeVisible();

  // Headless Chrome does not reliably retain localhost pages in BFCache, so
  // exercise the persisted pagehide branch explicitly before browser Back.
  await dispatchPersistedPagehide(page);
  await expect(viewer.locator('canvas').first()).toBeVisible();

  await page.goto('/');
  await page.goBack();

  await expect(page).toHaveURL(/\/docx\/?\?all$/);
  await expect(status).toBeHidden({ timeout: 60_000 });
  await expect(viewer.locator('canvas').first()).toBeVisible();
});

test('other live viewer screens survive persisted pagehide', async ({ page }) => {
  await page.goto('/review-ui/');
  await expect(page.locator('[data-built-in-comment-status]')).toBeHidden({ timeout: 60_000 });
  await expect(page.locator('[data-comment-list-loading]')).toBeHidden({ timeout: 60_000 });
  const builtInCanvas = page.locator('[data-built-in-comment-viewer] canvas').first();
  const listCanvas = page.locator('[data-comment-list-viewer] canvas').first();
  const listItem = page.locator('[data-comment-list-items] button').first();
  await expect(builtInCanvas).toBeVisible();
  await expect(listCanvas).toBeVisible();
  await expect(listItem).toBeVisible();
  await dispatchPersistedPagehide(page);
  await expect(builtInCanvas).toBeVisible();
  await expect(listCanvas).toBeVisible();
  await expect(listItem).toBeVisible();

  await page.goto('/selection-context/');
  const selectionCanvas = page.locator('[data-selection-context-demo] canvas').first();
  await expect(selectionCanvas).toBeVisible({ timeout: 60_000 });
  await dispatchPersistedPagehide(page);
  await expect(selectionCanvas).toBeVisible();

  await page.goto('/try/');
  await page.locator('#file').setInputFiles(docxSample);
  const tryCanvas = page.locator('#stage canvas').first();
  await expect(tryCanvas).toBeVisible({ timeout: 60_000 });
  await dispatchPersistedPagehide(page);
  await expect(tryCanvas).toBeVisible();
});

for (const format of ['csv', 'tsv'] as const) {
  test(`Try Yours opens a selected ${format.toUpperCase()} file in the sheet surface`, async ({
    page,
  }) => {
    const separator = format === 'csv' ? ',' : '\t';
    await page.goto('/try/');
    await page.locator('#file').setInputFiles({
      name: `table.${format}`,
      mimeType: format === 'csv' ? 'text/csv' : 'text/tab-separated-values',
      buffer: Buffer.from(`Code${separator}Value\n001${separator}alpha`),
    });

    await expect(page.locator('#stage canvas').first()).toBeVisible({ timeout: 60_000 });
    await expect(page.locator('#status')).toContainText('rendered in');
    await expect(page.locator('#wasm-badge')).toBeHidden();
    await expect(page.locator('#stage .xlsx-tab-strip')).toHaveCount(0);
  });
}

test('PPTX single-comment margin has no trailing scroll range', async ({ page }) => {
  await page.setViewportSize({ width: 1280, height: 720 });
  await page.goto('/pptx/?all');
  await expect(page.locator('[data-built-in-comment-status]')).toBeHidden({ timeout: 60_000 });

  const margin = page.locator('[data-ooxml-comment-ui="margin"]')
    .filter({ has: page.locator('.ooxml-comment-card') })
    .first();
  await expect(margin.locator('.ooxml-comment-card')).toHaveCount(1);

  const before = await margin.evaluate((element) => ({
    clientHeight: element.clientHeight,
    scrollHeight: element.scrollHeight,
    scrollTop: element.scrollTop,
  }));
  expect(before.scrollHeight).toBe(before.clientHeight);

  await margin.locator('.ooxml-comment-card').hover();
  await page.mouse.wheel(0, 100);
  await expect.poll(() => margin.evaluate((element) => element.scrollTop)).toBe(0);
});
