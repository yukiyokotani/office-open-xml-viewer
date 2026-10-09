import { expect, test, type Page } from '@playwright/test';
import { fileURLToPath } from 'node:url';

const docxSample = fileURLToPath(
  new URL('../../packages/docx/public/demo/sample-1.docx', import.meta.url),
);
const xlsxSample = fileURLToPath(
  new URL('../../packages/xlsx/public/demo/sample-1.xlsx', import.meta.url),
);

test('Try Yours XLSX chrome follows the site theme without recoloring cells', async ({ page }) => {
  // Keep unrelated network font completion from changing the pixel oracle.
  await page.route('https://fonts.googleapis.com/**', (route) => route.abort());
  await page.route('https://fonts.gstatic.com/**', (route) => route.abort());
  await page.addInitScript(() => localStorage.setItem('ooxml-theme', 'dark'));
  await page.goto('/try/');
  const canvas = page.locator('#stage canvas').first();
  const tabs = page.locator('#stage .xlsx-tab-strip button');
  const readCanvas = () => canvas.evaluate(async (element: HTMLCanvasElement) => {
    // Read a copy: repeated getImageData on the production canvas can make
    // Chromium switch its painting backend and change text antialiasing.
    const copy = document.createElement('canvas');
    copy.width = element.width;
    copy.height = element.height;
    const context = copy.getContext('2d', { willReadFrequently: true }) as CanvasRenderingContext2D;
    context.drawImage(element, 0, 0);
    const dpr = devicePixelRatio;
    // Corner pixels are chrome. The lower-right crop is entirely in the cell
    // area, away from row/column headers and selection borders.
    return {
      corner: Array.from(context.getImageData(4 * dpr, 4 * dpr, 1, 1).data),
      column: Array.from(context.getImageData(200 * dpr, 4 * dpr, 1, 1).data),
      cells: Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',
        context.getImageData(60 * dpr, 40 * dpr, 500 * dpr, 200 * dpr).data))),
    };
  });
  await page.locator('#file').setInputFiles(xlsxSample);
  await expect(page.locator('#status')).toContainText('rendered in', { timeout: 60_000 });
  await expect.poll(async () => (await readCanvas()).corner).toEqual([8, 13, 19, 255]);
  await expect(tabs.first()).toHaveCSS('background-color', 'rgb(8, 13, 19)');
  await expect(tabs.first()).toHaveCSS('color', 'rgb(237, 243, 250)');
  await expect(page.locator('#stage button[aria-label="Zoom in"]')).toHaveCSS('color', 'rgb(170, 183, 199)');
  const dark = await readCanvas();

  await page.locator('[data-theme-toggle]').first().click();
  await expect.poll(async () => (await readCanvas()).corner).toEqual([248, 249, 250, 255]);
  await expect(tabs.first()).toHaveCSS('background-color', 'rgb(255, 255, 255)');
  await expect(tabs.first()).toHaveCSS('color', 'rgb(0, 0, 0)');
  expect((await readCanvas()).cells).toEqual(dark.cells);

  await page.locator('[data-theme-toggle]').first().click();
  await expect.poll(async () => (await readCanvas()).corner).toEqual(dark.corner);
  expect((await readCanvas()).cells).toEqual(dark.cells);
  await page.locator('#stage [data-xlsx-viewport-input]').click({ position: { x: 200, y: 4 } });
  await expect.poll(async () => (await readCanvas()).column).toEqual([52, 69, 88, 255]);
  await page.locator('[data-theme-toggle]').first().click();
  await expect.poll(async () => (await readCanvas()).column).toEqual([202, 221, 246, 255]);
  await page.locator('[data-theme-toggle]').first().click();
  await expect.poll(async () => (await readCanvas()).column).toEqual([52, 69, 88, 255]);
  await tabs.nth(1).click();
  await expect(tabs.nth(1)).toHaveCSS('background-color', 'rgb(8, 13, 19)');
  await expect.poll(async () => (await readCanvas()).corner).toEqual(dark.corner);

  await page.locator('#file').setInputFiles(xlsxSample);
  await expect(page.locator('#status')).toContainText('rendered in', { timeout: 60_000 });
  await expect.poll(async () => (await readCanvas()).corner).toEqual(dark.corner);
  await page.reload();
  await page.locator('#file').setInputFiles(xlsxSample);
  await expect(page.locator('#status')).toContainText('rendered in', { timeout: 60_000 });
  await expect.poll(async () => (await readCanvas()).corner).toEqual(dark.corner);
});

test('Try Yours XLSX scroll buttons open a native sheet list on right-click', async ({ page }) => {
  // Font-network completion is unrelated to footer navigation.
  await page.route('https://fonts.googleapis.com/**', (route) => route.abort());
  await page.route('https://fonts.gstatic.com/**', (route) => route.abort());
  const pageErrors: string[] = [];
  page.on('pageerror', (error) => pageErrors.push(error.message));
  const status = page.locator('#status');
  const strip = page.locator('#stage .xlsx-tab-strip');
  const prev = page.locator('#stage [data-xlsx-tab-nav="prev"]');
  const next = page.locator('#stage [data-xlsx-tab-nav="next"]');
  const list = page.locator('#stage [data-xlsx-sheet-list]');
  const items = list.getByRole('button');
  const current = list.locator('[aria-current="true"]');
  const tab = (name: string) => strip.getByRole('button', { name, exact: true });
  const loadSample = async () => {
    await page.locator('#file').setInputFiles(xlsxSample);
    await expect(status).toContainText('rendered in', { timeout: 60_000 });
  };
  // A disabled scroll button (strip start/end, or no overflow) ignores pointer
  // events, so `force` right-clicks its position as a user would.
  const rightClick = (button: typeof prev) => button.click({ button: 'right', force: true });

  await page.goto('/try/');
  await loadSample();
  const names = await strip.locator('button').evaluateAll((tabs) => tabs
    .filter((element) => getComputedStyle(element).display !== 'none')
    .map((element) => element.textContent ?? ''));
  expect(names.length).toBeGreaterThan(1);
  await expect(list).toHaveCount(1);
  await expect(list).toBeHidden();

  // Start of the strip: the previous button is disabled.
  await prev.scrollIntoViewIfNeeded();
  await expect(prev).toHaveCSS('pointer-events', 'none');
  await rightClick(prev);
  await expect(list).toBeVisible();
  await expect(items).toHaveText(names);
  await expect(current).toHaveCount(1);
  await page.screenshot({ path: test.info().outputPath('sheet-list.png') });
  const [box, footer] = await Promise.all([list.boundingBox(), prev.boundingBox()]);
  const viewport = page.viewportSize() as { width: number; height: number };
  if (!box || !footer) throw new Error('sheet list or footer has no layout box');
  expect(box.y).toBeGreaterThanOrEqual(0);
  expect(box.y + box.height).toBeLessThanOrEqual(viewport.height);
  expect(box.y + box.height <= footer.y + 1 || box.y >= footer.y + footer.height - 1).toBe(true);

  const choice = list.locator('button:not([aria-current])').last();
  const chosen = (await choice.textContent()) ?? '';
  await choice.click();
  await expect(list).toBeHidden();
  await expect(tab(chosen)).toHaveCSS('font-weight', '600');

  // Normal clicks keep scrolling the strip and do not open the list.
  await strip.evaluate((element: HTMLElement) => {
    element.style.flex = '0 0 40px';
    element.scrollLeft = 0;
    element.dispatchEvent(new Event('scroll'));
  });
  await expect(next).toHaveCSS('pointer-events', 'auto');
  await next.click();
  await expect.poll(() => strip.evaluate((element) => element.scrollLeft)).toBeGreaterThan(0);
  await expect(list).toBeHidden();

  // End of the strip: the next button is disabled.
  await strip.evaluate((element: HTMLElement) => {
    element.scrollLeft = element.scrollWidth;
    element.dispatchEvent(new Event('scroll'));
  });
  await expect(next).toHaveCSS('pointer-events', 'none');
  await rightClick(next);
  await expect(list).toBeVisible();
  await expect(current).toHaveText(chosen);
  await page.keyboard.press('Escape');
  await expect(list).toBeHidden();
  await expect(next).toBeFocused();

  // Keyboard flow. The platform context-menu key is not portable across
  // headless platforms, so dispatch the contextmenu event it produces.
  await next.evaluate((button: HTMLElement) => {
    button.focus();
    button.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true }));
  });
  await expect(current).toBeFocused();
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('Enter');
  await expect(list).toBeHidden();
  await expect(next).toBeFocused();
  await expect(tab(names[(names.indexOf(chosen) + 1) % names.length])).toHaveCSS('font-weight', '600');

  // Light dismissal.
  await rightClick(prev);
  await expect(list).toBeVisible();
  await status.click();
  await expect(list).toBeHidden();

  // Replacing the file while the list is open tears it down with the viewer.
  await rightClick(prev);
  await expect(list).toBeVisible();
  await loadSample();
  await expect(list).toBeHidden();
  await expect(list).toHaveCount(1);

  await page.reload();
  await loadSample();
  await expect(list).toHaveCount(1);
  await rightClick(prev);
  await expect(items).toHaveText(names);
  await page.keyboard.press('Escape');
  await expect(list).toBeHidden();
  expect(pageErrors).toEqual([]);
});

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
