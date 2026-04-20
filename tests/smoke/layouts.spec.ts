import { expect, test } from '@playwright/test';

// Expected non-zero page/slide counts per sample used by the stories.
const EXPECTED = {
  pptx: 9,   // packages/pptx/public/demo/sample-1.pptx
  docx: 6,   // packages/docx/public/demo/sample-1.docx (see docx visual.spec.ts)
};

type StoryId =
  | 'pptxviewer-examples--scroll-view'
  | 'pptxviewer-examples--thumbnail-grid'
  | 'pptxviewer-examples--master-detail'
  | 'docxviewer-examples--scroll-view'
  | 'docxviewer-examples--thumbnail-grid'
  | 'docxviewer-examples--master-detail';

async function canvasHasInk(page: import('@playwright/test').Page, index = 0): Promise<boolean> {
  return page.evaluate((i) => {
    const canvases = Array.from(document.querySelectorAll('canvas')) as HTMLCanvasElement[];
    const c = canvases[i];
    if (!c) return false;
    const ctx = c.getContext('2d');
    if (!ctx) return false;
    const w = c.width, h = c.height;
    if (w === 0 || h === 0) return false;
    // Scan a 20x20 grid; count pixels that are neither transparent nor pure white.
    let inked = 0;
    for (let gy = 0; gy < 20; gy++) {
      for (let gx = 0; gx < 20; gx++) {
        const x = Math.floor(((gx + 0.5) / 20) * w);
        const y = Math.floor(((gy + 0.5) / 20) * h);
        const { data } = ctx.getImageData(x, y, 1, 1);
        const [r, g, b, a] = [data[0], data[1], data[2], data[3]];
        const notBlank = a > 0 && !(r >= 250 && g >= 250 && b >= 250);
        if (notBlank) inked++;
        if (inked >= 3) return true;
      }
    }
    return false;
  }, index);
}

async function waitForLoaded(page: import('@playwright/test').Page, text: RegExp): Promise<void> {
  // The Layouts stories write "Loaded N slides" / "Loaded N pages" to a status div.
  await page.waitForFunction(
    (re) => {
      const matcher = new RegExp(re);
      for (const el of Array.from(document.querySelectorAll('div'))) {
        if (matcher.test(el.textContent ?? '')) return true;
      }
      return false;
    },
    text.source,
    { timeout: 60_000 },
  );
}

async function openStory(page: import('@playwright/test').Page, id: StoryId): Promise<void> {
  const res = await page.goto(`/iframe.html?id=${id}&viewMode=story`);
  expect(res?.status(), `HTTP status for ${id}`).toBeLessThan(400);
}

test.describe('Layouts smoke — pptx', () => {
  test('ScrollView renders every slide', async ({ page }) => {
    await openStory(page, 'pptxviewer-examples--scroll-view');
    await waitForLoaded(page, new RegExp(`Loaded ${EXPECTED.pptx} slides`));
    const count = await page.locator('canvas').count();
    expect(count).toBe(EXPECTED.pptx);
    expect(await canvasHasInk(page, 0)).toBe(true);
    expect(await canvasHasInk(page, Math.floor(EXPECTED.pptx / 2))).toBe(true);
    expect(await canvasHasInk(page, EXPECTED.pptx - 1)).toBe(true);
  });

  test('ThumbnailGrid renders every slide', async ({ page }) => {
    await openStory(page, 'pptxviewer-examples--thumbnail-grid');
    await waitForLoaded(page, new RegExp(`Loaded ${EXPECTED.pptx} slides`));
    const count = await page.locator('canvas').count();
    expect(count).toBe(EXPECTED.pptx);
    expect(await canvasHasInk(page, 0)).toBe(true);
    expect(await canvasHasInk(page, EXPECTED.pptx - 1)).toBe(true);
  });

  test('MasterDetail renders thumbs + large preview and switches on click', async ({ page }) => {
    await openStory(page, 'pptxviewer-examples--master-detail');
    await waitForLoaded(page, new RegExp(`Loaded ${EXPECTED.pptx} slides`));
    const count = await page.locator('canvas').count();
    // thumbs + 1 detail
    expect(count).toBe(EXPECTED.pptx + 1);
    // detail canvas is the first one we appended (layout is detail after thumbs column)
    // — regardless of DOM order, all canvases must be inked
    expect(await canvasHasInk(page, 0)).toBe(true);
    expect(await canvasHasInk(page, count - 1)).toBe(true);

    // Click last thumbnail cell and ensure the detail canvas is still inked.
    const cells = page.locator('div[style*="cursor: pointer"]');
    await cells.nth(EXPECTED.pptx - 1).click();
    await page.waitForTimeout(500);
    // The detail canvas is the first canvas in DOM (layout appended detailCol last → but detailCanvas is inside detailCol, thumbs in thumbCol appended first).
    // Regardless: ensure every canvas still has content after the click.
    for (let i = 0; i < count; i++) {
      expect(await canvasHasInk(page, i), `canvas ${i} blank after click`).toBe(true);
    }
  });
});

// docx demo/sample-1 ends with a mostly-blank trailing page, so we require a
// majority (not every) canvas to contain ink. This still catches broken renders.
async function countInkedCanvases(page: import('@playwright/test').Page, total: number): Promise<number> {
  let n = 0;
  for (let i = 0; i < total; i++) {
    if (await canvasHasInk(page, i)) n++;
  }
  return n;
}

test.describe('Layouts smoke — docx', () => {
  test('ScrollView renders every page', async ({ page }) => {
    await openStory(page, 'docxviewer-examples--scroll-view');
    await waitForLoaded(page, new RegExp(`Loaded ${EXPECTED.docx} pages`));
    const count = await page.locator('canvas').count();
    expect(count).toBe(EXPECTED.docx);
    // first page must have ink; majority of pages must render non-blank.
    expect(await canvasHasInk(page, 0)).toBe(true);
    expect(await countInkedCanvases(page, count)).toBeGreaterThanOrEqual(count - 1);
  });

  test('ThumbnailGrid renders every page', async ({ page }) => {
    await openStory(page, 'docxviewer-examples--thumbnail-grid');
    await waitForLoaded(page, new RegExp(`Loaded ${EXPECTED.docx} pages`));
    const count = await page.locator('canvas').count();
    expect(count).toBe(EXPECTED.docx);
    expect(await canvasHasInk(page, 0)).toBe(true);
    expect(await countInkedCanvases(page, count)).toBeGreaterThanOrEqual(count - 1);
  });

  test('MasterDetail renders thumbs + large preview', async ({ page }) => {
    await openStory(page, 'docxviewer-examples--master-detail');
    await waitForLoaded(page, new RegExp(`Loaded ${EXPECTED.docx} pages`));
    const count = await page.locator('canvas').count();
    // N thumbs + 1 detail = N+1 canvases (but trailing page may be blank)
    expect(count).toBe(EXPECTED.docx + 1);
    expect(await canvasHasInk(page, 0)).toBe(true);
    expect(await countInkedCanvases(page, count)).toBeGreaterThanOrEqual(count - 1);
  });
});
