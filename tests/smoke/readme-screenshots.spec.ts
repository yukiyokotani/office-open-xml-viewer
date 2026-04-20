import { expect, test } from '@playwright/test';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

// Generates PNG screenshots referenced from the root README.
// Run manually: `npx playwright test readme-screenshots --reporter=list`

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const OUT_DIR = path.resolve(__dirname, '../../docs/images');

async function waitForLoaded(page: import('@playwright/test').Page): Promise<void> {
  await page.waitForFunction(
    () => {
      for (const el of Array.from(document.querySelectorAll('div, span'))) {
        if (/^Loaded/.test((el.textContent ?? '').trim())) return true;
      }
      return false;
    },
    null,
    { timeout: 60_000 },
  );
  // let the final render settle
  await page.waitForTimeout(400);
}

test.describe('README screenshots', () => {
  test('pptx', async ({ page }) => {
    const res = await page.goto('/iframe.html?id=pptxviewer-examples--demo&viewMode=story');
    expect(res?.status()).toBeLessThan(400);
    await waitForLoaded(page);
    const canvas = page.locator('canvas').first();
    await canvas.screenshot({ path: `${OUT_DIR}/pptx.png` });
  });

  test('docx', async ({ page }) => {
    const res = await page.goto('/iframe.html?id=docxviewer-examples--demo&viewMode=story');
    expect(res?.status()).toBeLessThan(400);
    await waitForLoaded(page);
    const canvas = page.locator('canvas').first();
    await canvas.screenshot({ path: `${OUT_DIR}/docx.png` });
  });

  test('xlsx', async ({ page }) => {
    // xlsx viewer is a full-viewport tab bar + grid, so capture the whole story body.
    await page.setViewportSize({ width: 1200, height: 720 });
    const res = await page.goto('/iframe.html?id=xlsxviewer-examples--demo&viewMode=story');
    expect(res?.status()).toBeLessThan(400);
    // XlsxViewer's status text is "Loaded — N sheet(s)" or "Sheet: Name"
    await page.waitForFunction(
      () => {
        for (const el of Array.from(document.querySelectorAll('div'))) {
          const t = (el.textContent ?? '').trim();
          if (/^(Loaded|Sheet:)/.test(t)) return true;
        }
        return false;
      },
      null,
      { timeout: 60_000 },
    );
    await page.waitForTimeout(600);
    await page.screenshot({ path: `${OUT_DIR}/xlsx.png`, fullPage: false });
  });
});
