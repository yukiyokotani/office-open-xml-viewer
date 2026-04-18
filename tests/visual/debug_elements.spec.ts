import { test } from '@playwright/test';

for (const [pptx, slide] of [['sample-2', 15]]) {
  test(`debug ${pptx} slide ${+slide+1}`, async ({ page }) => {
    await page.goto(`http://localhost:5173/tests/visual/debug_elements.html?pptx=${pptx}&slide=${slide}`);
    await page.waitForFunction(() => document.body.dataset.status === 'ready', { timeout: 15000 });
    const text = await page.evaluate(() => document.getElementById('output')!.textContent ?? '');
    console.log(text);
  });
}
