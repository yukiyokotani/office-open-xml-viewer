import { test, expect } from '@playwright/test';

// Main and worker mode both parse in a worker and rebuild the typed error on
// the main thread, so this is the only boundary where a regression could turn
// `OoxmlError('not-ooxml')` back into a plain Error or a placeholder document.
test('main and worker mode reject non-OOXML input with OoxmlError not-ooxml', async ({ page }) => {
  await page.goto('/tests/visual/non-ooxml-fixture.html');
  await page.waitForFunction(() => document.body.dataset.status === 'ready', { timeout: 60_000 });
  const results = await page.evaluate(
    () => (window as unknown as { __nonOoxmlResults: unknown[] }).__nonOoxmlResults,
  );
  expect(results).toHaveLength(8);
  for (const result of results) {
    expect(result).toMatchObject({ resolved: false, ooxmlError: true, code: 'not-ooxml' });
  }
});
