import { test, expect } from '@playwright/test';

interface PaintResult { mode: 'main' | 'worker'; variant: string; hash: string; png: string }
interface RecordResult { mode: 'main' | 'worker'; variant: 'records'; failures: unknown[]; calls: unknown[] }

// #1974 — a synthetic package with one malformed second slide. Main and worker mode
// paint the placeholder pixel-identically for the built-in text and for both
// option forms; the option changes the text, and the failure is listed and
// passed to the function on the main thread.
test('PPTX placeholder text and part failures match in main and worker mode', async ({ page }, testInfo) => {
  await page.goto('/tests/visual/part-failure-fixture.html');
  await page.waitForFunction(() => document.body.dataset.status === 'ready', { timeout: 60_000 });
  const error = await page.evaluate(() => (window as unknown as { __partFailureError?: string }).__partFailureError);
  expect(error).toBeUndefined();
  const results = await page.evaluate(
    () => (window as unknown as { __partFailureResults: (PaintResult | RecordResult)[] }).__partFailureResults,
  );
  const paint = (mode: string, variant: string) =>
    results.find((r): r is PaintResult => r.mode === mode && r.variant === variant)!;
  for (const variant of ['builtin', 'static', 'function']) {
    expect(paint('worker', variant).hash, variant).toBe(paint('main', variant).hash);
    await testInfo.attach(`${variant}.png`, {
      body: Buffer.from(paint('main', variant).png.split(',')[1]!, 'base64'),
      contentType: 'image/png',
    });
  }
  const hashes = new Set(['builtin', 'static', 'function'].map((variant) => paint('main', variant).hash));
  expect(hashes.size).toBe(3);
  for (const mode of ['main', 'worker']) {
    const records = results.find((r): r is RecordResult => r.mode === mode && r.variant === 'records')!;
    expect(records.failures).toMatchObject([{ kind: 'slide', slideIndex: 1, slideNumber: 2, message: expect.stringContaining('ppt/slides/slide2.xml') }]);
    expect(records.calls).toEqual(records.failures.filter((f) => (f as { kind: string }).kind !== 'workbook'));
  }
});
