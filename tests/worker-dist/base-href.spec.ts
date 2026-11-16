import { expect, test } from '@playwright/test';

/**
 * #2212 — a relative `wasmUrl` resolves against the page's base URL, which
 * honours `<base href>`, as `fetch()` would. The VS Code webview sets <base> to
 * the extension's dist directory while its location is the webview host; the
 * location-relative URL there named a file the host does not serve.
 */
test('a relative wasmUrl resolves against <base href> in main and worker mode', async ({ page }) => {
  const wasmRequests: string[] = [];
  page.context().on('request', (request) => {
    const pathname = new URL(request.url()).pathname;
    if (pathname.endsWith('.wasm')) wasmRequests.push(pathname);
  });
  await page.goto('/tests/worker-dist/base-href.html');
  await expect(page.locator('body')).toHaveAttribute('data-status', 'imported', { timeout: 60_000 });

  for (const format of ['docx', 'xlsx', 'pptx'] as const) {
    for (const mode of ['main', 'worker'] as const) {
      const effective = await page.evaluate(
        ([name, loadMode]) => (window as typeof window & {
          loadWithRelativeWasm(format: string, mode: string): Promise<string>;
        }).loadWithRelativeWasm(name, loadMode),
        [format, mode] as const,
      );
      expect(effective, `${format} ${mode} effective mode`).toBe(mode);
    }
  }
  expect([...new Set(wasmRequests)].sort()).toEqual([
    '/dist/docx_parser_bg.wasm',
    '/dist/pptx_parser_bg.wasm',
    '/dist/xlsx_parser_bg.wasm',
  ]);
});
