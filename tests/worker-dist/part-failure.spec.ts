import { expect, test } from '@playwright/test';
import { brokenDocx, brokenPptx, brokenXlsx } from '../../packages/core/src/testing/part-failure-packages';

/**
 * #2248 — the placeholder a failed part shows is code no healthy file needs.
 * The published entries fetch it with the first failed part, in main mode; a
 * render worker is one self-contained file and carries its own copy. Either
 * way the placeholder, the `partFailures` records and the `errorPlaceholder`
 * calls are the same.
 */
const PLACEHOLDER_CHUNK = /\/dist\/(error-page|error-placeholder-text|parse-error-placeholder|sheet-error-overlay)-[^/]+\.js$/;
const PACKAGES = { docx: brokenDocx, xlsx: brokenXlsx, pptx: brokenPptx } as const;
const PLACEHOLDER_OF = { docx: 'error-page', xlsx: 'sheet-error-overlay', pptx: 'parse-error-placeholder' } as const;

interface PartFailureRun {
  builtin: string;
  custom: string;
  failures: Array<{ kind: string; message: string }>;
  calls: unknown[];
}

test('published entries fetch the placeholder code with the first failed part', async ({ page }) => {
  const placeholderChunks: string[] = [];
  page.context().on('request', (request) => {
    const pathname = new URL(request.url()).pathname;
    if (PLACEHOLDER_CHUNK.test(pathname)) placeholderChunks.push(pathname);
  });
  await page.goto('/tests/worker-dist/part-failure.html');
  await expect(page.locator('body')).toHaveAttribute('data-status', 'imported', { timeout: 60_000 });
  expect(placeholderChunks, 'importing the entries fetches no placeholder code').toEqual([]);

  const run = (format: keyof typeof PACKAGES, mode: 'main' | 'worker') => page.evaluate(
    ([name, renderMode, bytes]) => (window as typeof window & {
      runPartFailure(format: string, mode: string, bytes: number[]): Promise<PartFailureRun>;
    }).runPartFailure(name, renderMode, bytes),
    [format, mode, [...PACKAGES[format]()]] as const,
  );

  for (const format of ['docx', 'xlsx', 'pptx'] as const) {
    // Worker first: the worker file inlines the code, so the page fetches none.
    const worker = await run(format, 'worker');
    expect(placeholderChunks.filter((path) => path.includes(PLACEHOLDER_OF[format])), `${format} worker mode`).toEqual([]);
    const main = await run(format, 'main');
    expect(placeholderChunks.filter((path) => path.includes(PLACEHOLDER_OF[format])), `${format} main mode`).toHaveLength(1);
    expect(placeholderChunks.filter((path) => path.includes('error-placeholder-text')), 'shared text code is fetched once')
      .toHaveLength(1);

    expect(worker.builtin, `${format} built-in placeholder`).toBe(main.builtin);
    expect(worker.custom, `${format} custom placeholder`).toBe(main.custom);
    expect(main.custom).not.toBe(main.builtin);
    for (const result of [main, worker]) {
      const painted = result.failures.filter((failure) => failure.kind !== 'workbook');
      expect(painted).toHaveLength(1);
      expect(painted[0]!.message).not.toBe('');
      expect(result.calls).toEqual(painted);
    }
    expect(worker.failures).toEqual(main.failures);
    // A second failed load reuses the loaded code.
    const again = await run(format, 'main');
    expect(again).toEqual(main);
    expect(placeholderChunks.filter((path) => path.includes(PLACEHOLDER_OF[format]))).toHaveLength(1);
  }
});
