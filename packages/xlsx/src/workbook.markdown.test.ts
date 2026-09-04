import { describe, it, expect, vi } from 'vitest';
import { XlsxWorkbook } from './workbook';
import type { WorkerRequest } from './types';

/**
 * `XlsxWorkbook.toMarkdown()` routes through the persistent worker via the
 * `toMarkdown` message (xlsx uses the `type` discriminant, like docx) and
 * returns the `markdownRendered` string. The archive was opened at `parse` and
 * stays in the worker, so no file bytes cross back — only the projected
 * markdown string.
 *
 * The constructor opens a real Worker, so we build the instance off-prototype
 * and inject a fake `bridge` whose `request` resolves a `markdownRendered`
 * response. This isolates the request-shape + string-passthrough contract from
 * the actual WASM projection (exercised end-to-end by the parser's own tests).
 */
/** The subset of XlsxWorkbook this test exercises. Kept separate from the class
 *  type so the off-prototype build doesn't intersect its private fields. */
interface ToMarkdownProbe {
  toMarkdown(): Promise<string>;
}

describe('XlsxWorkbook.toMarkdown', () => {
  function makeWorkbook(requestImpl: (req: WorkerRequest) => unknown) {
    const request = vi.fn((build: (id: number) => WorkerRequest) =>
      Promise.resolve(requestImpl(build(1))),
    );
    const instance = Object.create(XlsxWorkbook.prototype) as Record<string, unknown>;
    instance.bridge = { request };
    const wb = instance as unknown as ToMarkdownProbe;
    return { wb, request };
  }

  it('posts a toMarkdown request and returns the rendered string', async () => {
    const { wb, request } = makeWorkbook((req) => {
      expect(req.type).toBe('toMarkdown');
      return { type: 'markdownRendered', id: 1, markdown: '## Sheet1\n\n| A | B |' };
    });

    const md = await wb.toMarkdown();
    expect(md).toBe('## Sheet1\n\n| A | B |');
    expect(request).toHaveBeenCalledTimes(1);
  });
});
