import { describe, it, expect, vi } from 'vitest';
import { DocxDocument } from './document';
import type { WorkerRequest } from './types';

/**
 * `DocxDocument.toMarkdown()` routes through the persistent worker via the
 * `toMarkdown` message (docx uses the `type` discriminant) and returns the
 * `markdownRendered` string. The archive was opened at `parse` and stays in the
 * worker, so no file bytes cross back — only the projected markdown string.
 *
 * The constructor opens a real Worker, so we build the instance off-prototype
 * and inject a fake `_bridge` whose `request` resolves a `markdownRendered`
 * response. This isolates the request-shape + string-passthrough contract from
 * the actual WASM projection (exercised end-to-end by the parser's own tests).
 */
/** The subset of DocxDocument this test exercises. Kept separate from the class
 *  type so the off-prototype build doesn't intersect its private fields. */
interface ToMarkdownProbe {
  toMarkdown(): Promise<string>;
}

describe('DocxDocument.toMarkdown', () => {
  function makeDocument(requestImpl: (req: WorkerRequest) => unknown) {
    const request = vi.fn((build: (id: number) => WorkerRequest) =>
      Promise.resolve(requestImpl(build(1))),
    );
    const instance = Object.create(DocxDocument.prototype) as Record<string, unknown>;
    instance._bridge = { request };
    const doc = instance as unknown as ToMarkdownProbe;
    return { doc, request };
  }

  it('posts a toMarkdown request and returns the rendered string', async () => {
    const { doc, request } = makeDocument((req) => {
      expect(req.type).toBe('toMarkdown');
      return { type: 'markdownRendered', id: 1, markdown: '# Heading\n\nBody.' };
    });

    const md = await doc.toMarkdown();
    expect(md).toBe('# Heading\n\nBody.');
    expect(request).toHaveBeenCalledTimes(1);
  });
});
