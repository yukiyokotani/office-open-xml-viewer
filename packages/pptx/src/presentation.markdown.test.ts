import { describe, it, expect, vi } from 'vitest';
import { PptxPresentation } from './presentation';
import type { PptxWorkerRequest as WorkerRequest } from './worker-protocol';

/**
 * `PptxPresentation.toMarkdown()` routes through the persistent worker via the
 * `toMarkdown` message (pptx uses the `kind` discriminant) and returns the
 * `markdownRendered` string. The archive was opened at `parse` and stays in the
 * worker, so no file bytes cross back — only the projected markdown string.
 *
 * The constructor opens a real Worker, so we build the instance off-prototype
 * and inject a fake `_bridge` whose `request` resolves a `markdownRendered`
 * response. This isolates the request-shape + string-passthrough contract from
 * the actual WASM projection (exercised end-to-end by the parser's own tests).
 */
/** The subset of PptxPresentation this test exercises. Kept separate from the
 *  class type so the off-prototype build doesn't intersect its private fields. */
interface ToMarkdownProbe {
  toMarkdown(): Promise<string>;
}

describe('PptxPresentation.toMarkdown', () => {
  function makePresentation(requestImpl: (req: WorkerRequest) => unknown) {
    const request = vi.fn((build: (id: number) => WorkerRequest) =>
      Promise.resolve(requestImpl(build(1))),
    );
    const instance = Object.create(PptxPresentation.prototype) as Record<string, unknown>;
    instance._bridge = { request };
    const pres = instance as unknown as ToMarkdownProbe;
    return { pres, request };
  }

  it('posts a toMarkdown request and returns the rendered string', async () => {
    const { pres, request } = makePresentation((req) => {
      expect(req.kind).toBe('toMarkdown');
      return { kind: 'markdownRendered', id: 1, markdown: '# Slide 1\n\n- point' };
    });

    const md = await pres.toMarkdown();
    expect(md).toBe('# Slide 1\n\n- point');
    expect(request).toHaveBeenCalledTimes(1);
  });
});
