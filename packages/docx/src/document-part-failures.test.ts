import { afterEach, describe, expect, it, vi } from 'vitest';
import { DocxDocument } from './document.js';
import { attachDocumentLayoutRuntime } from './layout/runtime-state.js';

/**
 * #1974 — `DocxDocument.partFailures` and the main-thread resolution of
 * `errorPlaceholder`. Worker mode learns the failure from the clamped
 * `DocumentMeta.partFailure`; the option itself never reaches postMessage.
 */
function workerDocument(meta: Record<string, unknown> | null) {
  vi.stubGlobal('OffscreenCanvas', class { constructor(readonly width: number, readonly height: number) {} });
  const request = vi.fn(async () => ({ type: 'pageRendered', id: 1, bitmap: { close() {} }, runs: [], diagnostics: [] }));
  const instance = Object.create(DocxDocument.prototype) as Record<string, unknown>;
  Object.assign(instance, { _mode: 'worker', _meta: meta, _bridge: { request } });
  const document = instance as unknown as DocxDocument;
  attachDocumentLayoutRuntime(document, 0);
  return { document, request };
}

function sentOptions(request: ReturnType<typeof vi.fn>): Record<string, unknown> {
  const build = request.mock.calls[0]![0] as (id: number) => { opts: Record<string, unknown> };
  return build(1).opts;
}

afterEach(() => {
  vi.unstubAllGlobals();
});

describe('DocxDocument part failures (#1974)', () => {
  const meta = { pageCount: 1, pageSizes: [{ widthPt: 612, heightPt: 792 }], partFailure: { message: 'word/document.xml: bad' } };

  it('lists the failed main part from worker metadata', () => {
    const { document } = workerDocument(meta);
    expect(document.partFailures).toEqual([{ kind: 'document', message: 'word/document.xml: bad' }]);
    expect(Object.isFrozen(document.partFailures[0])).toBe(true);
    expect(workerDocument({ pageCount: 1, pageSizes: [] }).document.partFailures).toEqual([]);
  });

  it('resolves a function option on the main thread and sends only text', async () => {
    const { document, request } = workerDocument(meta);
    const resolver = vi.fn((failure: { message: string }) => ({ title: `Failed: ${failure.message}` }));
    await document.renderPageToBitmap(0, { dpr: 1, errorPlaceholder: resolver });
    expect(resolver).toHaveBeenCalledWith({ kind: 'document', message: 'word/document.xml: bad' });
    const opts = sentOptions(request);
    expect(opts).not.toHaveProperty('errorPlaceholder');
    expect(opts.errorPlaceholderText).toEqual({ state: 'resolved', title: 'Failed: word/document.xml: bad', detail: undefined });
    expect(() => structuredClone(opts)).not.toThrow();
  });

  it('never evaluates or sends the option for a healthy document', async () => {
    const { document, request } = workerDocument({ pageCount: 1, pageSizes: [] });
    const resolver = vi.fn(() => ({ title: 'x' }));
    await document.renderPageToBitmap(0, { dpr: 1, errorPlaceholder: resolver });
    expect(resolver).not.toHaveBeenCalled();
    const opts = sentOptions(request);
    expect(opts).not.toHaveProperty('errorPlaceholder');
    expect(opts).not.toHaveProperty('errorPlaceholderText');
  });

  it('sends the built-in state for a resolver returning undefined', async () => {
    const { document, request } = workerDocument(meta);
    await document.renderPageToBitmap(0, { dpr: 1, errorPlaceholder: () => undefined });
    expect(sentOptions(request)).not.toHaveProperty('errorPlaceholderText');
  });

  it('rejects the render for an invalid or throwing option', async () => {
    const { document } = workerDocument(meta);
    await expect(document.renderPageToBitmap(0, { errorPlaceholder: { title: 1 } as never })).rejects.toThrow(TypeError);
    const failure = new Error('resolver failed');
    await expect(document.renderPageToBitmap(0, { errorPlaceholder: () => { throw failure; } })).rejects.toBe(failure);
  });

  it('empties the list on destroy', () => {
    const { document } = workerDocument(meta);
    expect(document.partFailures).toHaveLength(1);
    Object.assign(document as unknown as Record<string, unknown>, { _bridge: { terminate() {} } });
    try { document.destroy(); } catch { /* off-prototype teardown of unrelated resources */ }
    expect(document.partFailures).toEqual([]);
  });

  it('never posts the option through collectPageRuns', async () => {
    const { document, request } = workerDocument(meta);
    request.mockImplementation(async () => ({ type: 'runsCollected', id: 1, runs: [] }) as never);
    const reused = { width: 600, errorPlaceholder: () => ({ title: 'x' }) };
    await document.collectPageRuns(0, reused as never);
    const opts = sentOptions(request);
    expect(opts).toEqual({ width: 600 });
    expect(() => structuredClone(opts)).not.toThrow();
  });
});

