import { describe, it, expect, afterEach, vi } from 'vitest';
import { DocxScrollViewer } from './scroll-viewer.js';
import { DocxDocument } from './document.js';
import { installDom, makeContainer, FakeDocxEngine, type FakeEl } from './scroll-viewer-test-dom.js';
import { publishDocxLayout } from './document-layout-events.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

const SIZE = [{ widthPt: 100, heightPt: 200 }];

/**
 * SC20 for the scroll viewer: a SELF-LOADED (non-injected) scroll viewer owns its
 * engine, so a second `load()` must destroy the previous one instead of orphaning
 * its worker + WASM. (An injected engine is caller-owned — load() throws there.)
 */
describe('DocxScrollViewer.load() — no orphaned engine on re-load (SC20)', () => {
  function build(opts = {}) {
    installDom();
    const container = makeContainer(200, 400);
    const v = new DocxScrollViewer(container as unknown as HTMLElement, { gap: 10, ...opts });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    return { v };
  }

  it('destroys the previous engine when a self-loaded viewer is re-loaded', async () => {
    const { v } = build();
    const first = new FakeDocxEngine(4, SIZE);
    const second = new FakeDocxEngine(4, SIZE);
    vi.spyOn(DocxDocument, 'load')
      .mockResolvedValueOnce(first.asDoc())
      .mockResolvedValueOnce(second.asDoc());

    await v.load('one.docx');
    expect(first.destroyed).toBe(false);

    await v.load('two.docx');
    expect(first.destroyed).toBe(true);
    expect(second.destroyed).toBe(false);

    v.destroy();
    expect(second.destroyed).toBe(true);
  });

  it('keeps the current engine when the re-load fails (atomic swap)', async () => {
    const onError = vi.fn();
    const { v } = build({ onError });
    const first = new FakeDocxEngine(4, SIZE);
    vi.spyOn(DocxDocument, 'load')
      .mockResolvedValueOnce(first.asDoc())
      .mockRejectedValueOnce(new Error('boom'));

    await v.load('one.docx');
    await expect(v.load('bad.docx')).rejects.toThrow('boom');
    expect(onError).not.toHaveBeenCalled();
    expect(first.destroyed).toBe(false);
    expect(v.pageCount).toBe(4);

    v.destroy();
    expect(first.destroyed).toBe(true);
  });

  it('a failed re-load does not disconnect the retained document\u2019s background layout', async () => {
    // The atomic swap keeps the previous document installed when acquisition
    // rejects. Progress therefore follows a subscription to the installed
    // document, not callbacks owned by whichever acquisition ran most recently.
    const onVisiblePageChange = vi.fn();
    const { v } = build({ progressiveLayout: true, onVisiblePageChange });
    const first = new FakeDocxEngine(2, SIZE);
    vi.spyOn(DocxDocument, 'load')
      .mockResolvedValueOnce(first.asDoc())
      .mockRejectedValueOnce(new Error('boom'));

    await v.load('one.docx');
    expect(v.pageCount).toBe(2);
    await expect(v.load('bad.docx')).rejects.toThrow('boom');

    // The retained document's background layout finishes AFTER the failed
    // swap; its completion must still reach this viewer.
    first.setPageCount(80);
    onVisiblePageChange.mockClear();
    publishDocxLayout(first.asDoc(), { pageCount: 80, exact: true, complete: true });
    expect(v.pageCount).toBe(80);
    expect(onVisiblePageChange).toHaveBeenCalledWith(0, 80, expect.anything());

    v.destroy();
    expect(first.destroyed).toBe(true);
  });

  it('keeps retained progressive callbacks after overlapping newer loads fail and stale-settle', async () => {
    const onVisiblePageChange = vi.fn();
    const { v } = build({ progressiveLayout: true, onVisiblePageChange });
    const retained = new FakeDocxEngine(2, SIZE);
    const stale = new FakeDocxEngine(4, SIZE);
    let resolveStale!: () => void;
    const staleLoad = new Promise<DocxDocument>((resolve) => {
      resolveStale = () => resolve(stale.asDoc());
    });

    vi.spyOn(DocxDocument, 'load')
      .mockResolvedValueOnce(retained.asDoc())
      .mockImplementationOnce(() => staleLoad)
      .mockRejectedValueOnce(new Error('newest failed'));

    await v.load('retained.docx');
    const superseded = v.load('stale.docx');
    await expect(v.load('newest.docx')).rejects.toThrow('newest failed');

    // The middle load is stale after the newest request, even though that
    // newest acquisition failed. Its candidate must not replace the retained
    // document when it eventually settles.
    resolveStale();
    await superseded;
    expect(stale.destroyed).toBe(true);
    expect(retained.destroyed).toBe(false);

    // The retained document remains the installed authority and must still be
    // allowed to publish the authoritative geometry from its background layout.
    retained.setPageCount(80);
    onVisiblePageChange.mockClear();
    publishDocxLayout(retained.asDoc(), { pageCount: 80, exact: true, complete: true });
    expect(v.pageCount).toBe(80);
    expect(onVisiblePageChange).toHaveBeenCalledWith(0, 80, expect.anything());

    v.destroy();
    expect(retained.destroyed).toBe(true);
  });

  it('rejects an initial window render failure without also calling onError', async () => {
    const onError = vi.fn();
    const { v } = build({ onError });
    const engine = new FakeDocxEngine(1, SIZE, 'main', true);
    const failure = new Error('initial page render failed');
    vi.spyOn(DocxDocument, 'load').mockResolvedValue(engine.asDoc());

    const loading = v.load('one.docx');
    await Promise.resolve();
    await Promise.resolve();
    expect(engine.renderCalls).toHaveLength(1);
    engine.renderCalls[0].reject(failure);

    await expect(loading).rejects.toBe(failure);
    expect(onError).not.toHaveBeenCalled();
    v.destroy();
  });
});
