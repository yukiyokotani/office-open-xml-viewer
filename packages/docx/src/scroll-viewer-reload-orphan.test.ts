import { describe, it, expect, afterEach, vi } from 'vitest';
import { DocxScrollViewer } from './scroll-viewer.js';
import { DocxDocument } from './document.js';
import { installDom, makeContainer, FakeDocxEngine, type FakeEl } from './scroll-viewer-test-dom.js';

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
    return { v, container };
  }

  /** The virtualization spacer whose height IS the document's scroll extent. */
  function spacerHeight(container: FakeEl): number {
    return parseFloat(
      ((container.children[0] as FakeEl).children[0] as FakeEl).children[0].style.height,
    );
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
    // rejects — but load() had already bumped the generation its progressive
    // callbacks captured. Without restoring it, the retained document's later
    // publications and completion were ignored forever: the engine reported the
    // full page count while the viewer stayed frozen at the preview prefix.
    const onVisiblePageChange = vi.fn();
    const { v } = build({ progressiveLayout: true, onVisiblePageChange });
    const first = new FakeDocxEngine(2, SIZE);
    let captured: Parameters<typeof DocxDocument.load>[1] | undefined;
    vi.spyOn(DocxDocument, 'load')
      .mockImplementationOnce((_source, opts) => {
        captured = opts;
        return Promise.resolve(first.asDoc());
      })
      .mockRejectedValueOnce(new Error('boom'));

    await v.load('one.docx');
    expect(v.pageCount).toBe(2);
    await expect(v.load('bad.docx')).rejects.toThrow('boom');

    // The retained document's background layout finishes AFTER the failed
    // swap; its completion must still reach this viewer.
    first.setPageCount(80);
    onVisiblePageChange.mockClear();
    captured?.onLayoutComplete?.();
    expect(v.pageCount).toBe(80);
    expect(onVisiblePageChange).toHaveBeenCalledWith(0, 80, expect.anything());

    v.destroy();
    expect(first.destroyed).toBe(true);
  });

  it('keeps the retained document\u2019s layout alive when an OVERLAPPING newer load fails', async () => {
    // Three-load overlap, which the single-reload case above cannot reach.
    // A is installed; B is still acquiring; C starts and rejects.
    // `TerminalResourceOwner`'s replacement generation is MONOTONIC, so C's
    // start already condemned B: when B finally resolves the owner discards it
    // and A stays installed. A viewer that answered "which document owns my
    // relayout?" with the request counter would restore the number B captured
    // — a document nobody installed — and A's own publications, captured one
    // number lower, would be refused forever: the engine reports 80 pages while
    // the spacer stays at the 2-page preview height.
    const { v, container } = build({ progressiveLayout: true });
    const retained = new FakeDocxEngine(2, SIZE);
    const superseded = new FakeDocxEngine(9, SIZE);
    let captured: Parameters<typeof DocxDocument.load>[1] | undefined;
    let resolveSuperseded!: () => void;
    const supersededLoad = new Promise<DocxDocument>((resolve) => {
      resolveSuperseded = () => { resolve(superseded.asDoc()); };
    });
    vi.spyOn(DocxDocument, 'load')
      .mockImplementationOnce((_source, opts) => {
        captured = opts;
        return Promise.resolve(retained.asDoc());
      })
      .mockImplementationOnce(() => supersededLoad)
      .mockRejectedValueOnce(new Error('boom'));

    await v.load('retained.docx');
    const provisionalHeight = spacerHeight(container);
    expect(v.pageCount).toBe(2);

    const pending = v.load('superseded.docx');
    await expect(v.load('failing.docx')).rejects.toThrow('boom');

    // B loses the swap and destroys itself; A is untouched and still on screen.
    resolveSuperseded();
    await pending;
    expect(superseded.destroyed).toBe(true);
    expect(retained.destroyed).toBe(false);
    expect(v.pageCount).toBe(2);

    // A's background layout finishes. Its publications must still reach the
    // viewer: it is the document being rendered. The spacer, not `pageCount`,
    // is what pins that — `pageCount` reads straight through to the engine and
    // so reports the growth whether or not the viewer ever relaid out.
    retained.setPageCount(40);
    captured?.onLayoutPartial?.({ pageCount: 40, exact: false });
    const partialHeight = spacerHeight(container);
    expect(partialHeight).toBeGreaterThan(provisionalHeight);

    retained.setPageCount(80);
    captured?.onLayoutComplete?.();
    expect(v.pageCount).toBe(80);
    expect(spacerHeight(container)).toBeGreaterThan(partialHeight);

    v.destroy();
    expect(retained.destroyed).toBe(true);
  });

  it('refuses a superseded load\u2019s publications while the retained document is on screen', async () => {
    // The mirror image of the test above: B lost the swap, so B's own
    // background layout may not relayout the viewer around A's geometry.
    const { v } = build({ progressiveLayout: true });
    const retained = new FakeDocxEngine(2, SIZE);
    const superseded = new FakeDocxEngine(9, SIZE);
    let supersededOpts: Parameters<typeof DocxDocument.load>[1] | undefined;
    let resolveSuperseded!: () => void;
    const supersededLoad = new Promise<DocxDocument>((resolve) => {
      resolveSuperseded = () => { resolve(superseded.asDoc()); };
    });
    vi.spyOn(DocxDocument, 'load')
      .mockImplementationOnce(() => Promise.resolve(retained.asDoc()))
      .mockImplementationOnce((_source, opts) => {
        supersededOpts = opts;
        return supersededLoad;
      })
      .mockRejectedValueOnce(new Error('boom'));

    await v.load('retained.docx');
    const pending = v.load('superseded.docx');
    await expect(v.load('failing.docx')).rejects.toThrow('boom');
    resolveSuperseded();
    await pending;

    // A publication from the discarded document changes nothing on screen: it
    // must not even repaint A's mounted slots, which is what a non-exact
    // publication does for the document it belongs to.
    const paintedBefore = retained.renderCalls.length;
    supersededOpts?.onLayoutPartial?.({ pageCount: 500, exact: false });
    supersededOpts?.onLayoutComplete?.();
    expect(retained.renderCalls.length).toBe(paintedBefore);
    expect(v.pageCount).toBe(2);

    v.destroy();
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
