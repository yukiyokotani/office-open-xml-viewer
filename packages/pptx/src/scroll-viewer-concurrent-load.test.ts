import { describe, it, expect, afterEach, vi } from 'vitest';
import { PptxScrollViewer } from './scroll-viewer.js';
import { PptxPresentation } from './presentation.js';
import { installDom, makeContainer, FakePptxEngine, type FakeEl } from './scroll-viewer-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

const SLIDE_W_EMU = 9525 * 200;
const SLIDE_H_EMU = 9525 * 150;

/**
 * Concurrent-load latch for the SELF-LOADED scroll viewer (composes with SC20's
 * success-after-swap). Two overlapping `load(A)`/`load(B)` calls race the WASM
 * parse / worker init; the stale one resolving LAST must NOT win the swap — it
 * destroys its own just-loaded engine and leaves the winner (its engine +
 * recycle/relayout post-load work) untouched. An injected engine can never orphan
 * (load() throws up-front there), so this only covers the self-loading path.
 */
describe('PptxScrollViewer.load() — concurrent-load latch', () => {
  function build(onError?: (error: Error) => void) {
    installDom();
    const container = makeContainer(200, 400);
    const v = new PptxScrollViewer(container as unknown as HTMLElement, { gap: 10, onError });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    return { v };
  }

  function deferredLoad(engine: FakePptxEngine): { resolve: () => void; promise: Promise<PptxPresentation> } {
    let resolve!: () => void;
    const promise = new Promise<PptxPresentation>((r) => {
      resolve = () => r(engine.asPres());
    });
    return { resolve, promise };
  }

  it('the later-started load winning first leaves the stale load a no-op (its engine destroyed)', async () => {
    const { v } = build();
    const a = new FakePptxEngine(3, SLIDE_W_EMU, SLIDE_H_EMU);
    const b = new FakePptxEngine(3, SLIDE_W_EMU, SLIDE_H_EMU);
    const da = deferredLoad(a);
    const db = deferredLoad(b);
    vi.spyOn(PptxPresentation, 'load')
      .mockImplementationOnce(() => da.promise)
      .mockImplementationOnce(() => db.promise);

    const pa = v.load('a.pptx'); // gen 1
    const pb = v.load('b.pptx'); // gen 2 — supersedes A

    db.resolve();
    await pb;
    expect(b.destroyed).toBe(false);
    expect(a.destroyed).toBe(false);

    da.resolve();
    await pa;
    expect(a.destroyed).toBe(true); // loser's engine cleaned up (no leak)
    expect(b.destroyed).toBe(false); // winner untouched — still current
    expect(v.slideCount).toBe(3);

    v.destroy();
    expect(b.destroyed).toBe(true);
    expect(a.destroyed).toBe(true);
  });

  it('resolving in start order (A then B) behaves like today — B wins normally', async () => {
    const { v } = build();
    const a = new FakePptxEngine(3, SLIDE_W_EMU, SLIDE_H_EMU);
    const b = new FakePptxEngine(3, SLIDE_W_EMU, SLIDE_H_EMU);
    const da = deferredLoad(a);
    const db = deferredLoad(b);
    vi.spyOn(PptxPresentation, 'load')
      .mockImplementationOnce(() => da.promise)
      .mockImplementationOnce(() => db.promise);

    const pa = v.load('a.pptx');
    const pb = v.load('b.pptx');

    da.resolve();
    await pa;
    expect(a.destroyed).toBe(true); // superseded loser cleaned up

    db.resolve();
    await pb;
    expect(b.destroyed).toBe(false);
    expect(v.slideCount).toBe(3);

    v.destroy();
    expect(b.destroyed).toBe(true);
  });

  it('does not report an old slot render rejected while a successful reload recycles it', async () => {
    const onError = vi.fn();
    const { v } = build(onError);
    const old = new FakePptxEngine(3, SLIDE_W_EMU, SLIDE_H_EMU);
    const next = new FakePptxEngine(3, SLIDE_W_EMU, SLIDE_H_EMU);
    vi.spyOn(PptxPresentation, 'load')
      .mockResolvedValueOnce(old.asPres())
      .mockResolvedValueOnce(next.asPres());

    await v.load('old.pptx');
    const internals = v as unknown as {
      _slots: Map<number, { renderedSlide: number }>;
      _renderSlot(index: number, slot: { renderedSlide: number }): void;
    };
    const [entry] = internals._slots.entries();
    expect(entry).toBeDefined();
    const [index, slot] = entry as [number, { renderedSlide: number }];
    (old as unknown as { deferred: boolean }).deferred = true;
    slot.renderedSlide = -1;
    internals._renderSlot(index, slot);
    const staleCall = old.renderCalls.at(-1);

    await v.load('next.pptx');
    staleCall?.reject(new Error('old worker terminated'));
    await Promise.resolve();

    expect(old.destroyed).toBe(true);
    expect(onError).not.toHaveBeenCalled();
    v.destroy();
  });

  it('rejects load after destroy without acquiring a presentation', async () => {
    const { v } = build();
    const load = vi.spyOn(PptxPresentation, 'load');
    v.destroy();

    await expect(v.load('late.pptx')).rejects.toThrow('PptxScrollViewer is destroyed');
    expect(load).not.toHaveBeenCalled();
  });
});
