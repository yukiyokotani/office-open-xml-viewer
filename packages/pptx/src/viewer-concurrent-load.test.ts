import { describe, it, expect, afterEach, vi } from 'vitest';
import { PptxViewer } from './viewer.js';
import { PptxPresentation } from './presentation.js';
import { installDom, makeEl, FakePptxEngine } from './scroll-viewer-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

const SLIDE_W_EMU = 9144000;
const SLIDE_H_EMU = 6858000;

/**
 * Concurrent-load latch (composes with SC20's success-after-swap): if a caller
 * fires `load(A)` and, before it resolves, `load(B)`, both loads race the WASM
 * parse / worker init concurrently. Whichever resolves LAST must NOT win the swap
 * when it is the stale one — the loser's freshly-loaded engine (never installed,
 * or installed then overwritten) must be destroyed, not leaked, and the winner's
 * engine must stay live and untouched. A generation token (`_loadGen`) closes it.
 */
describe('PptxViewer.load() — concurrent-load latch', () => {
  function mount() {
    installDom();
    return { canvas: makeEl('canvas') };
  }

  function deferredLoad(engine: FakePptxEngine): { resolve: () => void; promise: Promise<PptxPresentation> } {
    let resolve!: () => void;
    const promise = new Promise<PptxPresentation>((r) => {
      resolve = () => r(engine.asPres());
    });
    return { resolve, promise };
  }

  it('the later-started load winning first leaves the stale load a no-op (its engine destroyed)', async () => {
    const { canvas } = mount();
    const a = new FakePptxEngine(3, SLIDE_W_EMU, SLIDE_H_EMU);
    const b = new FakePptxEngine(3, SLIDE_W_EMU, SLIDE_H_EMU);
    const da = deferredLoad(a);
    const db = deferredLoad(b);
    vi.spyOn(PptxPresentation, 'load')
      .mockImplementationOnce(() => da.promise)
      .mockImplementationOnce(() => db.promise);

    const v = new PptxViewer(canvas as unknown as HTMLCanvasElement);
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

    v.destroy();
    expect(b.destroyed).toBe(true);
    expect(a.destroyed).toBe(true);
  });

  it('resolving in start order (A then B) behaves like today — B wins normally', async () => {
    const { canvas } = mount();
    const a = new FakePptxEngine(3, SLIDE_W_EMU, SLIDE_H_EMU);
    const b = new FakePptxEngine(3, SLIDE_W_EMU, SLIDE_H_EMU);
    const da = deferredLoad(a);
    const db = deferredLoad(b);
    vi.spyOn(PptxPresentation, 'load')
      .mockImplementationOnce(() => da.promise)
      .mockImplementationOnce(() => db.promise);

    const v = new PptxViewer(canvas as unknown as HTMLCanvasElement);
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

  it('does not double-destroy or leak when only one load runs (regression guard)', async () => {
    const { canvas } = mount();
    const only = new FakePptxEngine(3, SLIDE_W_EMU, SLIDE_H_EMU);
    vi.spyOn(PptxPresentation, 'load').mockResolvedValue(only.asPres());
    const v = new PptxViewer(canvas as unknown as HTMLCanvasElement);
    await v.load('one.pptx');
    expect(only.destroyed).toBe(false);
    v.destroy();
    expect(only.destroyed).toBe(true);
  });

  it('does not report a pending old-presentation render rejected by a successful reload', async () => {
    const { canvas } = mount();
    const onError = vi.fn();
    const old = new FakePptxEngine(3, SLIDE_W_EMU, SLIDE_H_EMU);
    const next = new FakePptxEngine(3, SLIDE_W_EMU, SLIDE_H_EMU);
    vi.spyOn(PptxPresentation, 'load')
      .mockResolvedValueOnce(old.asPres())
      .mockResolvedValueOnce(next.asPres());

    const v = new PptxViewer(canvas as unknown as HTMLCanvasElement, { onError });
    await v.load('old.pptx');
    (old as unknown as { deferred: boolean }).deferred = true;
    const staleNavigation = v.goToSlide(1);
    const staleCall = old.renderCalls.at(-1);
    expect(staleCall?.slide).toBe(1);

    await v.load('next.pptx');
    expect(old.destroyed).toBe(true);
    staleCall?.reject(new Error('old worker terminated'));
    await staleNavigation;

    expect(onError).not.toHaveBeenCalled();
    expect(next.destroyed).toBe(false);
    v.destroy();
  });

  it('rejects load after destroy without acquiring a presentation', async () => {
    const { canvas } = mount();
    const load = vi.spyOn(PptxPresentation, 'load');
    const v = new PptxViewer(canvas as unknown as HTMLCanvasElement);
    v.destroy();

    await expect(v.load('late.pptx')).rejects.toThrow('PptxViewer is destroyed');
    expect(load).not.toHaveBeenCalled();
  });
});
