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
 * SC20: calling `load()` a second time must not orphan the previously loaded
 * engine (and its worker + pinned WASM allocation). The viewer holds a single
 * engine reference; overwriting it without `destroy()` leaks the worker.
 *
 * The fix is an atomic swap: the new engine is loaded first, and only on success
 * is the OLD engine destroyed — so a failed re-load preserves the current one.
 */
describe('PptxViewer.load() — no orphaned engine on re-load (SC20)', () => {
  function mount() {
    installDom();
    const canvas = makeEl('canvas');
    return { canvas };
  }

  it('destroys the previous engine when load() is called again', async () => {
    const { canvas } = mount();
    const first = new FakePptxEngine(3, SLIDE_W_EMU, SLIDE_H_EMU);
    const second = new FakePptxEngine(3, SLIDE_W_EMU, SLIDE_H_EMU);
    const tiff = { render: vi.fn(async () => null) };
    const loadSpy = vi
      .spyOn(PptxPresentation, 'load')
      .mockResolvedValueOnce(first.asPres())
      .mockResolvedValueOnce(second.asPres());

    const v = new PptxViewer(canvas as unknown as HTMLCanvasElement, {
      password: 'secret',
      tiff,
    });
    await v.load('one.pptx');
    expect(loadSpy).toHaveBeenNthCalledWith(
      1,
      'one.pptx',
      expect.objectContaining({ password: 'secret', tiff }),
    );
    expect(first.destroyed).toBe(false); // still current after first load

    await v.load('two.pptx');
    expect(loadSpy).toHaveBeenCalledTimes(2);
    // First engine (and its worker) torn down; second is now current.
    expect(first.destroyed).toBe(true);
    expect(second.destroyed).toBe(false);

    v.destroy();
    expect(second.destroyed).toBe(true);
  });

  it('keeps the current engine when the re-load fails (atomic swap)', async () => {
    const { canvas } = mount();
    const first = new FakePptxEngine(3, SLIDE_W_EMU, SLIDE_H_EMU);
    vi.spyOn(PptxPresentation, 'load')
      .mockResolvedValueOnce(first.asPres())
      .mockRejectedValueOnce(new Error('boom'));

    const onError = vi.fn();
    const v = new PptxViewer(canvas as unknown as HTMLCanvasElement, { onError });
    await v.load('one.pptx');

    await expect(v.load('bad.pptx')).rejects.toThrow('boom');
    // The rejected reload leaves the first engine intact and is not duplicated.
    expect(onError).not.toHaveBeenCalled();
    expect(first.destroyed).toBe(false);
    expect(v.slideCount).toBe(3);

    v.destroy();
    expect(first.destroyed).toBe(true);
  });
});
