import { describe, it, expect, afterEach, vi } from 'vitest';
import { DocxViewer } from './viewer.js';
import { DocxDocument } from './document.js';
import { installDom, makeEl, FakeDocxEngine } from './scroll-viewer-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

const A4 = [{ widthPt: 595, heightPt: 842 }];

/**
 * SC20: a second `load()` must not orphan the previous engine (its worker +
 * pinned WASM). The fix is an atomic swap — load the new engine first, destroy
 * the old one only on success.
 */
describe('DocxViewer.load() — no orphaned engine on re-load (SC20)', () => {
  function mount() {
    installDom();
    const canvas = makeEl('canvas');
    return { canvas };
  }

  it('destroys the previous engine when load() is called again', async () => {
    const { canvas } = mount();
    const first = new FakeDocxEngine(2, A4);
    const second = new FakeDocxEngine(2, A4);
    const tiff = { render: vi.fn(async () => null) };
    const loadSpy = vi
      .spyOn(DocxDocument, 'load')
      .mockResolvedValueOnce(first.asDoc())
      .mockResolvedValueOnce(second.asDoc());

    const v = new DocxViewer(canvas as unknown as HTMLCanvasElement, {
      password: 'secret',
      tiff,
    });
    await v.load('one.docx');
    expect(loadSpy).toHaveBeenNthCalledWith(
      1,
      'one.docx',
      expect.objectContaining({ password: 'secret', tiff }),
    );
    expect(first.destroyed).toBe(false);

    await v.load('two.docx');
    expect(loadSpy).toHaveBeenCalledTimes(2);
    expect(first.destroyed).toBe(true);
    expect(second.destroyed).toBe(false);

    v.destroy();
    expect(second.destroyed).toBe(true);
  });

  it('keeps the current engine when the re-load fails (atomic swap)', async () => {
    const { canvas } = mount();
    const first = new FakeDocxEngine(2, A4);
    vi.spyOn(DocxDocument, 'load')
      .mockResolvedValueOnce(first.asDoc())
      .mockRejectedValueOnce(new Error('boom'));

    const onError = vi.fn();
    const v = new DocxViewer(canvas as unknown as HTMLCanvasElement, { onError });
    await v.load('one.docx');

    await expect(v.load('bad.docx')).rejects.toThrow('boom');
    expect(onError).not.toHaveBeenCalled();
    expect(first.destroyed).toBe(false);
    expect(v.pageCount).toBe(2);

    v.destroy();
    expect(first.destroyed).toBe(true);
  });
});
