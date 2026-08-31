import { describe, it, expect, afterEach, vi } from 'vitest';
import { DocxViewer } from './viewer.js';
import { installDom, makeEl, FakeDocxEngine, type FakeEl } from './scroll-viewer-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

/**
 * The viewer wraps the caller's `<canvas>` in a positioned wrapper by
 * REPARENTING it (`parent.insertBefore(wrapper, canvas)` → `wrapper.appendChild(canvas)`).
 * destroy() must UNDO that reparent so the caller's canvas is returned to its
 * original DOM position and can be reused (e.g. to re-create the viewer on the
 * same canvas). Removing the wrapper without returning the canvas would delete
 * the caller-owned element from the DOM — the bug these tests pin.
 */
describe('DocxViewer.destroy() — canvas reparent return', () => {
  /** parent → [before, canvas, after]; returns the pieces for assertions. */
  function mount() {
    installDom();
    const parent = makeEl('div');
    const before = makeEl('span');
    const canvas = makeEl('canvas');
    const after = makeEl('span');
    parent.appendChild(before);
    parent.appendChild(canvas);
    parent.appendChild(after);
    return { parent, before, canvas, after };
  }

  it('returns the canvas to its original parent and sibling position, and removes the wrapper', () => {
    const { parent, before, canvas, after } = mount();
    const v = new DocxViewer(canvas as unknown as HTMLCanvasElement);

    // After construction the canvas lives inside a wrapper that took its slot.
    const wrapper = parent.children[1] as FakeEl;
    expect(wrapper.tag).toBe('div');
    expect(wrapper.children).toContain(canvas);
    expect(canvas.parentElement).toBe(wrapper);

    v.destroy();

    // (b) wrapper removed from the DOM.
    expect(parent.children).not.toContain(wrapper);
    // (a) canvas returned to its ORIGINAL parent, between `before` and `after`.
    expect(canvas.parentElement).toBe(parent);
    expect(parent.children).toEqual([before, canvas, after]);
  });

  it('restores the canvas display style to its original (unset) value', () => {
    const { canvas } = mount();
    // The constructor forces display:block when unset; destroy must restore ''.
    expect(canvas.style.display).toBe('');
    const v = new DocxViewer(canvas as unknown as HTMLCanvasElement);
    expect(canvas.style.display).toBe('block'); // forced on
    v.destroy();
    expect(canvas.style.display).toBe(''); // restored to original
  });

  it('preserves a caller-set display style across construct/destroy', () => {
    const { canvas } = mount();
    canvas.style.display = 'inline-block'; // caller-provided value
    const v = new DocxViewer(canvas as unknown as HTMLCanvasElement);
    // Constructor must not overwrite an explicit display.
    expect(canvas.style.display).toBe('inline-block');
    v.destroy();
    expect(canvas.style.display).toBe('inline-block'); // untouched
  });

  it('handles a detached canvas (no original parent): destroy just unwraps, no throw', () => {
    installDom();
    const canvas = makeEl('canvas'); // never attached to a parent
    const v = new DocxViewer(canvas as unknown as HTMLCanvasElement);
    // The wrapper is the canvas's parent, itself detached (no grand-parent).
    const wrapper = canvas.parentElement as FakeEl;
    expect(wrapper.tag).toBe('div');
    expect(() => v.destroy()).not.toThrow();
    // Canvas is simply detached from the wrapper (original parent was null).
    expect(canvas.parentElement).toBe(null);
  });

  it('is safe to call destroy() twice', () => {
    const { canvas } = mount();
    const v = new DocxViewer(canvas as unknown as HTMLCanvasElement);
    v.destroy();
    expect(() => v.destroy()).not.toThrow();
  });

  it('falls back to appending when the recorded next-sibling was removed before destroy()', () => {
    const { parent, before, canvas, after } = mount();
    const v = new DocxViewer(canvas as unknown as HTMLCanvasElement);
    // The caller removes the recorded next-sibling while the viewer is alive.
    // insertBefore with that stale reference would throw NotFoundError in a
    // real browser — destroy() must detect it and append at the end instead.
    parent.removeChild(after);
    expect(() => v.destroy()).not.toThrow();
    expect(canvas.parentElement).toBe(parent);
    expect(parent.children).toEqual([before, canvas]);
  });

  it('borrows a document through fromDocument() with the same lifecycle contract as the scroll viewer', async () => {
    const { canvas } = mount();
    const engine = new FakeDocxEngine(2, [{ widthPt: 612, heightPt: 792 }]);
    const viewer = DocxViewer.fromDocument(
      canvas as unknown as HTMLCanvasElement,
      engine.asDoc(),
    );
    expect(viewer.pageCount).toBe(2);
    await expect((viewer as DocxViewer).load('x.docx')).rejects.toThrow(/fromDocument/);
    viewer.destroy();
    expect(engine.destroyed).toBe(false);
  });

  it('inherits the borrowed document\'s active layout view', async () => {
    const { canvas } = mount();
    const engine = new FakeDocxEngine(1, [{ widthPt: 612, heightPt: 792 }]);
    engine.layoutView = {
      showTrackedChanges: true,
      currentDate: 1_700_000_000_000,
    };
    const viewer = DocxViewer.fromDocument(
      canvas as unknown as HTMLCanvasElement,
      engine.asDoc(),
    );

    await viewer.goToPage(0);
    expect(engine.renderCalls.at(-1)).toMatchObject({
      showTrackedChanges: true,
      currentDate: 1_700_000_000_000,
    });

    await viewer.setShowTrackedChanges(false);
    expect(engine.layoutViews.at(-1)).toEqual({
      showTrackedChanges: false,
      currentDate: 1_700_000_000_000,
    });
    expect(engine.renderCalls.at(-1)).toMatchObject({
      showTrackedChanges: false,
      currentDate: 1_700_000_000_000,
    });
    viewer.destroy();
  });

  it('rejects a borrowed mode conflict before reparenting the canvas', () => {
    const { parent, canvas } = mount();
    const engine = new FakeDocxEngine(1, [{ widthPt: 612, heightPt: 792 }], 'worker');
    expect(() => DocxViewer.fromDocument(
      canvas as unknown as HTMLCanvasElement,
      engine.asDoc(),
      { mode: 'main' } as never,
    )).toThrow(/opts\.mode='main'.*mode='worker'/);
    expect(parent.children).toContain(canvas);
  });
});
