import { describe, it, expect, afterEach, vi } from 'vitest';
import { DocxViewer } from './viewer.js';
import { DocxDocument } from './document.js';
import { PT_TO_PX } from '@silurus/ooxml-core';
import { installDom, makeEl, FakeDocxEngine, type FakeEl } from './scroll-viewer-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

const PAGE = [{ widthPt: 600, heightPt: 800 }];
const NATURAL_W = 600 * PT_TO_PX; // 800 CSS px at 100%
const NATURAL_H = 800 * PT_TO_PX;

/** Mount a DocxViewer over a FakeDocxEngine (the render-error suite's seam), with
 *  a parent container of the given size so fit has something to measure. */
async function mount(opts: Record<string, unknown> = {}, containerSize?: { w: number; h: number }) {
  installDom();
  const canvas = makeEl('canvas');
  if (containerSize) {
    // Give the canvas a laid-out parent; the viewer reparents into it, so
    // `_wrapper.parentElement` becomes this element and `_fitContainer` sees it.
    const parent = makeEl('div');
    parent.clientWidth = containerSize.w;
    parent.clientHeight = containerSize.h;
    parent.appendChild(canvas);
  }
  const engine = new FakeDocxEngine(3, PAGE);
  vi.spyOn(DocxDocument, 'load').mockResolvedValue(engine.asDoc());
  const v = new DocxViewer(canvas as unknown as HTMLCanvasElement, opts);
  await v.load('x.docx');
  return { v, engine, canvas: canvas as FakeEl };
}

/** The `width` passed to the most recent renderPage call. */
function lastRenderWidth(engine: FakeDocxEngine): number | undefined {
  const w = engine.renderCalls;
  return w[w.length - 1]?.width;
}

describe('DocxViewer IX9 zoom contract — byte-identical default', () => {
  it('renders at opts.width verbatim when no zoom method is called', async () => {
    const { engine } = await mount({ width: 1000 });
    // The initial load render used opts.width unchanged (no zoom latched).
    expect(lastRenderWidth(engine)).toBe(1000);
  });

  it('renders at the natural width (undefined) when opts.width is unset', async () => {
    const { engine } = await mount();
    expect(lastRenderWidth(engine)).toBeUndefined();
  });

  it('getScale() reflects opts.width / natural before any zoom call', async () => {
    const { v } = await mount({ width: NATURAL_W * 2 });
    expect(v.getScale()).toBeCloseTo(2, 6);
  });

  it('getScale() is 1 when opts.width is unset (natural render)', async () => {
    const { v } = await mount();
    expect(v.getScale()).toBe(1);
  });

  it('worker bitmap clamping preserves the requested logical CSS page size', async () => {
    installDom();
    const canvas = makeEl('canvas');
    const engine = new FakeDocxEngine(1, PAGE, 'worker');
    vi.spyOn(DocxDocument, 'load').mockResolvedValue(engine.asDoc());
    const v = new DocxViewer(canvas as unknown as HTMLCanvasElement, {
      width: 1200,
      dpr: 2,
      mode: 'worker',
    });
    await v.load('x.docx');
    // The fake bitmap is only 600×800, modelling a backing store clamped below
    // the requested 2400×3200 device pixels. Its CSS box must retain the
    // requested 1200×1600 logical size rather than collapsing to bitmap/dpr.
    expect(canvas.width).toBe(600);
    expect(canvas.height).toBe(800);
    expect(canvas.style.width).toBe('1200px');
    expect(canvas.style.height).toBe('1600px');
    v.destroy();
  });
});

describe('DocxViewer IX9 zoom contract — setScale / steppers', () => {
  it('setScale renders at naturalWidth × scale and fires onScaleChange', async () => {
    const onScaleChange = vi.fn();
    const { v, engine } = await mount({ onScaleChange });
    await v.setScale(1.5);
    expect(v.getScale()).toBe(1.5);
    expect(lastRenderWidth(engine)).toBe(Math.round(NATURAL_W * 1.5));
    expect(onScaleChange).toHaveBeenCalledTimes(1);
    expect(onScaleChange).toHaveBeenCalledWith(1.5);
  });

  it('setScale clamps to [zoomMin, zoomMax]', async () => {
    const { v } = await mount({ zoomMin: 0.5, zoomMax: 2 });
    await v.setScale(10);
    expect(v.getScale()).toBe(2);
    await v.setScale(0.01);
    expect(v.getScale()).toBe(0.5);
  });

  it('setScale does not fire onScaleChange when the factor is unchanged', async () => {
    const onScaleChange = vi.fn();
    const { v } = await mount({ onScaleChange });
    await v.setScale(1.5);
    onScaleChange.mockClear();
    await v.setScale(1.5);
    expect(onScaleChange).not.toHaveBeenCalled();
  });

  it('zoomIn / zoomOut walk the shared ladder from 100%', async () => {
    const { v } = await mount();
    expect(v.getScale()).toBe(1);
    await v.zoomIn();
    expect(v.getScale()).toBe(1.1);
    await v.zoomIn();
    expect(v.getScale()).toBe(1.25);
    await v.zoomOut();
    expect(v.getScale()).toBe(1.1);
  });
});

describe('DocxViewer IX9 zoom contract — fitWidth / fitPage', () => {
  it('fitWidth sets the scale that spans the page width in the container', async () => {
    const { v } = await mount({}, { w: NATURAL_W / 2, h: 10000 });
    await v.fitWidth();
    // container width = naturalW/2 ⇒ fit factor 0.5.
    expect(v.getScale()).toBeCloseTo(0.5, 6);
  });

  it('fitPage takes the tighter of width/height', async () => {
    // Container: width fits at 1.0 (== natural), height fits at 0.5.
    const { v } = await mount({}, { w: NATURAL_W, h: NATURAL_H / 2 });
    await v.fitPage();
    expect(v.getScale()).toBeCloseTo(0.5, 6);
  });

  it('fitWidth respects an explicit opts.container over the DOM parent', async () => {
    const container = makeEl('div');
    container.clientWidth = NATURAL_W / 4; // ⇒ fit 0.25
    container.clientHeight = 10000;
    const { v } = await mount({ container: container as unknown as HTMLElement });
    await v.fitWidth();
    expect(v.getScale()).toBeCloseTo(0.25, 6);
  });

  it('fitWidth defers (no-op) with no container to measure', async () => {
    const onScaleChange = vi.fn();
    const { v } = await mount({ onScaleChange }); // detached canvas, no parent
    await v.fitWidth();
    expect(v.getScale()).toBe(1);
    expect(onScaleChange).not.toHaveBeenCalled();
  });
});

// IX9 F1 — family-unified pre-load setScale semantics (pinned across all five
// viewers): a setScale before load is LATCHED and applied to the first render.
describe('DocxViewer IX9 zoom contract — pre-load setScale latch (F1)', () => {
  it('setScale before load/layout is latched and applied once established (IX9 F1)', async () => {
    installDom();
    const canvas = makeEl('canvas');
    const engine = new FakeDocxEngine(3, PAGE);
    vi.spyOn(DocxDocument, 'load').mockResolvedValue(engine.asDoc());
    const v = new DocxViewer(canvas as unknown as HTMLCanvasElement, {});
    await v.setScale(1.5); // nothing loaded yet — latched, render no-ops
    expect(v.getScale()).toBe(1.5); // getScale reports the latched factor
    await v.load('x.docx');
    // The first real render already honours the latched factor.
    expect(lastRenderWidth(engine)).toBe(Math.round(NATURAL_W * 1.5));
    expect(v.getScale()).toBe(1.5);
    v.destroy();
  });

  it('a pre-load setScale latch is clamped to [zoomMin, zoomMax] (IX9 F1)', async () => {
    installDom();
    const canvas = makeEl('canvas');
    const engine = new FakeDocxEngine(3, PAGE);
    vi.spyOn(DocxDocument, 'load').mockResolvedValue(engine.asDoc());
    const v = new DocxViewer(canvas as unknown as HTMLCanvasElement, { zoomMin: 0.5, zoomMax: 3 });
    await v.setScale(100);
    expect(v.getScale()).toBe(3); // latched pre-clamped
    v.destroy();
  });
});
