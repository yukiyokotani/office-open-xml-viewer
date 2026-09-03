import { describe, it, expect, afterEach, vi } from 'vitest';
import { PptxScrollViewer } from './scroll-viewer.js';
import { installDom, makeContainer, makeBorrowedPptxScrollViewer, FakePptxEngine, type FakeEl } from './scroll-viewer-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

/**
 * IX9 — PptxScrollViewer's slice of the shared
 * {@link import('@silurus/ooxml-core').ZoomableViewer} contract. Mirrors the docx
 * scroll-viewer contract test: the pre-existing ABSOLUTE `setScale`, Ctrl-wheel
 * zoom and resize re-fit stay verbatim; IX9 adds getScale / zoomIn / zoomOut /
 * fitWidth / fitPage + onScaleChange over the same absolute `_scale` (a slide
 * draws at `slideWidth/EMU_PER_PX × _scale`).
 */
const SLIDE_W_EMU = 9525 * 200; // natural 200 px
const SLIDE_H_EMU = 9525 * 120; // natural 120 px

function setup(opts: Record<string, unknown> = {}, host = { w: 200, h: 400 }) {
  installDom();
  const container = makeContainer(host.w, host.h);
  const engine = new FakePptxEngine(20, SLIDE_W_EMU, SLIDE_H_EMU);
  const v = makeBorrowedPptxScrollViewer(container as unknown as HTMLElement, {
    presentation: engine.asPres(),
    gap: 10,
    paddingTop: 0,
    paddingBottom: 0,
    paddingLeft: 0,
    paddingRight: 0,
    zoomMin: 0.1,
    zoomMax: 4,
    ...opts,
  });
  const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
  scrollHost.clientHeight = host.h;
  scrollHost.clientWidth = host.w;
  v.relayout();
  return { v, scrollHost, engine, container };
}

describe('PptxScrollViewer IX9 zoom contract', () => {
  it('getScale() returns the absolute factor (the base fit after load)', () => {
    const { v } = setup();
    // base = 200 / 200 = 1.0
    expect(v.getScale()).toBeCloseTo(1.0, 6);
    expect(v.getScale()).toBeCloseTo(v.scaleForTest(), 10);
    v.destroy();
  });

  it('getScale() is 1 before a scale is established (no width yet)', () => {
    installDom();
    const engine = new FakePptxEngine(3, SLIDE_W_EMU, SLIDE_H_EMU);
    const v = makeBorrowedPptxScrollViewer(makeContainer(0, 0) as unknown as HTMLElement, {
      presentation: engine.asPres(),
    });
    expect(v.getScale()).toBe(1);
    v.destroy();
  });

  it('setScale fires onScaleChange with the new factor on a change only', () => {
    const onScaleChange = vi.fn();
    const { v } = setup({ onScaleChange });
    v.setScale(2);
    expect(v.getScale()).toBeCloseTo(2, 6);
    expect(onScaleChange).toHaveBeenCalledTimes(1);
    expect(onScaleChange).toHaveBeenCalledWith(2);
    onScaleChange.mockClear();
    v.setScale(2); // unchanged
    expect(onScaleChange).not.toHaveBeenCalled();
    v.destroy();
  });

  it('zoomIn / zoomOut walk the shared ladder from the base 1.0', () => {
    const { v } = setup();
    expect(v.getScale()).toBeCloseTo(1.0, 6);
    v.zoomIn();
    expect(v.getScale()).toBeCloseTo(1.1, 6);
    v.zoomIn();
    expect(v.getScale()).toBeCloseTo(1.25, 6);
    v.zoomOut();
    expect(v.getScale()).toBeCloseTo(1.1, 6);
    v.destroy();
  });

  it('zoomOut returns to an initial width fit below zoomMin', () => {
    // A poster-sized slide can need a scale below the ordinary user zoom floor.
    // The opening layout already permits that fit, so it must stay reachable
    // after the user zooms in. Here natural width 200px in a 10px host gives a
    // 0.05 fit, below zoomMin 0.1 and below the ladder's 0.25 bottom rung.
    const { v } = setup({}, { w: 10, h: 400 });
    expect(v.getScale()).toBeCloseTo(0.05, 6);

    v.zoomIn();
    expect(v.getScale()).toBeGreaterThan(0.05);
    v.zoomOut();

    expect(v.getScale()).toBeCloseTo(0.05, 6);
    v.destroy();
  });

  it('wheel zoom can round-trip to a poster fit below an explicit zoomMin', () => {
    const { v, scrollHost } = setup({ zoomMin: 0.5 }, { w: 10, h: 400 });
    const initialFit = v.getScale();
    expect(initialFit).toBeCloseTo(0.05, 6);

    scrollHost.dispatch('wheel', {
      ctrlKey: true,
      deltaY: -50,
      clientX: 5,
      clientY: 200,
      preventDefault() {},
    });
    expect(v.getScale()).toBeGreaterThan(initialFit);
    scrollHost.dispatch('wheel', {
      ctrlKey: true,
      deltaY: 50,
      clientX: 5,
      clientY: 200,
      preventDefault() {},
    });

    expect(v.getScale()).toBeCloseTo(initialFit, 10);
    v.destroy();
  });

  it('fitWidth restores the width-fit base after a zoom', () => {
    const { v } = setup();
    v.setScale(4);
    expect(v.getScale()).toBeCloseTo(4, 6);
    v.fitWidth();
    expect(v.getScale()).toBeCloseTo(1.0, 6);
    v.destroy();
  });

  it('fitPage takes the tighter of width/height fit', () => {
    // Container 200 wide × 60 tall. Natural slide 200 × 120.
    // widthfit = 200/200 = 1.0; heightfit = 60/120 = 0.5 ⇒ page-fit 0.5.
    const { v } = setup({}, { w: 200, h: 60 });
    v.fitPage();
    expect(v.getScale()).toBeCloseTo(0.5, 5);
    v.destroy();
  });

  it('the wheel-zoom path also notifies through onScaleChange', () => {
    const onScaleChange = vi.fn();
    const { v, scrollHost } = setup({ onScaleChange });
    scrollHost.dispatch('wheel', { ctrlKey: true, deltaY: -50, preventDefault() {} });
    expect(onScaleChange).toHaveBeenCalledTimes(1);
    expect(onScaleChange.mock.calls[0][0]).toBeGreaterThan(1.0);
    v.destroy();
  });

  // IX9 F1 — family-unified pre-load setScale semantics: a setScale before the
  // layout establishes must be LATCHED and applied once it does (the
  // single-canvas viewers honour a pre-load setScale on their first render; the
  // scroll viewers used to silently drop it).
  it('setScale before load/layout is latched and applied once established (IX9 F1)', () => {
    installDom();
    const container = makeContainer(0, 0); // zero-width ⇒ fit deferred, nothing established
    const engine = new FakePptxEngine(20, SLIDE_W_EMU, SLIDE_H_EMU);
    const onScaleChange = vi.fn();
    const v = makeBorrowedPptxScrollViewer(container as unknown as HTMLElement, {
      presentation: engine.asPres(),
      gap: 10,
      paddingTop: 0,
      paddingBottom: 0,
      paddingLeft: 0,
      paddingRight: 0,
      onScaleChange,
    });
    v.setScale(2); // pre-establishment: latched, not dropped
    expect(v.getScale()).toBe(2); // getScale reports the pending factor
    expect(onScaleChange).not.toHaveBeenCalled(); // fires at APPLICATION time
    // Container gains width ⇒ the next relayout establishes and applies the latch.
    container.clientWidth = 200;
    container.clientHeight = 400;
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientWidth = 200;
    scrollHost.clientHeight = 400;
    v.relayout();
    expect(v.scaleForTest()).toBeCloseTo(2, 6); // applied over the 1.0 base fit
    expect(v.getScale()).toBeCloseTo(2, 6);
    expect(onScaleChange).toHaveBeenCalledTimes(1);
    expect(onScaleChange).toHaveBeenCalledWith(2);
    // The base fit itself is still the true base (resize re-fit multiplier intact).
    expect(v.baseScaleForTest()).toBeCloseTo(1.0, 6);
    v.destroy();
  });

  it('a pre-establishment setScale latch is clamped to [zoomMin, zoomMax] (IX9 F1)', () => {
    installDom();
    const engine = new FakePptxEngine(3, SLIDE_W_EMU, SLIDE_H_EMU);
    const v = makeBorrowedPptxScrollViewer(makeContainer(0, 0) as unknown as HTMLElement, {
      presentation: engine.asPres(),
      zoomMin: 0.5,
      zoomMax: 3,
    });
    v.setScale(100);
    expect(v.getScale()).toBe(3); // latched pre-clamped
    v.destroy();
  });

  // Pointer-anchored ("zoom toward the cursor") Ctrl/⌘+wheel zoom: the content
  // point under the pointer before the zoom must be under the SAME viewport-y
  // after. getBoundingClientRect().top = 0 in the fake DOM, so clientY == the
  // scrollHost-viewport y.
  it('a Ctrl+wheel zoom keeps the content under the pointer fixed (zoom in)', () => {
    const { v, scrollHost } = setup();
    scrollHost.scrollTop = 600;
    const pointerY = 250;
    const before = v.contentAtViewportYForTest(pointerY);
    scrollHost.dispatch('wheel', {
      ctrlKey: true,
      deltaY: -60,
      clientX: 100,
      clientY: pointerY,
      preventDefault() {},
    });
    expect(v.getScale()).toBeGreaterThan(1);
    expect(v.viewportYOfForTest(before.slide, before.frac)).toBeCloseTo(pointerY, 4);
    v.destroy();
  });

  it('a Ctrl+wheel zoom keeps the content under the pointer fixed (zoom out)', () => {
    const { v, scrollHost } = setup();
    v.setScale(3);
    scrollHost.scrollTop = 500;
    const pointerY = 140;
    const before = v.contentAtViewportYForTest(pointerY);
    scrollHost.dispatch('wheel', {
      ctrlKey: true,
      deltaY: 60,
      clientX: 80,
      clientY: pointerY,
      preventDefault() {},
    });
    expect(v.getScale()).toBeLessThan(3);
    expect(v.viewportYOfForTest(before.slide, before.frac)).toBeCloseTo(pointerY, 4);
    v.destroy();
  });

  it('a non-gesture setScale still anchors on the viewport top (unchanged)', () => {
    const { v, scrollHost } = setup();
    scrollHost.scrollTop = 600;
    const topContent = v.contentAtViewportYForTest(0);
    v.setScale(3);
    expect(v.viewportYOfForTest(topContent.slide, topContent.frac)).toBeCloseTo(0, 4);
    v.destroy();
  });

  // HORIZONTAL pointer anchor with a non-zero left gutter. The gutter `padL` is
  // FIXED (does not scale), so the invariant is on the LOGICAL content-x under
  // the pointer: screen-x of content pixel c is `padL + c − scrollLeft`, hence
  // logicalX = (scrollLeft + x − padL) / scale must not move across the zoom.
  // Regression pin: subtracting/adding padL around the SCROLL as well (instead of
  // subtracting it from the anchor only) over-compensates by padL·(ratio−1).
  it('a Ctrl+wheel zoom keeps the content under the pointer fixed horizontally (padL > 0)', () => {
    const padL = 24;
    const { v, scrollHost } = setup({ paddingLeft: padL, paddingRight: padL });
    v.setScale(3); // slide 200 × 3 = 600px wide > the 200px viewport ⇒ h-scrollable
    // The fake DOM derives no layout from style; give the spacer a generous
    // laid-out width so the [0, maxLeft] clamp does not bind.
    const spacer = scrollHost.children[0] as FakeEl;
    spacer.offsetWidth = 100_000;
    scrollHost.scrollLeft = 120;
    const ax = 130; // pointer x in viewport px (over the slide, right of the gutter)
    const scaleBefore = v.getScale();
    const logicalXBefore = (scrollHost.scrollLeft + ax - padL) / scaleBefore;
    scrollHost.dispatch('wheel', {
      ctrlKey: true,
      deltaY: -20, // ratio e^0.2 ≈ 1.221 ⇒ 3 → ≈3.66, inside zoomMax 4 (unclamped)
      clientX: ax,
      clientY: 200,
      preventDefault() {},
    });
    const scaleAfter = v.getScale();
    expect(scaleAfter).toBeGreaterThan(scaleBefore);
    const logicalXAfter = (scrollHost.scrollLeft + ax - padL) / scaleAfter;
    expect(logicalXAfter).toBeCloseTo(logicalXBefore, 6);
    v.destroy();
  });

  // A wheel gesture whose setScale is a NO-OP (already clamped at zoomMax) must
  // NOT leak its pointer anchor into the next non-gesture setScale: the stepper
  // right after it still anchors on the viewport TOP.
  it('a no-op gesture (clamped at zoomMax) does not leak its anchor into the next setScale', () => {
    const { v, scrollHost } = setup();
    v.setScale(4); // pinned at zoomMax
    scrollHost.scrollTop = 600;
    // Ctrl+wheel zoom-IN at zoomMax ⇒ clamped to 4 ⇒ no-op; its anchor must be dropped.
    scrollHost.dispatch('wheel', {
      ctrlKey: true,
      deltaY: -60,
      clientX: 100,
      clientY: 250,
      preventDefault() {},
    });
    expect(v.getScale()).toBe(4); // confirmed no-op
    const topContent = v.contentAtViewportYForTest(0);
    v.zoomOut(); // non-gesture stepper — must keep the viewport-TOP anchor
    expect(v.viewportYOfForTest(topContent.slide, topContent.frac)).toBeCloseTo(0, 4);
    v.destroy();
  });
});
