import { describe, it, expect, afterEach, vi } from 'vitest';
import { DocxScrollViewer } from './scroll-viewer.js';
import { installDom, makeContainer, makeBorrowedDocxScrollViewer, FakeDocxEngine, type FakeEl } from './scroll-viewer-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

/**
 * IX9 — DocxScrollViewer's slice of the shared
 * {@link import('@silurus/ooxml-core').ZoomableViewer} contract. The viewer
 * already had an ABSOLUTE `setScale(scale)` (1 = 100%, the base fit ≠ 1),
 * Ctrl-wheel zoom, and a container-resize re-fit; IX9 keeps that verbatim and
 * layers on `getScale` / `zoomIn` / `zoomOut` / `fitWidth` / `fitPage` and the
 * `onScaleChange` notification. The absolute `_scale` IS the contract's
 * user-facing factor (a page draws at `widthPt × PT_TO_PX × _scale`), so the new
 * methods operate directly on it.
 */

/** A page 100pt × 200pt (natural CSS 133.33 × 266.67 px at 100%). */
const PAGE = { widthPt: 100, heightPt: 200 };

function setup(opts: Record<string, unknown> = {}, host = { w: 200, h: 400 }) {
  installDom();
  const container = makeContainer(host.w, host.h);
  const engine = new FakeDocxEngine(5, [PAGE]);
  const v = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, {
    document: engine.asDoc(),
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

describe('DocxScrollViewer IX9 zoom contract', () => {
  it('getScale() returns the absolute factor (the base fit after load)', () => {
    const { v } = setup();
    // base = 200 / (100 × 4/3) = 1.5
    expect(v.getScale()).toBeCloseTo(1.5, 5);
    expect(v.getScale()).toBeCloseTo(v.scaleForTest(), 10);
    v.destroy();
  });

  it('getScale() is 1 before a scale is established (no width yet)', () => {
    installDom();
    const engine = new FakeDocxEngine(3, [PAGE]);
    const v = makeBorrowedDocxScrollViewer(makeContainer(0, 0) as unknown as HTMLElement, {
      document: engine.asDoc(),
    });
    expect(v.getScale()).toBe(1);
    v.destroy();
  });

  it('setScale fires onScaleChange with the new factor on a change only', () => {
    const onScaleChange = vi.fn();
    const { v } = setup({ onScaleChange });
    v.setScale(2);
    expect(v.getScale()).toBeCloseTo(2, 5);
    expect(onScaleChange).toHaveBeenCalledTimes(1);
    expect(onScaleChange).toHaveBeenCalledWith(2);
    onScaleChange.mockClear();
    v.setScale(2); // unchanged
    expect(onScaleChange).not.toHaveBeenCalled();
    v.destroy();
  });

  it('zoomIn / zoomOut walk the shared ladder (off-base start snaps on)', () => {
    const { v } = setup();
    // Start at base 1.5 (an off-ladder value). First zoomIn snaps to 1.75.
    v.zoomIn();
    expect(v.getScale()).toBeCloseTo(1.75, 5);
    v.zoomIn();
    expect(v.getScale()).toBeCloseTo(2, 5);
    v.zoomOut();
    expect(v.getScale()).toBeCloseTo(1.75, 5);
    v.zoomOut();
    expect(v.getScale()).toBeCloseTo(1.5, 5); // ladder rung between 1.5 base
    v.destroy();
  });

  it('zoomOut returns to an initial width fit below zoomMin', () => {
    // Natural page width is 133.33px, so a 5px host establishes a 0.0375 fit.
    // That opening fit must remain reachable after crossing the zoom ladder.
    const { v } = setup({}, { w: 5, h: 400 });
    expect(v.getScale()).toBeCloseTo(0.0375, 6);

    v.zoomIn();
    expect(v.getScale()).toBeGreaterThan(0.0375);
    v.zoomOut();

    expect(v.getScale()).toBeCloseTo(0.0375, 6);
    v.destroy();
  });

  it('fitWidth restores the width-fit base after a zoom', () => {
    const { v } = setup();
    v.setScale(4); // zoom right in
    expect(v.getScale()).toBeCloseTo(4, 5);
    v.fitWidth();
    // Back to the width-fit base = 200 / (100 × 4/3) = 1.5.
    expect(v.getScale()).toBeCloseTo(1.5, 5);
    v.destroy();
  });

  it('fitWidth and the initial base use the widest page in the document', () => {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakeDocxEngine(2, [
      { widthPt: 100, heightPt: 200 },
      { widthPt: 100.05, heightPt: 200 },
    ]);
    const v = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      gap: 0,
      paddingTop: 0,
      paddingBottom: 0,
      paddingLeft: 0,
      paddingRight: 0,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientWidth = 200;
    scrollHost.clientHeight = 400;
    v.relayout();

    const widestFit = 200 / (100.05 * 4 / 3);
    expect(v.getScale()).toBeCloseTo(widestFit, 8);
    v.setScale(3);
    v.fitWidth();
    expect(v.getScale()).toBeCloseTo(widestFit, 8);
    v.destroy();
  });

  it('re-fits when progressive layout reveals a wider later page', () => {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakeDocxEngine(1, [
      { widthPt: 100, heightPt: 200 },
      { widthPt: 120, heightPt: 200 },
    ]);
    const v = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      gap: 0,
      paddingTop: 0,
      paddingBottom: 0,
      paddingLeft: 0,
      paddingRight: 0,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientWidth = 200;
    scrollHost.clientHeight = 400;
    v.relayout();
    expect(v.getScale()).toBeCloseTo(1.5, 8);

    engine.setPageCount(2);
    v.relayout();

    expect(v.getScale()).toBeCloseTo(1.25, 8);
    v.destroy();
  });

  it('fitPage takes the tighter of width/height fit', () => {
    // Container 200 wide × 200 tall. Natural page 133.33 × 266.67.
    // widthfit = 200/133.33 = 1.5; heightfit = 200/266.67 = 0.75 ⇒ page-fit 0.75.
    const { v } = setup({}, { w: 200, h: 200 });
    v.fitPage();
    expect(v.getScale()).toBeCloseTo(0.75, 4);
    v.destroy();
  });

  it('the wheel-zoom path also notifies through onScaleChange', () => {
    const onScaleChange = vi.fn();
    const { v, scrollHost } = setup({ onScaleChange });
    // Ctrl+wheel up (deltaY < 0 ⇒ zoom in). The handler routes through setScale.
    scrollHost.dispatch('wheel', { ctrlKey: true, deltaY: -50, preventDefault() {} });
    expect(onScaleChange).toHaveBeenCalledTimes(1);
    expect(onScaleChange.mock.calls[0][0]).toBeGreaterThan(1.5);
    v.destroy();
  });

  // IX9 F1 — family-unified pre-load setScale semantics: a setScale before the
  // layout establishes must be LATCHED and applied once it does (the
  // single-canvas viewers honour a pre-load setScale on their first render; the
  // scroll viewers used to silently drop it).
  it('setScale before load/layout is latched and applied once established (IX9 F1)', () => {
    installDom();
    const container = makeContainer(0, 0); // zero-width ⇒ fit deferred, nothing established
    const engine = new FakeDocxEngine(5, [PAGE]);
    const onScaleChange = vi.fn();
    const v = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
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
    expect(v.scaleForTest()).toBeCloseTo(2, 6); // applied over the 1.5 base fit
    expect(v.getScale()).toBeCloseTo(2, 6);
    expect(onScaleChange).toHaveBeenCalledTimes(1);
    expect(onScaleChange).toHaveBeenCalledWith(2);
    // The base fit itself is still the true base (resize re-fit multiplier intact).
    expect(v.baseScaleForTest()).toBeCloseTo(1.5, 5);
    v.destroy();
  });

  it('a pre-establishment setScale latch is clamped to [zoomMin, zoomMax] (IX9 F1)', () => {
    installDom();
    const engine = new FakeDocxEngine(5, [PAGE]);
    const v = makeBorrowedDocxScrollViewer(makeContainer(0, 0) as unknown as HTMLElement, {
      document: engine.asDoc(),
      zoomMin: 0.5,
      zoomMax: 3,
    });
    v.setScale(100);
    expect(v.getScale()).toBe(3); // latched pre-clamped
    v.destroy();
  });

  // Pointer-anchored ("zoom toward the cursor") Ctrl/⌘+wheel zoom: the content
  // point under the pointer before the zoom must be under the SAME viewport-y
  // after. The fake DOM reports getBoundingClientRect().top = 0, so the wheel
  // event's clientY is exactly the scrollHost-viewport y.
  it('a Ctrl+wheel zoom keeps the content under the pointer fixed (zoom in)', () => {
    const { v, scrollHost } = setup();
    scrollHost.scrollTop = 500; // scrolled into the document
    const pointerY = 250; // mid-viewport
    const before = v.contentAtViewportYForTest(pointerY);
    scrollHost.dispatch('wheel', {
      ctrlKey: true,
      deltaY: -60, // zoom in
      clientX: 100,
      clientY: pointerY,
      preventDefault() {},
    });
    expect(v.getScale()).toBeGreaterThan(1.5); // zoomed in past the base
    // The same content point (page + intra-page fraction) is back under pointerY.
    expect(v.viewportYOfForTest(before.page, before.frac)).toBeCloseTo(pointerY, 4);
    v.destroy();
  });

  it('a Ctrl+wheel zoom keeps the content under the pointer fixed (zoom out)', () => {
    const { v, scrollHost } = setup();
    v.setScale(3); // start zoomed in so there is room to zoom out
    scrollHost.scrollTop = 400;
    const pointerY = 120;
    const before = v.contentAtViewportYForTest(pointerY);
    scrollHost.dispatch('wheel', {
      ctrlKey: true,
      deltaY: 60, // zoom out
      clientX: 80,
      clientY: pointerY,
      preventDefault() {},
    });
    expect(v.getScale()).toBeLessThan(3);
    expect(v.viewportYOfForTest(before.page, before.frac)).toBeCloseTo(pointerY, 4);
    v.destroy();
  });

  it('a non-gesture setScale still anchors on the viewport top (unchanged)', () => {
    const { v, scrollHost } = setup();
    scrollHost.scrollTop = 500;
    const topContent = v.contentAtViewportYForTest(0); // what sits at the viewport top
    v.setScale(3); // public API, no pointer anchor
    // The viewport-top content is preserved (historical top-anchor behaviour).
    expect(v.viewportYOfForTest(topContent.page, topContent.frac)).toBeCloseTo(0, 4);
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
    v.setScale(3); // page 133.33 × 3 = 400px wide > the 200px viewport ⇒ h-scrollable
    // The fake DOM derives no layout from style; give the spacer a generous
    // laid-out width so the [0, maxLeft] clamp does not bind.
    const spacer = scrollHost.children[0] as FakeEl;
    spacer.offsetWidth = 100_000;
    scrollHost.scrollLeft = 120;
    const ax = 130; // pointer x in viewport px (over the page, right of the gutter)
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
    scrollHost.scrollTop = 500;
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
    expect(v.viewportYOfForTest(topContent.page, topContent.frac)).toBeCloseTo(0, 4);
    v.destroy();
  });
});
