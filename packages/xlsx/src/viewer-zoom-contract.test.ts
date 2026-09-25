import { describe, it, expect, afterEach, vi } from 'vitest';
import { XlsxViewer } from './viewer.js';
import { worksheetContentBounds } from './internal/worksheet-content-bounds.js';
import { installDom, makeContainer } from './viewer-destroy-test-dom.js';
import { HEADER_W, HEADER_H, colWidthToPx, rowHeightToPx, getMdwForWorksheet } from './renderer.js';
import type { Worksheet } from './types.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

/**
 * IX9 — the XlsxViewer's slice of the shared {@link import('@silurus/ooxml-core').ZoomableViewer}
 * contract. XlsxViewer already had `setScale` + a slider + Ctrl-wheel zoom; IX9
 * adds `getScale` / `zoomIn` / `zoomOut` / `fitWidth` / `fitPage` and the
 * `onScaleChange` notification, WITHOUT changing the existing slider / setScale
 * clamp-and-snap behaviour (non-regression is pinned below).
 *
 * The viewer is driven through its real API on the hand-rolled fake DOM (the repo
 * has no jsdom). A worksheet is injected into the private `currentWorksheet` field
 * — the same technique the hit-test suite uses — so scale + fit run against real
 * geometry. `renderCurrentSheet` early-returns on the fake DOM's 0-sized canvas,
 * so these tests observe scale STATE and the callback, not pixels.
 */

/** A small used range: 3 custom columns + defaults, 2 custom rows. */
function makeSheet(): Worksheet {
  return {
    name: 'Sheet1',
    rows: [{ index: 5, cells: [{ row: 5, col: 8, value: 'x' }] }],
    colWidths: { 1: 12, 2: 20, 3: 30 },
    rowHeights: { 1: 25, 2: 40 },
    defaultColWidth: 8.43,
    defaultRowHeight: 15,
    mergeCells: [],
    freezeRows: 0,
    freezeCols: 0,
    conditionalFormats: [],
    images: [],
    charts: [],
  } as unknown as Worksheet;
}

interface FakeScrollHost {
  scrollTop: number;
  scrollLeft: number;
  clientWidth: number;
  clientHeight: number;
  scrollWidth: number;
  scrollHeight: number;
}
interface Priv {
  currentWorksheet: Worksheet | null;
  canvasArea: { clientWidth: number; clientHeight: number; getBoundingClientRect(): DOMRect };
  scrollHost: FakeScrollHost;
  navPrev: { parentElement: { style: { width: string } } | null };
  tabBar: { style: { flexDirection: string } };
  tabStrip: {
    style: { marginLeft: string; marginRight: string };
    scrollLeft: number;
    clientWidth: number;
    scrollWidth: number;
  };
  tabList: { style: { flexDirection: string; minWidth: string; width: string } };
  tabs: Array<{ offsetLeft: number; offsetWidth: number }>;
  updateFooterDirection(): void;
  scrollTabs(direction: -1 | 1): void;
  _pendingZoomAnchor: { x: number; y: number } | null;
}

/** Unscaled (cs=1) frozen band extent of `ws` — mirrors getCellAt's inline
 *  loops. Used only to spell out the full logical-coordinate oracle below. */
function frozenExtent(ws: Worksheet): { frozenW: number; frozenH: number } {
  const mdw = getMdwForWorksheet(ws);
  let frozenH = 0;
  for (let r = 1; r <= (ws.freezeRows ?? 0); r++) {
    frozenH += rowHeightToPx(ws.rowHeights[r] ?? ws.defaultRowHeight);
  }
  let frozenW = 0;
  for (let c = 1; c <= (ws.freezeCols ?? 0); c++) {
    frozenW += colWidthToPx(ws.colWidths[c] ?? ws.defaultColWidth, mdw);
  }
  return { frozenW, frozenH };
}

/** Construct a viewer, inject `ws`, and size the canvas area so fit math has a
 *  laid-out container (the fake DOM defaults geometry to 0). */
function mount(
  ws: Worksheet | null,
  opts: ConstructorParameters<typeof XlsxViewer>[1] = {},
  container = { cw: 400, ch: 300 },
): { v: XlsxViewer; priv: Priv } {
  const v = new XlsxViewer(makeContainer() as unknown as HTMLElement, opts);
  const priv = v as unknown as Priv;
  priv.currentWorksheet = ws;
  priv.canvasArea.clientWidth = container.cw;
  priv.canvasArea.clientHeight = container.ch;
  return { v, priv };
}

/** Natural (cs=1) used-range extent oracle, mirroring the viewer's private
 *  `_naturalContentExtent` (header + used cols/rows, no scroll headroom). */
function naturalExtent(ws: Worksheet): { width: number; height: number } {
  const mdw = getMdwForWorksheet(ws);
  const { maxRow, maxCol } = worksheetContentBounds(ws);
  let width = HEADER_W;
  for (let c = 1; c <= maxCol; c++) width += colWidthToPx(ws.colWidths[c] ?? ws.defaultColWidth, mdw);
  let height = HEADER_H;
  for (let r = 1; r <= maxRow; r++) height += rowHeightToPx(ws.rowHeights[r] ?? ws.defaultRowHeight);
  return { width, height };
}

describe('XlsxViewer IX9 zoom contract', () => {
  it('extends the scrollable used range through cell-anchored drawings', () => {
    const ws = makeSheet();
    ws.charts = [{
      fromCol: 9, fromColOff: 0, fromRow: 95, fromRowOff: 0,
      toCol: 17, toColOff: 0, toRow: 113, toRowOff: 0,
      chart: {} as never,
    }];
    expect(worksheetContentBounds(ws)).toEqual({ maxRow: 114, maxCol: 26 });
  });

  it('getScale() is 1 (100%) by default and reflects the cellScale option', () => {
    installDom();
    expect(mount(makeSheet()).v.getScale()).toBe(1);
    expect(mount(makeSheet(), { cellScale: 1.5 }).v.getScale()).toBe(1.5);
  });

  it('setScale fires onScaleChange with the snapped factor exactly once', () => {
    installDom();
    const onScaleChange = vi.fn();
    const { v } = mount(makeSheet(), { onScaleChange });
    v.setScale(1.5);
    expect(v.getScale()).toBe(1.5);
    expect(onScaleChange).toHaveBeenCalledTimes(1);
    expect(onScaleChange).toHaveBeenCalledWith(1.5);
  });

  it('setScale does NOT fire onScaleChange when the scale is unchanged', () => {
    installDom();
    const onScaleChange = vi.fn();
    const { v } = mount(makeSheet(), { onScaleChange, cellScale: 2 });
    v.setScale(2); // already 200%
    expect(onScaleChange).not.toHaveBeenCalled();
  });

  it('setScale clamps to [zoomMin, zoomMax]', () => {
    installDom();
    const { v } = mount(makeSheet(), { zoomMin: 0.5, zoomMax: 2 });
    v.setScale(10);
    expect(v.getScale()).toBe(2);
    v.setScale(0.01);
    expect(v.getScale()).toBe(0.5);
  });

  it('keeps the footer tab-navigation width fixed at its 100% size', () => {
    installDom();
    const { v, priv } = mount(makeSheet(), { cellScale: 2 });

    // The footer is viewer chrome, not sheet content. Its two navigation
    // buttons therefore keep the row-header width at 100%, even when the
    // workbook is initially opened or subsequently rendered at another zoom.
    expect(priv.navPrev.parentElement?.style.width).toBe(`${HEADER_W}px`);
    v.setScale(0.5);
    expect(priv.navPrev.parentElement?.style.width).toBe(`${HEADER_W}px`);
  });

  it('mirrors footer control order for an RTL worksheet', () => {
    installDom();
    const rtlSheet = { ...makeSheet(), rightToLeft: true };
    const { priv } = mount(rtlSheet);

    priv.updateFooterDirection();

    expect(priv.tabBar.style.flexDirection).toBe('row-reverse');
    expect(priv.tabStrip.style.marginLeft).toBe('0');
    expect(priv.tabStrip.style.marginRight).toBe('1px');
    expect(priv.tabList.style.flexDirection).toBe('row-reverse');
    expect(priv.tabList.style.minWidth).toBe('100%');
    expect(priv.tabList.style.width).toBe('max-content');

    priv.currentWorksheet = makeSheet();
    priv.updateFooterDirection();

    expect(priv.tabBar.style.flexDirection).toBe('row');
    expect(priv.tabStrip.style.marginLeft).toBe('1px');
    expect(priv.tabStrip.style.marginRight).toBe('0');
    expect(priv.tabList.style.flexDirection).toBe('row');
  });

  it('scrolls to the nearest clipped tab by geometry regardless of visual order', () => {
    installDom();
    const { priv } = mount(makeSheet());
    priv.tabStrip.clientWidth = 100;
    priv.tabStrip.scrollWidth = 200;
    // Visual RTL order: the first logical tab has the greatest offset.
    priv.tabs = [
      { offsetLeft: 150, offsetWidth: 50 },
      { offsetLeft: 100, offsetWidth: 50 },
      { offsetLeft: 50, offsetWidth: 50 },
      { offsetLeft: 0, offsetWidth: 50 },
    ];

    priv.tabStrip.scrollLeft = 0;
    priv.scrollTabs(1);
    expect(priv.tabStrip.scrollLeft).toBe(50);

    priv.tabStrip.scrollLeft = 100;
    priv.scrollTabs(-1);
    expect(priv.tabStrip.scrollLeft).toBe(50);
  });

  it('zoomIn / zoomOut walk the shared ladder', () => {
    installDom();
    const { v } = mount(makeSheet());
    expect(v.getScale()).toBe(1);
    v.zoomIn();
    expect(v.getScale()).toBe(1.1); // 100% → next ladder rung
    v.zoomIn();
    expect(v.getScale()).toBe(1.25);
    v.zoomOut();
    expect(v.getScale()).toBe(1.1);
    v.zoomOut();
    expect(v.getScale()).toBe(1);
  });

  it('zoomIn from an off-ladder (wheel-zoomed) scale snaps onto the ladder', () => {
    installDom();
    const { v } = mount(makeSheet(), { cellScale: 1.03 });
    v.zoomIn();
    expect(v.getScale()).toBe(1.1);
  });

  it('fitWidth sets the scale that spans the used-range width in the container', () => {
    installDom();
    const ws = makeSheet();
    const { width } = naturalExtent(ws);
    const cw = 400;
    const { v } = mount(ws, {}, { cw, ch: 300 });
    v.fitWidth();
    // fitScale = cw / width, then snapped to whole percent by setScale.
    const expected = Math.round((cw / width) * 100) / 100;
    expect(v.getScale()).toBe(expected);
  });

  it('fitPage takes the tighter of the width/height fit', () => {
    installDom();
    const ws = makeSheet();
    const { width, height } = naturalExtent(ws);
    const cw = 400;
    const ch = 300;
    const { v } = mount(ws, {}, { cw, ch });
    v.fitPage();
    const raw = Math.min(cw / width, ch / height);
    const expected = Math.round(raw * 100) / 100;
    expect(v.getScale()).toBe(expected);
    // fitPage must never exceed fitWidth (height can only tighten).
    const wRaw = cw / width;
    expect(raw).toBeLessThanOrEqual(wRaw);
  });

  it('fitWidth is a no-op (defers) with no sheet or an unlaid-out container', () => {
    installDom();
    const onScaleChange = vi.fn();
    // No worksheet.
    const a = mount(null, { onScaleChange });
    a.v.fitWidth();
    expect(a.v.getScale()).toBe(1);
    // Zero-width container ⇒ fitScale returns 0 ⇒ defer.
    const b = mount(makeSheet(), { onScaleChange }, { cw: 0, ch: 0 });
    b.v.fitWidth();
    expect(b.v.getScale()).toBe(1);
    expect(onScaleChange).not.toHaveBeenCalled();
  });

  // IX9 F1 — family-unified pre-load setScale semantics (pinned across all five
  // viewers): a setScale before load is LATCHED and applied to the first render
  // (cellScale is read by every subsequent sheet render).
  it('setScale before load/layout is latched and applied once established (IX9 F1)', () => {
    installDom();
    // No worksheet loaded yet — setScale must latch (renderCurrentSheet no-ops).
    const v = new XlsxViewer(makeContainer() as unknown as HTMLElement, {});
    v.setScale(1.5);
    expect(v.getScale()).toBe(1.5); // latched; the first showSheet renders at it
  });

  it('a pre-load setScale latch is clamped to [zoomMin, zoomMax] (IX9 F1)', () => {
    installDom();
    const v = new XlsxViewer(makeContainer() as unknown as HTMLElement, { zoomMin: 0.5, zoomMax: 3 });
    v.setScale(100);
    expect(v.getScale()).toBe(3); // latched pre-clamped
  });

  // Pointer-anchored ("zoom toward the cursor") zoom, both axes, past the fixed
  // header + frozen band. The logical cell under the pointer is invariant across
  // the zoom. From getCellAt: logical content-Y under screen-y `py` is
  // `(py + scrollTop)/cs − (HEADER_H + frozenH)`; the x mirror-image holds via the
  // effective (start-anchored) scrollLeft. We drive setScale with an injected
  // gesture anchor and a generously-sized scroll host so the clamps do not bind.
  it('a gesture-anchored zoom keeps the logical cell under the pointer fixed', () => {
    installDom();
    const ws = makeSheet(); // no frozen panes ⇒ frozen extent 0
    const { v, priv } = mount(ws, { cellScale: 1 });
    // A roomy scroll host so neither axis clamps at an edge.
    priv.scrollHost.clientWidth = 400;
    priv.scrollHost.clientHeight = 300;
    priv.scrollHost.scrollWidth = 100000;
    priv.scrollHost.scrollHeight = 100000;
    priv.scrollHost.scrollTop = 800;
    priv.scrollHost.scrollLeft = 600;

    const { frozenW, frozenH } = frozenExtent(ws);
    const py = 180; // pointer screen-y within the grid
    const px = 260; // pointer screen-x
    const csOld = 1;
    // Logical content coordinate under the pointer BEFORE the zoom.
    const logicalYBefore = (py + priv.scrollHost.scrollTop) / csOld - (HEADER_H + frozenH);
    const logicalXBefore = (px + priv.scrollHost.scrollLeft) / csOld - (HEADER_W + frozenW);

    // Inject the gesture anchor exactly as the Ctrl/⌘+wheel handler would, then
    // zoom in via setScale (the same path the handler calls).
    priv._pendingZoomAnchor = { x: px, y: py };
    v.setScale(1.5);
    const csNew = v.getScale();
    expect(csNew).toBe(1.5);

    const logicalYAfter = (py + priv.scrollHost.scrollTop) / csNew - (HEADER_H + frozenH);
    const logicalXAfter = (px + priv.scrollHost.scrollLeft) / csNew - (HEADER_W + frozenW);
    expect(logicalYAfter).toBeCloseTo(logicalYBefore, 2);
    expect(logicalXAfter).toBeCloseTo(logicalXBefore, 2);
  });

  // FROZEN PANES: the header + frozen band is a SCALING lead-in (drawn at ×cs),
  // so it cancels out of the anchor equation — the raw pointer is the anchor and
  // the SAME logical-cell invariance holds with freezeRows/freezeCols > 0.
  it('a gesture-anchored zoom keeps the logical cell fixed with frozen panes', () => {
    installDom();
    const ws = makeSheet();
    (ws as { freezeRows?: number }).freezeRows = 2;
    (ws as { freezeCols?: number }).freezeCols = 1;
    const { v, priv } = mount(ws, { cellScale: 1 });
    priv.scrollHost.clientWidth = 400;
    priv.scrollHost.clientHeight = 300;
    priv.scrollHost.scrollWidth = 100000;
    priv.scrollHost.scrollHeight = 100000;
    priv.scrollHost.scrollTop = 800;
    priv.scrollHost.scrollLeft = 600;

    const { frozenW, frozenH } = frozenExtent(ws);
    expect(frozenH).toBeGreaterThan(0); // the frozen band is real in this fixture
    expect(frozenW).toBeGreaterThan(0);
    const py = 220; // pointer inside the scrollable region (below header + frozen)
    const px = 300;
    const csOld = 1;
    const logicalYBefore = (py + priv.scrollHost.scrollTop) / csOld - (HEADER_H + frozenH);
    const logicalXBefore = (px + priv.scrollHost.scrollLeft) / csOld - (HEADER_W + frozenW);

    priv._pendingZoomAnchor = { x: px, y: py };
    v.setScale(1.5);
    const csNew = v.getScale();
    expect(csNew).toBe(1.5);

    const logicalYAfter = (py + priv.scrollHost.scrollTop) / csNew - (HEADER_H + frozenH);
    const logicalXAfter = (px + priv.scrollHost.scrollLeft) / csNew - (HEADER_W + frozenW);
    expect(logicalYAfter).toBeCloseTo(logicalYBefore, 2);
    expect(logicalXAfter).toBeCloseTo(logicalXBefore, 2);
  });

  // Near the sheet START the pointer-pinning offset may legitimately fall BELOW
  // the scaled lead-in (header + frozen) — the clamp must be the native
  // [0, maxScroll], not a K·cs floor (regression pin for the virtual-scroll
  // detour, which clamped scrollTop up to K·cs_new near the top).
  it('a gesture zoom near the sheet start is not floored at the scaled lead-in', () => {
    installDom();
    const ws = makeSheet();
    const { v, priv } = mount(ws, { cellScale: 1 });
    priv.scrollHost.clientWidth = 400;
    priv.scrollHost.clientHeight = 300;
    priv.scrollHost.scrollWidth = 100000;
    priv.scrollHost.scrollHeight = 100000;
    priv.scrollHost.scrollTop = 0; // at the very top
    priv.scrollHost.scrollLeft = 0;

    const py = 100;
    priv._pendingZoomAnchor = { x: 50, y: py };
    v.setScale(1.2);
    // Exact pointer-pinned offset: ratio·(scrollTop + py) − py = 1.2·100 − 100.
    expect(priv.scrollHost.scrollTop).toBeCloseTo(20, 6);
    // The K·cs_new floor would have been HEADER_H·1.2 = 24 (> 20).
    expect(priv.scrollHost.scrollTop).toBeLessThan(HEADER_H * 1.2);
  });

  it('a non-gesture setScale preserves the START-anchored top-left (unchanged)', () => {
    installDom();
    const ws = makeSheet();
    const { v, priv } = mount(ws, { cellScale: 1 });
    priv.scrollHost.clientWidth = 400;
    priv.scrollHost.clientHeight = 300;
    priv.scrollHost.scrollWidth = 100000;
    priv.scrollHost.scrollHeight = 100000;
    priv.scrollHost.scrollLeft = 600;
    // No _pendingZoomAnchor ⇒ the historical start-anchored branch runs; the
    // effective (LTR) scroll position is preserved verbatim (scrollLeft unchanged
    // for an LTR sheet).
    v.setScale(1.5);
    expect(priv.scrollHost.scrollLeft).toBe(600);
  });

  // A gesture whose setScale is a NO-OP (the whole-percent snap swallows a small
  // deltaY, or the scale is pinned at zoomMin/zoomMax) must NOT leak its pointer
  // anchor into the next non-gesture setScale: the stepper right after it still
  // preserves the START-anchored (top-left) position.
  it('a no-op gesture (percent-snap) does not leak its anchor into the next setScale', () => {
    installDom();
    const ws = makeSheet();
    const { v, priv } = mount(ws, { cellScale: 1 });
    priv.scrollHost.clientWidth = 400;
    priv.scrollHost.clientHeight = 300;
    priv.scrollHost.scrollWidth = 100000;
    priv.scrollHost.scrollHeight = 100000;
    priv.scrollHost.scrollLeft = 600;
    // Inject the anchor exactly as the wheel handler would, then a setScale that
    // snaps to the SAME whole percent (1.004 → 100%) ⇒ no-op; anchor must drop.
    priv._pendingZoomAnchor = { x: 260, y: 180 };
    v.setScale(1.004);
    expect(v.getScale()).toBe(1); // confirmed no-op
    v.zoomIn(); // non-gesture stepper (1 → 1.1) — must run the START-anchored branch
    expect(v.getScale()).toBe(1.1);
    expect(priv.scrollHost.scrollLeft).toBe(600); // unchanged for an LTR sheet
  });
});

/**
 * Issue #842 — the viewer's OWN +/- chrome must step exactly like the IX9
 * contract methods: along the shared ladder (`ZOOM_STEP_LADDER` via
 * `nextZoomStep`/`prevZoomStep`), not the pre-IX9 linear ±0.1. A host wiring
 * its own buttons to `zoomIn`/`zoomOut` and a user clicking the built-in
 * buttons now land on the same scales.
 */
describe('XlsxViewer built-in +/- buttons follow the shared ladder (issue #842)', () => {
  function mountWithButtons(opts: ConstructorParameters<typeof XlsxViewer>[1] = {}) {
    installDom();
    const container = makeContainer();
    const v = new XlsxViewer(container as unknown as HTMLElement, opts);
    const priv = v as unknown as Priv;
    priv.currentWorksheet = makeSheet();
    priv.canvasArea.clientWidth = 400;
    priv.canvasArea.clientHeight = 300;
    const minus = container.querySelector('button[aria-label="Zoom out"]');
    const plus = container.querySelector('button[aria-label="Zoom in"]');
    if (!minus || !plus) throw new Error('built-in zoom buttons not found');
    return { v, minus, plus };
  }

  it('the + button steps to the next ladder rung, not +0.1', () => {
    const { v, plus } = mountWithButtons({ cellScale: 1.25 });
    plus.dispatch('click');
    expect(v.getScale()).toBe(1.5); // linear would give 1.35
    plus.dispatch('click');
    expect(v.getScale()).toBe(1.75);
  });

  it('the − button steps to the previous ladder rung, not −0.1', () => {
    const { v, minus } = mountWithButtons();
    minus.dispatch('click');
    expect(v.getScale()).toBe(0.9);
    minus.dispatch('click');
    expect(v.getScale()).toBe(0.75); // linear would give 0.8
  });

  it('an off-ladder (wheel-zoomed) scale snaps onto the ladder', () => {
    const { v, plus } = mountWithButtons({ cellScale: 1.03 });
    plus.dispatch('click');
    expect(v.getScale()).toBe(1.1); // linear would give 1.13
  });

  it('the − button holds at the bottom rung instead of sliding below it', () => {
    const { v, minus } = mountWithButtons({ cellScale: 0.25 });
    minus.dispatch('click');
    expect(v.getScale()).toBe(0.25); // linear would give 0.15
  });

  it('button steps and contract zoomIn land on identical scales', () => {
    const a = mountWithButtons({ cellScale: 1.25 });
    const b = mountWithButtons({ cellScale: 1.25 });
    a.plus.dispatch('click');
    b.v.zoomIn();
    expect(a.v.getScale()).toBe(b.v.getScale());
  });

  it('uses non-submit buttons so a surrounding form is not submitted (issue #1082)', () => {
    const { minus, plus } = mountWithButtons();
    expect(minus.type).toBe('button');
    expect(plus.type).toBe('button');
  });
});

/**
 * Slider interaction and non-regression coverage for the pre-IX9 position ↔
 * scale mapping (PR #315's "100% dead-center piecewise-linear" behaviour).
 */
describe('XlsxViewer zoom slider', () => {
  it('slider position 50 maps to 100% for any bounds', () => {
    installDom();
    const { v } = mount(makeSheet());
    const priv = v as unknown as {
      zoomPosToScale(p: number, min: number, max: number): number;
      zoomScaleToPos(s: number, min: number, max: number): number;
    };
    expect(priv.zoomPosToScale(50, 0.1, 4)).toBeCloseTo(1, 10);
    expect(priv.zoomScaleToPos(1, 0.1, 4)).toBeCloseTo(50, 10);
    // Each half is its own linear segment.
    expect(priv.zoomPosToScale(0, 0.1, 4)).toBeCloseTo(0.1, 10);
    expect(priv.zoomPosToScale(100, 0.1, 4)).toBeCloseTo(4, 10);
  });

  it('magnetically snaps the thumb to 100% while dragged near the center', () => {
    installDom();
    const container = makeContainer();
    const v = new XlsxViewer(container as unknown as HTMLElement, { cellScale: 0.5 });
    const slider = container.querySelector('input[aria-label="Zoom"]');
    if (!slider) throw new Error('built-in zoom slider not found');

    // The attraction zone is expressed in slider-position units so it feels
    // equally wide on both sides of the asymmetric 10%→100%→400% mapping.
    slider.value = '48';
    slider.dispatch('input');
    expect(v.getScale()).toBe(1);
    expect(slider.value).toBe('50');

    v.setScale(1.5);
    slider.value = '52';
    slider.dispatch('input');
    expect(v.getScale()).toBe(1);
    expect(slider.value).toBe('50');
  });

  it('keeps continuous slider zoom outside the 100% attraction zone', () => {
    installDom();
    const container = makeContainer();
    const v = new XlsxViewer(container as unknown as HTMLElement, { cellScale: 0.5 });
    const slider = container.querySelector('input[aria-label="Zoom"]');
    if (!slider) throw new Error('built-in zoom slider not found');

    slider.value = '47';
    slider.dispatch('input');
    expect(v.getScale()).toBe(0.95);

    slider.value = '53';
    slider.dispatch('input');
    expect(v.getScale()).toBe(1.18);
  });
});
