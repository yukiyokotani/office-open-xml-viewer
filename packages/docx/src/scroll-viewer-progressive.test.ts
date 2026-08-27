import { afterEach, describe, expect, it, vi } from 'vitest';
import { DocxScrollViewer } from './scroll-viewer.js';
import {
  FakeDocxEngine,
  installDom,
  makeContainer,
  type FakeEl,
} from './scroll-viewer-test-dom.js';

// ─────────────────────────────────────────────────────────────────────────────
// Progressive layout hands the viewer a document whose page count GROWS: it
// mounts the provisional opening pages, then relays out when the authoritative
// layout lands. The virtualization math already takes the heights array fresh
// on every pass, so what needs pinning is the viewer's side of that contract —
// the scroll extent tracks the new page count, the mounted window is unchanged
// for pages the user is already looking at, and scroll position survives.
// ─────────────────────────────────────────────────────────────────────────────

const PAGE = [{ widthPt: 612, heightPt: 792 }];

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

function spacerOf(container: FakeEl): FakeEl {
  return container.children[0].children[0].children[0];
}

describe('DocxScrollViewer — growing page count', () => {
  it('extends the scroll region when layout completes', () => {
    installDom();
    const container = makeContainer(700, 500);
    // Two provisional pages, as a preview publishes.
    const engine = new FakeDocxEngine(2, PAGE);
    const viewer = DocxScrollViewer.fromDocument(
      container as unknown as HTMLElement,
      engine.asDoc(),
    );
    const provisionalHeight = parseFloat(spacerOf(container).style.height);
    expect(viewer.pageCount).toBe(2);
    expect(provisionalHeight).toBeGreaterThan(0);

    // The authoritative layout arrives.
    engine.setPageCount(80);
    viewer.relayout();

    expect(viewer.pageCount).toBe(80);
    const finalHeight = parseFloat(spacerOf(container).style.height);
    expect(finalHeight).toBeGreaterThan(provisionalHeight);
    viewer.destroy();
  });

  it('keeps the pages already on screen mounted across the handover', () => {
    installDom();
    const container = makeContainer(700, 500);
    const engine = new FakeDocxEngine(2, PAGE);
    const viewer = DocxScrollViewer.fromDocument(
      container as unknown as HTMLElement,
      engine.asDoc(),
    );
    const mountedBefore = viewer.topVisiblePage;
    expect(mountedBefore).toBe(0);

    engine.setPageCount(80);
    viewer.relayout();

    // Growing the document must not scroll the user somewhere else.
    expect(viewer.topVisiblePage).toBe(mountedBefore);
    viewer.destroy();
  });

  it('repaints pages in place when the layout underneath them is replaced', () => {
    installDom();
    const container = makeContainer(700, 500);
    const engine = new FakeDocxEngine(2, PAGE);
    const viewer = DocxScrollViewer.fromDocument(
      container as unknown as HTMLElement,
      engine.asDoc(),
    );
    const paintedProvisionally = engine.renderCalls.length;
    expect(paintedProvisionally).toBeGreaterThan(0);

    // A plain relayout must NOT repaint: that guard is what keeps scrolling
    // cheap.
    viewer.relayout();
    expect(engine.renderCalls.length).toBe(paintedProvisionally);

    // Replacing the layout must, because a page's content can change without
    // its index changing (a footer's PAGE/NUMPAGES total, for one).
    (viewer as unknown as { _invalidateRenderedSlots(): void })._invalidateRenderedSlots();
    engine.setPageCount(80);
    viewer.relayout();
    expect(engine.renderCalls.length).toBeGreaterThan(paintedProvisionally);
    viewer.destroy();
  });

  it('reports a borrowed engine as fully laid out', async () => {
    // fromDocument borrows an already-loaded document, and an engine injected by
    // an integrator may predate these members entirely; neither may throw.
    installDom();
    const engine = new FakeDocxEngine(3, PAGE);
    const viewer = DocxScrollViewer.fromDocument(
      makeContainer(700, 500) as unknown as HTMLElement,
      engine.asDoc(),
    );
    expect(viewer.layoutComplete).toBe(true);
    await expect(viewer.whenLayoutComplete()).resolves.toBeUndefined();
    viewer.destroy();
  });

  it('re-fires onVisiblePageChange when the total grows without the index moving', () => {
    installDom();
    const fires: Array<[number, number, boolean]> = [];
    const engine = new FakeDocxEngine(2, PAGE);
    const viewer = DocxScrollViewer.fromDocument(
      makeContainer(700, 500) as unknown as HTMLElement,
      engine.asDoc(),
      { onVisiblePageChange: (top, total, complete) => { fires.push([top, total, complete]); } },
    );
    expect(fires).toEqual([[0, 2, true]]);

    // The user has not scrolled — topIndex is still 0 — but the document grew.
    // An index-only latch would strand the indicator on the preview count.
    engine.setPageCount(80);
    viewer.relayout();
    expect(fires).toEqual([[0, 2, true], [0, 80, true]]);
    viewer.destroy();
  });

  it('does not fire when neither the index nor the total changed', () => {
    installDom();
    const fires: Array<[number, number]> = [];
    const engine = new FakeDocxEngine(4, PAGE);
    const viewer = DocxScrollViewer.fromDocument(
      makeContainer(700, 500) as unknown as HTMLElement,
      engine.asDoc(),
      { onVisiblePageChange: (top, total) => { fires.push([top, total]); } },
    );
    const initial = fires.length;
    viewer.relayout();
    viewer.relayout();
    expect(fires.length).toBe(initial);
    viewer.destroy();
  });

  it('re-fires when the document shrinks', () => {
    // A tracked-changes view can have FEWER pages than the final view, so the
    // count moves down as well as up.
    installDom();
    const fires: Array<[number, number]> = [];
    const engine = new FakeDocxEngine(40, PAGE);
    const viewer = DocxScrollViewer.fromDocument(
      makeContainer(700, 500) as unknown as HTMLElement,
      engine.asDoc(),
      { onVisiblePageChange: (top, total) => { fires.push([top, total]); } },
    );
    fires.length = 0;
    engine.setPageCount(3);
    viewer.relayout();
    expect(fires).toEqual([[0, 3]]);
    viewer.destroy();
  });

  it('moves the document to the markup variant when tracked changes are toggled', async () => {
    installDom();
    const engine = new FakeDocxEngine(20, PAGE);
    const viewer = DocxScrollViewer.fromDocument(
      makeContainer(700, 500) as unknown as HTMLElement,
      engine.asDoc(),
    );
    expect(engine.layoutViews).toEqual([]);

    await viewer.setShowTrackedChanges(true);
    expect(engine.layoutViews).toEqual([
      { showTrackedChanges: true, currentDate: undefined },
    ]);
    viewer.destroy();
  });

  it('reads no geometry from the markup variant until the document installs it', async () => {
    // Worker mode builds the variant in the worker, so the switch is a
    // round-trip. While it is in flight the document still answers for the
    // variant on screen, and the viewer must keep measuring against THAT —
    // adopting the new view early is the paint/geometry split all over again.
    installDom();
    const container = makeContainer(700, 500);
    const engine = new FakeDocxEngine(20, PAGE, 'worker');
    engine.setVariantPageCounts(20, 4);
    engine.deferLayoutViews();
    const viewer = DocxScrollViewer.fromDocument(
      container as unknown as HTMLElement,
      engine.asDoc(),
    );
    const finalHeight = parseFloat(spacerOf(container).style.height);
    engine.bitmapCalls.length = 0;

    const switching = viewer.setShowTrackedChanges(true);
    await Promise.resolve();
    // Nothing has moved: no repaint at the new variant, no new scroll extent.
    expect(engine.bitmapCalls).toHaveLength(0);
    expect(viewer.pageCount).toBe(20);
    expect(parseFloat(spacerOf(container).style.height)).toBe(finalHeight);

    // And a repaint forced WHILE the switch is in flight (a scroll, a resize,
    // a progressive publication) must still name the variant the document's
    // geometry describes — painting markup against the final view's page count
    // is precisely the split this switch exists to avoid.
    (viewer as unknown as { _invalidateRenderedSlots(): void })._invalidateRenderedSlots();
    viewer.relayout();
    expect(engine.bitmapCalls.length).toBeGreaterThan(0);
    for (const call of engine.bitmapCalls) {
      expect(call.showTrackedChanges ?? false).toBe(false);
    }
    expect(viewer.pageCount).toBe(20);
    engine.bitmapCalls.length = 0;

    expect(engine.pendingLayoutViews).toHaveLength(1);
    engine.pendingLayoutViews[0]!();
    await switching;

    expect(viewer.pageCount).toBe(4);
    expect(parseFloat(spacerOf(container).style.height)).toBeLessThan(finalHeight);
    expect(engine.bitmapCalls.length).toBeGreaterThan(0);
    for (const call of engine.bitmapCalls) expect(call.showTrackedChanges).toBe(true);
    viewer.destroy();
  });

  it('lands on the last requested view when the toggle is flipped twice mid-switch', async () => {
    // The viewer's own flag lags the request in worker mode, so a toggle back
    // has to compare against what was ASKED for. Comparing against what is
    // painted makes the second call a no-op and the first one wins.
    installDom();
    const engine = new FakeDocxEngine(20, PAGE, 'worker');
    engine.setVariantPageCounts(20, 4);
    engine.deferLayoutViews();
    const viewer = DocxScrollViewer.fromDocument(
      makeContainer(700, 500) as unknown as HTMLElement,
      engine.asDoc(),
    );

    const toMarkup = viewer.setShowTrackedChanges(true);
    const backToFinal = viewer.setShowTrackedChanges(false);
    expect(engine.layoutViews.map((view) => view.showTrackedChanges)).toEqual([true, false]);

    for (const release of [...engine.pendingLayoutViews]) release();
    await Promise.all([toMarkup, backToFinal]);

    expect(viewer.pageCount).toBe(20);
    engine.bitmapCalls.length = 0;
    viewer.relayout();
    (viewer as unknown as { _invalidateRenderedSlots(): void })._invalidateRenderedSlots();
    viewer.relayout();
    for (const call of engine.bitmapCalls) {
      expect(call.showTrackedChanges ?? false).toBe(false);
    }
    viewer.destroy();
  });

  it('recycles out-of-range slots when the markup variant is shorter', async () => {
    // Hiding vs showing deletions changes the page count. Toggling to a SHORTER
    // variant while scrolled deep used to leave slots asking for pages that no
    // longer exist, which surfaced as a RangeError and a blank page.
    installDom();
    const engine = new FakeDocxEngine(60, PAGE);
    const container = makeContainer(700, 500);
    const viewer = DocxScrollViewer.fromDocument(
      container as unknown as HTMLElement,
      engine.asDoc(),
    );
    viewer.scrollToPage(50);
    expect(viewer.topVisiblePage).toBeGreaterThan(0);

    // The toggle shortens the document under the reader — the variant switch
    // itself is what repaginates it, exactly as a real document does.
    engine.setVariantPageCounts(60, 3);
    engine.renderCalls.length = 0;
    await viewer.setShowTrackedChanges(true);

    expect(viewer.pageCount).toBe(3);
    // Every page requested AFTER the toggle must exist in the shorter variant.
    expect(engine.renderCalls.length).toBeGreaterThan(0);
    for (const call of engine.renderCalls) {
      expect(call.page).toBeLessThan(3);
    }
    viewer.destroy();
  });

  it('does not mount the whole document just because it grew', () => {
    installDom();
    const container = makeContainer(700, 500);
    const engine = new FakeDocxEngine(2, PAGE);
    const viewer = DocxScrollViewer.fromDocument(
      container as unknown as HTMLElement,
      engine.asDoc(),
      { overscan: 1 },
    );
    engine.setPageCount(400);
    viewer.relayout();

    const scrollHost = container.children[0].children[0];
    const canvases = scrollHost.children.filter(
      (child: FakeEl) => child.children.some((nested: FakeEl) => nested.tag === 'canvas'),
    );
    // Virtualization still applies: a 400-page document mounts a viewport's
    // worth of slots, not 400.
    expect(canvases.length).toBeGreaterThan(0);
    expect(canvases.length).toBeLessThan(10);
    viewer.destroy();
  });
});
