import { afterEach, describe, expect, it, vi } from 'vitest';
import { DocxScrollViewer } from './scroll-viewer.js';
import { publishDocxLayout } from './document-layout-events.js';
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

function scrollHostOf(container: FakeEl): FakeEl {
  return container.children[0].children[0];
}

async function settlePromises(): Promise<void> {
  await Promise.resolve();
  await Promise.resolve();
}

describe('DocxScrollViewer — growing page count', () => {
  it('admits the first non-empty prefix immediately', () => {
    installDom();
    const container = makeContainer(700, 500);
    const engine = new FakeDocxEngine(0, PAGE);
    const doc = engine.asDoc();
    const viewer = DocxScrollViewer.fromDocument(
      container as unknown as HTMLElement,
      doc,
    );
    expect(parseFloat(spacerOf(container).style.height || '0')).toBe(0);

    engine.setLayoutComplete(false);
    engine.setPageCount(3);
    publishDocxLayout(doc, { pageCount: 3, exact: false, complete: false });

    expect(parseFloat(spacerOf(container).style.height)).toBeGreaterThan(0);
    expect(scrollHostOf(container).children.some((child) =>
      child.children.some((nested) => nested.tag === 'canvas'))).toBe(true);
    viewer.destroy();
  });

  it('publishes background page growth without waiting for reader scroll', () => {
    installDom();
    const container = makeContainer(700, 500);
    const engine = new FakeDocxEngine(4, PAGE);
    const doc = engine.asDoc();
    const fires: Array<[number, number, boolean]> = [];
    const viewer = DocxScrollViewer.fromDocument(
      container as unknown as HTMLElement,
      doc,
      { onVisiblePageChange: (top, total, complete) => fires.push([top, total, complete]) },
    );
    const initialHeight = parseFloat(spacerOf(container).style.height);

    engine.setLayoutComplete(false);
    engine.setPageCount(8);
    publishDocxLayout(doc, { pageCount: 8, exact: false, complete: false });
    engine.setPageCount(16);
    publishDocxLayout(doc, { pageCount: 16, exact: false, complete: false });

    // The document, callback and native scroll extent all expose the newest
    // paintable prefix while the reader remains idle at the top.
    expect(viewer.pageCount).toBe(16);
    expect(parseFloat(spacerOf(container).style.height)).toBeGreaterThan(initialHeight);
    expect(fires.at(-1)).toEqual([0, 16, false]);
    viewer.destroy();
  });

  it('reveals a coalesced prefix when programmatic navigation targets it', () => {
    installDom();
    const container = makeContainer(700, 500);
    const engine = new FakeDocxEngine(4, PAGE);
    const doc = engine.asDoc();
    const viewer = DocxScrollViewer.fromDocument(
      container as unknown as HTMLElement,
      doc,
    );
    const initialHeight = parseFloat(spacerOf(container).style.height);

    engine.setLayoutComplete(false);
    engine.setPageCount(16);
    publishDocxLayout(doc, { pageCount: 16, exact: false, complete: false });
    viewer.scrollToPage(12);

    expect(parseFloat(spacerOf(container).style.height)).toBeGreaterThan(initialHeight);
    expect(viewer.topVisiblePage).toBe(12);
    viewer.destroy();
  });

  it('reveals growth immediately when the reader is already at the presented tail', () => {
    installDom();
    const container = makeContainer(700, 500);
    const engine = new FakeDocxEngine(4, PAGE);
    const doc = engine.asDoc();
    const viewer = DocxScrollViewer.fromDocument(
      container as unknown as HTMLElement,
      doc,
    );
    const initialHeight = parseFloat(spacerOf(container).style.height);
    const scrollHost = scrollHostOf(container);
    scrollHost.scrollTop = initialHeight - scrollHost.clientHeight;
    scrollHost.dispatch('scroll');

    engine.setLayoutComplete(false);
    engine.setPageCount(12);
    publishDocxLayout(doc, { pageCount: 12, exact: false, complete: false });

    // A wheel gesture at a native scroll maximum may not emit `scroll`, so the
    // publication itself must open the next prefix for a reader already there.
    expect(parseFloat(spacerOf(container).style.height)).toBeGreaterThan(initialHeight);
    viewer.destroy();
  });

  it('keeps the painted canvas visible until an authoritative main-mode refresh is ready', async () => {
    installDom();
    const container = makeContainer(700, 500);
    const engine = new FakeDocxEngine(2, PAGE, 'main', true);
    const doc = engine.asDoc();
    const viewer = DocxScrollViewer.fromDocument(
      container as unknown as HTMLElement,
      doc,
    );
    for (const call of [...engine.renderCalls]) call.resolve();
    await settlePromises();

    const scrollHost = scrollHostOf(container);
    const firstWrapper = scrollHost.children.find((child) =>
      child.children.some((nested) => nested.tag === 'canvas'))!;
    const oldCanvas = firstWrapper.children.find((child) => child.tag === 'canvas')!;
    const oldResizeCount = oldCanvas._deviceResizes.length;
    const initialHeight = parseFloat(spacerOf(container).style.height);
    engine.renderCalls.length = 0;

    engine.setPageCount(6);
    engine.setLayoutComplete(true);
    publishDocxLayout(doc, { pageCount: 6, exact: true, complete: true });

    const refresh = engine.renderCalls.find((call) => call.page === 0)!;
    expect(parseFloat(spacerOf(container).style.height)).toBeGreaterThan(initialHeight);
    expect(refresh.canvas).not.toBe(oldCanvas);
    expect(firstWrapper.children).toContain(oldCanvas);
    expect(oldCanvas._deviceResizes).toHaveLength(oldResizeCount);

    refresh.resolve();
    await settlePromises();
    expect(firstWrapper.children).not.toContain(oldCanvas);
    expect(firstWrapper.children).toContain(refresh.canvas);
    viewer.destroy();
  });

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

  it('repaints pages atomically when the layout underneath them is replaced', () => {
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
    engine.setPageCount(80);
    publishDocxLayout(engine.asDoc(), { pageCount: 80, exact: true, complete: true });
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
    await expect(viewer.waitUntilLayoutComplete()).resolves.toBeUndefined();
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

  it('re-fires when layout completes without changing the page count', () => {
    installDom();
    const fires: Array<[number, number, boolean]> = [];
    const engine = new FakeDocxEngine(4, PAGE);
    engine.setLayoutComplete(false);
    const doc = engine.asDoc();
    const viewer = DocxScrollViewer.fromDocument(
      makeContainer(700, 500) as unknown as HTMLElement,
      doc,
      { onVisiblePageChange: (top, total, complete) => fires.push([top, total, complete]) },
    );
    expect(fires).toEqual([[0, 4, false]]);

    engine.setLayoutComplete(true);
    publishDocxLayout(doc, { pageCount: 4, exact: true, complete: true });
    expect(fires).toEqual([[0, 4, false], [0, 4, true]]);
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
      { showTrackedChanges: true, currentDate: 0 },
    ]);
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

    // The toggle shortens the document under the reader.
    engine.setPageCount(3);
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

  it('does not start a deferred full-document find after clearFind()', async () => {
    installDom();
    const engine = new FakeDocxEngine(1, PAGE);
    engine.setLayoutComplete(false);
    const doc = engine.asDoc();
    const viewer = DocxScrollViewer.fromDocument(
      makeContainer(700, 500) as unknown as HTMLElement,
      doc,
    ) as DocxScrollViewer;
    const findController = (viewer as unknown as { _find: { find: unknown } })._find;
    const findSpy = vi.spyOn(findController as { find: (query: string) => Promise<unknown[]> }, 'find');

    const pending = viewer.findText('later');
    viewer.clearFind();
    engine.setPageCount(3);
    engine.setLayoutComplete(true);
    publishDocxLayout(doc, { pageCount: 3, exact: true, complete: true });

    await expect(pending).resolves.toEqual([]);
    expect(findSpy).not.toHaveBeenCalled();
    viewer.destroy();
  });
});
