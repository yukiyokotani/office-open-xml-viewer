import { describe, it, expect, afterEach, vi } from 'vitest';
import { DocxScrollViewer } from './scroll-viewer.js';
import { DocxDocument } from './document.js';
import { installDom, makeContainer, makeEl, makeBorrowedDocxScrollViewer, FakeDocxEngine, type FakeEl } from './scroll-viewer-test-dom.js';
import * as docxIndex from './index.js';
import type { DocxElementContext } from './selection-context.js';
import type { CommentAnchorRange } from './comments.js';
import type { DocxTextRunInfo } from './renderer.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

async function waitForBuiltInCommentUi(viewer: unknown): Promise<void> {
  await vi.waitFor(() => {
    expect((viewer as { _commentUi: unknown })._commentUi).not.toBeNull();
  });
}

describe('DocxScrollViewer — skeleton (T1)', () => {
  it('builds the wrapper → scrollHost → spacer DOM inside the container', () => {
    installDom();
    const container = makeContainer();
    const engine = new FakeDocxEngine(3, [{ widthPt: 612, heightPt: 792 }]);
    const v = DocxScrollViewer.fromDocument(container as unknown as HTMLElement, engine.asDoc());
    // container → wrapper
    const wrapper = container.children[0];
    expect(wrapper.tag).toBe('div');
    expect(wrapper.style.position).toBe('relative');
    // wrapper → scrollHost
    const scrollHost = wrapper.children[0];
    expect(scrollHost.style.overflow).toBe('auto');
    expect(scrollHost.style.scrollbarGutter).toBe('stable');
    // scrollHost → spacer
    const spacer = scrollHost.children[0];
    expect(spacer.style.position).toBe('absolute');
    v.destroy();
  });

  it('exposes pageCount from the borrowed engine', () => {
    installDom();
    const engine = new FakeDocxEngine(5, [{ widthPt: 612, heightPt: 792 }]);
    const v = DocxScrollViewer.fromDocument(
      makeContainer() as unknown as HTMLElement,
      engine.asDoc(),
    );
    expect(v.pageCount).toBe(5);
    v.destroy();
  });

  it('load() is unsupported when an engine is borrowed', async () => {
    installDom();
    const engine = new FakeDocxEngine(1, [{ widthPt: 612, heightPt: 792 }]);
    const v = DocxScrollViewer.fromDocument(
      makeContainer() as unknown as HTMLElement,
      engine.asDoc(),
    );
    await expect((v as DocxScrollViewer).load('x.docx')).rejects.toThrow(/fromDocument/i);
    v.destroy();
  });

  it('destroy() removes the DOM and does NOT destroy an borrowed engine', () => {
    installDom();
    const container = makeContainer();
    const engine = new FakeDocxEngine(1, [{ widthPt: 612, heightPt: 792 }]);
    const v = DocxScrollViewer.fromDocument(container as unknown as HTMLElement, engine.asDoc());
    expect(container.children.length).toBe(1); // wrapper mounted
    v.destroy();
    expect(container.children.length).toBe(0); // wrapper removed
    expect(engine.destroyed).toBe(false); // borrowed engine preserved (caller owns it)
  });

  it('pageCount is 0 before load resolves (no borrowed engine)', () => {
    installDom();
    const v = new DocxScrollViewer(makeContainer() as unknown as HTMLElement, {});
    expect(v.pageCount).toBe(0);
    v.destroy();
  });

  // A <canvas> type-checks as HTMLElement (the pager API takes one), but canvas
  // children never render — the viewer would come up silently blank. Rejected
  // loudly at construction instead.
  it('throws when the container is a <canvas> (pager-API confusion)', () => {
    installDom();
    expect(
      () => new DocxScrollViewer(makeEl('canvas') as unknown as HTMLElement, {}),
    ).toThrow(/container element .* not a <canvas>/i);
    // A plain div (the documented contract) constructs fine.
    const v = new DocxScrollViewer(makeContainer() as unknown as HTMLElement, {});
    v.destroy();
  });

  // O1 (design §11): an borrowed engine's own `mode` is authoritative. An
  // EXPLICITLY conflicting opts.mode is a mis-configuration rejected at
  // construction; a matching or absent opts.mode constructs fine.
  it('throws when opts.mode conflicts with an borrowed worker-mode engine', () => {
    installDom();
    const engine = new FakeDocxEngine(1, [{ widthPt: 612, heightPt: 792 }], 'worker');
    expect(
      () =>
        makeBorrowedDocxScrollViewer(makeContainer() as unknown as HTMLElement, {
          document: engine.asDoc(),
          mode: 'main',
        }),
    ).toThrow(/mode/i);
  });

  it('does NOT throw when opts.mode matches an borrowed worker-mode engine', () => {
    installDom();
    const engine = new FakeDocxEngine(1, [{ widthPt: 612, heightPt: 792 }], 'worker');
    const v = makeBorrowedDocxScrollViewer(makeContainer() as unknown as HTMLElement, {
      document: engine.asDoc(),
      mode: 'worker',
    });
    expect(v.pageCount).toBe(1);
    v.destroy();
    // Borrowed engine is caller-owned even in the worker case: destroy() leaves it intact.
    expect(engine.destroyed).toBe(false);
  });

  it('constructs a default-main borrowed engine with absent opts.mode (load still rejects; destroy preserves engine)', async () => {
    installDom();
    // Default mode is 'main'; opts.mode is absent ⇒ no conflict, resolved path is main.
    const engine = new FakeDocxEngine(2, [{ widthPt: 612, heightPt: 792 }]);
    const v = makeBorrowedDocxScrollViewer(makeContainer() as unknown as HTMLElement, {
      document: engine.asDoc(),
    });
    expect(v.pageCount).toBe(2);
    await expect(v.load('x.docx')).rejects.toThrow(/borrowed/i);
    v.destroy();
    // Borrowed engine is caller-owned: destroy() must not tear it down.
    expect(engine.destroyed).toBe(false);
  });

  // `background` paints the scroll surface (the "desk") visible behind/between
  // pages. It applies to the viewer-owned scrollHost; pages keep their own white.
  it('applies opts.background to the scrollHost element', () => {
    installDom();
    const container = makeContainer();
    const engine = new FakeDocxEngine(1, [{ widthPt: 612, heightPt: 792 }]);
    const v = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      background: '#525659',
    });
    const scrollHost = container.children[0].children[0]; // wrapper → scrollHost
    expect(scrollHost.style.background).toBe('#525659');
    v.destroy();
  });

  it('sets no background on the scrollHost by default (transparent — container shows through)', () => {
    installDom();
    const container = makeContainer();
    const engine = new FakeDocxEngine(1, [{ widthPt: 612, heightPt: 792 }]);
    const v = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
    });
    const scrollHost = container.children[0].children[0];
    expect(scrollHost.style.background).toBe('');
    v.destroy();
  });
});

describe('DocxScrollViewer — opt-in comment cards', () => {
  it('can keep authored range highlighting while an application owns the comment list', async () => {
    installDom();
    const engine = new FakeDocxEngine(1, [{ widthPt: 612, heightPt: 792 }]);
    const source = { story: 'body', storyInstance: 'body', path: [0] } as const;
    engine.comments = [{ id: 'external-list', author: 'Ada', text: 'Review this' }];
    engine.commentAnchors = [{
      commentId: 'external-list', source, startRunIndex: 0, endRunIndex: 1,
      reference: { source, runIndex: 1, affinity: 'preceding' },
    }] as CommentAnchorRange[];
    engine.feedTextRuns = [{
      text: 'anchored', source, sourceRunIndex: 0,
      x: 20, y: 30, w: 80, h: 14, fontSize: 12, font: '12px sans-serif',
    }];
    const container = makeContainer();
    const viewer = DocxScrollViewer.fromDocument(
      container as unknown as HTMLElement,
      engine.asDoc(),
      { comments: { cards: false, markers: false } },
    );
    await vi.waitFor(() => expect(engine.renderCalls).toHaveLength(1));
    await waitForBuiltInCommentUi(viewer);

    const scrollHost = container.children[0]!.children[0]!;
    const page = scrollHost.children.find((child) => child !== scrollHost.children[0])!;
    expect(page.children.some((child) => child.style.cssText.includes('overflow-y:auto'))).toBe(false);
    const tintLayer = page.children.find((child) =>
      child.children.some((entry) => entry.dataset.ooxmlCommentHighlight !== undefined))!;
    expect(tintLayer).toBeDefined();
    expect(tintLayer.children[0]?.dataset.active).toBe('false');

    await expect(viewer.goToComment('external-list')).resolves.toBe(true);
    expect(tintLayer.children[0]?.dataset.active).toBe('true');
    viewer.destroy();
  });

  it('navigates from an application-owned comment list and caches the resolved page', async () => {
    installDom();
    const engine = new FakeDocxEngine(4, [{ widthPt: 612, heightPt: 792 }]);
    const source = { story: 'body', storyInstance: 'body', path: [0] } as const;
    engine.comments = [{ id: '7', author: 'Ada', text: 'Review this' }];
    engine.commentAnchors = [{
      commentId: '7',
      source,
      startRunIndex: 0,
      endRunIndex: 1,
      reference: { source, runIndex: 1, affinity: 'preceding' },
    }] as CommentAnchorRange[];
    const collectPageRuns = vi.fn(async (page: number) => page === 2 ? [{
      text: 'anchored', source, sourceRunIndex: 0,
      x: 20, y: 500, w: 80, h: 14, fontSize: 12, font: '12px sans-serif',
    }] : []);
    engine.collectPageRuns = collectPageRuns;
    const container = makeContainer();
    const viewer = DocxScrollViewer.fromDocument(
      container as unknown as HTMLElement,
      engine.asDoc(),
    );

    const scrollHost = container.children[0]!.children[0]!;
    viewer.scrollToPage(2);
    const pageTop = scrollHost.scrollTop;
    viewer.scrollToPage(0);
    await expect(viewer.goToComment('7')).resolves.toBe(true);
    expect(scrollHost.scrollTop).toBeGreaterThan(pageTop);
    expect(viewer.getSelectionContext()).toMatchObject({
      kind: 'comment', commentId: '7', pageIndex: 2,
    });
    expect(collectPageRuns.mock.calls.map(([page]) => page)).toEqual([0, 1, 2]);

    await expect(viewer.goToComment('7')).resolves.toBe(true);
    expect(collectPageRuns.mock.calls.map(([page]) => page)).toEqual([0, 1, 2]);
    await expect(viewer.goToComment('missing')).resolves.toBe(false);
    viewer.destroy();
  });

  it('shares a progressive comment-page scan and lets the latest list click win', async () => {
    installDom();
    const engine = new FakeDocxEngine(3, [{ widthPt: 612, heightPt: 792 }]);
    const sourceA = { story: 'body', storyInstance: 'body', path: [0] } as const;
    const sourceB = { story: 'body', storyInstance: 'body', path: [1] } as const;
    engine.comments = [
      { id: 'a', author: 'Ada', text: 'First' },
      { id: 'b', author: 'Grace', text: 'Second' },
    ];
    engine.commentAnchors = [
      {
        commentId: 'a', source: sourceA, startRunIndex: 0, endRunIndex: 1,
        reference: { source: sourceA, runIndex: 1, affinity: 'preceding' },
      },
      {
        commentId: 'b', source: sourceB, startRunIndex: 0, endRunIndex: 1,
        reference: { source: sourceB, runIndex: 1, affinity: 'preceding' },
      },
    ] as CommentAnchorRange[];
    let resolveFirstPage!: (runs: DocxTextRunInfo[]) => void;
    const firstPage = new Promise<DocxTextRunInfo[]>((resolve) => { resolveFirstPage = resolve; });
    const collectPageRuns = vi.fn((page: number) => {
      if (page === 0) return firstPage;
      const source = page === 1 ? sourceA : sourceB;
      return Promise.resolve([{
        text: page === 1 ? 'first' : 'second', source, sourceRunIndex: 0,
        x: 20, y: 200, w: 80, h: 14, fontSize: 12, font: '12px sans-serif',
      }]);
    });
    engine.collectPageRuns = collectPageRuns;
    const viewer = DocxScrollViewer.fromDocument(
      makeContainer() as unknown as HTMLElement,
      engine.asDoc(),
    );

    const older = viewer.goToComment('a');
    await Promise.resolve();
    const newer = viewer.goToComment('b');
    resolveFirstPage([]);

    await expect(older).resolves.toBe(false);
    await expect(newer).resolves.toBe(true);
    expect(collectPageRuns.mock.calls.map(([page]) => page)).toEqual([0, 1, 2]);
    expect(viewer.getSelectionContext()).toMatchObject({
      kind: 'comment', commentId: 'b', pageIndex: 2,
    });

    // Page 1 was indexed while the second request advanced to page 2. Navigating
    // back to its comment reuses that shared scan and its retained run geometry.
    await expect(viewer.goToComment('a')).resolves.toBe(true);
    expect(collectPageRuns.mock.calls.map(([page]) => page)).toEqual([0, 1, 2]);
    viewer.destroy();
  });

  it('navigates to a requested rendered occurrence instead of the first page cache', async () => {
    installDom();
    const engine = new FakeDocxEngine(3, [{ widthPt: 612, heightPt: 792 }]);
    const source = { story: 'header', storyInstance: 'default', path: [0] } as const;
    engine.comments = [{ id: 'repeated', author: 'Ada', text: 'Repeated header' }];
    engine.commentAnchors = [{
      commentId: 'repeated',
      source,
      startRunIndex: 0,
      endRunIndex: 1,
      reference: { source, runIndex: 1, affinity: 'preceding' },
    }] as CommentAnchorRange[];
    const collectPageRuns = vi.fn(async (page: number) => page === 1 || page === 2 ? [{
      text: 'header', source, sourceRunIndex: 0,
      x: 20, y: 40, w: 80, h: 14, fontSize: 12, font: '12px sans-serif',
    }] : []);
    engine.collectPageRuns = collectPageRuns;
    const viewer = DocxScrollViewer.fromDocument(
      makeContainer() as unknown as HTMLElement,
      engine.asDoc(),
    );

    await expect(viewer.goToComment('repeated')).resolves.toBe(true);
    expect(viewer.getSelectionContext()).toMatchObject({ pageIndex: 1 });

    await expect(viewer.goToComment('repeated', { pageIndex: 2 })).resolves.toBe(true);
    expect(viewer.getSelectionContext()).toMatchObject({ pageIndex: 2 });
    await expect(viewer.goToComment('repeated', { pageIndex: 0 })).resolves.toBe(false);
    await expect(viewer.goToComment('repeated', { pageIndex: 99 })).resolves.toBe(false);
    viewer.destroy();
  });

  it('does not collect comment geometry or add comment DOM by default', async () => {
    installDom();
    const engine = new FakeDocxEngine(1, [{ widthPt: 612, heightPt: 792 }]);
    engine.comments = [{ id: '1', text: 'Present but not requested' }];
    const container = makeContainer();
    const viewer = DocxScrollViewer.fromDocument(
      container as unknown as HTMLElement,
      engine.asDoc(),
    );
    await vi.waitFor(() => expect(engine.renderCalls).toHaveLength(1));
    expect(engine.renderCalls[0]?.hasTextRunCallback).toBe(false);
    const scrollHost = container.children[0]!.children[0]!;
    const page = scrollHost.children.find((child) => child !== scrollHost.children[0])!;
    expect(page.children.some((child) => child.style.cssText.includes('overflow-y:auto'))).toBe(false);
    viewer.destroy();
  });

  it('does not reserve an empty margin when every root thread is resolved', async () => {
    installDom();
    const engine = new FakeDocxEngine(1, [{ widthPt: 612, heightPt: 792 }]);
    engine.comments = [{ id: 'resolved', text: 'Done', resolved: true }];
    const container = makeContainer();
    const viewer = DocxScrollViewer.fromDocument(
      container as unknown as HTMLElement,
      engine.asDoc(),
      { comments: true },
    );
    await vi.waitFor(() => expect(engine.renderCalls).toHaveLength(1));
    const scrollHost = container.children[0]!.children[0]!;
    const page = scrollHost.children.find((child) => child !== scrollHost.children[0])!;
    expect(page.children.some((child) => child.style.cssText.includes('overflow-y:auto'))).toBe(false);
    viewer.destroy();
  });

  it('does not reserve an empty margin for a comment without a displayable anchor', async () => {
    installDom();
    const engine = new FakeDocxEngine(1, [{ widthPt: 612, heightPt: 792 }]);
    engine.comments = [{ id: 'orphan', text: 'No range or point anchor' }];
    const container = makeContainer();
    const viewer = DocxScrollViewer.fromDocument(
      container as unknown as HTMLElement,
      engine.asDoc(),
      { comments: true },
    );
    await vi.waitFor(() => expect(engine.renderCalls).toHaveLength(1));
    const scrollHost = container.children[0]!.children[0]!;
    const page = scrollHost.children.find((child) => child !== scrollHost.children[0])!;
    expect(page.children.some((child) => child.style.cssText.includes('overflow-y:auto'))).toBe(false);
    viewer.destroy();
  });

  it('renders a themeable built-in card without a connector and clears selection outside it', async () => {
    const dom = installDom();
    const engine = new FakeDocxEngine(1, [{ widthPt: 612, heightPt: 792 }]);
    const source = { story: 'body', storyInstance: 'body', path: [0] } as const;
    engine.comments = [{ id: '7', author: 'Ada', text: 'Review this' }];
    engine.commentAnchors = [{
      commentId: '7',
      source,
      startRunIndex: 0,
      endRunIndex: 1,
      reference: { source, runIndex: 1, affinity: 'preceding' },
    }] as CommentAnchorRange[];
    engine.feedTextRuns = [{
      text: 'anchored', source, sourceRunIndex: 0,
      x: 20, y: 30, w: 80, h: 14,
      highlightBounds: { x: 20, y: 32, width: 80, height: 10 },
      fontSize: 12, font: '12px sans-serif',
    }];
    const container = makeContainer();
    const viewer = DocxScrollViewer.fromDocument(
      container as unknown as HTMLElement,
      engine.asDoc(),
      { comments: true },
    );
    await vi.waitFor(() => expect(engine.renderCalls).toHaveLength(1));
    await waitForBuiltInCommentUi(viewer);

    const scrollHost = container.children[0]!.children[0]!;
    const page = scrollHost.children.find((child) => child !== scrollHost.children[0])!;
    const margin = page.children.find((child) => child.style.cssText.includes('overflow-y:auto'))!;
    const tintLayer = page.children.find((child) =>
      child.children.some((entry) => entry.dataset.ooxmlCommentHighlight !== undefined))!;
    const tint = tintLayer.children[0]!;
    expect(parseFloat(tint.style.top) / parseFloat(tint.style.height)).toBeCloseTo(3.2);
    expect(margin.style.background).toBe('');
    expect(margin.dataset.ooxmlCommentUi).toBe('margin');
    const card = margin.children[0]!.children[0]!;
    const geometry = vi.spyOn(
      viewer as unknown as { _scheduleCommentGeometry(page: number, slot: unknown): void },
      '_scheduleCommentGeometry',
    );
    dom.resizeCb()?.();
    expect(geometry).not.toHaveBeenCalled();
    const frame = card.children.find((child) => child.dataset.ooxmlCommentPart === 'frame')!;
    expect(card.dataset.ooxmlCommentCard).toBe('');
    expect(card.className).toBe('ooxml-comment-card');
    expect(card.style.cssText).toContain('--ooxml-comment-author-accent:');
    expect(card.style.cssText).not.toContain('background:');
    expect(card.style.cssText).not.toContain('border-radius:');
    expect(frame.style.cssText).toBe('');
    expect(tint.style.cssText).toContain('--ooxml-comment-author-accent:');
    expect(tint.style.background).toBe('');
    expect(tint.dataset.active).toBe('false');
    const marker = tintLayer.children.find((child) =>
      child.dataset.ooxmlCommentMarker !== undefined)!;
    expect(marker.className).toBe('ooxml-comment-marker');
    expect(parseFloat(marker.style.left)).toBeGreaterThan(
      parseFloat(tint.style.left) + parseFloat(tint.style.width),
    );
    expect(card.children[0]?.children[0]?.children[1]?.textContent).toBe('Review this');
    card.dispatch('click');
    expect(card.dataset.active).toBe('true');
    expect(viewer.getSelectionContext()).toMatchObject({
      format: 'docx', kind: 'comment', pageIndex: 0, commentId: '7',
      thread: { root: { author: 'Ada', text: 'Review this' }, replies: [] },
    });
    dom.dispatchDocument('pointerdown', { target: card });
    expect(card.dataset.active).toBe('true');
    dom.dispatchDocument('pointerdown', { target: container });
    expect(card.dataset.active).toBe('false');
    expect(viewer.getSelectionContext()).toBeNull();
    expect(page.children.some((child) =>
      child.dataset.ooxmlCommentConnectors !== undefined)).toBe(false);
    viewer.destroy();
  });

  it('updates only connector geometry when the comment margin scrolls', async () => {
    installDom();
    const engine = new FakeDocxEngine(1, [{ widthPt: 612, heightPt: 792 }]);
    const source = { story: 'body', storyInstance: 'body', path: [0] } as const;
    engine.comments = [{ id: 'connector', author: 'Ada', text: 'Review this' }];
    engine.commentAnchors = [{
      commentId: 'connector', source, startRunIndex: 0, endRunIndex: 1,
      reference: { source, runIndex: 1, affinity: 'preceding' },
    }] as CommentAnchorRange[];
    engine.feedTextRuns = [{
      text: 'anchored', source, sourceRunIndex: 0,
      x: 20, y: 300, w: 80, h: 14, fontSize: 12, font: '12px sans-serif',
    }];
    const container = makeContainer();
    const viewer = DocxScrollViewer.fromDocument(
      container as unknown as HTMLElement,
      engine.asDoc(),
      { comments: { connectors: {} } },
    );
    await vi.waitFor(() => expect(engine.renderCalls).toHaveLength(1));
    await waitForBuiltInCommentUi(viewer);
    const scrollHost = container.children[0]!.children[0]!;
    const page = scrollHost.children.find((child) => child !== scrollHost.children[0])!;
    expect(page.children.some((child) =>
      child.dataset.ooxmlCommentConnectors !== undefined)).toBe(true);
    const margin = page.children.find((child) => child.style.cssText.includes('overflow-y:auto'))!;
    const tintLayer = page.children.find((child) =>
      child.children.some((entry) => entry.dataset.ooxmlCommentHighlight !== undefined))!;
    const highlight = tintLayer.children.find((child) =>
      child.dataset.ooxmlCommentHighlight !== undefined)!;
    const marker = tintLayer.children.find((child) =>
      child.dataset.ooxmlCommentMarker !== undefined)!;
    const card = margin.children[0]!.children[0]!;
    const fullRedraw = vi.spyOn(
      viewer as unknown as { _redrawSlotComments(page: number, slot: unknown): void },
      '_redrawSlotComments',
    );
    const connectorRedraw = vi.spyOn(
      viewer as unknown as { _redrawSlotCommentConnectors(page: number, slot: unknown): void },
      '_redrawSlotCommentConnectors',
    );
    margin.scrollTop = 24;
    margin.dispatch('scroll');
    await Promise.resolve();

    expect(fullRedraw).toHaveBeenCalledTimes(0);
    expect(connectorRedraw).toHaveBeenCalledTimes(1);
    expect(tintLayer.children.filter((child) =>
      child.dataset.ooxmlCommentHighlight !== undefined)).toHaveLength(1);
    expect(tintLayer.children.find((child) =>
      child.dataset.ooxmlCommentHighlight !== undefined)).toBe(highlight);
    expect(tintLayer.children.find((child) =>
      child.dataset.ooxmlCommentMarker !== undefined)).toBe(marker);
    expect(margin.children[0]!.children[0]).toBe(card);
    viewer.destroy();
  });

  it('places the margin on the left for an RTL host and permits an explicit override', async () => {
    installDom();
    const engine = new FakeDocxEngine(1, [{ widthPt: 612, heightPt: 792 }]);
    const source = { story: 'body', storyInstance: 'body', path: [0] } as const;
    engine.comments = [{ id: 'rtl', author: 'Ada', text: 'Review this' }];
    engine.commentAnchors = [{
      commentId: 'rtl', source, startRunIndex: 0, endRunIndex: 1,
      reference: { source, runIndex: 1, affinity: 'preceding' },
    }] as CommentAnchorRange[];
    engine.feedTextRuns = [{
      text: 'anchored', source, sourceRunIndex: 0,
      x: 20, y: 30, w: 80, h: 14, fontSize: 12, font: '12px sans-serif',
    }];
    const rtlContainer = makeContainer();
    rtlContainer.style.direction = 'rtl';
    const rtlViewer = DocxScrollViewer.fromDocument(
      rtlContainer as unknown as HTMLElement,
      engine.asDoc(),
      { comments: { side: 'auto', connectors: {} } },
    );
    await vi.waitFor(() => expect(engine.renderCalls).toHaveLength(1));
    await waitForBuiltInCommentUi(rtlViewer);
    const rtlHost = rtlContainer.children[0]!.children[0]!;
    const rtlPage = rtlHost.children.find((child) => child !== rtlHost.children[0])!;
    const rtlMargin = rtlPage.children.find((child) => child.style.cssText.includes('overflow-y:auto'))!;
    const rtlConnectors = rtlPage.children.find((child) =>
      child.dataset.ooxmlCommentConnectors !== undefined)!;
    expect(rtlMargin.style.right).toContain('calc(100% +');
    expect(rtlMargin.style.left).toBe('');
    expect(parseFloat(rtlConnectors.style.left)).toBeLessThan(0);
    rtlViewer.destroy();

    const rightContainer = makeContainer();
    rightContainer.style.direction = 'rtl';
    const rightViewer = DocxScrollViewer.fromDocument(
      rightContainer as unknown as HTMLElement,
      engine.asDoc(),
      { comments: { side: 'right', connectors: {} } },
    );
    await vi.waitFor(() => expect(engine.renderCalls).toHaveLength(2));
    await waitForBuiltInCommentUi(rightViewer);
    const rightHost = rightContainer.children[0]!.children[0]!;
    const rightPage = rightHost.children.find((child) => child !== rightHost.children[0])!;
    const rightMargin = rightPage.children.find((child) => child.style.cssText.includes('overflow-y:auto'))!;
    const rightConnectors = rightPage.children.find((child) =>
      child.dataset.ooxmlCommentConnectors !== undefined)!;
    expect(rightMargin.style.left).toContain('calc(100% +');
    expect(rightMargin.style.right).toBe('');
    expect(rightConnectors.style.left).toBe('0px');
    rightViewer.destroy();
  });

  it('smoothly previews comment geometry while the settled zoom render is pending', async () => {
    installDom();
    const engine = new FakeDocxEngine(1, [{ widthPt: 612, heightPt: 792 }]);
    const source = { story: 'body', storyInstance: 'body', path: [0] } as const;
    engine.comments = [{
      id: '8', author: 'Ada', date: '2026-08-20T09:00:00Z', text: 'Review this',
    }];
    engine.commentAnchors = [{
      commentId: '8', source, startRunIndex: 0, endRunIndex: 1,
      reference: { source, runIndex: 1, affinity: 'preceding' },
    }] as CommentAnchorRange[];
    engine.feedTextRuns = [{
      text: 'anchored', source, sourceRunIndex: 0,
      x: 20, y: 30, w: 80, h: 14, fontSize: 12, font: '12px sans-serif',
    }];
    const container = makeContainer();
    const viewer = DocxScrollViewer.fromDocument(
      container as unknown as HTMLElement,
      engine.asDoc(),
      { comments: { connectors: {} } },
    );
    await vi.waitFor(() => expect(engine.renderCalls).toHaveLength(1));
    await waitForBuiltInCommentUi(viewer);

    const scrollHost = container.children[0]!.children[0]!;
    const page = scrollHost.children.find((child) => child !== scrollHost.children[0])!;
    const margin = page.children.find((child) => child.style.cssText.includes('overflow-y:auto'))!;
    const tint = page.children.find((child) => child.style.cssText.includes('inset:0'))!;
    const decoration = page.children.find((child) =>
      child.style.cssText.includes('overflow:visible'))!;
    const item = margin.children[0]!;
    const card = item.children[0]!;
    margin.clientHeight = 1000;
    const comment = card.children[0]!;
    const content = comment.children[0]!;
    const identity = content.children[0]!;
    const baseScale = viewer.getScale();
    expect(parseFloat(margin.style.width)).toBeCloseTo(280 * baseScale);
    expect(parseFloat(margin.style.fontSize)).toBeCloseTo(13);
    expect(item.style.transform).toBe(`scale(${baseScale})`);
    expect(card.dataset.focused).toBe('false');
    card.dispatch('focus');
    expect(card.dataset.focused).toBe('true');
    expect(card.style.boxShadow).toBe('');
    card.dispatch('blur');
    expect(card.dataset.focused).toBe('false');
    card.dispatch('click');
    const baseTop = parseFloat(item.style.top);
    expect(baseTop).toBeGreaterThan(0);
    expect(card.dataset.active).toBe('true');
    expect(identity.children[0]?.dataset.ooxmlCommentPart).toBe('author');
    expect(identity.children[0]?.className).toBe('ooxml-comment-card__author');
    expect(identity.children[1]?.dataset.ooxmlCommentPart).toBe('date');
    expect(identity.children[1]?.className).toBe('ooxml-comment-card__date');
    expect(content.children[1]?.dataset.ooxmlCommentPart).toBe('body');
    expect(content.children[1]?.className).toBe('ooxml-comment-card__body');

    viewer.setScale(baseScale * 1.5);
    expect(parseFloat(margin.style.width)).toBeCloseTo(280 * baseScale * 1.5);
    expect(parseFloat(margin.style.fontSize)).toBeCloseTo(13);
    expect(parseFloat(item.style.top)).toBeCloseTo(baseTop * 1.5);
    expect(item.style.transform).toBe(`scale(${baseScale * 1.5})`);
    expect(tint.style.visibility).toBe('');
    expect(margin.style.visibility).toBe('');
    expect(decoration.style.visibility).toBe('');
    await vi.waitFor(() => {
      expect(parseFloat(item.style.top)).toBeCloseTo(baseTop * 1.5);
      expect(item.style.transform).toBe(`scale(${baseScale * 1.5})`);
      expect(tint.style.visibility).toBe('');
      expect(margin.style.visibility).toBe('');
      expect(decoration.style.visibility).toBe('');
    });
    viewer.destroy();
  });

  it('joins consecutive comment highlights on one line without crossing a line break', async () => {
    installDom();
    const engine = new FakeDocxEngine(1, [{ widthPt: 612, heightPt: 792 }]);
    const source = { story: 'body', storyInstance: 'body', path: [0] } as const;
    engine.comments = [{ id: '10', author: 'Ada', text: 'One line' }];
    engine.commentAnchors = [{
      commentId: '10', source, startRunIndex: 0, endRunIndex: 4,
      reference: { source, runIndex: 4, affinity: 'preceding' },
    }] as CommentAnchorRange[];
    engine.feedTextRuns = [
      { text: 'among', source, sourceRunIndex: 0, direction: 'rtl', x: 20, y: 30, w: 35, h: 14, fontSize: 12, font: '12px sans-serif' },
      { text: 'older', source, sourceRunIndex: 1, direction: 'rtl', x: 60, y: 30, w: 30, h: 14, fontSize: 12, font: '12px sans-serif' },
      { text: 'things', source, sourceRunIndex: 2, direction: 'rtl', x: 96, y: 30, w: 36, h: 14, fontSize: 12, font: '12px sans-serif' },
      { text: 'below', source, sourceRunIndex: 3, direction: 'rtl', x: 20, y: 50, w: 34, h: 14, fontSize: 12, font: '12px sans-serif' },
    ];
    const container = makeContainer();
    const viewer = DocxScrollViewer.fromDocument(
      container as unknown as HTMLElement,
      engine.asDoc(),
      { comments: true },
    );
    await vi.waitFor(() => expect(engine.renderCalls).toHaveLength(1));
    await waitForBuiltInCommentUi(viewer);

    const scrollHost = container.children[0]!.children[0]!;
    const page = scrollHost.children.find((child) => child !== scrollHost.children[0])!;
    const tint = page.children.find((child) => child.style.cssText.includes('inset:0'))!;
    const highlights = tint.children.filter((child) =>
      child.dataset.ooxmlCommentHighlight !== undefined);
    expect(highlights).toHaveLength(2);
    const pageWidth = parseFloat(page.style.width);
    expect(highlights[0]?.style.left).toBe(`${20 / pageWidth * 100}%`);
    expect(highlights[0]?.style.width).toBe(`${112 / pageWidth * 100}%`);
    const marker = tint.children.find((child) =>
      child.dataset.ooxmlCommentMarker !== undefined)!;
    expect(parseFloat(marker.style.left)).toBeLessThan(parseFloat(highlights[0]!.style.left));
    viewer.destroy();
  });

  it('does not paint an orphan highlight for a hidden resolved thread', async () => {
    installDom();
    const engine = new FakeDocxEngine(1, [{ widthPt: 612, heightPt: 792 }]);
    const source = { story: 'body', storyInstance: 'body', path: [0] } as const;
    engine.comments = [
      { id: 'active', author: 'Ada', text: 'Visible' },
      { id: 'resolved', author: 'Grace', text: 'Hidden', resolved: true },
    ];
    engine.commentAnchors = [
      {
        commentId: 'active', source, startRunIndex: 0, endRunIndex: 1,
        reference: { source, runIndex: 1, affinity: 'preceding' },
      },
      {
        commentId: 'resolved', source, startRunIndex: 1, endRunIndex: 2,
        reference: { source, runIndex: 2, affinity: 'preceding' },
      },
    ] as CommentAnchorRange[];
    engine.feedTextRuns = [
      { text: 'active', source, sourceRunIndex: 0, x: 20, y: 30, w: 40, h: 14, fontSize: 12, font: '12px sans-serif' },
      { text: 'resolved', source, sourceRunIndex: 1, x: 20, y: 60, w: 50, h: 14, fontSize: 12, font: '12px sans-serif' },
    ];
    const container = makeContainer();
    const viewer = DocxScrollViewer.fromDocument(
      container as unknown as HTMLElement,
      engine.asDoc(),
      { comments: { markers: false } },
    );
    await vi.waitFor(() => expect(engine.renderCalls).toHaveLength(1));
    await waitForBuiltInCommentUi(viewer);

    const scrollHost = container.children[0]!.children[0]!;
    const page = scrollHost.children.find((child) => child !== scrollHost.children[0])!;
    const tint = page.children.find((child) => child.style.cssText.includes('inset:0'))!;
    const margin = page.children.find((child) => child.style.cssText.includes('overflow-y:auto'))!;
    expect(tint.children.filter((child) =>
      child.dataset.ooxmlCommentHighlight !== undefined)).toHaveLength(1);
    expect(tint.children.filter((child) =>
      child.dataset.ooxmlCommentMarker !== undefined)).toHaveLength(0);
    expect(margin.children).toHaveLength(1);
    viewer.destroy();
  });

  it('keeps the page and comment scale stable across repeated same-size resize observations', async () => {
    const dom = installDom();
    const engine = new FakeDocxEngine(1, [{ widthPt: 612, heightPt: 792 }]);
    engine.comments = [{ id: '9', author: 'Ada', text: 'Stable size' }];
    const source = { story: 'body', storyInstance: 'body', path: [0] } as const;
    engine.commentAnchors = [{
      commentId: '9', source, startRunIndex: 0, endRunIndex: 1,
      reference: { source, runIndex: 1, affinity: 'preceding' },
    }] as CommentAnchorRange[];
    engine.feedTextRuns = [{
      text: 'anchored', source, sourceRunIndex: 0,
      x: 20, y: 30, w: 80, h: 14, fontSize: 12, font: '12px sans-serif',
    }];
    const container = makeContainer();
    const viewer = DocxScrollViewer.fromDocument(
      container as unknown as HTMLElement,
      engine.asDoc(),
      { comments: true },
    );
    await vi.waitFor(() => expect(engine.renderCalls).toHaveLength(1));
    const scrollHost = container.children[0]!.children[0]!;
    const page = scrollHost.children.find((child) => child !== scrollHost.children[0])!;
    const margin = page.children.find((child) => child.style.cssText.includes('overflow-y:auto'))!;
    const scale = viewer.getScale();
    const width = margin.style.width;

    dom.resizeCb()?.();
    dom.resizeCb()?.();

    expect(viewer.getScale()).toBeCloseTo(scale);
    expect(margin.style.width).toBe(width);
    viewer.destroy();
  });
});

describe('DocxScrollViewer — layout + virtualization (T2)', () => {
  // Uniform 100pt×200pt pages, dpr:1, PT_TO_PX applied in the viewer. To keep
  // the assertions in pt-independent terms we drive scrollTop/clientHeight in the
  // SAME px units the viewer computes heights in. Use pageSize heightPt directly
  // and set base scale by mapping the FIRST page's width to the container width
  // (width:undefined ⇒ fit the container).
  function setup(pageCount: number, opts = {}) {
    const dom = installDom();
    const container = makeContainer(200, 400); // clientWidth 200, clientHeight 400
    const engine = new FakeDocxEngine(
      pageCount,
      // Full per-page array (fixed harness treats a single-element array as
      // uniform, but T2 passes the full array explicitly).
      Array.from({ length: pageCount }, () => ({ widthPt: 100, heightPt: 200 })),
    );
    const v = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      gap: 10,
      overscan: 1,
      width: undefined, // fit container width (200) → base scale below
      // Flush horizontal gutters: paddingLeft/Right default to `gap`, which would
      // shrink the fit width to 180 and break the base-scale-1.5 geometry the T2
      // assertions assume (page width px = 200). Pin them to 0 so the fit is the
      // full 200 container width (same pattern the vertical commit used).
      paddingLeft: 0,
      paddingRight: 0,
      ...opts,
    });
    const wrapper = container.children[0] as FakeEl;
    const scrollHost = wrapper.children[0] as FakeEl;
    const spacer = scrollHost.children[0] as FakeEl;
    // Drive layout: container width 200, page width 100pt → base scale maps
    // 100pt*PT_TO_PX to 200px. Provide the viewport height.
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    return { dom, container, engine, v, wrapper, scrollHost, spacer };
  }

  it('sizes the spacer to computeVisibleRange.totalHeight and mounts only the visible window', () => {
    const { v, scrollHost, spacer } = setup(10);
    v.relayout(); // T2 exposes an explicit relayout() the viewer calls after load/resize
    // Spacer height > 0 and equals Σ page-heights + (n-1)*gap in px.
    expect(parseFloat(spacer.style.height)).toBeGreaterThan(0);
    // At scrollTop 0 with a 400px viewport and pages ~ (100pt fit to 200px wide →
    // 200pt page becomes 400px tall) only page 0 is fully visible; overscan 1
    // mounts page 1 too. Assert the mounted slot count is small and bounded.
    const mounted = scrollHost.children.filter((c) => c !== spacer);
    expect(mounted.length).toBeGreaterThanOrEqual(1);
    expect(mounted.length).toBeLessThanOrEqual(3); // topIndex 0 .. lastVisible+overscan
  });

  it('spacer height is exact for variable page heights + gap', () => {
    const dom = installDom();
    const container = makeContainer(200, 400);
    // Variable heights: widths equal so base scale is one value; heights differ.
    const sizes = [
      { widthPt: 100, heightPt: 200 },
      { widthPt: 100, heightPt: 300 },
      { widthPt: 100, heightPt: 150 },
    ];
    const engine = new FakeDocxEngine(3, sizes);
    const v = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      gap: 10,
      overscan: 1,
      paddingLeft: 0, // full-width fit (see T2 setup note) → base scale 1.5
      paddingRight: 0,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    const spacer = scrollHost.children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    // base scale = 200 / (100 * PT_TO_PX) = 200 / (400/3) = 1.5.
    const PT_TO_PX = 4 / 3;
    const scale = 200 / (100 * PT_TO_PX);
    const heights = sizes.map((s) => s.heightPt * PT_TO_PX * scale);
    // New default: paddingTop/paddingBottom each default to `gap` (10), so the
    // spacer is padTop + Σheights + (n-1)*gap + padBottom.
    const expected = 10 + heights.reduce((a, b) => a + b, 0) + (sizes.length - 1) * 10 + 10;
    expect(parseFloat(spacer.style.height)).toBeCloseTo(expected, 3);
    v.destroy();
    void dom;
  });

  it('recycles slots on scroll without unbounded canvas growth (pool reuse)', () => {
    const { v, scrollHost } = setup(50);
    v.relayout();
    const initialMount = scrollHost.children.length;
    // Scroll far down and fire the scroll listener repeatedly.
    for (let top = 0; top <= 4000; top += 400) {
      scrollHost.scrollTop = top;
      scrollHost.dispatch('scroll');
    }
    // The DOM child count (spacer + mounted slots) must stay bounded — the pool
    // reuses slots rather than appending a new canvas per page.
    expect(scrollHost.children.length).toBeLessThanOrEqual(initialMount + 2);
  });

  it('scrolling far then back reuses pooled slot wrappers (bounded distinct allocations)', () => {
    const { v, scrollHost } = setup(50);
    v.relayout();
    // Track EVERY distinct wrapper element the viewer ever appends. If slots are
    // pooled (recycled), scrolling across the whole document then back reuses a
    // small fixed set of wrappers rather than allocating one per visited page.
    const seen = new Set<FakeEl>();
    const collect = () => {
      for (const c of scrollHost.children) if (c.tag === 'div') seen.add(c);
    };
    collect();
    // Sweep deep into the document and back to the top.
    for (const top of [8000, 16000, 8000, 0]) {
      scrollHost.scrollTop = top;
      scrollHost.dispatch('scroll');
      collect();
    }
    // 50 pages were visited across the sweep; a per-page allocator would have
    // created dozens of wrappers. The pool must keep the distinct-wrapper count
    // tiny: spacer(1) + at most a couple of window-sized generations of slots.
    // Window here is ~2-4 slots; allow generous headroom but far below 50.
    expect(seen.size).toBeLessThanOrEqual(8);
    // Coming back to the top mounts page 0 again, drawn from the pool.
    expect(v.mountedPageIndicesForTest()).toContain(0);
  });

  it('mounts the correct window for a mid-document scrollTop', () => {
    const { v, scrollHost } = setup(20);
    v.relayout();
    // Jump to a scrollTop deep in the doc; the mounted slots must include the
    // page under the viewport top (topVisiblePage) and its overscan neighbours.
    scrollHost.scrollTop = 2000;
    scrollHost.dispatch('scroll');
    const top = v.topVisiblePage;
    const mountedPages = v.mountedPageIndicesForTest(); // T2 test hook
    expect(mountedPages).toContain(top);
    expect(Math.max(...mountedPages) - Math.min(...mountedPages)).toBeLessThanOrEqual(4);
  });
});

describe('DocxScrollViewer — rendering (T3)', () => {
  it("main mode: calls renderPage once per mounted slot with that page's px width", () => {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakeDocxEngine(10, [{ widthPt: 100, heightPt: 200 }]);
    const v = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      gap: 10,
      paddingLeft: 0, // full-width fit → page width px = 200 (asserted below)
      paddingRight: 0,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    // Every mounted page got exactly one renderPage call, no more no less.
    const mounted = v.mountedPageIndicesForTest().sort((a, b) => a - b);
    const rendered = engine.renderCalls.map((c) => c.page).sort((a, b) => a - b);
    expect(rendered).toEqual(mounted);
    // Each renderPage call carried THIS page's own px width (uniform px-per-pt
    // scale, §7). base scale = 200/(100*PT_TO_PX); page width px = 200 for all.
    const widths = engine.renderPageWidths();
    for (const w of widths) expect(w).toBeCloseTo(200, 3);
    v.destroy();
  });

  it('main mode: passes each page its OWN px width for MIXED page sizes (uniform px-per-pt scale, §7)', () => {
    installDom();
    const container = makeContainer(200, 400);
    // Mixed physical widths: page 0 is 100pt wide, page 1 is 200pt wide. The
    // uniform document scale is fixed by fitting the FIRST page to the container
    // (base scale = 200 / (100 * PT_TO_PX) = 1.5), so page 1 — twice as wide —
    // must render at TWICE the px width, not the same fit-width. A vacuous impl
    // that hands every page a constant fit-width would give [200, 200].
    const engine = new FakeDocxEngine(2, [
      { widthPt: 100, heightPt: 100 },
      { widthPt: 200, heightPt: 100 },
    ]);
    const v = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      gap: 10,
      paddingLeft: 0, // full-width fit → base scale 1.5 (page widths 200 / 400)
      paddingRight: 0,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    // Both pages mount (page 0 height px = 100*PT_TO_PX*1.5 = 200; page 1 sits at
    // top 210 and intersects the 400px viewport), so both get exactly one render.
    const mounted = v.mountedPageIndicesForTest().sort((a, b) => a - b);
    expect(mounted).toEqual([0, 1]);
    // Per-page px widths in call order: page 0 → 200, page 1 → 400.
    const widths = engine.renderCalls
      .slice()
      .sort((a, b) => a.page - b.page)
      .map((c) => c.width ?? NaN);
    expect(widths[0]).toBeCloseTo(200, 3);
    expect(widths[1]).toBeCloseTo(400, 3);
    v.destroy();
  });

  it('does not re-render a mounted slot for the same page on a no-op scroll', () => {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakeDocxEngine(10, [{ widthPt: 100, heightPt: 200 }]);
    const v = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      gap: 10,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    const before = engine.renderCalls.length;
    const position = vi.spyOn(
      v as unknown as { _positionSlot(slot: unknown, page: number, range: unknown): void },
      '_positionSlot',
    );
    scrollHost.scrollTop = 0; // unchanged window
    scrollHost.dispatch('scroll');
    expect(engine.renderCalls.length).toBe(before); // no duplicate renders
    expect(position).not.toHaveBeenCalled(); // retained geometry is scroll-invariant
    v.destroy();
  });

  it('worker mode: never calls renderPage — routes every slot through renderPageToBitmap', async () => {
    installDom();
    const container = makeContainer(200, 400);
    // mode:'worker' — the real DocxDocument.renderPage THROWS synchronously in
    // worker mode; a viewer that mis-routed would blow up (and renderCalls would
    // record the attempt). The direct _mode routing must never touch renderPage.
    const engine = new FakeDocxEngine(10, [{ widthPt: 100, heightPt: 200 }], 'worker');
    const v = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      gap: 10,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    await Promise.resolve();
    await Promise.resolve();
    expect(engine.renderCalls).toHaveLength(0); // renderPage never touched
    // Every mounted page dispatched exactly one bitmap render.
    const mounted = v.mountedPageIndicesForTest().sort((a, b) => a - b);
    const bitmapPages = [...new Set(engine.bitmapCalls.map((c) => c.page))].sort((a, b) => a - b);
    expect(bitmapPages).toEqual(mounted);
    v.destroy();
  });

  it('worker mode: paints a resolved bitmap into the slot canvas (transfer)', async () => {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakeDocxEngine(10, [{ widthPt: 100, heightPt: 200 }], 'worker');
    const v = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      gap: 10,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    await Promise.resolve();
    await Promise.resolve();
    // A bitmap was produced and transferred into the mounted slot's canvas.
    expect(engine.createdBitmaps.length).toBeGreaterThan(0);
    const slotWrapper = scrollHost.children.find((c) =>
      c.children.some((k) => k.tag === 'canvas'),
    ) as FakeEl;
    const canvas = slotWrapper.children.find((k) => k.tag === 'canvas') as FakeEl;
    // The bitmaprenderer ctx received the bitmap (fake records lastBitmap).
    expect(canvas._bitmapCtx?.lastBitmap).toBe(engine.createdBitmaps[0]);
    v.destroy();
  });

  it('worker mode: closes the ImageBitmap on recycle (deferred — render in flight when the slot recycles)', async () => {
    // Bitmap-close observability path (plan open question): the primary path
    // (deterministic slot.bitmap hold under synchronous resolve) does NOT hold —
    // in non-deferred mode transferFromImageBitmap consumes the bitmap and nulls
    // slot.bitmap synchronously BEFORE any scroll-away, so a later recycle has
    // nothing to close. We use the DOCUMENTED FALLBACK: a deferred fake so the
    // render is genuinely in flight when the slot recycles, and the on-resolution
    // stale-check closes the orphan (design §11).
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakeDocxEngine(50, [{ widthPt: 100, heightPt: 200 }], 'worker', true);
    const v = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      gap: 10,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    const firstBatch = engine.bitmapCalls.length;
    expect(firstBatch).toBeGreaterThan(0);
    // Scroll far away so the first pages' slots recycle while their renders are
    // still in flight (deferred → not yet resolved).
    scrollHost.scrollTop = 12000;
    scrollHost.dispatch('scroll');
    // Now resolve the early (now-stale) renders — the viewer must close them.
    for (const call of engine.bitmapCalls.slice(0, firstBatch)) call.resolve();
    await Promise.resolve();
    await Promise.resolve();
    expect(engine.createdBitmaps.some((b) => b.close.mock.calls.length > 0)).toBe(true);
    v.destroy();
  });

  it('worker mode: drops a stale in-flight render — no paint + bitmap closed when the slot moved on', async () => {
    installDom();
    const container = makeContainer(200, 400);
    // Deferred: renderPageToBitmap resolves only when the test calls resolve(),
    // so we can scroll a slot's page out of the window BEFORE the bitmap arrives.
    const engine = new FakeDocxEngine(50, [{ widthPt: 100, heightPt: 200 }], 'worker', true);
    const v = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      gap: 10,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    // Page 0's bitmap is dispatched but NOT yet resolved.
    const page0 = engine.bitmapCalls.find((c) => c.page === 0);
    expect(page0).toBeDefined();
    // Scroll far away so page 0's slot recycles while its render is in flight.
    scrollHost.scrollTop = 12000;
    scrollHost.dispatch('scroll');
    expect(v.mountedPageIndicesForTest()).not.toContain(0);
    // Now the stale render for page 0 resolves — the viewer must NOT paint it and
    // must close the orphaned bitmap.
    const bmp0 = engine.createdBitmaps[engine.bitmapCalls.indexOf(page0!)];
    page0!.resolve();
    await Promise.resolve();
    await Promise.resolve();
    expect(bmp0.close.mock.calls.length).toBeGreaterThan(0); // orphan closed
    v.destroy();
  });

  it('worker mode: a page that recycles then re-mounts while in flight still gets a fresh render', async () => {
    // Coalescing keys on page index; a naive Set<number> would swallow the
    // re-mounted slot's render (the stale resolve clears in-flight AFTER the
    // remount coalesced away → the new slot never paints). The viewer must
    // re-dispatch the live slot's render so the page never stays blank.
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakeDocxEngine(50, [{ widthPt: 100, heightPt: 200 }], 'worker', true);
    const v = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      gap: 10,
      paddingTop: 0, // flush top so page 0's slot sits at top:0px (asserted below)
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    const page0Initial = engine.bitmapCalls.find((c) => c.page === 0);
    expect(page0Initial).toBeDefined();
    const dispatchesForPage0 = () => engine.bitmapCalls.filter((c) => c.page === 0).length;
    expect(dispatchesForPage0()).toBe(1);
    // Scroll away (page 0 recycles, render still in flight) then back to the top
    // (page 0 re-mounts on a fresh slot) — all while the initial render is deferred.
    scrollHost.scrollTop = 12000;
    scrollHost.dispatch('scroll');
    expect(v.mountedPageIndicesForTest()).not.toContain(0);
    scrollHost.scrollTop = 0;
    scrollHost.dispatch('scroll');
    expect(v.mountedPageIndicesForTest()).toContain(0);
    // Coalescing pin: page 0's render is STILL in flight (initial deferred call
    // not yet resolved), so the re-mount must NOT dispatch a second render for
    // page 0 — the in-flight guard swallows it. Exactly one dispatch so far.
    expect(dispatchesForPage0()).toBe(1);
    // The initial (now stale) render resolves — the orphan is dropped and the
    // re-mounted slot's render is (re-)dispatched.
    page0Initial!.resolve();
    await Promise.resolve();
    await Promise.resolve();
    expect(dispatchesForPage0()).toBeGreaterThanOrEqual(2); // fresh render issued
    // Resolve the fresh render and confirm it paints into the live slot.
    for (const c of engine.bitmapCalls.filter((c) => c.page === 0)) c.resolve();
    await Promise.resolve();
    await Promise.resolve();
    const slot0 = scrollHost.children.find(
      (c) => c.tag === 'div' && c.children.some((k) => k.tag === 'canvas') && c.style.top === '0px',
    ) as FakeEl | undefined;
    expect(slot0).toBeDefined();
    const canvas0 = slot0!.children.find((k) => k.tag === 'canvas') as FakeEl;
    expect(canvas0._bitmapCtx?.lastBitmap).toBeTruthy(); // painted, not blank
    v.destroy();
  });

  it('worker mode: a PLAIN render rejection (slot still live, epoch unchanged) does NOT re-dispatch — no retry storm', async () => {
    // B1 regression: the finally re-dispatch must gate on STALENESS, not merely
    // `!painted`. When `renderPageToBitmap` rejects while the slot is still live
    // and the epoch is unchanged, `painted` is false and `live === slot`, but
    // NEITHER staleness test fires, so we must NOT re-dispatch. A `!painted`-only
    // gate would loop reject → re-dispatch → reject … unbounded (empirically 1→2
    // →3→4 with onError every round). The onError contract leaves the page blank.
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakeDocxEngine(50, [{ widthPt: 100, heightPt: 200 }], 'worker', true);
    const onError = vi.fn();
    const v = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      gap: 10,
      overscan: 1,
      onError,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();

    const page0 = engine.bitmapCalls.find((c) => c.page === 0);
    expect(page0).toBeDefined();
    expect(v.mountedPageIndicesForTest()).toContain(0); // slot still live
    const dispatchesForPage0 = () => engine.bitmapCalls.filter((c) => c.page === 0).length;
    expect(dispatchesForPage0()).toBe(1);

    // Reject the render with the slot STILL LIVE and no scale change (epoch fixed).
    page0!.reject(new Error('worker render failed'));
    // Flush microtasks generously — a retry storm would keep queuing dispatches.
    for (let k = 0; k < 8; k++) await Promise.resolve();

    // Dispatch count for page 0 stayed at 1 — the storm is gone.
    expect(dispatchesForPage0()).toBe(1);
    // onError fired exactly once (the single failure), never again.
    expect(onError).toHaveBeenCalledTimes(1);
    // Still no further dispatches after more microtask flushing.
    for (let k = 0; k < 8; k++) await Promise.resolve();
    expect(dispatchesForPage0()).toBe(1);
    expect(onError).toHaveBeenCalledTimes(1);
    v.destroy();
  });

  it('worker mode: destroy() mid-flight closes the resolving bitmap and does not fire onError post-destroy', async () => {
    installDom();
    const container = makeContainer(200, 400);
    // Deferred so a render is genuinely in flight when we destroy the viewer.
    const engine = new FakeDocxEngine(50, [{ widthPt: 100, heightPt: 200 }], 'worker', true);
    const onError = vi.fn();
    const v = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      gap: 10,
      onError,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    const page0 = engine.bitmapCalls.find((c) => c.page === 0);
    expect(page0).toBeDefined();
    // Tear down while page 0's render is still in flight. destroy() recycles the
    // slot (no bitmap held yet — none received), so the on-resolution stale-check
    // must close the orphan; the identity guard fails (slot no longer live).
    v.destroy();
    const bmp0 = engine.createdBitmaps[engine.bitmapCalls.indexOf(page0!)];
    page0!.resolve();
    await Promise.resolve();
    await Promise.resolve();
    expect(bmp0.close.mock.calls.length).toBeGreaterThan(0); // orphan closed, no leak
    expect(onError).not.toHaveBeenCalled(); // no error surfaced after teardown
  });
});

describe('DocxScrollViewer — zoom (T4)', () => {
  // PT_TO_PX = 4/3. Uniform 100pt×200pt pages, container 200×400, width:undefined
  // ⇒ base fit maps page-0 width (100pt) to 200px:
  //   base = 200 / (100 * 4/3) = 1.5
  //   page height px = 200 * 4/3 * base = 400
  //   offset[i] = i * (400 + gap) = i * 410
  const PT_TO_PX = 4 / 3;

  function setup(pageCount = 20, opts = {}) {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakeDocxEngine(
      pageCount,
      Array.from({ length: pageCount }, () => ({ widthPt: 100, heightPt: 200 })),
    );
    const v = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      gap: 10,
      // Flush top/bottom (paddingTop/Bottom default to `gap`; the T4 offset
      // arithmetic below is written for offset[0]===0). This exercises the
      // "explicit 0 ⇒ old flush behavior reachable" contract.
      paddingTop: 0,
      paddingBottom: 0,
      // Flush left/right so the fit width is the full 200 container ⇒ base 1.5.
      paddingLeft: 0,
      paddingRight: 0,
      overscan: 1,
      zoomMin: 0.5,
      zoomMax: 3,
      ...opts,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    return { v, scrollHost, engine, container };
  }

  it('base fit scale maps the first page width to the container width', () => {
    const { v } = setup();
    // base = 200 / (100 * 4/3) = 1.5
    expect(v.scaleForTest()).toBeCloseTo(1.5, 5);
    expect(v.baseScaleForTest()).toBeCloseTo(1.5, 5);
    v.destroy();
  });

  it('re-anchors so the page under the viewport top stays fixed across a zoom (intraFrac 0)', () => {
    const { v, scrollHost } = setup();
    // page 3 top offset at base (1.5): 3 * (400 + 10) = 1230
    scrollHost.scrollTop = 1230;
    scrollHost.dispatch('scroll');
    expect(v.topVisiblePage).toBe(3);
    const cur = v.scaleForTest(); // 1.5
    v.setScale(cur * 2); // 3.0 (== zoomMax, no clamp)
    expect(v.scaleForTest()).toBeCloseTo(3, 5);
    // page 3 stays the top page; new offset = 3 * (800 + 10) = 2430
    expect(v.topVisiblePage).toBe(3);
    expect(Math.abs(scrollHost.scrollTop - 2430)).toBeLessThan(2);
    v.destroy();
  });

  it('re-anchors preserving the intra-page fraction (intraFrac ≠ 0)', () => {
    const { v, scrollHost } = setup();
    // Scroll so HALF of page 3 (200 of its 400px) has passed above the viewport
    // top: scrollTop = offset[3] + 0.5*400 = 1230 + 200 = 1430 → intraFrac 0.5.
    scrollHost.scrollTop = 1430;
    scrollHost.dispatch('scroll');
    expect(v.topVisiblePage).toBe(3);
    v.setScale(v.scaleForTest() * 2); // 1.5 → 3.0; heights' = 800
    // newScrollTop = offset'[3] + 0.5 * 800 = 2430 + 400 = 2830
    expect(Math.abs(scrollHost.scrollTop - 2830)).toBeLessThan(2);
    // Page 3 is still under the viewport top: offset'[3]=2430 <= 2830 < 3240.
    expect(v.topVisiblePage).toBe(3);
    v.destroy();
  });

  it('clamps setScale to the absolute [zoomMin, zoomMax] px-per-pt bounds', () => {
    const { v } = setup();
    v.setScale(100); // above zoomMax 3 (absolute, NOT a multiple of base)
    expect(v.scaleForTest()).toBeCloseTo(3, 5);
    v.setScale(0.01); // below zoomMin 0.5
    expect(v.scaleForTest()).toBeCloseTo(0.5, 5);
    v.destroy();
  });

  it('setScale defaults clamp to [0.1, 4] when zoomMin/zoomMax are unset', () => {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakeDocxEngine(5, [{ widthPt: 100, heightPt: 200 }]);
    const v = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      gap: 10,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    v.setScale(999);
    expect(v.scaleForTest()).toBeCloseTo(4, 5); // default zoomMax
    v.setScale(0.0001);
    expect(v.scaleForTest()).toBeCloseTo(0.1, 5); // default zoomMin
    v.destroy();
  });

  it('Ctrl+wheel zooms in and calls preventDefault; bare wheel does not zoom or preventDefault', () => {
    const { v, scrollHost } = setup();
    const before = v.scaleForTest();

    // Bare wheel: no zoom, native scroll (preventDefault NOT called).
    const barePrevent = vi.fn();
    scrollHost.dispatch('wheel', { deltaY: -100, ctrlKey: false, metaKey: false, preventDefault: barePrevent });
    expect(v.scaleForTest()).toBe(before);
    expect(barePrevent).not.toHaveBeenCalled();

    // Ctrl+wheel (deltaY<0 = zoom in): scale increases, preventDefault called.
    const ctrlPrevent = vi.fn();
    scrollHost.dispatch('wheel', { deltaY: -100, ctrlKey: true, metaKey: false, preventDefault: ctrlPrevent });
    expect(v.scaleForTest()).toBeGreaterThan(before);
    expect(ctrlPrevent).toHaveBeenCalledTimes(1);
    v.destroy();
  });

  it('Cmd(meta)+wheel also zooms', () => {
    const { v, scrollHost } = setup();
    const before = v.scaleForTest();
    scrollHost.dispatch('wheel', { deltaY: -100, ctrlKey: false, metaKey: true, preventDefault() {} });
    expect(v.scaleForTest()).toBeGreaterThan(before);
    v.destroy();
  });

  it('positive deltaY (ctrl+wheel down) zooms out', () => {
    const { v, scrollHost } = setup();
    const before = v.scaleForTest();
    scrollHost.dispatch('wheel', { deltaY: 100, ctrlKey: true, metaKey: false, preventDefault() {} });
    expect(v.scaleForTest()).toBeLessThan(before);
    v.destroy();
  });

  it('enableZoom:false installs no wheel handler — ctrl+wheel is inert', () => {
    const { v, scrollHost } = setup(20, { enableZoom: false });
    const before = v.scaleForTest();
    // No wheel listener was registered at all.
    expect(scrollHost._listeners.has('wheel')).toBe(false);
    const prevent = vi.fn();
    scrollHost.dispatch('wheel', { deltaY: -100, ctrlKey: true, preventDefault: prevent });
    expect(v.scaleForTest()).toBe(before);
    expect(prevent).not.toHaveBeenCalled();
    v.destroy();
  });

  it('a wheel with deltaY 0 is a no-op even with ctrl held', () => {
    const { v, scrollHost } = setup();
    const before = v.scaleForTest();
    const prevent = vi.fn();
    scrollHost.dispatch('wheel', { deltaY: 0, ctrlKey: true, preventDefault: prevent });
    expect(v.scaleForTest()).toBe(before);
    v.destroy();
  });

  // ⚠ MANDATORY render-epoch test (T3 review finding I-3). Worker path, deferred:
  // dispatch at scale A, setScale(B) mid-flight, resolve — the old-scale bitmap is
  // closed, NOT painted, and a fresh dispatch at scale B repaints the slot.
  it('render epoch: a bitmap dispatched at the old scale is dropped (closed, not painted) after setScale, and re-dispatched at the new scale', async () => {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakeDocxEngine(20, [{ widthPt: 100, heightPt: 200 }], 'worker', true);
    const v = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      gap: 10,
      paddingTop: 0, // flush top so page 0's slot sits at top:0px (asserted below)
      overscan: 1,
      zoomMin: 0.5,
      zoomMax: 3,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();

    // Page 0's bitmap is dispatched at the base (old) scale but NOT yet resolved.
    const page0Old = engine.bitmapCalls.find((c) => c.page === 0);
    expect(page0Old).toBeDefined();
    const oldDispatchCount = engine.bitmapCalls.filter((c) => c.page === 0).length;
    expect(oldDispatchCount).toBe(1);
    const oldWidth = page0Old!.width;

    // Zoom mid-flight: bumps the render epoch and re-mounts. Page 0's re-mount is
    // coalesced away by the in-flight guard (same index still in flight).
    v.setScale(v.scaleForTest() * 2);
    expect(v.mountedPageIndicesForTest()).toContain(0); // page 0 still visible at top
    expect(engine.bitmapCalls.filter((c) => c.page === 0).length).toBe(1); // no double-dispatch yet

    // The OLD-scale render resolves. Epoch moved ⇒ STALE: close, do not paint,
    // then re-dispatch page 0 at the NEW scale.
    const bmpOld = engine.createdBitmaps[engine.bitmapCalls.indexOf(page0Old!)];
    page0Old!.resolve();
    await Promise.resolve();
    await Promise.resolve();
    expect(bmpOld.close.mock.calls.length).toBeGreaterThan(0); // old bitmap closed
    // A fresh dispatch for page 0 was issued at the new scale.
    const page0Calls = engine.bitmapCalls.filter((c) => c.page === 0);
    expect(page0Calls.length).toBeGreaterThanOrEqual(2);
    const freshCall = page0Calls[page0Calls.length - 1];
    // The fresh dispatch's width reflects the NEW scale (double the old width).
    expect(freshCall.width).toBeGreaterThan(oldWidth ?? 0);

    // The stale old bitmap was NOT transferred; only the fresh one paints.
    freshCall.resolve();
    await Promise.resolve();
    await Promise.resolve();
    const slot0 = scrollHost.children.find(
      (c) => c.tag === 'div' && c.children.some((k) => k.tag === 'canvas') && c.style.top === '0px',
    ) as FakeEl | undefined;
    expect(slot0).toBeDefined();
    const canvas0 = slot0!.children.find((k) => k.tag === 'canvas') as FakeEl;
    expect(canvas0._bitmapCtx?.lastBitmap).not.toBe(bmpOld); // never painted the stale bitmap
    expect(canvas0._bitmapCtx?.lastBitmap).toBeTruthy(); // painted the fresh one
    v.destroy();
  });

  // B1: epoch-then-reject must stay BOUNDED. setScale bumps the epoch mid-flight,
  // so the OLD dispatch is stale on resolution; whether it resolves or REJECTS,
  // the finally must issue exactly ONE fresh dispatch at the new epoch. The fresh
  // dispatch captures the new epoch, so if IT later rejects (same epoch, still
  // live) nothing re-dispatches — the retry is bounded to a single fresh attempt.
  it('render epoch: rejecting the OLD-scale dispatch after setScale still yields exactly ONE fresh dispatch at the new scale (bounded)', async () => {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakeDocxEngine(20, [{ widthPt: 100, heightPt: 200 }], 'worker', true);
    const onError = vi.fn();
    const v = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      gap: 10,
      paddingTop: 0, // flush top so page 0's slot sits at top:0px (asserted below)
      overscan: 1,
      zoomMin: 0.5,
      zoomMax: 3,
      onError,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();

    const page0Old = engine.bitmapCalls.find((c) => c.page === 0);
    expect(page0Old).toBeDefined();
    const dispatchesForPage0 = () => engine.bitmapCalls.filter((c) => c.page === 0).length;
    expect(dispatchesForPage0()).toBe(1);
    const oldWidth = page0Old!.width;

    // Zoom mid-flight: bumps the epoch. Page 0's re-mount is coalesced away.
    v.setScale(v.scaleForTest() * 2);
    expect(v.mountedPageIndicesForTest()).toContain(0);
    expect(dispatchesForPage0()).toBe(1);

    // REJECT the old-scale dispatch. Epoch moved ⇒ stale ⇒ re-dispatch (not a plain
    // failure retry). Exactly ONE fresh dispatch at the new epoch.
    page0Old!.reject(new Error('old-scale render failed'));
    for (let k = 0; k < 8; k++) await Promise.resolve();
    const afterReject = engine.bitmapCalls.filter((c) => c.page === 0);
    expect(afterReject.length).toBe(2); // one fresh dispatch, no storm
    const freshCall = afterReject[afterReject.length - 1];
    expect(freshCall.width).toBeGreaterThan(oldWidth ?? 0); // at the NEW (larger) scale

    // The fresh dispatch then SUCCEEDS and paints — the page is not left blank.
    freshCall.resolve();
    for (let k = 0; k < 8; k++) await Promise.resolve();
    expect(dispatchesForPage0()).toBe(2); // still exactly two; success does not re-dispatch
    const slot0 = scrollHost.children.find(
      (c) => c.tag === 'div' && c.children.some((k) => k.tag === 'canvas') && c.style.top === '0px',
    ) as FakeEl | undefined;
    expect(slot0).toBeDefined();
    const canvas0 = slot0!.children.find((k) => k.tag === 'canvas') as FakeEl;
    expect(canvas0._bitmapCtx?.lastBitmap).toBeTruthy(); // painted the fresh bitmap
    v.destroy();
  });

  it('setScale to the SAME scale is a no-op (no re-anchor, no epoch bump churn)', () => {
    const { v, scrollHost } = setup();
    scrollHost.scrollTop = 1230;
    scrollHost.dispatch('scroll');
    const topBefore = scrollHost.scrollTop;
    v.setScale(v.scaleForTest()); // identical scale
    expect(scrollHost.scrollTop).toBe(topBefore); // unchanged
    v.destroy();
  });
});

describe('DocxScrollViewer — self-load path (T7 story)', () => {
  it('load(url) lays out and mounts the first window (relayout wired into load)', async () => {
    installDom();
    const container = makeContainer(200, 400);
    // Mock the static loader so load() resolves to a fake engine WITHOUT touching
    // a real Worker / WASM. The viewer must call relayout() after assignment so a
    // self-loaded (non-borrowed) viewer is not left blank (I-2).
    const engine = new FakeDocxEngine(10, [{ widthPt: 100, heightPt: 200 }]);
    const loadSpy = vi.spyOn(DocxDocument, 'load').mockResolvedValue(engine.asDoc());
    const v = new DocxScrollViewer(container as unknown as HTMLElement, { gap: 10, password: 'secret' });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    await v.load('sample.docx');
    expect(loadSpy).toHaveBeenCalledTimes(1);
    expect(loadSpy).toHaveBeenCalledWith(
      'sample.docx',
      expect.objectContaining({ password: 'secret' }),
    );
    // Layout happened: slots mounted and the spacer was sized.
    expect(v.mountedPageIndicesForTest().length).toBeGreaterThan(0);
    const spacer = scrollHost.children[0] as FakeEl;
    expect(parseFloat(spacer.style.height)).toBeGreaterThan(0);
    // The mounted pages were actually rendered (main mode → renderPage).
    expect(engine.renderCalls.length).toBeGreaterThan(0);
    v.destroy();
  });
});

describe('DocxScrollViewer — text selection (T5)', () => {
  const RUN = { text: 'Hi', x: 1, y: 2, w: 10, h: 12, fontSize: 12, font: '12px serif' };

  function findSlotWrapper(scrollHost: FakeEl): FakeEl {
    return scrollHost.children.find((c) => c.children.some((k) => k.tag === 'canvas')) as FakeEl;
  }
  function findTextLayer(slotWrapper: FakeEl): FakeEl {
    return slotWrapper.children.find((c) => c.tag === 'div') as FakeEl;
  }

  it('main mode: builds an overlay span per run for each visible slot', async () => {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakeDocxEngine(3, [{ widthPt: 100, heightPt: 200 }]);
    engine.feedTextRuns = [RUN];
    const v = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      enableTextSelection: true,
      gap: 10,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    v.relayout();
    await Promise.resolve();
    await Promise.resolve();
    // Every mounted slot got its overlay built from that page's render.
    const mounted = v.mountedPageIndicesForTest();
    expect(mounted.length).toBeGreaterThan(0);
    for (const wrapper of scrollHost.children.filter((c) => c.children.some((k) => k.tag === 'canvas'))) {
      const textLayer = findTextLayer(wrapper);
      expect(textLayer.children.length).toBe(1);
      expect(textLayer.children[0].textContent).toBe('Hi');
    }
    v.destroy();
  });

  it('leaves the overlay container at 100% so it tracks the slot box', async () => {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakeDocxEngine(1, [{ widthPt: 100, heightPt: 200 }]);
    engine.feedTextRuns = [RUN];
    const v = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      enableTextSelection: true,
      gap: 10,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    v.relayout();
    await Promise.resolve();
    await Promise.resolve();
    const wrapper = findSlotWrapper(scrollHost);
    const textLayer = findTextLayer(wrapper);
    // buildDocxTextLayer no longer pins the overlay to a literal px size: the
    // container keeps its creation `width:100%;height:100%` so it tracks the
    // slot wrapper's actual box (which the slot sizes in px), and every
    // %-placed span scales with a canvas a consumer shrinks via external CSS.
    expect(textLayer.style.width).toBe('100%');
    expect(textLayer.style.height).toBe('100%');
    // The wrapper IS sized to the page CSS box, so `100%` resolves to it.
    expect(wrapper.style.width).toBe('180px'); // 100pt × PT_TO_PX (1.8) — see _pageWidthPx
    v.destroy();
  });

  it('main mode: recycling a slot clears its overlay so the free pool holds no stale spans', async () => {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakeDocxEngine(50, [{ widthPt: 100, heightPt: 200 }]);
    engine.feedTextRuns = [RUN];
    const v = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      enableTextSelection: true,
      gap: 10,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    v.relayout();
    // The overlay is built in the renderPage .then() microtask.
    await Promise.resolve();
    await Promise.resolve();
    // A page-0 slot exists with a built overlay.
    const wrapper = findSlotWrapper(scrollHost);
    const textLayer = findTextLayer(wrapper);
    expect(textLayer.children.length).toBe(1);
    // Scroll far away so page 0 recycles into the free pool.
    scrollHost.scrollTop = 8000;
    scrollHost.dispatch('scroll');
    // The recycled slot's overlay was cleared (no stale spans linger in the pool).
    expect(textLayer.children.length).toBe(0);
    v.destroy();
  });

  it('worker mode: builds an overlay span per run (IX6, no warning)', async () => {
    installDom();
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    const container = makeContainer(200, 400);
    const engine = new FakeDocxEngine(3, [{ widthPt: 100, heightPt: 200 }], 'worker');
    engine.feedTextRuns = [RUN];
    const v = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      enableTextSelection: true,
      gap: 10,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    v.relayout();
    // The overlay is built in the renderPageToBitmap resolution microtask; give
    // the bitmap promise chain a couple of turns to settle.
    await Promise.resolve();
    await Promise.resolve();
    await Promise.resolve();
    // IX6 — worker mode no longer warns; the runs ride back beside the bitmap so
    // the overlay is populated identically to main mode.
    expect(warn).not.toHaveBeenCalled();
    // The worker path (renderPageToBitmap), NOT renderPage, was used.
    expect(engine.renderCalls.length).toBe(0);
    expect(engine.bitmapCalls.length).toBeGreaterThan(0);
    // Every mounted slot got its overlay built from the worker-shipped runs.
    const mounted = v.mountedPageIndicesForTest();
    expect(mounted.length).toBeGreaterThan(0);
    for (const wrapper of scrollHost.children.filter((c) => c.children.some((k) => k.tag === 'canvas'))) {
      const textLayer = wrapper.children.find((c) => c.tag === 'div') as FakeEl;
      expect(textLayer.children.length).toBe(1);
      expect(textLayer.children[0].textContent).toBe('Hi');
    }
    warn.mockRestore();
    v.destroy();
  });

  it('stale main render (epoch moved by setScale mid-flight) does NOT rebuild the overlay', async () => {
    installDom();
    const container = makeContainer(200, 400);
    // Deferred main mode: the test resolves renderPage manually so setScale can
    // move the epoch WHILE page 0's first render is in flight.
    const engine = new FakeDocxEngine(20, [{ widthPt: 100, heightPt: 200 }], 'main', true);
    engine.feedTextRuns = [RUN];
    const v = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      enableTextSelection: true,
      gap: 10,
      zoomMin: 0.1,
      zoomMax: 4,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    v.relayout();
    // Grab the first (stale) page-0 render call before resolving it. Capture the
    // slot wrapper it targets so we can inspect its overlay afterwards.
    const staleWrapper = scrollHost.children.find((c) => c.children.some((k) => k.tag === 'canvas')) as FakeEl;
    const staleLayer = staleWrapper.children.find((c) => c.tag === 'div') as FakeEl;
    const staleCall = engine.renderCalls.find((c) => c.page === 0)!;
    // Zoom mid-flight: bumps the epoch and force-re-mounts every slot (which
    // re-dispatches a fresh page-0 render at the new scale).
    v.setScale(v.scaleForTest() * 2);
    // Now resolve the STALE (old-epoch) render. Its .then must bail on the epoch
    // guard and NOT build the overlay.
    staleCall.resolve();
    await Promise.resolve();
    await Promise.resolve();
    // The stale render built nothing (epoch guard). The fresh render is still
    // deferred (not resolved), so its overlay isn't built either — the stale
    // slot/layer holds zero spans from the superseded render.
    expect(staleLayer.children.length).toBe(0);
    v.destroy();
  });
});

describe('DocxScrollViewer — element context', () => {
  function elementContext(): DocxElementContext {
    return {
      format: 'docx', kind: 'element', pageIndex: 0, elementIndex: 0,
      elementType: 'chart', point: { xPt: 50, yPt: 100 },
      bounds: { xPt: 10, yPt: 20, widthPt: 80, heightPt: 60 },
      source: { story: 'body', storyInstance: 'body', path: [0, 0] },
      text: 'Revenue', seriesCount: 1, truncated: false,
      truncationReasons: [], textCharacters: 7, maxTextCharacters: 16_384,
    };
  }

  async function setupElementContext(opts: ConstructorParameters<typeof DocxScrollViewer>[1]) {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakeDocxEngine(1, [{ widthPt: 100, heightPt: 200 }]);
    engine.elementContext = elementContext();
    const v = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(), gap: 10, paddingLeft: 0, paddingRight: 0, ...opts,
    });
    const scrollHost = container.children[0].children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    await Promise.resolve();
    await Promise.resolve();
    const wrapper = scrollHost.children.find((child) =>
      child.children.some((candidate) => candidate.tag === 'canvas')) as FakeEl;
    const canvas = wrapper.children.find((child) => child.tag === 'canvas') as FakeEl;
    canvas.clientWidth = parseFloat(wrapper.style.width);
    canvas.clientHeight = parseFloat(wrapper.style.height);
    return { canvas, engine, scrollHost, v, wrapper };
  }

  it('maps contextmenu to the mounted page while keeping the native event synchronous', async () => {
    const received: Array<{ originalEvent: MouseEvent; getContext(): Promise<unknown> }> = [];
    const mounted = await setupElementContext({
      enableElementSelection: true,
      onContextMenu(event) { received.push(event); },
    });
    const originalEvent = {
      target: mounted.canvas, button: 2,
      clientX: mounted.canvas.clientWidth / 2,
      clientY: mounted.canvas.clientHeight / 2,
      defaultPrevented: false,
    } as unknown as MouseEvent;

    mounted.scrollHost.dispatch('contextmenu', originalEvent);

    expect(received).toHaveLength(1);
    expect(received[0].originalEvent).toBe(originalEvent);
    await expect(received[0].getContext()).resolves.toMatchObject({
      format: 'docx', kind: 'element', pageIndex: 0,
    });
    expect(mounted.engine.elementContextCalls).toHaveLength(1);
    mounted.v.destroy();
    expect(mounted.scrollHost._listeners.get('contextmenu') ?? []).toHaveLength(0);
  });

  it('rejects contextmenu context lookup failures without also calling onError', async () => {
    const failure = new Error('element lookup failed');
    const onError = vi.fn();
    let received: { getContext(): Promise<unknown> } | undefined;
    const mounted = await setupElementContext({
      enableElementSelection: true,
      onError,
      onContextMenu(event) { received = event; },
    });
    mounted.engine.getElementContextAt = vi.fn().mockRejectedValue(failure);

    mounted.scrollHost.dispatch('contextmenu', {
      target: mounted.canvas, button: 2,
      clientX: mounted.canvas.clientWidth / 2,
      clientY: mounted.canvas.clientHeight / 2,
      defaultPrevented: false,
    });

    await expect(received?.getContext()).rejects.toBe(failure);
    expect(onError).not.toHaveBeenCalled();
    mounted.v.destroy();
  });

  it('does not activate object hit-testing from the callback alone', async () => {
    const onSelectionContextChange = vi.fn();
    const mounted = await setupElementContext({ onSelectionContextChange });

    mounted.scrollHost.dispatch('click', {
      target: mounted.canvas, button: 0, clientX: 100, clientY: 200,
      defaultPrevented: false,
    });
    await Promise.resolve();

    expect(mounted.engine.elementContextCalls).toEqual([]);
    expect(onSelectionContextChange).not.toHaveBeenCalled();
    expect(mounted.v.getSelectionContext()).toBeNull();
    mounted.v.destroy();
  });

  it('supports getter-only element context when explicitly enabled', async () => {
    const currentDate = new Date('2026-08-09T12:00:00Z');
    const mounted = await setupElementContext({ enableElementSelection: true, currentDate });

    mounted.scrollHost.dispatch('click', {
      target: mounted.canvas, button: 0,
      clientX: mounted.canvas.clientWidth / 2,
      clientY: mounted.canvas.clientHeight / 2,
      defaultPrevented: false,
    });
    await Promise.resolve();
    await Promise.resolve();

    expect(mounted.engine.elementContextCalls).toEqual([{
      pageIndex: 0,
      point: { xPt: 50, yPt: 100 },
      options: { currentDate, maxTextCharacters: 65_536 },
    }]);
    expect(mounted.v.getSelectionContext()).toMatchObject({
      format: 'docx', kind: 'element', elementType: 'chart',
    });
    expect(mounted.wrapper.children.at(-1)?.children).toHaveLength(1);
    mounted.v.destroy();
  });

  it('uses the same currentDate for mounted rendering and element hit-testing', async () => {
    const currentDate = Date.parse('2026-08-09T12:00:00Z');
    const mounted = await setupElementContext({ enableElementSelection: true, currentDate });

    expect(mounted.engine.renderCalls[0]?.currentDate).toBe(currentDate);
    mounted.scrollHost.dispatch('click', {
      target: mounted.canvas, button: 0,
      clientX: mounted.canvas.clientWidth / 2,
      clientY: mounted.canvas.clientHeight / 2,
      defaultPrevented: false,
    });
    await Promise.resolve();
    await Promise.resolve();

    expect(mounted.engine.elementContextCalls[0]?.options.currentDate).toBe(currentDate);
    mounted.v.destroy();
  });
});

describe('DocxScrollViewer — full-document find', () => {
  const RUN = {
    text: 'Alpha beta',
    x: 4,
    y: 6,
    w: 60,
    h: 12,
    fontSize: 12,
    font: '12px serif',
  };

  it('finds case-insensitively beyond the mounted window, navigates, and clears', async () => {
    installDom();
    const container = makeContainer(200, 120);
    const engine = new FakeDocxEngine(
      4,
      Array.from({ length: 4 }, () => ({ widthPt: 100, heightPt: 200 })),
    );
    engine.feedTextRuns = [RUN];
    const v = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      findHighlightColors: {
        match: 'rgba(1, 2, 3, 0.4)',
        active: 'rgba(4, 5, 6, 0.7)',
      },
      gap: 0,
      overscan: 0,
      paddingTop: 0,
      paddingBottom: 0,
      paddingLeft: 0,
      paddingRight: 0,
    });
    const scrollHost = container.children[0].children[0] as FakeEl;
    scrollHost.clientHeight = 120;
    v.relayout();

    const matches = await v.findText('ALPHA');
    expect(matches).toHaveLength(4);
    expect(matches.map((match) => match.location.page)).toEqual([0, 1, 2, 3]);

    const initiallyMountedHighlightLayers = scrollHost.children
      .filter((child) => child.children.some((nested) => nested.tag === 'canvas'))
      .map((slot) => slot.children.find((child) => child.tag === 'div') as FakeEl);
    expect(
      initiallyMountedHighlightLayers
        .flatMap((layer) => layer.children)
        .find((box) => box.style.background === 'rgba(1, 2, 3, 0.4)'),
    ).toBeDefined();

    await v.findNext();
    const second = await v.findNext();
    expect(second?.location.page).toBe(1);
    expect(scrollHost.scrollTop).toBeGreaterThan(0);

    const mountedHighlightLayers = scrollHost.children
      .filter((child) => child.children.some((nested) => nested.tag === 'canvas'))
      .map((slot) => slot.children.find((child) => child.tag === 'div') as FakeEl);
    expect(mountedHighlightLayers.some((layer) => layer.children.length > 0)).toBe(true);
    expect(
      mountedHighlightLayers
        .flatMap((layer) => layer.children)
        .find((box) => box.style.background === 'rgba(4, 5, 6, 0.7)'),
    ).toBeDefined();

    v.clearFind();
    expect(mountedHighlightLayers.every((layer) => layer.children.length === 0)).toBe(true);
    v.destroy();
  });

  it.each(['main', 'worker'] as const)(
    'refreshes cached find geometry after a zoom settle render in %s mode',
    async (mode) => {
    vi.useFakeTimers();
    try {
      installDom();
      const container = makeContainer(200, 120);
      const engine = new FakeDocxEngine(1, [{ widthPt: 100, heightPt: 200 }], mode);
      engine.feedTextRuns = [RUN];
      const v = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, {
        document: engine.asDoc(),
        gap: 0,
        paddingTop: 0,
        paddingBottom: 0,
        paddingLeft: 0,
        paddingRight: 0,
      });
      const scrollHost = container.children[0].children[0] as FakeEl;
      scrollHost.clientHeight = 120;
      v.relayout();
      await v.findText('alpha');
      await Promise.resolve();

      const slot = scrollHost.children.find(
        (child) => child.children.some((nested) => nested.tag === 'canvas'),
      ) as FakeEl;
      const highlightLayer = slot.children.find((child) => child.tag === 'div') as FakeEl;
      expect(highlightLayer.children[0].style.left).toBe('2%');

      engine.feedTextRuns = [{ ...RUN, x: 40 }];
      v.setScale(v.getScale() * 2);
      vi.advanceTimersByTime(200);
      await Promise.resolve();
      await Promise.resolve();

      expect(highlightLayer.children[0].style.left).toBe('10%');
      v.destroy();
    } finally {
      vi.useRealTimers();
    }
    },
  );

  it.each(['main', 'worker'] as const)(
    'keeps zoom-settle geometry when an older pending find completes in %s mode',
    async (mode) => {
      vi.useFakeTimers();
      try {
        installDom();
        const container = makeContainer(200, 120);
        const engine = new FakeDocxEngine(
          2,
          Array.from({ length: 2 }, () => ({ widthPt: 100, heightPt: 200 })),
          mode,
        );
        engine.feedTextRuns = [RUN];
        let resolveSecond!: (runs: NonNullable<typeof engine.feedTextRuns>) => void;
        engine.collectPageRuns = (page) => page === 1
          ? new Promise((resolve) => { resolveSecond = resolve; })
          : Promise.resolve([RUN]);
        const v = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, {
          document: engine.asDoc(),
          gap: 0,
          overscan: 0,
          paddingTop: 0,
          paddingBottom: 0,
          paddingLeft: 0,
          paddingRight: 0,
        });
        const scrollHost = container.children[0].children[0] as FakeEl;
        scrollHost.clientHeight = 120;
        v.relayout();
        await Promise.resolve();
        await Promise.resolve();

        const pending = v.findText('alpha');
        await Promise.resolve();
        await Promise.resolve();

        engine.feedTextRuns = [{ ...RUN, x: 40 }];
        v.setScale(v.getScale() * 2);
        vi.advanceTimersByTime(200);
        await Promise.resolve();
        await Promise.resolve();

        resolveSecond([RUN]);
        await expect(pending).resolves.toHaveLength(2);

        const slot = scrollHost.children.find(
          (child) => child.children.some((nested) => nested.tag === 'canvas'),
        ) as FakeEl;
        const highlightLayer = slot.children.find((child) => child.tag === 'div') as FakeEl;
        expect(highlightLayer.children[0].style.left).toBe('10%');
        v.destroy();
      } finally {
        vi.useRealTimers();
      }
    },
  );

  it.each(['clear', 'destroy'] as const)(
    'does not restore a pending find after %s',
    async (action) => {
      installDom();
      const engine = new FakeDocxEngine(1, [{ widthPt: 100, heightPt: 200 }]);
      let resolveRuns!: (runs: NonNullable<typeof engine.feedTextRuns>) => void;
      engine.collectPageRuns = () => new Promise((resolve) => { resolveRuns = resolve; });
      const v = makeBorrowedDocxScrollViewer(makeContainer(200, 120) as unknown as HTMLElement, {
        document: engine.asDoc(),
      });
      const pending = v.findText('alpha');

      if (action === 'clear') v.clearFind();
      else v.destroy();
      resolveRuns([RUN]);

      await expect(pending).resolves.toEqual([]);
      await expect(v.findNext()).resolves.toBeNull();
      if (action === 'clear') v.destroy();
    },
  );

  it('does not restore a pending find after reload', async () => {
    installDom();
    const first = new FakeDocxEngine(1, [{ widthPt: 100, heightPt: 200 }]);
    const second = new FakeDocxEngine(1, [{ widthPt: 100, heightPt: 200 }]);
    let resolveRuns!: (runs: NonNullable<typeof first.feedTextRuns>) => void;
    first.collectPageRuns = () => new Promise((resolve) => { resolveRuns = resolve; });
    vi.spyOn(DocxDocument, 'load')
      .mockResolvedValueOnce(first.asDoc())
      .mockResolvedValueOnce(second.asDoc());
    const container = makeContainer(200, 120);
    const v = new DocxScrollViewer(container as unknown as HTMLElement);
    const scrollHost = container.children[0].children[0] as FakeEl;
    scrollHost.clientHeight = 120;
    await v.load('first.docx');
    const pending = v.findText('alpha');

    await v.load('second.docx');
    resolveRuns([RUN]);

    await expect(pending).resolves.toEqual([]);
    await expect(v.findNext()).resolves.toBeNull();
    v.destroy();
  });
});

describe('DocxScrollViewer — navigation, resize, empty (T6)', () => {
  const PT_TO_PX = 4 / 3;
  // container fit width 200, page widthPt 100 → base scale = 200/(100*PT_TO_PX)=1.5.
  // page height px = heightPt * PT_TO_PX * scale = 200 * (4/3) * 1.5 = 400.
  // (The plan text mislabelled this as 200px; the viewer fits WIDTH, which scales
  //  height too — verified against the T2/T4 blocks' identical geometry.)
  const BASE = 200 / (100 * PT_TO_PX);
  const PAGE_H = 200 * PT_TO_PX * BASE; // 400
  const GAP = 10;
  const STRIDE = PAGE_H + GAP; // 410 — the top-to-top distance between pages

  function setup(pageCount = 20, opts = {}) {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakeDocxEngine(
      pageCount,
      Array.from({ length: pageCount }, () => ({ widthPt: 100, heightPt: 200 })),
    );
    const v = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      gap: GAP,
      // Flush top/bottom: the T6 STRIDE/offset arithmetic assumes offset[0]===0.
      // paddingTop/Bottom default to `gap`; pinning them to 0 keeps the pre-padding
      // geometry (and exercises the "explicit 0 ⇒ old flush behavior" contract).
      paddingTop: 0,
      paddingBottom: 0,
      // Flush left/right: the fit width must be the full 200 container so BASE = 1.5
      // and PAGE_H/STRIDE hold (the default `gap` gutters would shrink the fit).
      paddingLeft: 0,
      paddingRight: 0,
      ...opts,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    return { v, scrollHost, container, engine };
  }

  it('scrollToPage sets scrollTop to the page offset and clamps out-of-range', () => {
    const { v, scrollHost } = setup();
    // Pick a clientHeight that makes maxTop STRICTLY LESS than the last page's top
    // offset, so the `Math.min(maxTop, target)` clamp is actually exercised. With the
    // default clientHeight 400, offsets[19] === maxTop and the clamp is a degenerate
    // no-op. totalHeight = 8190, offsets[19] = 19*STRIDE = 7790.
    scrollHost.clientHeight = 500; // maxTop = 8190 − 500 = 7690 < offsets[19] 7790
    v.scrollToPage(3);
    expect(Math.abs(scrollHost.scrollTop - 3 * STRIDE)).toBeLessThan(2);
    v.scrollToPage(999); // clamps to last page index (19)
    // Page 19's top offset (7790) exceeds the max scroll top (7690), so scrollToPage
    // pins to the max (totalHeight − viewportHeight). This asserts the Math.min clamp
    // fired: scrollTop is maxTop (7690), STRICTLY below offsets[19] (7790).
    const totalHeight = 20 * PAGE_H + 19 * GAP; // Σheights + Σgaps = 8190
    const maxTop = totalHeight - 500; // − viewport height = 7690
    const lastOffset = 19 * STRIDE; // 7790
    expect(maxTop).toBeLessThan(lastOffset); // precondition: clamp is exercised
    expect(Math.abs(scrollHost.scrollTop - maxTop)).toBeLessThan(2);
    v.scrollToPage(-5); // clamps to 0
    expect(scrollHost.scrollTop).toBe(0);
    v.destroy();
  });

  it('scrollToPage: viewport taller than total content ⇒ scrollTop pinned to 0 (negative maxTop guard)', () => {
    // A 2-page document is shorter than a very tall viewport, so
    // totalHeight − clientHeight is NEGATIVE; the `Math.max(0, …)` maxTop guard must
    // pin scrollTop to 0 rather than a negative top.
    const { v, scrollHost } = setup(2);
    scrollHost.clientHeight = 5000; // >> totalHeight (2*400 + 10 = 810)
    v.scrollToPage(1); // last page; its offset (410) is above maxTop (0)
    expect(scrollHost.scrollTop).toBe(0);
    v.destroy();
  });

  it('onVisiblePageChange fires only when topIndex changes', () => {
    const changes: number[] = [];
    const { v, scrollHost } = setup(20, { onVisiblePageChange: (i: number) => changes.push(i) });
    scrollHost.scrollTop = STRIDE; // page 1 top
    scrollHost.dispatch('scroll');
    scrollHost.scrollTop = STRIDE + 10; // still within page 1
    scrollHost.dispatch('scroll');
    scrollHost.scrollTop = 3 * STRIDE; // page 3 top
    scrollHost.dispatch('scroll');
    // Deduped: 0 (initial) → 1 → 3, no repeat for the STRIDE+10 no-op.
    expect(changes).toEqual([0, 1, 3]);
    v.destroy();
  });

  it('scrollToPage does not re-fire onVisiblePageChange for the same top page', () => {
    const changes: number[] = [];
    const { v } = setup(20, { onVisiblePageChange: (i: number) => changes.push(i) });
    // Initial mount fired 0. scrollToPage(0) lands on the same top page — no new fire.
    v.scrollToPage(0);
    expect(changes).toEqual([0]);
    // Navigate to page 5 → one fire; navigate there again → no duplicate.
    v.scrollToPage(5);
    v.scrollToPage(5);
    expect(changes).toEqual([0, 5]);
    v.destroy();
  });

  it('ResizeObserver re-fit preserves the zoom multiplier', () => {
    // zoomMax headroom: base 1.5, 2× zoom → 3.0; after the width doubles the new
    // base is 3.0 so the preserved scale is 6.0. zoomMin/zoomMax are ABSOLUTE
    // px-per-pt bounds (design §8.1), so a low zoomMax would legitimately CLAMP
    // the re-fit and break preservation. Give the multiplier room to survive.
    const { v, container } = setup(20, { zoomMax: 10 });
    v.setScale(v.baseScaleForTest() * 2); // 2× zoom
    const zoomMultiplier = v.scaleForTest() / v.baseScaleForTest();
    // Container widens; the observer callback re-fits base and re-applies the mult.
    container.clientWidth = 400;
    (container.children[0] as FakeEl).children[0].clientWidth = 400;
    v.resizeForTest(); // fires the observed resize path
    expect(v.scaleForTest() / v.baseScaleForTest()).toBeCloseTo(zoomMultiplier, 5);
    v.destroy();
  });

  it('refitOnResize false preserves the absolute scale when the container width changes', () => {
    const onScaleChange = vi.fn();
    const { v, container } = setup(20, { refitOnResize: false, onScaleChange });
    v.setScale(1);
    onScaleChange.mockClear();
    const epochBefore = v.renderEpochForTest();

    container.clientWidth = 400;
    (container.children[0] as FakeEl).children[0].clientWidth = 400;
    v.resizeForTest();

    expect(v.scaleForTest()).toBe(1);
    expect(v.renderEpochForTest()).toBe(epochBefore);
    expect(onScaleChange).not.toHaveBeenCalled();
    v.destroy();
  });

  it('ResizeObserver re-fit bumps the render epoch (resize is an epoch event)', () => {
    const { v, container } = setup();
    const before = v.renderEpochForTest();
    container.clientWidth = 400;
    (container.children[0] as FakeEl).children[0].clientWidth = 400;
    v.resizeForTest();
    // A width change re-fits the base scale (routed through setScale), which bumps
    // the epoch so any bitmap in flight at the old scale is treated as stale.
    expect(v.renderEpochForTest()).toBeGreaterThan(before);
    v.destroy();
  });

  it('ResizeObserver re-fit with an UNCHANGED width does not re-fit or bump the epoch', () => {
    const { v } = setup();
    const scaleBefore = v.scaleForTest();
    const epochBefore = v.renderEpochForTest();
    v.resizeForTest(); // same width — no-op re-fit
    expect(v.scaleForTest()).toBe(scaleBefore);
    expect(v.renderEpochForTest()).toBe(epochBefore);
    v.destroy();
  });

  it('height-only resize (width unchanged) re-mounts the newly-revealed window without a scroll, no epoch bump, no callback', () => {
    // F1 regression: a height-only grow used to early-return before `_mountVisible`,
    // leaving the newly-revealed rows blank until the user scrolled. At scrollTop 0
    // the initially-mounted window is [0,1] (viewport 400 = one page; +1 overscan).
    // Growing the viewport to 1600 must mount [0..4] purely from the resize — no
    // scroll event, no epoch bump (geometry/scale unchanged so cached canvases stay
    // valid), and no onVisiblePageChange fire (topIndex stays 0 at scrollTop 0).
    const changes: number[] = [];
    const { v, scrollHost } = setup(20, { onVisiblePageChange: (i: number) => changes.push(i) });
    const mountedBefore = v.mountedPageIndicesForTest().slice().sort((a, b) => a - b);
    expect(mountedBefore).toEqual([0, 1]);
    expect(changes).toEqual([0]); // initial mount fired 0
    const epochBefore = v.renderEpochForTest();
    const scaleBefore = v.scaleForTest();

    // Grow HEIGHT only; width (and thus the fit-width base scale) is unchanged.
    scrollHost.clientHeight = 1600;
    v.resizeForTest(); // NO scroll event dispatched

    const mountedAfter = v.mountedPageIndicesForTest().slice().sort((a, b) => a - b);
    // The revealed window grew: bottom edge 1600 now intersects pages 0..3 (+1
    // overscan ⇒ end 4). The mounted span strictly grows.
    expect(mountedAfter).toEqual([0, 1, 2, 3, 4]);
    expect(mountedAfter.length).toBeGreaterThan(mountedBefore.length);
    // No epoch bump — nothing was re-scaled, so no in-flight render is stale.
    expect(v.renderEpochForTest()).toBe(epochBefore);
    expect(v.scaleForTest()).toBe(scaleBefore);
    // topIndex is still 0 at scrollTop 0 ⇒ callback did NOT fire again.
    expect(changes).toEqual([0]);
    v.destroy();
  });

  it('clamped resize (zoomMax clamp + width & height grow) still refreshes the revealed window', () => {
    // F1 clamped-path variant: pin zoomMax to the current base so a width grow's
    // re-fit (newBase × mult) clamps back to the SAME scale, making setScale a no-op
    // (which skips its own force-re-render mount). The post-setScale `_mountVisible`
    // must still mount the rows the taller viewport revealed. base = 1.5, no user
    // zoom ⇒ mult 1; after width→400 newBase 3.0 clamps to zoomMax 1.5 (== _scale).
    const changes: number[] = [];
    const { v, scrollHost, container } = setup(20, {
      zoomMax: BASE, // 1.5 — clamps the re-fit back to the current scale
      onVisiblePageChange: (i: number) => changes.push(i),
    });
    expect(v.scaleForTest()).toBe(BASE);
    const mountedBefore = v.mountedPageIndicesForTest().slice().sort((a, b) => a - b);
    expect(mountedBefore).toEqual([0, 1]);
    const epochBefore = v.renderEpochForTest();

    // Grow BOTH width and height. The width grow triggers the re-fit branch, but the
    // clamp pins the scale unchanged ⇒ setScale no-ops. Heights are therefore
    // unchanged (scale unchanged), so the taller viewport reveals more of the SAME
    // geometry: mounted grows to [0..4].
    container.clientWidth = 400;
    (container.children[0] as FakeEl).children[0].clientWidth = 400;
    scrollHost.clientWidth = 400;
    scrollHost.clientHeight = 1600;
    v.resizeForTest();

    const mountedAfter = v.mountedPageIndicesForTest().slice().sort((a, b) => a - b);
    expect(mountedAfter).toEqual([0, 1, 2, 3, 4]);
    expect(mountedAfter.length).toBeGreaterThan(mountedBefore.length);
    // Scale stayed pinned at the clamp, so setScale no-oped ⇒ no epoch bump.
    expect(v.scaleForTest()).toBe(BASE);
    expect(v.renderEpochForTest()).toBe(epochBefore);
    // topIndex still 0 ⇒ no extra callback.
    expect(changes).toEqual([0]);
    v.destroy();
  });

  it('zero-width container defers, then lays out on first non-zero resize', () => {
    installDom();
    const container = makeContainer(0, 0); // unlaid-out
    const engine = new FakeDocxEngine(5, Array.from({ length: 5 }, () => ({ widthPt: 100, heightPt: 200 })));
    const v = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, { document: engine.asDoc() });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    expect(v.mountedPageIndicesForTest().length).toBe(0); // deferred
    container.clientWidth = 300;
    scrollHost.clientWidth = 300;
    scrollHost.clientHeight = 400;
    v.resizeForTest();
    expect(v.mountedPageIndicesForTest().length).toBeGreaterThan(0);
    v.destroy();
  });

  it('empty document: pageCount 0 ⇒ spacer 0, no slots, scrollToPage no-op', () => {
    installDom();
    const container = makeContainer(300, 400);
    const engine = new FakeDocxEngine(0, []);
    const v = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, { document: engine.asDoc() });
    v.relayout();
    const spacer = (container.children[0] as FakeEl).children[0].children[0] as FakeEl;
    expect(parseFloat(spacer.style.height || '0')).toBe(0);
    expect(v.mountedPageIndicesForTest().length).toBe(0);
    v.scrollToPage(0); // no throw
    v.destroy();
  });

  it('empty document: resize does not crash and fires no callback', () => {
    installDom();
    const container = makeContainer(0, 0);
    const engine = new FakeDocxEngine(0, []);
    const changes: number[] = [];
    const v = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      onVisiblePageChange: (i: number) => changes.push(i),
    });
    container.clientWidth = 300;
    (container.children[0] as FakeEl).children[0].clientWidth = 300;
    v.resizeForTest(); // no pages ⇒ no-op
    expect(v.mountedPageIndicesForTest().length).toBe(0);
    expect(changes).toEqual([]);
    v.destroy();
  });

  it('destroy() disconnects the ResizeObserver', () => {
    installDom();
    let disconnected = 0;
    class SpyRO {
      cb: () => void;
      constructor(cb: () => void) {
        this.cb = cb;
      }
      observe(): void {}
      disconnect(): void {
        disconnected++;
      }
    }
    vi.stubGlobal('ResizeObserver', SpyRO);
    const container = makeContainer(200, 400);
    const engine = new FakeDocxEngine(3, [{ widthPt: 100, heightPt: 200 }]);
    const v = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, { document: engine.asDoc() });
    v.destroy();
    expect(disconnected).toBe(1);
  });
});

describe('DocxScrollViewer — paddingTop/paddingBottom (desk margin)', () => {
  const PT_TO_PX = 4 / 3;
  const BASE = 200 / (100 * PT_TO_PX); // 1.5
  const PAGE_H = 200 * PT_TO_PX * BASE; // 400
  const GAP = 10;
  const STRIDE = PAGE_H + GAP; // 410

  function setup(opts = {}, pageCount = 20) {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakeDocxEngine(
      pageCount,
      Array.from({ length: pageCount }, () => ({ widthPt: 100, heightPt: 200 })),
    );
    const v = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      gap: GAP,
      // This block tests the VERTICAL desk margin; pin the horizontal gutters to 0
      // so the fit width is the full 200 container ⇒ BASE 1.5 / PAGE_H 400 / STRIDE
      // 410 (the default `gap` gutters would otherwise shrink the fit).
      paddingLeft: 0,
      paddingRight: 0,
      ...opts,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    return { v, scrollHost, container, engine };
  }

  /** The wrapper currently mounted for page `i` (identified by its top offset). */
  function slotTopFor(scrollHost: FakeEl, i: number, stride: number, padTop: number): FakeEl | undefined {
    const want = `${padTop + i * stride}px`;
    return scrollHost.children.find(
      (c) => c.tag === 'div' && c.children.some((k) => k.tag === 'canvas') && c.style.top === want,
    ) as FakeEl | undefined;
  }

  it('explicit paddingTop mounts the first slot at top = paddingTop px', () => {
    const { v, scrollHost } = setup({ paddingTop: 24, paddingBottom: 24 });
    expect(v.mountedPageIndicesForTest()).toContain(0);
    // Page 0's wrapper sits at top = paddingTop (not flush 0).
    expect(slotTopFor(scrollHost, 0, STRIDE, 24)).toBeDefined();
    v.destroy();
  });

  it('spacer height = padTop + Σheights + (n-1)*gap + padBottom', () => {
    const { scrollHost } = setup({ paddingTop: 24, paddingBottom: 40 });
    const spacer = scrollHost.children[0] as FakeEl;
    const expected = 24 + 20 * PAGE_H + 19 * GAP + 40; // 8254
    expect(parseFloat(spacer.style.height)).toBeCloseTo(expected, 3);
  });

  it('DEFAULT paddingTop/paddingBottom = gap (uniform rhythm: no options ⇒ first slot at gap px)', () => {
    // No paddingTop/paddingBottom → each defaults to `gap` (10). Page 0 sits at
    // top:10px, NOT flush 0 (this is the sanctioned pre-release default change).
    const { v, scrollHost } = setup();
    expect(v.mountedPageIndicesForTest()).toContain(0);
    expect(slotTopFor(scrollHost, 0, STRIDE, GAP)).toBeDefined();
    // And the spacer includes both default pads.
    const spacer = scrollHost.children[0] as FakeEl;
    const expected = GAP + 20 * PAGE_H + 19 * GAP + GAP;
    expect(parseFloat(spacer.style.height)).toBeCloseTo(expected, 3);
    v.destroy();
  });

  it('explicit 0 ⇒ flush (old behavior reachable): first slot at top 0, spacer has no pad', () => {
    const { v, scrollHost } = setup({ paddingTop: 0, paddingBottom: 0 });
    expect(slotTopFor(scrollHost, 0, STRIDE, 0)).toBeDefined();
    const spacer = scrollHost.children[0] as FakeEl;
    const expected = 20 * PAGE_H + 19 * GAP; // no pad
    expect(parseFloat(spacer.style.height)).toBeCloseTo(expected, 3);
    v.destroy();
  });

  it('scrollToPage(0) lands on offsets[0] (= paddingTop, not 0)', () => {
    const { v, scrollHost } = setup({ paddingTop: 24, paddingBottom: 24 });
    // Move away first, then navigate to page 0.
    scrollHost.scrollTop = 3 * STRIDE;
    scrollHost.dispatch('scroll');
    v.scrollToPage(0);
    // offsets[0] = paddingTop (24), so the top edge of page 0 sits below the pad.
    expect(scrollHost.scrollTop).toBe(24);
    v.destroy();
  });

  it('scrollToPage(k) lands on paddingTop + k*stride', () => {
    const { v, scrollHost } = setup({ paddingTop: 24, paddingBottom: 24 });
    v.scrollToPage(3);
    expect(Math.abs(scrollHost.scrollTop - (24 + 3 * STRIDE))).toBeLessThan(2);
    v.destroy();
  });

  it('re-anchor keeps the page under the viewport top fixed WITH padding intact after setScale', () => {
    const { v, scrollHost } = setup({ paddingTop: 24, paddingBottom: 24, zoomMin: 0.5, zoomMax: 3 });
    // Scroll so page 3's top sits at the viewport top: offset[3] = 24 + 3*410 = 1254.
    scrollHost.scrollTop = 24 + 3 * STRIDE;
    scrollHost.dispatch('scroll');
    expect(v.topVisiblePage).toBe(3);
    v.setScale(v.scaleForTest() * 2); // 1.5 → 3.0; PAGE_H' = 800, STRIDE' = 810
    expect(v.scaleForTest()).toBeCloseTo(3, 5);
    // Page 3 stays the top page; padding is intact so offset'[3] = 24 + 3*810 = 2454.
    expect(v.topVisiblePage).toBe(3);
    expect(Math.abs(scrollHost.scrollTop - (24 + 3 * 810))).toBeLessThan(2);
    v.destroy();
  });

  it('keeps the leading desk padding visible when setScale runs at the top', () => {
    const { v, scrollHost } = setup({ paddingTop: 24, paddingBottom: 24, zoomMin: 0.5, zoomMax: 3 });
    expect(scrollHost.scrollTop).toBe(0);
    v.setScale(v.scaleForTest() * 0.75);
    expect(scrollHost.scrollTop).toBe(0);
    expect(slotTopFor(scrollHost, 0, PAGE_H * 0.75 + GAP, 24)).toBeDefined();
    v.destroy();
  });
});

describe('DocxScrollViewer — paddingLeft/paddingRight (horizontal desk gutters)', () => {
  const PT_TO_PX = 4 / 3;

  /** Build a viewer over a container of `cw`×400, uniform 100pt×200pt pages. */
  function setup(cw: number, opts = {}, pageCount = 20) {
    installDom();
    const container = makeContainer(cw, 400);
    const engine = new FakeDocxEngine(
      pageCount,
      Array.from({ length: pageCount }, () => ({ widthPt: 100, heightPt: 200 })),
    );
    const v = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      gap: 16,
      ...opts,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = cw;
    v.relayout();
    return { v, scrollHost, container, engine };
  }

  /** The wrapper mounted for page 0 (flush top ⇒ it sits at top:0px unless a
   *  vertical pad is set; callers that keep the default gap pass the expected top). */
  function slot(scrollHost: FakeEl, top: string): FakeEl | undefined {
    return scrollHost.children.find(
      (c) => c.tag === 'div' && c.children.some((k) => k.tag === 'canvas') && c.style.top === top,
    ) as FakeEl | undefined;
  }

  it('fit subtracts the DEFAULT gutters: container 232 with default gap-16 gutters ⇒ fit 200 ⇒ base 1.5', () => {
    // 232 − 16 (padL) − 16 (padR) = 200, the same fit the old 200-container tests
    // used, so the base scale matches their 1.5 (page widthPt 100 → 200/(100×4/3)).
    const { v } = setup(232);
    expect(v.baseScaleForTest()).toBeCloseTo(1.5, 5);
    expect(v.scaleForTest()).toBeCloseTo(1.5, 5);
    v.destroy();
  });

  it('explicit paddingLeft:0/paddingRight:0 ⇒ old full-width fit reachable (base uses the FULL container width)', () => {
    // With the gutters pinned to 0 the fit is the full 200 container width ⇒ base
    // 1.5, exactly the pre-feature behavior.
    const { v } = setup(200, { paddingLeft: 0, paddingRight: 0 });
    expect(v.baseScaleForTest()).toBeCloseTo(1.5, 5);
    v.destroy();
  });

  it('explicit opts.width is NOT reduced by the gutters (it is the page CSS-width contract)', () => {
    // opts.width 150 stays 150 regardless of the gutters: base = 150/(100×4/3) =
    // 1.125. The gutters still apply to PLACEMENT (asserted in the centering test),
    // just not to the width.
    const { v } = setup(400, { width: 150, paddingLeft: 24, paddingRight: 24 });
    expect(v.baseScaleForTest()).toBeCloseTo(150 / (100 * PT_TO_PX), 5);
    v.destroy();
  });

  it('slot left = paddingLeft when the page fills the fit exactly (symmetric gutters ⇒ centre lands on the floor)', () => {
    // Container 232, default gutters 16 ⇒ fit 200, page px = 200. scrollHost
    // clientWidth 232, so centre = (232 − 200)/2 = 16 = padL ⇒ left pinned at padL.
    const { v, scrollHost } = setup(232);
    // Default vertical padding = gap 16 ⇒ page 0 sits at top:16px.
    const s0 = slot(scrollHost, '16px');
    expect(s0).toBeDefined();
    expect(s0!.style.left).toBe('16px'); // = paddingLeft
    v.destroy();
  });

  it('slot is CENTERED when the page is narrower than the viewport (left = (cw − pw)/2 > paddingLeft)', () => {
    // Explicit width 100 (NOT reduced) ⇒ page px = 100, far narrower than the 400
    // viewport. centre = (400 − 100)/2 = 150, which exceeds padL 16, so the page is
    // centred at 150 rather than pinned to the gutter floor.
    const { v, scrollHost } = setup(400, { width: 100, paddingLeft: 16, paddingRight: 16 });
    const s0 = slot(scrollHost, '16px'); // default vertical pad = gap 16
    expect(s0).toBeDefined();
    expect(parseFloat(s0!.style.left)).toBeCloseTo(150, 3);
    expect(parseFloat(s0!.style.left)).toBeGreaterThan(16);
    v.destroy();
  });

  it('zoomed-in (page wider than viewport): left pins to paddingLeft and the spacer width = pageW + padL + padR', () => {
    // Explicit width 400 in a 200 viewport ⇒ page px 400 > cw 200. centre =
    // (200 − 400)/2 = −100, so the floor pins left at padL 24. The horizontal scroll
    // extent (spacer width) is the page width plus both gutters: 400 + 24 + 24 = 448.
    const { v, scrollHost } = setup(200, { width: 400, paddingLeft: 24, paddingRight: 24 });
    const s0 = slot(scrollHost, '16px'); // default vertical pad = gap 16
    expect(s0).toBeDefined();
    expect(s0!.style.left).toBe('24px'); // pinned to paddingLeft
    const spacer = scrollHost.children[0] as FakeEl;
    expect(parseFloat(spacer.style.width)).toBeCloseTo(400 + 24 + 24, 3);
    v.destroy();
  });

  it('a Ctrl+wheel zoom that widens the page past the viewport updates the spacer width and pins left to paddingLeft', () => {
    // Container 200 pinned gutters 0 ⇒ fit 200, base 1.5, page px 200 (== viewport).
    // Zoom ×2 ⇒ page px 400 > 200 viewport. left floors to padL (0 here) and the
    // spacer width tracks the new page width + gutters (400 + 0 + 0 = 400).
    const { v, scrollHost } = setup(200, {
      paddingLeft: 0,
      paddingRight: 0,
      paddingTop: 0,
      zoomMin: 0.5,
      zoomMax: 4,
    });
    expect(v.scaleForTest()).toBeCloseTo(1.5, 5);
    v.setScale(v.scaleForTest() * 2); // page px 200 → 400
    const s0 = scrollHost.children.find(
      (c) => c.tag === 'div' && c.children.some((k) => k.tag === 'canvas') && c.style.top === '0px',
    ) as FakeEl | undefined;
    expect(s0).toBeDefined();
    expect(s0!.style.left).toBe('0px'); // paddingLeft 0
    const spacer = scrollHost.children[0] as FakeEl;
    expect(parseFloat(spacer.style.width)).toBeCloseTo(400, 3);
    v.destroy();
  });

  it('spacer width uses the WIDEST page over mixed page widths (+ both gutters)', () => {
    installDom();
    const container = makeContainer(200, 400);
    // Page 0 100pt wide (base fit target), page 1 200pt wide (twice as wide).
    const engine = new FakeDocxEngine(2, [
      { widthPt: 100, heightPt: 100 },
      { widthPt: 200, heightPt: 100 },
    ]);
    const v = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      gap: 10,
      paddingLeft: 12,
      paddingRight: 12,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    // fit = 200 − 12 − 12 = 176, base = 176/(100×4/3) = 1.32. Page widths px: page 0
    // = 176, page 1 = 200×4/3×1.32 = 352 (the widest). Spacer width = 352 + 12 + 12.
    const base = 176 / (100 * PT_TO_PX);
    const widest = 200 * PT_TO_PX * base; // 352
    const spacer = scrollHost.children[0] as FakeEl;
    expect(parseFloat(spacer.style.width)).toBeCloseTo(widest + 12 + 12, 3);
    v.destroy();
  });

  it('fits to the scrollbar-reduced scrollport width without horizontal overflow', () => {
    const { v, container, scrollHost } = setup(232);
    // A classic vertical scrollbar consumes 17px inside the scrollport while the
    // outer container remains 232px wide. Re-fit must use 215, yielding page 183
    // + two 16px gutters = exactly 215px, not the stale outer width.
    scrollHost.clientWidth = 215;
    v.resizeForTest();
    const spacer = scrollHost.children[0] as FakeEl;
    expect(parseFloat(spacer.style.width)).toBeCloseTo(215, 3);
    expect(container.clientWidth).toBe(232);
    v.destroy();
  });

  it('gutters wider than the container defer layout (non-positive fit ⇒ zero-width deferral)', () => {
    // padL + padR = 300 > container 200 ⇒ fit ≤ 0 ⇒ deferred (no slots mounted),
    // mirroring the zero-width container deferral.
    const { v } = setup(200, { paddingLeft: 150, paddingRight: 150 });
    expect(v.mountedPageIndicesForTest().length).toBe(0);
    v.destroy();
  });

  it('the slot wrapper carries NO CSS auto-centering (explicit JS left replaces margin:0 auto)', () => {
    // The old wrapper used `left:0;right:0;margin:0 auto`; that would fight the
    // explicit per-mount `left`. Assert the auto-centering margin is gone.
    const { v, scrollHost } = setup(232);
    const s0 = slot(scrollHost, '16px');
    expect(s0).toBeDefined();
    expect(s0!.style.margin).toBe(''); // no `0 auto`
    expect(s0!.style.right).toBe(''); // no right:0 pinning
    v.destroy();
  });
});

describe('DocxScrollViewer — flicker-free zoom (T8)', () => {
  // Flicker-free zoom (design §7): setScale must NOT blank a visible page.
  // Three mechanisms — CSS preview (immediate, no device-buffer resize / no
  // recycle of in-window slots), settle re-render (debounced ZOOM_SETTLE_MS),
  // double-buffer swap (main-mode settle renders into a SPARE off-DOM canvas and
  // swaps it in). PT_TO_PX 4/3, uniform 100pt×200pt pages, container 200×400,
  // flush pads ⇒ base 1.5, PAGE_H 400.
  const PT_TO_PX = 4 / 3;

  function setup(pageCount = 20, opts = {}, mode: 'main' | 'worker' = 'main', deferred = false) {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakeDocxEngine(
      pageCount,
      Array.from({ length: pageCount }, () => ({ widthPt: 100, heightPt: 200 })),
      mode,
      deferred,
    );
    const v = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      gap: 10,
      paddingTop: 0,
      paddingBottom: 0,
      paddingLeft: 0,
      paddingRight: 0,
      overscan: 1,
      zoomMin: 0.5,
      zoomMax: 4,
      ...opts,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    return { v, scrollHost, engine, container };
  }

  /** The wrapper mounted for the page currently at top:`top`px. */
  function slotAtTop(scrollHost: FakeEl, top: string): FakeEl | undefined {
    return scrollHost.children.find(
      (c) => c.tag === 'div' && c.children.some((k) => k.tag === 'canvas') && c.style.top === top,
    ) as FakeEl | undefined;
  }

  // (a) setScale does NOT unmount in-window slots (same wrapper before/after) and
  //     does NOT resize any on-screen canvas device buffer during the preview.
  it('CSS preview: setScale keeps the SAME in-window slot wrappers and never resizes their device buffer', () => {
    const { v, scrollHost } = setup();
    // The page-0 slot mounted at top:0 with its canvas.
    const before = slotAtTop(scrollHost, '0px');
    expect(before).toBeDefined();
    const beforeCanvas = before!.children.find((k) => k.tag === 'canvas') as FakeEl;
    const resizesBefore = beforeCanvas._deviceResizes.length;

    v.setScale(v.scaleForTest() * 2); // zoom in

    // The very SAME wrapper element is still mounted for page 0 (not recycled +
    // re-mounted). Its canvas is the SAME element (identity preserved).
    const after = slotAtTop(scrollHost, '0px');
    expect(after).toBe(before);
    const afterCanvas = after!.children.find((k) => k.tag === 'canvas') as FakeEl;
    expect(afterCanvas).toBe(beforeCanvas);
    // No device-buffer resize during the CSS preview (the flicker cause). The
    // preview only sets style.width/height.
    expect(afterCanvas._deviceResizes.length).toBe(resizesBefore);
    v.destroy();
  });

  it('CSS preview: setScale CSS-resizes the slot canvas (style.width/height) to the new layout size immediately', () => {
    const { v, scrollHost } = setup();
    const slot = slotAtTop(scrollHost, '0px')!;
    const canvas = slot.children.find((k) => k.tag === 'canvas') as FakeEl;
    // base 1.5 ⇒ page px width 200. Zoom ×2 ⇒ CSS width 400.
    v.setScale(v.scaleForTest() * 2);
    expect(canvas.style.width).toBe('400px');
    // height: 200pt × PT_TO_PX × 3.0 = 800.
    expect(canvas.style.height).toBe('800px');
    v.destroy();
  });

  it('CSS preview: text layer gets a transform: scale(ratio) matching newScale / renderedScale', async () => {
    const { v, scrollHost, engine } = setup(20, { enableTextSelection: true });
    engine.feedTextRuns = [{ text: 'Hi', x: 1, y: 2, w: 10, h: 12, fontSize: 12, font: '12px serif' }];
    // The overlay is built in the renderPage .then() microtask.
    await Promise.resolve();
    await Promise.resolve();
    const slot = slotAtTop(scrollHost, '0px')!;
    const textLayer = slot.children.find((k) => k.tag === 'div') as FakeEl;
    // Zoom ×2 — the overlay was built at the base scale (1.5); the preview scales
    // it by newScale/renderedScale = 3.0/1.5 = 2.
    v.setScale(v.scaleForTest() * 2);
    expect(textLayer.style.transform).toBe('scale(2)');
    expect(textLayer.style.transformOrigin).toBe('0 0');
    v.destroy();
  });

  // (c) NO render dispatch during a burst; ONE dispatch after ZOOM_SETTLE_MS.
  it('debounce: a burst of setScale calls dispatches NO settle render until ZOOM_SETTLE_MS elapses, then exactly one per slot', () => {
    vi.useFakeTimers();
    try {
      const { v, engine } = setup();
      const dispatchesBefore = engine.renderCalls.length; // initial mount renders
      // Three rapid setScale calls within the settle window.
      v.setScale(v.scaleForTest() * 1.1);
      vi.advanceTimersByTime(50);
      v.setScale(v.scaleForTest() * 1.1);
      vi.advanceTimersByTime(50);
      v.setScale(v.scaleForTest() * 1.1);
      // Still inside the settle window (150ms): NO settle re-render dispatched.
      vi.advanceTimersByTime(50); // total 150 since first, but timer reset each tick
      expect(engine.renderCalls.length).toBe(dispatchesBefore);
      // Advance past the settle timeout measured from the LAST setScale.
      vi.advanceTimersByTime(150);
      // Now the visible window re-rendered exactly once each (no per-tick storm).
      const mounted = v.mountedPageIndicesForTest();
      const settleRenders = engine.renderCalls.length - dispatchesBefore;
      expect(settleRenders).toBe(mounted.length);
    } finally {
      vi.useRealTimers();
    }
    // no v.destroy() under real timers needed — installDom stubs are cleaned in afterEach
  });

  // (d) settle resolution swaps the canvas without a blank: main-mode settle
  //     renders into a SPARE canvas (device-buffer resize lands on the spare, not
  //     the on-screen canvas), then swaps it into the wrapper. The old canvas
  //     stays attached until the fresh one is painted.
  it('double-buffer (main): settle renders into a SPARE canvas and swaps it in — the on-screen canvas is never device-resized in place', () => {
    vi.useFakeTimers();
    try {
      const { v, scrollHost, engine } = setup(20, {}, 'main', true);
      // Resolve the initial mount renders so the on-screen canvas is "painted".
      for (const c of engine.renderCalls.slice()) c.resolve();
      const slot = slotAtTop(scrollHost, '0px')!;
      const onScreenCanvas = slot.children.find((k) => k.tag === 'canvas') as FakeEl;
      const onScreenUid = onScreenCanvas._uid;
      const resizesOnScreenBefore = onScreenCanvas._deviceResizes.length;

      v.setScale(v.scaleForTest() * 2);
      vi.advanceTimersByTime(200); // past ZOOM_SETTLE_MS

      // A settle render was dispatched for page 0 into a canvas that is NOT the
      // on-screen one (a fresh spare with a higher uid).
      const settleCall = engine.renderCalls.filter((c) => c.page === 0).pop()!;
      expect(settleCall.canvas).toBeDefined();
      expect(settleCall.canvas!._uid).not.toBe(onScreenUid);
      // The on-screen canvas's device buffer was NOT resized in place by the settle
      // (the render — and its synchronous clear — targeted the spare).
      expect(onScreenCanvas._deviceResizes.length).toBe(resizesOnScreenBefore);
      // The on-screen canvas is still attached (blank-free): the spare has not yet
      // resolved, so no swap happened.
      expect(slot.children).toContain(onScreenCanvas);

      // Resolve the spare render → swap it into the wrapper, retire the old canvas.
      settleCall.resolve();
      // The swap runs in the render .then() microtask.
      return Promise.resolve()
        .then(() => Promise.resolve())
        .then(() => {
          const nowCanvas = slot.children.find((k) => k.tag === 'canvas') as FakeEl;
          expect(nowCanvas._uid).toBe(settleCall.canvas!._uid); // spare swapped in
          expect(slot.children).not.toContain(onScreenCanvas); // old retired
          v.destroy();
        });
    } finally {
      vi.useRealTimers();
    }
  });

  it('double-buffer (worker): settle renders into the SAME canvas (sync resize+transfer paints no blank frame)', async () => {
    vi.useFakeTimers();
    const { v, scrollHost, engine } = setup(20, {}, 'worker', true);
    // Resolve the initial mount bitmap renders.
    for (const c of engine.bitmapCalls.slice()) c.resolve();
    await Promise.resolve();
    await Promise.resolve();
    const slot = slotAtTop(scrollHost, '0px')!;
    const canvas = slot.children.find((k) => k.tag === 'canvas') as FakeEl;
    const canvasUid = canvas._uid;
    const bitmapsBefore = engine.bitmapCalls.length;

    v.setScale(v.scaleForTest() * 2);
    vi.advanceTimersByTime(200); // past ZOOM_SETTLE_MS

    // Worker settle re-dispatched a bitmap render for the visible window (no spare
    // canvas created — the sync resize+transfer never paints an intermediate blank).
    expect(engine.bitmapCalls.length).toBeGreaterThan(bitmapsBefore);
    const settleCall = engine.bitmapCalls.filter((c) => c.page === 0).pop()!;
    // createdBitmaps is index-parallel to bitmapCalls (both pushed in the same
    // renderPageToBitmap call), so this is the bitmap for THIS settle dispatch.
    const settleBitmap = engine.createdBitmaps[engine.bitmapCalls.indexOf(settleCall)];
    settleCall.resolve();
    await Promise.resolve();
    await Promise.resolve();
    // The bitmap was transferred into the SAME on-screen canvas element (identity
    // preserved — worker mode keeps the same canvas).
    const nowCanvas = slot.children.find((k) => k.tag === 'canvas') as FakeEl;
    expect(nowCanvas._uid).toBe(canvasUid);
    expect(nowCanvas._bitmapCtx?.lastBitmap).toBe(settleBitmap);
    vi.useRealTimers();
    v.destroy();
  });

  // (e) A settle whose epoch is superseded (another setScale during the settle
  //     render) is discarded per the existing epoch gate.
  it('stale settle: an epoch bump during the settle render discards the settle (no swap, spare not attached)', () => {
    vi.useFakeTimers();
    try {
      const { v, scrollHost, engine } = setup(20, {}, 'main', true);
      for (const c of engine.renderCalls.slice()) c.resolve();
      const slot = slotAtTop(scrollHost, '0px')!;
      const onScreenCanvas = slot.children.find((k) => k.tag === 'canvas') as FakeEl;

      v.setScale(v.scaleForTest() * 2);
      vi.advanceTimersByTime(200); // settle dispatched (deferred, in flight)
      const settleCall = engine.renderCalls.filter((c) => c.page === 0).pop()!;

      // Another setScale bumps the epoch WHILE the settle render is in flight.
      v.setScale(v.scaleForTest() * 1.1);

      // The old settle now resolves — epoch moved ⇒ STALE ⇒ no swap.
      settleCall.resolve();
      return Promise.resolve()
        .then(() => Promise.resolve())
        .then(() => {
          const nowCanvas = slot.children.find((k) => k.tag === 'canvas') as FakeEl;
          // The stale spare was NOT swapped in; the on-screen canvas still stands.
          expect(nowCanvas).toBe(onScreenCanvas);
          expect(nowCanvas._uid).not.toBe(settleCall.canvas!._uid);
          v.destroy();
        });
    } finally {
      vi.useRealTimers();
    }
  });

  // (f) destroy during a pending settle timer: no crash, timer cleared, no
  //     post-destroy dispatch.
  it('destroy during a pending settle timer clears it — no post-destroy render dispatch, no crash', () => {
    vi.useFakeTimers();
    try {
      const { v, engine } = setup();
      const dispatchesBefore = engine.renderCalls.length;
      v.setScale(v.scaleForTest() * 2); // schedules a settle timer
      v.destroy(); // must clear the pending timer
      vi.advanceTimersByTime(500); // fire any surviving timer
      // No settle render dispatched after destroy.
      expect(engine.renderCalls.length).toBe(dispatchesBefore);
    } finally {
      vi.useRealTimers();
    }
  });

  it('destroy during an in-flight settle render does not swap or fire onError post-destroy', () => {
    vi.useFakeTimers();
    const onError = vi.fn();
    try {
      const { v, scrollHost, engine } = setup(20, { onError }, 'main', true);
      for (const c of engine.renderCalls.slice()) c.resolve();
      const slot = slotAtTop(scrollHost, '0px')!;
      const onScreenCanvas = slot.children.find((k) => k.tag === 'canvas') as FakeEl;
      v.setScale(v.scaleForTest() * 2);
      vi.advanceTimersByTime(200);
      const settleCall = engine.renderCalls.filter((c) => c.page === 0).pop()!;
      v.destroy(); // tear down while the settle render is in flight
      settleCall.resolve();
      return Promise.resolve()
        .then(() => Promise.resolve())
        .then(() => {
          // No crash, no error surfaced after teardown, and the (now detached)
          // wrapper never received a swapped spare.
          expect(onError).not.toHaveBeenCalled();
          expect(slot.children).toContain(onScreenCanvas);
        });
    } finally {
      vi.useRealTimers();
    }
  });

  it('settle clears the preview transform on the text layer (overlay rebuilt at full resolution)', async () => {
    vi.useFakeTimers();
    const { v, scrollHost, engine } = setup(20, { enableTextSelection: true }, 'main', true);
    engine.feedTextRuns = [{ text: 'Hi', x: 1, y: 2, w: 10, h: 12, fontSize: 12, font: '12px serif' }];
    for (const c of engine.renderCalls.slice()) c.resolve();
    await Promise.resolve();
    await Promise.resolve();
    const slot = slotAtTop(scrollHost, '0px')!;
    const textLayer = slot.children.find((k) => k.tag === 'div') as FakeEl;

    v.setScale(v.scaleForTest() * 2);
    expect(textLayer.style.transform).toBe('scale(2)'); // preview transform applied
    vi.advanceTimersByTime(200);
    const settleCall = engine.renderCalls.filter((c) => c.page === 0).pop()!;
    settleCall.resolve();
    await Promise.resolve();
    await Promise.resolve();
    // After settle, the preview transform is cleared (the overlay is rebuilt at the
    // full resolution so it no longer needs the scale()).
    expect(textLayer.style.transform).toBe('');
    vi.useRealTimers();
    v.destroy();
  });
});

describe('DocxScrollViewer — pageShadow (T9)', () => {
  // pageShadow paints a CSS box-shadow on every page CANVAS (not the wrapper).
  // Default = the recipe drop shadow; `false` disables; a custom string is
  // honoured verbatim. It must reach EVERY canvas: fresh mounts, the settle
  // double-buffer SPARE, and recycled+re-mounted slots. Reuses the T8 setup
  // (deferred engine + fake timers) for the spare-canvas case.
  const DEFAULT_PAGE_SHADOW = '0 1px 3px rgba(0,0,0,0.2)';

  function setup(opts = {}, mode: 'main' | 'worker' = 'main', deferred = false) {
    installDom();
    const container = makeContainer(200, 400);
    const engine = new FakeDocxEngine(
      20,
      Array.from({ length: 20 }, () => ({ widthPt: 100, heightPt: 200 })),
      mode,
      deferred,
    );
    const v = makeBorrowedDocxScrollViewer(container as unknown as HTMLElement, {
      document: engine.asDoc(),
      gap: 10,
      paddingTop: 0,
      paddingBottom: 0,
      paddingLeft: 0,
      paddingRight: 0,
      overscan: 1,
      zoomMin: 0.5,
      zoomMax: 4,
      ...opts,
    });
    const scrollHost = (container.children[0] as FakeEl).children[0] as FakeEl;
    scrollHost.clientHeight = 400;
    scrollHost.clientWidth = 200;
    v.relayout();
    return { v, scrollHost, engine };
  }

  /** Every mounted page canvas (each slot wrapper's canvas child). */
  function mountedCanvases(scrollHost: FakeEl): FakeEl[] {
    return scrollHost.children
      .filter((c) => c.tag === 'div' && c.children.some((k) => k.tag === 'canvas'))
      .map((w) => w.children.find((k) => k.tag === 'canvas') as FakeEl);
  }

  it('applies the default recipe shadow to every mounted page canvas', () => {
    const { v, scrollHost } = setup();
    const canvases = mountedCanvases(scrollHost);
    expect(canvases.length).toBeGreaterThan(0);
    for (const c of canvases) expect(c.style.boxShadow).toBe(DEFAULT_PAGE_SHADOW);
    v.destroy();
  });

  it('pageShadow:false sets no box-shadow on the page canvas', () => {
    const { v, scrollHost } = setup({ pageShadow: false });
    const canvases = mountedCanvases(scrollHost);
    expect(canvases.length).toBeGreaterThan(0);
    for (const c of canvases) expect(c.style.boxShadow).toBe('');
    v.destroy();
  });

  it('honours a custom pageShadow string verbatim (e.g. a spread-ring border look)', () => {
    const ring = '0 0 0 1px #c8ccd0';
    const { v, scrollHost } = setup({ pageShadow: ring });
    const canvases = mountedCanvases(scrollHost);
    expect(canvases.length).toBeGreaterThan(0);
    for (const c of canvases) expect(c.style.boxShadow).toBe(ring);
    v.destroy();
  });

  it('does NOT paint the shadow on the wrapper (only the canvas)', () => {
    const { v, scrollHost } = setup();
    const wrapper = scrollHost.children.find(
      (c) => c.tag === 'div' && c.children.some((k) => k.tag === 'canvas'),
    ) as FakeEl;
    expect(wrapper.style.boxShadow).toBe('');
    v.destroy();
  });

  it('the settle double-buffer SPARE canvas carries the shadow (swapped-in canvas keeps it)', () => {
    vi.useFakeTimers();
    try {
      const { v, scrollHost, engine } = setup({}, 'main', true);
      // Resolve the initial mount renders so the on-screen canvases are painted.
      for (const c of engine.renderCalls.slice()) c.resolve();

      v.setScale(v.scaleForTest() * 2); // zoom → schedules a settle
      vi.advanceTimersByTime(200); // past ZOOM_SETTLE_MS → settle dispatched

      // The settle render for page 0 targets a fresh SPARE canvas — it must
      // already carry the shadow BEFORE the swap.
      const settleCall = engine.renderCalls.filter((c) => c.page === 0).pop()!;
      const spare = settleCall.canvas as unknown as FakeEl;
      expect(spare).toBeDefined();
      expect(spare.style.boxShadow).toBe(DEFAULT_PAGE_SHADOW);

      // Resolve → swap the spare in. The now-on-screen canvas still has the shadow.
      settleCall.resolve();
      const slot = scrollHost.children.find(
        (c) => c.tag === 'div' && c.style.top === '0px' && c.children.some((k) => k.tag === 'canvas'),
      ) as FakeEl;
      return Promise.resolve()
        .then(() => Promise.resolve())
        .then(() => {
          const nowCanvas = slot.children.find((k) => k.tag === 'canvas') as FakeEl;
          expect(nowCanvas._uid).toBe(spare._uid); // spare swapped in
          expect(nowCanvas.style.boxShadow).toBe(DEFAULT_PAGE_SHADOW);
          v.destroy();
        });
    } finally {
      vi.useRealTimers();
    }
  });

  it('a recycled + re-mounted slot still carries the shadow', () => {
    const { v, scrollHost } = setup();
    // Scroll far down so page 0's slot recycles into the free pool, then back to
    // the top so a POOLED slot is re-mounted for page 0.
    scrollHost.scrollTop = 5000;
    scrollHost.dispatch('scroll');
    scrollHost.scrollTop = 0;
    scrollHost.dispatch('scroll');
    const canvases = mountedCanvases(scrollHost);
    expect(canvases.length).toBeGreaterThan(0);
    for (const c of canvases) expect(c.style.boxShadow).toBe(DEFAULT_PAGE_SHADOW);
    v.destroy();
  });
});

describe('DocxScrollViewer — barrel export (T7)', () => {
  it('is exported from the package entry', () => {
    expect(typeof docxIndex.DocxScrollViewer).toBe('function');
  });
});
