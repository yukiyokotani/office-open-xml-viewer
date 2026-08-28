import { afterEach, describe, expect, it, vi } from 'vitest';
import { DocxDocument } from './document.js';
import { publishDocxLayout } from './document-layout-events.js';
import { DocxViewer } from './viewer.js';
import { FakeDocxEngine, installDom, makeEl } from './scroll-viewer-test-dom.js';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

const PAGE = [{ widthPt: 612, heightPt: 792 }];

describe('DocxViewer progressive layout', () => {
  it('renders the first publication and waits visibly for a requested page to arrive', async () => {
    installDom();
    const engine = new FakeDocxEngine(1, PAGE);
    let complete = false;
    let settle!: () => void;
    const completion = new Promise<void>((resolve) => { settle = resolve; });
    const doc = engine.asDoc() as DocxDocument;
    Object.defineProperties(doc, {
      layoutComplete: { configurable: true, get: () => complete },
      waitUntilLayoutComplete: { configurable: true, value: () => completion },
    });
    const load = vi.spyOn(DocxDocument, 'load').mockResolvedValue(doc);
    const pageChanges: Array<[number, number, boolean]> = [];
    const viewer = new DocxViewer(makeEl('canvas') as unknown as HTMLCanvasElement, {
      progressiveLayout: true,
      onPageChange: (page, total, layoutComplete) => {
        pageChanges.push([page, total, layoutComplete]);
      },
    });

    await viewer.load('large.docx');
    expect(load.mock.calls[0]?.[1]).toMatchObject({ progressiveLayout: true });
    expect(pageChanges.at(-1)).toEqual([0, 1, false]);

    const navigation = viewer.nextPage();
    await Promise.resolve();
    expect((viewer as unknown as { _loadingLayer: { style: Record<string, string> } })
      ._loadingLayer.style.display).toBe('flex');

    engine.setPageCount(2);
    publishDocxLayout(doc, { pageCount: 2, exact: false, complete: false });
    await navigation;

    expect(viewer.currentPage).toBe(1);
    expect(engine.renderCalls.at(-1)?.page).toBe(1);
    expect(pageChanges.at(-1)).toEqual([1, 2, false]);
    expect((viewer as unknown as { _loadingLayer: { style: Record<string, string> } })
      ._loadingLayer.style.display).toBe('none');

    complete = true;
    engine.setPageCount(5);
    settle();
    publishDocxLayout(doc, { pageCount: 5, exact: true, complete: true });
    expect(viewer.layoutComplete).toBe(true);
    await expect(viewer.waitUntilLayoutComplete()).resolves.toBeUndefined();
    viewer.destroy();
  });

  it('stops listening to a document after a replacement commits', async () => {
    installDom();
    const first = new FakeDocxEngine(1, PAGE);
    const second = new FakeDocxEngine(2, PAGE);
    vi.spyOn(DocxDocument, 'load')
      .mockResolvedValueOnce(first.asDoc())
      .mockResolvedValueOnce(second.asDoc());
    const viewer = new DocxViewer(makeEl('canvas') as unknown as HTMLCanvasElement, {
      progressiveLayout: true,
    });
    await viewer.load('first.docx');
    await viewer.load('second.docx');
    const renders = second.renderCalls.length;

    first.setPageCount(40);
    publishDocxLayout(first.asDoc(), { pageCount: 40, exact: false, complete: false });
    await Promise.resolve();

    expect(viewer.pageCount).toBe(2);
    expect(second.renderCalls).toHaveLength(renders);
    viewer.destroy();
  });

  it('keeps listening to the installed document when a replacement fails', async () => {
    installDom();
    const retained = new FakeDocxEngine(1, PAGE);
    vi.spyOn(DocxDocument, 'load')
      .mockResolvedValueOnce(retained.asDoc())
      .mockRejectedValueOnce(new Error('replacement failed'));
    const viewer = new DocxViewer(makeEl('canvas') as unknown as HTMLCanvasElement, {
      progressiveLayout: true,
    });
    await viewer.load('retained.docx');
    await expect(viewer.load('bad.docx')).rejects.toThrow('replacement failed');
    const renders = retained.renderCalls.length;

    retained.setPageCount(3);
    publishDocxLayout(retained.asDoc(), { pageCount: 3, exact: false, complete: false });
    await Promise.resolve();

    expect(viewer.pageCount).toBe(3);
    expect(retained.renderCalls.length).toBeGreaterThan(renders);
    viewer.destroy();
  });

  it('settles an unavailable-page request when newer navigation supersedes it', async () => {
    installDom();
    const engine = new FakeDocxEngine(1, PAGE);
    const doc = engine.asDoc() as DocxDocument;
    Object.defineProperties(doc, {
      layoutComplete: { configurable: true, get: () => false },
      waitUntilLayoutComplete: { configurable: true, value: () => new Promise<void>(() => {}) },
    });
    vi.spyOn(DocxDocument, 'load').mockResolvedValue(doc);
    const viewer = new DocxViewer(makeEl('canvas') as unknown as HTMLCanvasElement, {
      progressiveLayout: true,
    });
    await viewer.load('large.docx');

    const superseded = viewer.goToPage(10);
    await Promise.resolve();
    expect((viewer as unknown as { _loadingLayer: { style: Record<string, string> } })
      ._loadingLayer.style.display).toBe('flex');

    await viewer.goToPage(0);
    await superseded;
    expect(viewer.currentPage).toBe(0);
    expect((viewer as unknown as { _loadingLayer: { style: Record<string, string> } })
      ._loadingLayer.style.display).toBe('none');
    viewer.destroy();
  });

  it('waits for authoritative pages before a full-document find', async () => {
    installDom();
    const engine = new FakeDocxEngine(1, PAGE);
    engine.setLayoutComplete(false);
    engine.feedTextRuns = [{
      text: 'opening', x: 0, y: 0, w: 10, h: 10, fontSize: 10, font: '10px serif',
    }];
    const doc = engine.asDoc();
    const viewer = DocxViewer.fromDocument(
      makeEl('canvas') as unknown as HTMLCanvasElement,
      doc,
    ) as DocxViewer;
    const findController = (viewer as unknown as { _find: { find: unknown } })._find;
    const laterMatch = {
      matchIndex: 0,
      text: 'later',
      location: { page: 2, runIndex: 0, startOffset: 0, endOffset: 5 },
    };
    const findSpy = vi.spyOn(
      findController as { find: (query: string) => Promise<unknown[]> },
      'find',
    ).mockResolvedValue([laterMatch]);
    const find = viewer.findText('later');
    await Promise.resolve();
    expect(engine.renderCalls.filter((call) => call.page > 0)).toHaveLength(0);

    engine.feedTextRuns = [{
      text: 'later', x: 0, y: 0, w: 10, h: 10, fontSize: 10, font: '10px serif',
    }];
    engine.setPageCount(3);
    engine.setLayoutComplete(true);
    publishDocxLayout(doc, { pageCount: 3, exact: true, complete: true });
    await expect(find).resolves.toEqual([laterMatch]);
    expect(findSpy).toHaveBeenCalledWith('later', {});
    viewer.destroy();
  });

  it('does not start a deferred find after clearFind()', async () => {
    installDom();
    const engine = new FakeDocxEngine(1, PAGE);
    engine.setLayoutComplete(false);
    const doc = engine.asDoc();
    const viewer = DocxViewer.fromDocument(
      makeEl('canvas') as unknown as HTMLCanvasElement,
      doc,
    ) as DocxViewer;
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
