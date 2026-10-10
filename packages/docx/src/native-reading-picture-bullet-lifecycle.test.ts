import { afterEach, describe, expect, it, vi } from 'vitest';
import { deserializeWorkerError, serializeWorkerError } from '@silurus/ooxml-core/worker';
import { ProgressiveLayoutLifecycle } from '@silurus/ooxml-core/internal/progressive-layout-lifecycle';
import { DocxDocument } from './document.js';
import { attachDocumentLayoutVariants } from './layout/document-layout-variants.js';
import { requireLayoutPage } from './layout/variant-store.js';
import { attachDocumentLayoutRuntime, documentLayoutRuntimeOf } from './layout/runtime-state.js';
import { layoutSourceStore } from './layout-source-model-adapter.js';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { syntheticDocxModel, installStubCanvas } from './testing/synthetic-document.js';
import { inventedReadingPictureBulletNumbering } from './testing/native-reading-picture-bullet.js';
import { nativeReadingNotices, noReadingNotices, updateNativeReadingNotice } from './native-reading-notice.js';
import { makeContainer, installDom } from './scroll-viewer-test-dom.js';

const originalCanvas = Object.getOwnPropertyDescriptor(globalThis, 'OffscreenCanvas');
afterEach(() => {
  vi.unstubAllGlobals(); vi.restoreAllMocks();
  if (originalCanvas) Object.defineProperty(globalThis, 'OffscreenCanvas', originalCanvas);
  else Reflect.deleteProperty(globalThis, 'OffscreenCanvas');
});

// The shared synthetic canvas is a measurement stub without a sizing constructor.
// Explicit caller targets need real dimensions to witness prepaint immutability.
function targetCanvas(width = 1, height = 1): OffscreenCanvas {
  const target = new OffscreenCanvas(width, height);
  target.width = width; target.height = height;
  return target;
}
function readingFixture() {
  installStubCanvas();
  const model = syntheticDocxModel('plain', { paragraphs: 1, wordsPerParagraph: 2 });
  const paragraph = model.body[0];
  if (paragraph.type !== 'paragraph') throw new Error('invented paragraph absent');
  paragraph.numbering = inventedReadingPictureBulletNumbering();
  const source = layoutSourceStore(model), services = createLayoutServices(source);
  return layoutDocument(source, services);
}
function worker(request: (make: (id: number) => { type: string }) => Promise<unknown>) {
  const layout = readingFixture();
  const notices = nativeReadingNotices(layout, { contour: false, wordBreaking: false, pictureBullets: true });
  const doc = Object.create(DocxDocument.prototype) as DocxDocument;
  Object.assign(doc, { _mode: 'worker', _layoutLifecycle: new ProgressiveLayoutLifecycle(),
    _layoutCompletion: null, _layoutViewGeneration: 0, _document: null, _source: null,
    _meta: { pageCount: layout.pages.length, pageSizes: [{ widthPt: 100, heightPt: 200 }],
      bookmarkPages: [], nativeReadingRequested: true, readingNotices: notices },
    _bridge: { request, terminate: vi.fn() }, _rawParts: { clear: vi.fn() }, _fetchImage: async () => new Blob(),
    _embeddedFontFaces: [], _officeFontFaces: [], _googleFontFaces: [],
  });
  attachDocumentLayoutRuntime(doc, 0);
  documentLayoutRuntimeOf(doc).activeLayoutOptions = Object.freeze({ currentDateMs: 0 });
  return { doc, notices, layout };
}
function mainReadingFixture() {
  installStubCanvas();
  const model = syntheticDocxModel('plain', { paragraphs: 1, wordsPerParagraph: 2 });
  const paragraph = model.body[0];
  if (paragraph.type !== 'paragraph') throw new Error('invented paragraph absent');
  paragraph.numbering = inventedReadingPictureBulletNumbering();
  const source = layoutSourceStore(model), services = createLayoutServices(source);
  const layout = layoutDocument(source, services);
  attachDocumentLayoutVariants({ source, services, defaultCurrentDateMs: 0, buildLayout: () => layout });
  const doc = Object.create(DocxDocument.prototype) as DocxDocument;
  Object.assign(doc, { _mode: 'main', _nativeReadingRequested: true, _source: source, _document: model,
    _meta: null, _layoutLifecycle: new ProgressiveLayoutLifecycle(), _layoutViewGeneration: 0 });
  attachDocumentLayoutRuntime(doc, 0); documentLayoutRuntimeOf(doc).services = services;
  documentLayoutRuntimeOf(doc).activeLayoutOptions = Object.freeze({ currentDateMs: 0 });
  return doc;
}
describe('picture-bullet reading publication lifetime', () => {
  it('preserves reading pages and disclosure after an invalid caller canvas state', async () => {
    const doc = mainReadingFixture(), target = targetCanvas();
    const before = await doc.collectPageRuns(0), notices = doc.readingNotices, count = doc.pageCount;
    vi.spyOn(target, 'getContext').mockImplementation(() => {
      throw new DOMException('transferred caller target', 'InvalidStateError');
    });
    await expect(doc.renderPage(target, 0)).rejects.toHaveProperty('code', 'docx-caller-input');
    expect(doc.pageCount).toBe(count); expect(doc.readingNotices).toEqual(notices);
    await expect(doc.collectPageRuns(0)).resolves.toEqual(before);
    expect(target.width).toBe(1); expect(target.height).toBe(1);
  });
  it('preserves picture reading publication and literal runs for a proven owner-realm caller error', async () => {
    const doc = mainReadingFixture();
    class ForeignCanvas extends OffscreenCanvas {
      readonly style = { width: '', height: '', display: '' };
    }
    class ForeignDOMException extends Error {
      constructor(message = '', name = 'Error') { super(message); this.name = name; }
    }
    const target = Object.assign(new ForeignCanvas(1, 1), { width: 1, height: 1 });
    Object.defineProperty(target, 'ownerDocument', { value: {
      defaultView: { HTMLCanvasElement: ForeignCanvas, DOMException: ForeignDOMException },
    } });
    const before = await doc.collectPageRuns(0), notices = doc.readingNotices, count = doc.pageCount;
    const failure = new ForeignDOMException('foreign transferred caller canvas', 'InvalidStateError');
    expect(failure).not.toBeInstanceOf(DOMException);
    vi.spyOn(target, 'getContext').mockImplementation(() => { throw failure; });
    await expect(doc.renderPage(target, 0)).rejects.toMatchObject({ code: 'docx-caller-input', cause: failure });
    expect(doc.pageCount).toBe(count); expect(doc.readingNotices).toEqual(notices);
    await expect(doc.collectPageRuns(0)).resolves.toEqual(before);
    expect(target.width).toBe(1); expect(target.height).toBe(1);
  });
  it.each(['ownerDocument', 'defaultView'] as const)(
    'retains picture reading pages and literal runs when proven caller state has unavailable %s', async lookup => {
      const doc = mainReadingFixture(), target = targetCanvas();
      const before = await doc.collectPageRuns(0), notices = doc.readingNotices, count = doc.pageCount;
      const originalFailure = new DOMException('caller target', 'InvalidStateError');
      const lookupFailure = new Error('optional caller realm unavailable');
      vi.spyOn(target, 'getContext').mockImplementation(() => { throw originalFailure; });
      if (lookup === 'ownerDocument') Object.defineProperty(target, 'ownerDocument', { get() { throw lookupFailure; } });
      else Object.defineProperty(target, 'ownerDocument', { value: {
        get defaultView() { throw lookupFailure; },
      } });
      await expect(doc.renderPage(target, 0)).rejects.toMatchObject({ code: 'docx-caller-input', cause: originalFailure });
      expect(doc.pageCount).toBe(count); expect(doc.readingNotices).toEqual(notices);
      await expect(doc.collectPageRuns(0)).resolves.toEqual(before);
      expect(target.width).toBe(1); expect(target.height).toBe(1);
    },
  );
  it.each(['callerResource', 'internalBitmap'] as const)(
    'retains original %s failure and revokes picture publication despite failed realm lookup', async kind => {
      const doc = mainReadingFixture(), target = targetCanvas();
      const originalFailure = kind === 'callerResource' ? new RangeError('original canvas resource failure')
        : new DOMException('original internal context failure', 'InvalidStateError');
      const lookupFailure = new Error('realm lookup failure');
      Object.defineProperty(target, 'ownerDocument', { get() { throw lookupFailure; } });
      vi.spyOn(target, 'getContext').mockImplementation(() => { throw originalFailure; });
      if (kind === 'internalBitmap') vi.stubGlobal('OffscreenCanvas', class { constructor() { return target; } });
      await expect(kind === 'callerResource' ? doc.renderPage(target, 0) : doc.renderPageToBitmap(0))
        .rejects.toBe(originalFailure);
      expect(doc.pageCount).toBe(0); expect(doc.readingNotices).toEqual(noReadingNotices);
      await expect(doc.collectPageRuns(0)).rejects.toThrow('revoked');
      expect(target.width).toBe(1); expect(target.height).toBe(1);
    },
  );
  it('revokes publication for a failed internally acquired bitmap target', async () => {
    const doc = mainReadingFixture(), target = targetCanvas();
    vi.spyOn(target, 'getContext').mockReturnValue(null);
    vi.stubGlobal('OffscreenCanvas', class { constructor() { return target; } });
    await expect(doc.renderPageToBitmap(0)).rejects.toThrow('internally acquired');
    expect(doc.pageCount).toBe(0); expect(doc.readingNotices).toEqual(noReadingNotices);
    await expect(doc.collectPageRuns(0)).rejects.toThrow('revoked');
  });
  it('does not exempt a caller target resource allocation failure', async () => {
    const doc = mainReadingFixture(), target = targetCanvas();
    const failure = new RangeError('target resource allocation failed');
    vi.spyOn(target, 'getContext').mockImplementation(() => { throw failure; });
    await expect(doc.renderPage(target, 0)).rejects.toBe(failure);
    expect(doc.pageCount).toBe(0); expect(doc.readingNotices).toEqual(noReadingNotices);
  });
  it('preserves acquired publication when a typed caller page error crosses the ordinary worker wire', async () => {
    const { doc, layout, notices } = worker(async make => {
      const request = make(1);
      if ('pageIndex' in request && request.pageIndex === 99) {
        let error: unknown;
        try { requireLayoutPage(layout, 99); } catch (caught) { error = caught; }
        throw deserializeWorkerError(structuredClone(serializeWorkerError(error)));
      }
      return { type: 'runsCollected', runs: [{ text: 'retained literal' }] };
    });
    await expect(doc.collectPageRuns(99)).rejects.toBeInstanceOf(RangeError);
    expect(doc.pageCount).toBe(layout.pages.length); expect(doc.readingNotices).toEqual(notices);
    await expect(doc.collectPageRuns(0)).resolves.toEqual([{ text: 'retained literal' }]);
  });
  it('revokes pages and disclosure for a genuine resource RangeError', async () => {
    const { doc } = worker(async () => { throw new RangeError('owned marker resource failed'); });
    await expect(doc.collectPageRuns(0)).rejects.toThrow('owned marker resource failed');
    expect(doc.pageCount).toBe(0); expect(doc.readingNotices).toEqual(noReadingNotices);
    await expect(doc.collectPageRuns(0)).rejects.toThrow('revoked');
  });
  it.each(['revoke', 'destroy'] as const)('cannot publish a pending worker variant after %s', async action => {
    let reply!: (value: unknown) => void;
    const pendingReply = new Promise(resolve => { reply = resolve; });
    const { doc, notices } = worker(async () => pendingReply);
    const pendingSelection = doc.setLayoutView({ currentDate: 1 });
    if (action === 'destroy') doc.destroy();
    else doc._invalidateReadingLayout(new Error('owned marker acquisition failed'));
    reply({ type: 'layoutViewSelected', meta: { pageCount: 7, pageSizes: [], bookmarkPages: [],
      nativeReadingRequested: true, readingNotices: notices } });
    await pendingSelection;
    expect(doc.pageCount).toBe(0); expect(doc.readingNotices).toEqual(noReadingNotices);
  });
  it('settles a completion waiter on revocation even if worker completion never settles', async () => {
    const { doc } = worker(async () => ({}));
    Object.assign(doc, { _layoutCompletion: new Promise<void>(() => {}) });
    const pending = doc.waitUntilLayoutComplete(), error = new Error('marker decode failed');
    doc._invalidateReadingLayout(error);
    await expect(pending).rejects.toBe(error);
  });
  it('preserves valid acquired pages when a text callback fails and releases its bitmap', async () => {
    const previous = Object.getOwnPropertyDescriptor(globalThis, 'OffscreenCanvas');
    const { doc, notices, layout } = worker(async make => make(1).type === 'collectRuns'
      ? { type: 'runsCollected', runs: [{ text: 'retained literal' }] }
      : { type: 'pageRendered', bitmap: { close, width: 1, height: 1 }, runs: [{ text: 'first' }, { text: 'second' }] });
    const close = vi.fn(), error = new TypeError('caller callback failed');
    vi.stubGlobal('OffscreenCanvas', class {});
    try {
      const callback = vi.fn(() => { throw error; });
      await expect(doc.renderPageToBitmap(0, { onTextRun: callback })).rejects.toBe(error);
      expect(callback).toHaveBeenCalledTimes(1); expect(close).toHaveBeenCalledTimes(1);
      expect(doc.pageCount).toBe(layout.pages.length); expect(doc.readingNotices).toEqual(notices);
    } finally {
      vi.unstubAllGlobals();
      if (previous) Object.defineProperty(globalThis, 'OffscreenCanvas', previous);
      else Reflect.deleteProperty(globalThis, 'OffscreenCanvas');
    }
  });
  it('does not repeat a live announcement when the disclosure did not change', () => {
    const layout = readingFixture(), notices = nativeReadingNotices(layout, { contour: false, wordBreaking: false, pictureBullets: true });
    installDom();
    const container = makeContainer(200, 400);
    const element = updateNativeReadingNotice(container as unknown as HTMLElement, null, notices)!;
    const text = element.textContent;
    let writes = 0;
    Object.defineProperty(element, 'textContent', { configurable: true, get: () => text, set: () => { writes++; } });
    expect(updateNativeReadingNotice(container as unknown as HTMLElement, element, notices)).toBe(element);
    expect(writes).toBe(0);
    expect(updateNativeReadingNotice(container as unknown as HTMLElement, element, noReadingNotices)).toBeNull();
  });
});
