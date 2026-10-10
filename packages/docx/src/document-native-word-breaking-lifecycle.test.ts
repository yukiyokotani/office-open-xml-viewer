import { deserializeWorkerError, serializeWorkerError } from '@silurus/ooxml-core/worker';
import { requireLayoutPage } from './layout/variant-store.js';
import { layoutDocument } from './document-layout.js';
import { createLayoutServices } from './layout-runtime.js';
import { subscribeDocxLayoutView } from './document-layout-view.js';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { DocxDocument } from './document.js';
import { projectRenderWorkerLayoutMeta } from './render-worker-metadata.js';
import { subscribeDocxLayout } from './document-layout-events.js';
import { attachDocumentLayoutRuntime, documentLayoutRuntimeOf } from './layout/runtime-state.js';
import { attachDocumentLayoutVariants } from './layout/document-layout-variants.js';
import { layoutSourceStore } from './layout-source-model-adapter.js';
import { installStubCanvas, syntheticDocxModel } from './testing/synthetic-document.js';
import { nativeReadingNotices, noReadingNotices, hasNativeReadingRequests, nativeReadingRequests } from './native-reading-notice.js';
import { ProgressiveLayoutLifecycle } from '@silurus/ooxml-core/internal/progressive-layout-lifecycle';
import type { DocumentLayout, LayoutServices } from './layout/types.js';

const readingLayout = { pages: [{ layers: { body: [{ kind: 'paragraph',  }] } }], diagnostics: [] } as unknown as DocumentLayout;
const notices = nativeReadingNotices(readingLayout, { contour: false, wordBreaking: true, pictureBullets: false });
const originalCanvas = Object.getOwnPropertyDescriptor(globalThis, 'OffscreenCanvas');
afterEach(() => {
  vi.unstubAllGlobals(); vi.restoreAllMocks();
  if (originalCanvas) Object.defineProperty(globalThis, 'OffscreenCanvas', originalCanvas);
  else Reflect.deleteProperty(globalThis, 'OffscreenCanvas');
});

function worker(request: (make: (id: number) => { type: string }) => Promise<unknown>, readingNotices = notices) {
  const doc = Object.create(DocxDocument.prototype) as DocxDocument;
  Object.assign(doc, { _mode: 'worker', _layoutLifecycle: new ProgressiveLayoutLifecycle(), _layoutCompletion: null, _layoutViewGeneration: 0, _document: null, _source: null,
    _meta: { pageCount: 2, pageSizes: [{ widthPt: 100, heightPt: 200 }], bookmarkPages: [], nativeReadingRequested: true, readingNotices },
    _bridge: { request, terminate: vi.fn() }, _rawParts: { clear: vi.fn() }, _fetchImage: async () => new Blob(),
    _embeddedFontFaces: [], _officeFontFaces: [], _googleFontFaces: [],
  });
  attachDocumentLayoutRuntime(doc, 0);
  documentLayoutRuntimeOf(doc).activeLayoutOptions = Object.freeze({ currentDateMs: 0 });
  return doc;
}

function mainWithFailingNextVariant() {
  const model = syntheticDocxModel('plain', { paragraphs: 1, wordsPerParagraph: 1 });
  const paragraph = model.body[0];
  if (paragraph.type !== 'paragraph') throw new Error('invalid invented paragraph');
  const textRun = paragraph.runs.find(run => run.type === 'text');
  if (!textRun || textRun.type !== 'text') throw new Error('invented text absent');
  Object.assign(textRun, { __nativeReadingWordBreaking: { rawHres: 0, rawChHres: 1 } });
  const source = layoutSourceStore(model);
  const services = Object.freeze({ text: { fingerprint: 'text' }, images: { fingerprint: 'images' }, math: { fingerprint: 'math' } }) as LayoutServices;
  const error = new Error('new reading variant acquisition failed');
  const build = vi.fn((options: { readonly currentDateMs: number }) => { if (options.currentDateMs === 1) throw error; return readingLayout; });
  attachDocumentLayoutVariants({ source, services, defaultCurrentDateMs: 0, buildLayout: build });
  const terminate = vi.fn(), clear = vi.fn();
  const doc = Object.create(DocxDocument.prototype) as DocxDocument;
  Object.assign(doc, { _mode: 'main', _layoutLifecycle: new ProgressiveLayoutLifecycle(), _layoutCompletion: null, _layoutViewGeneration: 0, _nativeReadingRequested: hasNativeReadingRequests(source), _nativeReadingRequests: nativeReadingRequests(source), _document: model, _source: source, _meta: null,
    _bridge: { terminate }, _rawParts: { clear }, _embeddedFontFaces: [], _officeFontFaces: [], _googleFontFaces: [],
    _fetchImage: async () => new Blob(),
  });
  attachDocumentLayoutRuntime(doc, 0); documentLayoutRuntimeOf(doc).services = services;
  documentLayoutRuntimeOf(doc).activeLayoutOptions = Object.freeze({ currentDateMs: 0 });
  return { doc, error, build, terminate, clear, source };
}

function mainWithActualReadingLayout() {
    installStubCanvas();
    const model = syntheticDocxModel('plain', { paragraphs: 1, wordsPerParagraph: 2 });
    const paragraph = model.body[0];
    if (paragraph.type !== 'paragraph' || paragraph.runs[0].type !== 'text') throw new Error('invented text absent');
    Object.assign(paragraph.runs[0], { __nativeReadingWordBreaking: { rawHres: 0, rawChHres: 1 } });
    const source = layoutSourceStore(model), services = createLayoutServices(source);
    const actual = layoutDocument(model, services);
    attachDocumentLayoutVariants({ source, services, defaultCurrentDateMs: 0, buildLayout: () => actual });
    const doc = Object.create(DocxDocument.prototype) as DocxDocument;
    Object.assign(doc, { _mode: 'main', _nativeReadingRequested: true, _source: source, _document: model, _meta: null,
      _layoutLifecycle: new ProgressiveLayoutLifecycle(), _layoutViewGeneration: 0 });
    attachDocumentLayoutRuntime(doc, 0); documentLayoutRuntimeOf(doc).services = services;
    documentLayoutRuntimeOf(doc).activeLayoutOptions = Object.freeze({ currentDateMs: 0 });
    return doc;
}

describe('native reading publication lifecycle', () => {
  it('keeps valid reading pages and copy after a caller requests an out-of-range page', async () => {
    const doc = mainWithActualReadingLayout();
    const before = await doc.collectPageRuns(0), disclosure = doc.readingNotices, count = doc.pageCount;
    await expect(doc.collectPageRuns(count)).rejects.toBeInstanceOf(RangeError);
    expect(doc.pageCount).toBe(count); expect(doc.readingNotices).toEqual(disclosure);
    await expect(doc.collectPageRuns(0)).resolves.toEqual(before);
    const target = new OffscreenCanvas(1, 1);
    const unavailable = vi.spyOn(target, 'getContext').mockReturnValue(null);
    await expect(doc.renderPage(target, 0)).rejects.toThrow('2D canvas is unavailable');
    expect(doc.pageCount).toBe(count); expect(doc.readingNotices).toEqual(disclosure);
    await expect(doc.collectPageRuns(0)).resolves.toEqual(before);
    unavailable.mockRestore();
    const invalidState = new DOMException('transferred caller canvas', 'InvalidStateError');
    const throwing = vi.spyOn(target, 'getContext').mockImplementation(() => { throw invalidState; });
    await expect(doc.renderPage(target, 0)).rejects.toHaveProperty('code', 'docx-caller-input');
    expect(doc.pageCount).toBe(count); expect(doc.readingNotices).toEqual(disclosure);
    await expect(doc.collectPageRuns(0)).resolves.toEqual(before);
    throwing.mockRestore();
    await doc.renderPage(target, 0);
    expect(doc.pageCount).toBe(count);
  });
  it('keeps publication after a current-realm InvalidStateError without inspecting a throwing owner getter, then renders a valid retry', async () => {
    const doc = mainWithActualReadingLayout(), target = new OffscreenCanvas(1, 1);
    const before = await doc.collectPageRuns(0), disclosure = doc.readingNotices, count = doc.pageCount;
    const failure = new DOMException('caller transferred canvas', 'InvalidStateError');
    const ownerLookup = vi.fn(() => { throw new Error('unneeded owner lookup'); });
    Object.defineProperty(target, 'ownerDocument', { configurable: true, get: ownerLookup });
    const failing = vi.spyOn(target, 'getContext').mockImplementation(() => { throw failure; });
    let targetError: unknown;
    try { await doc.renderPage(target, 0); } catch (error) { targetError = error; }
    expect(typeof targetError === 'object' && targetError !== null
      && 'code' in targetError && targetError.code === 'docx-caller-input').toBe(true);
    expect(typeof targetError === 'object' && targetError !== null
      && 'cause' in targetError && targetError.cause === failure).toBe(true);
    expect(ownerLookup).not.toHaveBeenCalled();
    expect(doc.pageCount).toBe(count); expect(doc.readingNotices).toEqual(disclosure);
    await expect(doc.collectPageRuns(0)).resolves.toEqual(before);
    failing.mockRestore();
    expect(Reflect.deleteProperty(target, 'ownerDocument')).toBe(true);
    const text = vi.fn();
    await doc.renderPage(target, 0, { onTextRun: text });
    expect(text).toHaveBeenCalled();
    expect(doc.pageCount).toBe(count); expect(doc.readingNotices).toEqual(disclosure);
    await expect(doc.collectPageRuns(0)).resolves.toEqual(before);
  });

  it('keeps document publication and copy after an owner-realm InvalidStateError, then renders a valid retry', async () => {
    const doc = mainWithActualReadingLayout();
    class ForeignCanvas extends OffscreenCanvas {
      readonly style = { width: '', height: '', display: '' };
    }
    class ForeignDOMException extends Error {
      constructor(message = '', name = 'Error') { super(message); this.name = name; }
    }
    const target = new ForeignCanvas(1, 1);
    Object.defineProperty(target, 'ownerDocument', { value: {
      defaultView: { HTMLCanvasElement: ForeignCanvas, DOMException: ForeignDOMException },
    } });
    const before = await doc.collectPageRuns(0), disclosure = doc.readingNotices, count = doc.pageCount;
    const failure = new ForeignDOMException('foreign transferred caller target', 'InvalidStateError');
    expect(failure).not.toBeInstanceOf(DOMException);
    const failing = vi.spyOn(target, 'getContext').mockImplementation(() => { throw failure; });
    await expect(doc.renderPage(target, 0)).rejects.toMatchObject({ code: 'docx-caller-input', cause: failure });
    expect(doc.pageCount).toBe(count); expect(doc.readingNotices).toEqual(disclosure);
    await expect(doc.collectPageRuns(0)).resolves.toEqual(before);
    failing.mockRestore();
    const text = vi.fn();
    await doc.renderPage(target, 0, { onTextRun: text });
    expect(text).toHaveBeenCalled();
    expect(doc.pageCount).toBe(count); expect(doc.readingNotices).toEqual(disclosure);
    await expect(doc.collectPageRuns(0)).resolves.toEqual(before);
  });

  it('retains original resource failure identity and revokes publication when target owner lookup also fails', async () => {
    const doc = mainWithActualReadingLayout(), target = new OffscreenCanvas(1, 1);
    const failure = new RangeError('original canvas allocation failure');
    const lookupFailure = new Error('owner lookup failure');
    Object.defineProperty(target, 'ownerDocument', { get() { throw lookupFailure; } });
    vi.spyOn(target, 'getContext').mockImplementation(() => { throw failure; });
    await expect(doc.renderPage(target, 0)).rejects.toBe(failure);
    expect(doc.pageCount).toBe(0); expect(doc.readingNotices).toEqual(noReadingNotices);
    await expect(doc.collectPageRuns(0)).rejects.toThrow('revoked');
  });

  it('revokes reading publication when its internally allocated bitmap target has no 2D context', async () => {
    const doc = mainWithActualReadingLayout();
    const target = new OffscreenCanvas(1, 1);
    vi.spyOn(target, 'getContext').mockReturnValue(null);
    vi.stubGlobal('OffscreenCanvas', class { constructor() { return target; } });
    await expect(doc.renderPageToBitmap(0)).rejects.toThrow('internally acquired');
    expect(doc.pageCount).toBe(0); expect(doc.readingNotices).toEqual(noReadingNotices);
    await expect(doc.collectPageRuns(0)).rejects.toThrow('revoked');
  });
  it('keeps a caller target resource failure terminal instead of classifying every getContext throw as input', async () => {
    const doc = mainWithActualReadingLayout(), target = new OffscreenCanvas(1, 1);
    const error = new RangeError('canvas resource allocation failed');
    vi.spyOn(target, 'getContext').mockImplementation(() => { throw error; });
    await expect(doc.renderPage(target, 0)).rejects.toBe(error);
    expect(doc.pageCount).toBe(0); expect(doc.readingNotices).toEqual(noReadingNotices);
  });
  it('preserves reading publication after a typed caller error crosses the real worker error wire', async () => {
    let callerError: unknown;
    try { requireLayoutPage(readingLayout, 1); } catch (error) { callerError = error; }
    const transported = deserializeWorkerError(structuredClone(serializeWorkerError(callerError)));
    const doc = worker(async make => {
      const request = make(1);
      if (request.type === 'collectRuns' && 'pageIndex' in request && request.pageIndex === 2) throw transported;
      return { type: 'runsCollected', runs: [{ text: 'retained' }] };
    });
    await expect(doc.collectPageRuns(2)).rejects.toBeInstanceOf(RangeError);
    expect(doc.pageCount).toBe(2); expect(doc.readingNotices).toEqual(notices);
    await expect(doc.collectPageRuns(0)).resolves.toEqual([{ text: 'retained' }]);
  });
  it('does not exempt genuine RangeError acquisition failures from revocation', async () => {
    const doc = worker(async () => { throw new RangeError('native resource acquisition range failed'); });
    await expect(doc.collectPageRuns(0)).rejects.toThrow('native resource acquisition range failed');
    expect(doc.pageCount).toBe(0); expect(doc.readingNotices).toEqual(noReadingNotices);
  });

  it.each(['revoke', 'destroy'] as const)('does not publish a pending worker selection after %s', async action => {
    let reply!: (value: unknown) => void;
    const waiting = new Promise(resolve => { reply = resolve; });
    const doc = worker(async () => waiting);
    const selected = doc.setLayoutView({ currentDate: 1 });
    if (action === 'destroy') doc.destroy();
    else doc._invalidateReadingLayout(new Error('active publication failed'));
    const publications: number[] = [];
    const stop = subscribeDocxLayoutView(doc, () => publications.push(doc.pageCount), () => {});
    reply({ type: 'layoutViewSelected', meta: { pageCount: 7, pageSizes: [], bookmarkPages: [], readingNotices: notices } });
    await selected;
    expect(publications).toEqual([]); expect(doc.pageCount).toBe(0); expect(doc.readingNotices).toEqual(noReadingNotices);
    stop();
  });

  it('settles a reading completion waiter on revocation even if transport completion never settles', async () => {
    const doc = worker(async () => ({}));
    Object.assign(doc, { _layoutCompletion: new Promise<void>(() => {}) });
    const waiting = doc.waitUntilLayoutComplete(), error = new Error('terminal acquisition failure');
    doc._invalidateReadingLayout(error);
    await expect(waiting).rejects.toBe(error);
  });

  it('releases the returned bitmap on a consumer callback failure without revoking valid acquired pages', async () => {
    vi.stubGlobal('OffscreenCanvas', class {});
    const close = vi.fn(), error = new TypeError('consumer callback failed');
    const doc = worker(async make => make(1).type === 'collectRuns'
      ? { type: 'runsCollected', runs: [{ text: 'retained' }] }
      : { type: 'pageRendered', bitmap: { close, width: 1, height: 1 }, runs: [{ text: 'first' }, { text: 'second' }] });
    const callback = vi.fn(() => { throw error; });
    await expect(doc.renderPageToBitmap(0, { onTextRun: callback })).rejects.toBe(error);
    expect(callback).toHaveBeenCalledTimes(1); expect(close).toHaveBeenCalledTimes(1);
    expect(doc.pageCount).toBe(2); expect(doc.readingNotices).toEqual(notices);
    await expect(doc.collectPageRuns(0)).resolves.toEqual([{ text: 'retained' }]);
  });

  it('ignores a presentation failure from an older selected publication', async () => {
    const doc = worker(async () => ({ type: 'layoutViewSelected', meta: { pageCount: 1, pageSizes: [], bookmarkPages: [], readingNotices: notices } }));
    const old = doc._readingPublicationToken!;
    await doc.setLayoutView({ currentDate: 1 });
    doc._invalidateReadingLayout(new Error('obsolete presentation failed'), old);
    expect(doc.pageCount).toBe(1); expect(doc.readingNotices).toEqual(notices);
    const current = doc._readingPublicationToken!;
    doc._invalidateReadingLayout(new Error('current presentation failed'), current);
    expect(doc.pageCount).toBe(0);
  });
  it('revokes page geometry and disclosure when worker surface construction fails before transport', async () => {
    const error = new Error('surface allocation failed'), request = vi.fn(async () => ({}));
    vi.stubGlobal('OffscreenCanvas', class { constructor() { throw error; } });
    const doc = worker(request), publications: number[] = [];
    const stop = subscribeDocxLayout(doc, () => ({ pageCount: doc.pageCount, exact: true, complete: true }), p => publications.push(p.pageCount), () => {});
    await expect(doc.renderPageToBitmap(0)).rejects.toBe(error);
    expect(doc.pageCount).toBe(0); expect(doc.readingNotices).toEqual(noReadingNotices);
    expect(request).not.toHaveBeenCalled(); expect(publications).toEqual([2, 0]); stop();
  });

  it('a no-op selection keeps a pending render failure bound to the same reading publication', async () => {
    vi.stubGlobal('OffscreenCanvas', class {});
    let rejectRender!: (error: Error) => void;
    const waiting = new Promise<never>((_, reject) => { rejectRender = reject; });
    const doc = worker(async () => waiting), pending = doc.renderPageToBitmap(0, { dpr: 1 });
    await doc.setLayoutView({ currentDate: 0 });
    const error = new Error('same-publication decoding failed'); rejectRender(error);
    await expect(pending).rejects.toBe(error);
    expect(doc.pageCount).toBe(0); expect(doc.layoutComplete).toBe(false); expect(doc.readingNotices).toEqual(noReadingNotices);
    await expect(doc.waitUntilLayoutComplete()).rejects.toBe(error);
    await expect(doc.collectPageRuns(0)).rejects.toThrow('revoked');
    await expect(doc.getElementContextAt(0, { xPt: 1, yPt: 1 })).rejects.toThrow('revoked');
    expect(doc.commentAnchorRanges()).toEqual([]); expect(doc.revisionAnchorRanges()).toEqual([]);
    expect(doc.pageSize(0)).toEqual({ widthPt: 0, heightPt: 0 });
  });

  it('releases a successful pending bitmap after another render revokes the complete layout', async () => {
    vi.stubGlobal('OffscreenCanvas', class {});
    let resolveRender!: (reply: unknown) => void;
    const waiting = new Promise(resolve => { resolveRender = resolve; });
    const close = vi.fn(), bitmap = { close, width: 10, height: 10 } as unknown as ImageBitmap;
    const error = new Error('another page failed');
    let calls = 0;
    const doc = worker(async () => { if (++calls === 1) return waiting; throw error; });
    const onTextRun = vi.fn();
    const pending = doc.renderPageToBitmap(0, { dpr: 1, onTextRun });
    await expect(doc.renderPageToBitmap(1, { dpr: 1 })).rejects.toBe(error);
    resolveRender({ type: 'pageRendered', bitmap, runs: [{ text: 'stale text' }] });
    await expect(pending).rejects.toThrow('revoked');
    expect(close).toHaveBeenCalledTimes(1); expect(onTextRun).not.toHaveBeenCalled(); expect(doc.pageCount).toBe(0);
  });

  it('does not let a failed older render revoke the successfully selected later variant', async () => {
    vi.stubGlobal('OffscreenCanvas', class {});
    let rejectOld!: (error: Error) => void;
    const old = new Promise<never>((_, reject) => { rejectOld = reject; });
    const doc = worker(async make => make(1).type === 'renderPage' ? old : ({ type: 'layoutViewSelected', id: 1,
      meta: { pageCount: 1, pageSizes: [{ widthPt: 100, heightPt: 200 }], bookmarkPages: [] } }));
    const pending = doc.renderPageToBitmap(0, { dpr: 1 });
    await doc.setLayoutView({ currentDate: 1 });
    const error = new Error('older render failed'); rejectOld(error);
    await expect(pending).rejects.toBe(error);
    expect(doc.pageCount).toBe(1); expect(doc.readingNotices).toEqual(noReadingNotices);
  });

  it('discloses source-owned simplified word breaking after complete page acquisition', () => {
    const { source } = mainWithFailingNextVariant();
    const layout = { pages: [{ pageIndex: 0, bookmarkStarts: [], geometry: { widthPt: 100, heightPt: 200 }, layers: { body: [] } }], diagnostics: [] } as unknown as DocumentLayout;
    const meta = projectRenderWorkerLayoutMeta(layout, source, { comments: [], revisions: [] });
    expect(meta.nativeReadingRequested).toBe(true); expect(meta.readingNotices).toEqual(notices);
    const strict = layoutSourceStore(syntheticDocxModel('plain', { paragraphs: 1, wordsPerParagraph: 1 }));
    expect(projectRenderWorkerLayoutMeta(layout, strict, { comments: [], revisions: [] }).nativeReadingRequested).toBeUndefined();
  });

  it('owns a notice-free view of a reading document and rejects older geometry after variant replacement', async () => {
    let resolveRuns!: (result: unknown) => void;
    const waiting = new Promise(resolve => { resolveRuns = resolve; });
    const doc = worker(async make => make(1).type === 'collectRuns' ? waiting : ({ type: 'layoutViewSelected', id: 1,
      meta: { pageCount: 1, pageSizes: [], bookmarkPages: [], nativeReadingRequested: true, readingNotices: notices } }), noReadingNotices);
    expect(doc.readingNotices).toEqual(noReadingNotices);
    const pending = doc.collectPageRuns(0);
    await doc.setLayoutView({ currentDate: 1 });
    resolveRuns({ type: 'runsCollected', runs: [{ text: 'older view' }] });
    await expect(pending).rejects.toThrow('superseded');
    expect(doc.pageCount).toBe(1); expect(doc.readingNotices).toEqual(notices);
  });

  it('revokes the active reading publication on a selected-variant acquisition failure', async () => {
    const error = new Error('variant acquisition failed');
    const doc = worker(async () => { throw error; });
    await expect(doc.setLayoutView({ currentDate: 1 })).rejects.toBe(error);
    expect(doc.pageCount).toBe(0); expect(doc.readingNotices).toEqual(noReadingNotices);
    await expect(doc.renderPageToBitmap(0)).rejects.toThrow('revoked');
  });

  it('acquires a main reading variant before publication and rejects the exact acquisition failure', async () => {
    const { doc, error } = mainWithFailingNextVariant();
    expect(doc.pageCount).toBe(1); expect(doc.readingNotices).toEqual(notices);
    await doc.setLayoutView({ currentDate: 2 });
    expect(doc.pageCount).toBe(1); expect(doc.readingNotices).toEqual(notices);
    await expect(doc.setLayoutView({ currentDate: 1 })).rejects.toBe(error);
    expect(doc.pageCount).toBe(0); expect(doc.readingNotices).toEqual(noReadingNotices);
    await expect(doc.waitUntilLayoutComplete()).rejects.toBe(error);
  });

  it('cancels without acquiring a failing lazy variant and still releases resource owners', async () => {
    const { doc, build, terminate, clear } = mainWithFailingNextVariant();
    expect(doc.pageCount).toBe(1);
    // Also cover cleanup after source acquisition but before a requested
    // initial reading variant has successfully primed.
    documentLayoutRuntimeOf(doc).activeLayoutOptions = Object.freeze({ currentDateMs: 1 });
    expect(() => doc.destroy()).not.toThrow();
    expect(build).toHaveBeenCalledTimes(1); expect(terminate).toHaveBeenCalledTimes(1); expect(clear).toHaveBeenCalledTimes(1);
    expect(doc.pageCount).toBe(0); expect(doc.readingNotices).toEqual(noReadingNotices);
  });
});
