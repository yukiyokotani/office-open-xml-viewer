import { afterEach, describe, expect, it, vi } from 'vitest';
import { WorkerBridge, deserializeWorkerError, type WorkerLike } from '@silurus/ooxml-core/worker';
import { acquireDocumentCanvasTargetContext } from './paint/canvas-document.js';
import { DocxDocument } from './document.js';
import { attachDocumentLayoutRuntime } from './layout/runtime-state.js';
import type { RenderWorkerResponse } from './worker-protocol.js';

/**
 * Worker-mode bitmap rendering must paint into a surface created by the caller
 * and transferred, exactly like the main-mode path paints into one created in
 * the caller. The loopback hands the posted object to the worker unchanged,
 * standing in for the transfer; parsing and layout are stubbed so the real
 * worker entries can be driven without generated WASM.
 */

const mocks = vi.hoisted(() => ({
  render: vi.fn(async (..._args: unknown[]) => undefined),
  store: {},
  paginationFailure: undefined as 'before-publication' | 'after-publication' | 'abort' | undefined,
  registeredFonts: new Set<FontFace>(),
  nativeSentinels: [] as FontFace[] | null,
}));

class FakeDocxArchive {
  constructor(_bytes: Uint8Array) {}
  assert_healthy(): void {}
  resource_usage(): Uint8Array {
    return new TextEncoder().encode(JSON.stringify({
      archiveEntryCount: 1, declaredInflatedBytes: 1, distinctInflatedBytes: 1, operationInflatedBytes: 1,
    }));
  }
  free(): void {}
}

vi.mock('./wasm/docx_parser.js', () => ({
  default: async () => undefined, reinit: async () => undefined, DocxArchive: FakeDocxArchive,
}));
vi.mock('@silurus/ooxml-core', async (load) => ({
  ...await load<typeof import('@silurus/ooxml-core')>(),
  loadOfficeFontFallbacks: async () => {
    const face = {} as FontFace;
    mocks.registeredFonts.add(face);
    return { faces: [face], routes: {} };
  },
  unloadOfficeFontFallbacks: (faces: FontFace[]) => {
    for (const face of faces) mocks.registeredFonts.delete(face);
  },
  loadNativeRouteSentinels: async () => mocks.nativeSentinels,
}));
vi.mock('./document-pull-worker.js', async (load) => ({
  ...await load<typeof import('./document-pull-worker.js')>(),
  DocumentPullWorker: class { open(): void {} async reset(): Promise<void> {} },
  createLocalDocumentPullTransport: () => ({}),
}));
vi.mock('./document-pull-client.js', async (load) => ({
  ...await load<typeof import('./document-pull-client.js')>(),
  materializeDocumentPullOwnedModelsSession: async () => ({ document: {}, ownedLayoutDocument: {} }),
}));
vi.mock('./vertical-render-capability.js', async (load) => ({
  ...await load<typeof import('./vertical-render-capability.js')>(),
  documentRequiresDomVerticalGlyphLayout: () => false,
}));
vi.mock('./layout-source-model-adapter.js', () => ({
  layoutSourceModelAdapterFromOwnedModel: () => ({
    document: {}, source: { fatalParse: null, mathOccurrences: [], blocks: { sources: [] } },
  }),
}));
vi.mock('./font-plan.js', async (load) => ({
  ...await load<typeof import('./font-plan.js')>(),
  docxOfficeFontFallbackRequests: () => [],
}));
vi.mock('./layout-runtime.js', async (load) => ({
  ...await load<typeof import('./layout-runtime.js')>(),
  createLayoutServices: () => ({}),
  nativeRouteLineRequests: () => [],
}));
vi.mock('./render-worker-layout.js', () => ({
  retainRenderWorkerDocumentLayout: () => ({
    defaultCurrentDateMs: 0, layoutServices: {},
    layoutVariants: { layoutFor: () => ({ pages: [] }) },
  }),
}));
vi.mock('./render-worker-metadata.js', () => ({
  projectRenderWorkerLayoutMeta: () => ({ pageCount: 1, pageSizes: [], bookmarkPages: [] }),
  renderWorkerLayoutMeta: () => ({ pageCount: 1, pageSizes: [], bookmarkPages: [] }),
}));
vi.mock('./render-worker-progressive.js', async () => ({
  paginateRenderWorkerDocumentProgressively: async (
    _doc: unknown, _source: unknown, sink: { publish(value: unknown): void },
  ) => {
    if (mocks.paginationFailure === 'abort') {
      const { PaginationAbortError } = await import('./layout/pagination-scheduler.js');
      throw new PaginationAbortError();
    }
    if (mocks.paginationFailure === 'before-publication') throw new Error('pagination failure');
    sink.publish({ pageCount: 1, unchangedPages: 0 });
    if (mocks.paginationFailure === 'after-publication') throw new Error('pagination failure');
    return 1;
  },
}));
vi.mock('./layout/runtime-state.js', async (load) => ({
  ...await load<typeof import('./layout/runtime-state.js')>(),
  layoutSourceStoreOf: () => mocks.store,
}));
vi.mock('./text-run-projection.js', async (load) => ({
  ...await load<typeof import('./text-run-projection.js')>(),
  textRunsForSelectedPage: () => [],
}));
vi.mock('./renderer.js', async (load) => ({
  ...await load<typeof import('./renderer.js')>(),
  renderLayoutSourceToCanvas: mocks.render,
}));

class FakeOffscreenCanvas {
  static created: FakeOffscreenCanvas[] = [];
  contextRequests = 0;
  readonly bitmap = { close: vi.fn() } as unknown as ImageBitmap;
  constructor(public width: number, public height: number) {
    FakeOffscreenCanvas.created.push(this);
  }
  getContext(): null { this.contextRequests += 1; return null; }
  transferToImageBitmap(): ImageBitmap { return this.bitmap; }
}

interface Posted { message: Record<string, unknown>; transfer: unknown[]; surfaces: unknown[] }

/** Host-side Worker whose messages reach the worker module's `self`. */
class LoopbackWorker implements WorkerLike {
  readonly posted: Posted[] = [];
  readonly received: unknown[] = [];
  private readonly listeners = new Set<(event: MessageEvent) => void>();
  constructor(private readonly scope: { onmessage: ((event: MessageEvent) => unknown) | null }) {}
  postMessage(message: unknown, transfer: Transferable[] = []): void {
    // Snapshot transferred surfaces at the moment of transfer.
    const surfaces = transfer.map((item) => item instanceof FakeOffscreenCanvas
      ? { width: item.width, height: item.height, contextRequests: item.contextRequests }
      : item);
    this.posted.push({ message: message as Record<string, unknown>, transfer, surfaces });
    void this.scope.onmessage?.({ data: message } as MessageEvent);
  }
  receive(data: unknown): void {
    this.received.push(data);
    for (const listener of this.listeners) listener({ data } as MessageEvent);
  }
  addEventListener(type: 'message' | 'messageerror' | 'error', listener: (event: never) => void): void {
    if (type === 'message') this.listeners.add(listener as (event: MessageEvent) => void);
  }
  removeEventListener(_type: 'message' | 'messageerror' | 'error', listener: (event: never) => void): void {
    this.listeners.delete(listener as (event: MessageEvent) => void);
  }
  terminate(): void {}
}

async function loadWorker(parse = true) {
  vi.resetModules();
  const scope = {
    // A realm FontFaceSet; registrations themselves are mocked above.
    fonts: {} as FontFaceSet,
    onmessage: null as ((event: MessageEvent) => Promise<void>) | null,
    postMessage: (message: unknown) => worker.receive(message),
  };
  const worker = new LoopbackWorker(scope);
  vi.stubGlobal('self', scope);
  vi.stubGlobal('OffscreenCanvas', FakeOffscreenCanvas);
  await import('./render-worker.js');
  const dispatch = (data: unknown) => scope.onmessage!({ data } as MessageEvent);
  await dispatch({ type: 'init', wasmUrl: 'x' });
  if (parse) await dispatch({
    type: 'parse', id: 100, data: new ArrayBuffer(1), defaultCurrentDateMs: 0,
    resourcePolicy: { maxArchiveEntryBytes: null, maxTotalInflatedBytes: null, maxArchiveEntries: null },
  });
  const bridge = new WorkerBridge<RenderWorkerResponse>(worker, {
    correlate: (response) => ('id' in response ? response.id : undefined),
    toError: (response) => 'type' in response && response.type === 'error' ? deserializeWorkerError(response) : undefined,
  });
  const document = Object.create(DocxDocument.prototype) as DocxDocument;
  Object.assign(document, { _mode: 'worker', _bridge: bridge });
  attachDocumentLayoutRuntime(document, 0);
  return { document, worker, dispatch };
}

afterEach(() => {
  vi.unstubAllGlobals();
  mocks.render.mockReset();
  mocks.paginationFailure = undefined;
  mocks.registeredFonts.clear();
  mocks.nativeSentinels = [];
  FakeOffscreenCanvas.created = [];
});

describe('DOCX render worker bitmap surface', () => {
  it('adopts the parse request locale for what the worker paints, and keeps its own without one', async () => {
    const parse = (id: number, extra: Record<string, unknown>) => ({
      type: 'parse', id, data: new ArrayBuffer(1), defaultCurrentDateMs: 0,
      resourcePolicy: { maxArchiveEntryBytes: null, maxTotalInflatedBytes: null, maxArchiveEntries: null },
      ...extra,
    });
    // A page paint in the worker realm: the painter formats a built-in short
    // date (chart category and cell text) with the worker's adopted locale.
    const painted: Array<{ date: string }> = [];
    mocks.render.mockImplementation(async () => {
      const core = await import('@silurus/ooxml-core');
      painted.push({ date: core.formatLocalizedExcelShortDate(45_659, false) });
    });
    const { document, dispatch } = await loadWorker(false);
    await dispatch(parse(101, {}));
    await document.renderPageToBitmap(0, { dpr: 1 });
    await dispatch(parse(102, { hostLocale: 'fr-FR' }));
    await document.renderPageToBitmap(0, { dpr: 1 });
    expect(painted.map(entry => entry.date)).toEqual(['1/2/2025', '02/01/2025']);
  });

  it('retains ordinary progressive page reuse metadata after source dispatch is shared', async () => {
    const { worker, dispatch } = await loadWorker();
    await dispatch({
      type: 'parse', id: 101, data: new ArrayBuffer(1), defaultCurrentDateMs: 0,
      resourcePolicy: { maxArchiveEntryBytes: null, maxTotalInflatedBytes: null, maxArchiveEntries: null },
      progressiveLayout: true,
    });
    expect(worker.received).toContainEqual(expect.objectContaining({
      type: 'layoutPartial', forId: 101,
      partial: expect.objectContaining({ pageCount: 1, unchangedPages: 0 }),
    }));
    expect(worker.received).toContainEqual(expect.objectContaining({
      type: 'parsedMeta', id: 101, unchangedPages: 1,
    }));
  });

  it('keeps published pages and their fonts usable after a late pagination failure until replacement', async () => {
    const { document, worker, dispatch } = await loadWorker(false);
    mocks.paginationFailure = 'after-publication';
    await dispatch({
      type: 'parse', id: 101, data: new ArrayBuffer(1), defaultCurrentDateMs: 0,
      resourcePolicy: { maxArchiveEntryBytes: null, maxTotalInflatedBytes: null, maxArchiveEntries: null },
      progressiveLayout: true,
    });
    expect(worker.received).toContainEqual(expect.objectContaining({ type: 'layoutPartial', forId: 101 }));
    expect(worker.received).toContainEqual(expect.objectContaining({
      type: 'error', id: 101, message: 'pagination failure',
    }));
    mocks.render.mockImplementation(async () => {
      if (mocks.registeredFonts.size !== 1) throw new Error('Published fonts were released');
    });
    await expect(document.renderPageToBitmap(0, { dpr: 1 })).resolves.toBeDefined();

    const publishedFonts = [...mocks.registeredFonts];
    mocks.paginationFailure = undefined;
    await dispatch({
      type: 'parse', id: 102, data: new ArrayBuffer(1), defaultCurrentDateMs: 0,
      resourcePolicy: { maxArchiveEntryBytes: null, maxTotalInflatedBytes: null, maxArchiveEntries: null },
    });
    expect(worker.received).toContainEqual(expect.objectContaining({ type: 'parsedMeta', id: 102 }));
    expect(publishedFonts.every((face) => !mocks.registeredFonts.has(face))).toBe(true);
    await expect(document.renderPageToBitmap(0, { dpr: 1 })).resolves.toBeDefined();
  });

  it('releases unpublished pages and fonts when pagination fails before publication', async () => {
    const { document, worker, dispatch } = await loadWorker(false);
    mocks.paginationFailure = 'before-publication';
    await dispatch({
      type: 'parse', id: 101, data: new ArrayBuffer(1), defaultCurrentDateMs: 0,
      resourcePolicy: { maxArchiveEntryBytes: null, maxTotalInflatedBytes: null, maxArchiveEntries: null },
      progressiveLayout: true,
    });
    expect(worker.received).toContainEqual(expect.objectContaining({ type: 'error', id: 101 }));
    expect(worker.received).not.toContainEqual(expect.objectContaining({ type: 'layoutPartial', forId: 101 }));
    expect(mocks.registeredFonts.size).toBe(0);
    await expect(document.renderPageToBitmap(0, { dpr: 1 })).rejects.toThrow('Document not loaded');
  });

  it('continues on the main thread when this realm cannot register the route sentinels', async () => {
    // Main mode admits measured native-route lines with its own sentinels; a
    // worker that cannot register them must not lay the document out without
    // them, or the two modes would paginate differently.
    const { worker, dispatch } = await loadWorker(false);
    mocks.nativeSentinels = null;
    await dispatch({
      type: 'parse', id: 101, data: new ArrayBuffer(1), defaultCurrentDateMs: 0,
      resourcePolicy: { maxArchiveEntryBytes: null, maxTotalInflatedBytes: null, maxArchiveEntries: null },
    });
    expect(worker.received).toContainEqual(expect.objectContaining({ type: 'mainThreadVerticalFallback', id: 101 }));
    expect(worker.received).not.toContainEqual(expect.objectContaining({ type: 'parsedMeta', id: 101 }));
  });

  it('suppresses ordinary progressive cancellation responses', async () => {
    const { worker, dispatch } = await loadWorker(false);
    mocks.paginationFailure = 'abort';
    await dispatch({
      type: 'parse', id: 101, data: new ArrayBuffer(1), defaultCurrentDateMs: 0,
      resourcePolicy: { maxArchiveEntryBytes: null, maxTotalInflatedBytes: null, maxArchiveEntries: null },
      progressiveLayout: true,
    });
    expect(worker.received).not.toContainEqual(expect.objectContaining({ type: 'error', id: 101 }));
    expect(worker.received).not.toContainEqual(expect.objectContaining({ type: 'parsedMeta', id: 101 }));
  });

  it('rejects resource usage before an ordinary document is loaded', async () => {
    const { worker, dispatch } = await loadWorker(false);
    await dispatch({ type: 'resourceUsage', id: 101 });
    expect(worker.received).toContainEqual(expect.objectContaining({
      type: 'error', id: 101, message: 'No docx loaded',
    }));
  });

  it('paints renderPageToBitmap into the contextless 1x1 canvas transferred from the caller', async () => {
    const { document, worker } = await loadWorker();

    const bitmap = await document.renderPageToBitmap(0, { dpr: 1 });

    expect(FakeOffscreenCanvas.created).toHaveLength(1);
    const [canvas] = FakeOffscreenCanvas.created;
    const request = worker.posted.find((post) => post.message.type === 'renderPage')!;
    expect(request.message.canvas).toBe(canvas);
    expect(request.transfer).toHaveLength(1);
    expect(request.transfer[0]).toBe(canvas);
    expect(request.surfaces).toEqual([{ width: 1, height: 1, contextRequests: 0 }]);
    expect(mocks.render).toHaveBeenCalledExactlyOnceWith(mocks.store, canvas, 0, expect.anything());
    expect(bitmap).toBe(canvas!.bitmap);
  });

  it('keeps a worker-created context failure terminal on the ordinary worker error wire', async () => {
    const { worker, dispatch } = await loadWorker();
    mocks.render.mockImplementation(async (_source, canvas, _page, options) => {
      acquireDocumentCanvasTargetContext(canvas as OffscreenCanvas,
        (options as { callerOwnedCanvasTarget?: boolean }).callerOwnedCanvasTarget !== false);
    });
    await dispatch({ type: 'renderPage', id: 17, pageIndex: 0, opts: { dpr: 1 } });
    const failure = worker.received.find((value) => typeof value === 'object' && value !== null
      && 'type' in value && value.type === 'error' && 'id' in value && value.id === 17);
    expect(failure).toBeDefined();
    expect(failure).toHaveProperty('message', '2D canvas is unavailable for internally acquired DOCX paint surface');
    expect(failure).not.toHaveProperty('code', 'docx-caller-input');
    expect(worker.received).not.toContainEqual(expect.objectContaining({ type: 'pageRendered', id: 17 }));
    expect(FakeOffscreenCanvas.created[0]!.width).toBe(1);
    expect(FakeOffscreenCanvas.created[0]!.height).toBe(1);
  });
  it('keeps a worker-local surface for a request that carries no canvas', async () => {
    const { worker, dispatch } = await loadWorker();

    await dispatch({ type: 'renderPage', id: 7, pageIndex: 0, opts: { dpr: 1 } });

    expect(FakeOffscreenCanvas.created).toHaveLength(1);
    const [canvas] = FakeOffscreenCanvas.created;
    expect(mocks.render).toHaveBeenCalledExactlyOnceWith(mocks.store, canvas, 0, expect.anything());
    expect(worker.received).toContainEqual(expect.objectContaining({
      type: 'pageRendered', id: 7, bitmap: canvas!.bitmap,
    }));
  });
});
