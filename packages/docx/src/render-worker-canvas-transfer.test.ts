import { afterEach, describe, expect, it, vi } from 'vitest';
import { WorkerBridge, type WorkerLike } from '@silurus/ooxml-core/worker';
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
  loadOfficeFontFallbacks: async () => ({ faces: [], routes: {} }),
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
    document: {}, source: { fatalParse: null, mathOccurrences: [] },
  }),
}));
vi.mock('./google-fonts.js', async (load) => ({
  ...await load<typeof import('./google-fonts.js')>(),
  docxOfficeFontFallbackRequests: () => [],
}));
vi.mock('./layout-runtime.js', async (load) => ({
  ...await load<typeof import('./layout-runtime.js')>(),
  createLayoutServices: () => ({}),
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

async function loadWorker(entry: 'render-worker' | 'render-worker-source') {
  vi.resetModules();
  const scope = {
    onmessage: null as ((event: MessageEvent) => Promise<void>) | null,
    postMessage: (message: unknown) => worker.receive(message),
  };
  const worker = new LoopbackWorker(scope);
  vi.stubGlobal('self', scope);
  vi.stubGlobal('OffscreenCanvas', FakeOffscreenCanvas);
  if (entry === 'render-worker') await import('./render-worker.js');
  else await import('./render-worker-source.js');
  const dispatch = (data: unknown) => scope.onmessage!({ data } as MessageEvent);
  await dispatch({ type: 'init', wasmUrl: 'x' });
  await dispatch({
    type: 'parse', id: 100, data: new ArrayBuffer(1), defaultCurrentDateMs: 0,
    resourcePolicy: { maxArchiveEntryBytes: null, maxTotalInflatedBytes: null, maxArchiveEntries: null },
  });
  const bridge = new WorkerBridge<RenderWorkerResponse>(worker, {
    correlate: (response) => ('id' in response ? response.id : undefined),
  });
  const document = Object.create(DocxDocument.prototype) as DocxDocument;
  Object.assign(document, { _mode: 'worker', _bridge: bridge });
  attachDocumentLayoutRuntime(document, 0);
  return { document, worker, dispatch };
}

afterEach(() => {
  vi.unstubAllGlobals();
  mocks.render.mockClear();
  FakeOffscreenCanvas.created = [];
});

describe.each(['render-worker', 'render-worker-source'] as const)('DOCX %s bitmap surface', (entry) => {
  it('paints renderPageToBitmap into the contextless 1x1 canvas transferred from the caller', async () => {
    const { document, worker } = await loadWorker(entry);

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

  it('keeps a worker-local surface for a request that carries no canvas', async () => {
    const { worker, dispatch } = await loadWorker(entry);

    await dispatch({ type: 'renderPage', id: 7, pageIndex: 0, opts: { dpr: 1 } });

    expect(FakeOffscreenCanvas.created).toHaveLength(1);
    const [canvas] = FakeOffscreenCanvas.created;
    expect(mocks.render).toHaveBeenCalledExactlyOnceWith(mocks.store, canvas, 0, expect.anything());
    expect(worker.received).toContainEqual(expect.objectContaining({
      type: 'pageRendered', id: 7, bitmap: canvas!.bitmap,
    }));
  });
});
