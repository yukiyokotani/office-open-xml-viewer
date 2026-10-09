import { afterEach, describe, expect, it, vi } from 'vitest';
import { WorkerBridge, type WorkerLike } from '@silurus/ooxml-core/worker';
import { XlsxWorkbook } from './workbook.js';
import type { RenderWorkerResponse } from './worker-protocol.js';
import type { ViewportRange } from './types.js';

/**
 * Worker-mode viewport painting must use a surface created by the caller and
 * transferred, exactly like the main-mode path paints into one created in the
 * caller. The loopback hands the posted object to the worker unchanged,
 * standing in for the transfer. A delimited-text sheet gives the real worker
 * entries a retained worksheet without generated WASM.
 */

const mocks = vi.hoisted(() => ({
  render: vi.fn(async (..._args: unknown[]) => undefined),
}));

vi.mock('./wasm/xlsx_parser.js', () => ({
  default: async () => undefined, reinit: async () => undefined, XlsxArchive: class {},
}));
vi.mock('@silurus/ooxml-core', async (load) => ({
  ...await load<typeof import('@silurus/ooxml-core')>(),
  loadOfficeFontFallbacks: async () => ({ faces: [], routes: {}, checked: [] }),
}));
vi.mock('./render-orchestrator.js', async (load) => ({
  ...await load<typeof import('./render-orchestrator.js')>(),
  renderWorksheetViewport: mocks.render,
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

const VIEWPORT = { row: 1, col: 1, rows: 1, cols: 1 } as unknown as ViewportRange;

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
  await dispatch({
    type: 'parseDelimitedText', id: 100, data: new TextEncoder().encode('a,b\n').buffer,
    options: { delimiter: ',', encoding: 'utf-8', sheetName: 'Sheet1' },
  });
  expect(worker.received).toContainEqual(expect.objectContaining({ type: 'delimitedTextParsed', id: 100 }));
  const bridge = new WorkerBridge<RenderWorkerResponse>(worker, {
    correlate: (response) => ('id' in response ? response.id : undefined),
  });
  const instance = Object.create(XlsxWorkbook.prototype) as Record<string, unknown>;
  Object.assign(instance, {
    _mode: 'worker', bridge, resourceFailure: null,
    parsedWorkbook: { workbook: { sheets: [{ name: 'Sheet1' }] }, sharedStrings: [], styles: {} },
    sheetCache: new Map([[0, { name: 'Sheet1', rows: [] }]]), sheetCacheUsage: new Map(),
    sheetLeases: new Map(), evictingSheets: new Map(), sheetLoads: new Map(),
    retainedSheetUsage: { rows: 0, cells: 0, ownedUtf8Bytes: 0, jsonBytes: 0 },
  });
  return { workbook: instance as unknown as XlsxWorkbook, worker, dispatch };
}

afterEach(() => {
  vi.unstubAllGlobals();
  mocks.render.mockClear();
  FakeOffscreenCanvas.created = [];
});

describe.each(['render-worker', 'render-worker-source'] as const)('XLSX %s bitmap surface', (entry) => {
  it('paints renderViewportToBitmap into the contextless 1x1 canvas transferred from the caller', async () => {
    const { workbook, worker } = await loadWorker(entry);

    const bitmap = await workbook.renderViewportToBitmap(0, VIEWPORT, { width: 100, height: 80, dpr: 1 });

    expect(FakeOffscreenCanvas.created).toHaveLength(1);
    const [canvas] = FakeOffscreenCanvas.created;
    const request = worker.posted.find((post) => post.message.type === 'renderViewport')!;
    expect(request.message.canvas).toBe(canvas);
    expect(request.transfer).toHaveLength(1);
    expect(request.transfer[0]).toBe(canvas);
    expect(request.surfaces).toEqual([{ width: 1, height: 1, contextRequests: 0 }]);
    expect(mocks.render).toHaveBeenCalledOnce();
    expect(mocks.render.mock.calls[0]![1]).toBe(canvas);
    expect(bitmap).toBe(canvas!.bitmap);
  });

  it('keeps a worker-local surface for a request that carries no canvas', async () => {
    const { worker, dispatch } = await loadWorker(entry);

    await dispatch({ type: 'renderViewport', id: 7, sheetIndex: 0, viewport: VIEWPORT, opts: { width: 100, height: 80, dpr: 1 } });

    expect(FakeOffscreenCanvas.created).toHaveLength(1);
    const [canvas] = FakeOffscreenCanvas.created;
    expect(mocks.render.mock.calls[0]![1]).toBe(canvas);
    expect(worker.received).toContainEqual(expect.objectContaining({
      type: 'viewportRendered', id: 7, bitmap: canvas!.bitmap,
    }));
  });
});
