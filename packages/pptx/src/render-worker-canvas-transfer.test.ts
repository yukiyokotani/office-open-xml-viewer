import { afterEach, describe, expect, it, vi } from 'vitest';
import { WorkerBridge, type WorkerLike } from '@silurus/ooxml-core/worker';
import { PptxPresentation } from './presentation.js';
import type { RenderWorkerResponse } from './worker-protocol.js';

/**
 * Worker-mode slide painting must use a surface created by the caller and
 * transferred, exactly like the main-mode path paints into one created in the
 * caller. Run collection paints too, so it follows the same rule. The loopback
 * hands the posted object to the worker unchanged, standing in for the
 * transfer; the parser is stubbed so the real worker entries run without WASM.
 */

const mocks = vi.hoisted(() => ({
  render: vi.fn(async (..._args: unknown[]) => undefined),
  elements: [] as unknown[],
}));

class FakePptxArchive {
  constructor(_bytes: Uint8Array) {}
  presentation_bootstrap(): Uint8Array {
    return new TextEncoder().encode(JSON.stringify({
      slideCount: 1, slideWidth: 914400, slideHeight: 914400,
      defaultTextColor: null, majorFont: null, minorFont: null,
      hlinkColor: null, folHlinkColor: null, embeddedFonts: [],
      slides: [{ index: 0, partName: 'ppt/slides/slide1.xml' }],
    }));
  }
  pull_slide(): Uint8Array {
    return new TextEncoder().encode(JSON.stringify({
      index: 0, slideNumber: 1, partName: 'ppt/slides/slide1.xml',
      background: null, elements: mocks.elements, notes: null, hidden: false,
    }));
  }
  slide_cursor_resource_usage(): Uint8Array {
    return new TextEncoder().encode(JSON.stringify({
      archiveEntryCount: 1, declaredInflatedBytes: 1, distinctInflatedBytes: 1, operationInflatedBytes: 1,
    }));
  }
  acknowledge_slide(): void {}
  cancel_slide(): void {}
  close_presentation_session(): void {}
  assert_healthy(): void {}
  free(): void {}
}

vi.mock('./wasm/pptx_parser.js', () => ({
  default: async () => undefined, reinit: async () => undefined, PptxArchive: FakePptxArchive,
}));
vi.mock('./font-requests', async (load) => ({
  ...await load<typeof import('./font-requests')>(),
  pptxSlideOfficeFontRequests: () => [],
}));
vi.mock('./renderer', () => ({ renderSlideWithEmbeddedFonts: mocks.render, dropImageBitmapCache: vi.fn() }));

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

async function loadWorker(variant: 'parser' | 'model-source' | 'odp-source' = 'parser') {
  vi.resetModules();
  const scope = {
    onmessage: null as ((event: MessageEvent) => Promise<void>) | null,
    postMessage: (message: unknown) => worker.receive(message),
  };
  const worker = new LoopbackWorker(scope);
  vi.stubGlobal('self', scope);
  vi.stubGlobal('OffscreenCanvas', FakeOffscreenCanvas);
  await import('./render-worker.js');
  const dispatch = (data: unknown) => scope.onmessage!({ data } as MessageEvent);
  await dispatch({ kind: 'init', wasmUrl: 'x' });
  await dispatch({
    kind: 'parse', id: 100, buffer: new ArrayBuffer(1),
    resourcePolicy: { maxArchiveEntryBytes: null, maxTotalInflatedBytes: null, maxArchiveEntries: null },
    ...(variant !== 'parser' ? {
      source: { protocol: 'ooxml-model-source-module/v2', target: 'pptx', moduleUrl: 'x', config: {} },
      sourceOwnerUrl: new URL('./test-support/worker-source-owner.ts', import.meta.url).href,
    } : {}),
    ...(variant === 'odp-source' ? { odfTables: true } : {}),
  });
  expect(worker.received).toContainEqual(expect.objectContaining({ kind: 'presentationReady', id: 100 }));
  const bridge = new WorkerBridge<RenderWorkerResponse>(worker, {
    correlate: (response) => ('id' in response ? response.id : undefined),
  });
  const instance = Object.create(PptxPresentation.prototype) as Record<string, unknown>;
  Object.assign(instance, {
    _mode: 'worker', _bridge: bridge, _resourceFailure: null, _destroyed: false,
    _bootstrap: { slideCount: 1 }, _availableSlideCount: 1,
  });
  return { presentation: instance as unknown as PptxPresentation, worker, dispatch };
}

function expectCallerSurface(worker: LoopbackWorker, kind: string): FakeOffscreenCanvas {
  expect(FakeOffscreenCanvas.created).toHaveLength(1);
  const [canvas] = FakeOffscreenCanvas.created;
  const request = worker.posted.find((post) => post.message.kind === kind)!;
  expect(request.message.canvas).toBe(canvas);
  expect(request.transfer).toHaveLength(1);
  expect(request.transfer[0]).toBe(canvas);
  expect(request.surfaces).toEqual([{ width: 1, height: 1, contextRequests: 0 }]);
  expect(mocks.render).toHaveBeenCalledOnce();
  expect(mocks.render.mock.calls[0]![0]).toBe(canvas);
  return canvas!;
}

afterEach(() => {
  vi.unstubAllGlobals();
  mocks.render.mockReset();
  mocks.render.mockImplementation(async () => undefined);
  mocks.elements = [];
  FakeOffscreenCanvas.created = [];
});

describe('PPTX render worker bitmap surface', () => {
  it('paints renderSlideToBitmap into the contextless 1x1 canvas transferred from the caller', async () => {
    const { presentation, worker } = await loadWorker();

    const bitmap = await presentation.renderSlideToBitmap(0, { width: 320, dpr: 1 });

    expect(bitmap).toBe(expectCallerSurface(worker, 'renderSlide').bitmap);
  });

  it('paints collectSlideRuns into the contextless 1x1 canvas transferred from the caller', async () => {
    const { presentation, worker } = await loadWorker();
    mocks.render.mockImplementationOnce(async (...args: unknown[]) => {
      (args[5] as (run: unknown) => void)({ text: 'run' });
    });

    await expect(presentation.collectSlideRuns(0, 320)).resolves.toEqual([{ text: 'run' }]);

    expectCallerSurface(worker, 'collectRuns');
  });

  it('keeps a worker-local surface for requests that carry no canvas', async () => {
    const { worker, dispatch } = await loadWorker();

    await dispatch({ kind: 'renderSlide', id: 7, slideIndex: 0, width: 320, dpr: 1 });
    await dispatch({ kind: 'collectRuns', id: 8, slideIndex: 0, width: 320 });

    expect(FakeOffscreenCanvas.created).toHaveLength(2);
    expect(mocks.render.mock.calls[0]![0]).toBe(FakeOffscreenCanvas.created[0]);
    expect(mocks.render.mock.calls[1]![0]).toBe(FakeOffscreenCanvas.created[1]);
    expect(worker.received).toContainEqual(expect.objectContaining({
      kind: 'slideRendered', id: 7, bitmap: FakeOffscreenCanvas.created[0]!.bitmap,
    }));
    expect(worker.received).toContainEqual(expect.objectContaining({ kind: 'runsCollected', id: 8 }));
  });
});

describe('PPTX render worker ODP table sizing', () => {
  // The model-source loader marks an LoadOptions.odp parse; the render
  // worker carries that to every render so the source's tables keep the ODF
  // row-height rule without a parser-emitted field.
  it.each([['parser', false], ['model-source', false], ['odp-source', true]] as const)(
    'passes odfTables for %s slides when painting and collecting runs', async (variant, expected) => {
      const { presentation } = await loadWorker(variant);

      await presentation.renderSlideToBitmap(0, { width: 320, dpr: 1 });
      await presentation.collectSlideRuns(0, 320);

      expect(mocks.render).toHaveBeenCalledTimes(2);
      for (const call of mocks.render.mock.calls) {
        expect((call[4] as { odfTables?: boolean }).odfTables === true).toBe(expected);
      }
    },
  );
});

describe('PPTX render worker parsed text-colour chain', () => {
  // Issue #1989's hard black default needs the built-in parser's resolved
  // colour chain; a model source keeps the 0.89.7 fallback.
  it.each([['parser', true], ['model-source', false]] as const)(
    'passes parsedTextChain for %s slides when painting and collecting runs', async (variant, expected) => {
      const { presentation } = await loadWorker(variant);

      await presentation.renderSlideToBitmap(0, { width: 320, dpr: 1 });
      await presentation.collectSlideRuns(0, 320);

      expect(mocks.render).toHaveBeenCalledTimes(2);
      for (const call of mocks.render.mock.calls) {
        expect((call[4] as { parsedTextChain?: boolean }).parsedTextChain).toBe(expected);
      }
    },
  );
});

describe('PPTX render worker table hit testing', () => {
  // A table whose rows grow paints beyond its authored frame; the worker hit
  // tests against the extent its last render reported, and sizes the tables
  // without painting when the slide was never rendered.
  const EMU = 12700;
  const table = {
    type: 'table', x: 0, y: 0, width: 100 * EMU, height: 20 * EMU, rotation: 0, flipH: false, flipV: false,
    cols: [100 * EMU], rows: [{ height: 20 * EMU, cells: [] }],
  };
  const reportGrownRow = async (...args: unknown[]): Promise<undefined> => {
    (args[4] as { onTableExtent?: (index: number, height: number) => void }).onTableExtent?.(0, 24 * EMU);
    return undefined;
  };
  const hitAt = async (dispatch: (data: unknown) => Promise<void>, worker: LoopbackWorker, id: number, y: number) => {
    await dispatch({ kind: 'hitTestElement', id, slideIndex: 0, point: { x: 10 * EMU, y }, options: {} });
    const reply = worker.received.find((message) => (message as { id?: number }).id === id) as
      { context: { elementType: string; bounds: { height: number } } | null };
    return reply.context;
  };

  it('hits the painted extent of a grown table after a render', async () => {
    mocks.elements = [table];
    mocks.render.mockImplementation(reportGrownRow);
    const { worker, dispatch } = await loadWorker();
    await dispatch({ kind: 'renderSlide', id: 7, slideIndex: 0, width: 320, dpr: 1 });
    const hit = await hitAt(dispatch, worker, 8, 22 * EMU);
    expect(hit?.elementType).toBe('table');
    expect(hit?.bounds.height).toBe(20 * EMU);
    expect(await hitAt(dispatch, worker, 9, 25 * EMU)).toBeNull();
    expect(mocks.render).toHaveBeenCalledOnce();
  });

  it('sizes the tables without painting when the slide was never rendered', async () => {
    mocks.elements = [table];
    mocks.render.mockImplementation(reportGrownRow);
    const { worker, dispatch } = await loadWorker();
    expect((await hitAt(dispatch, worker, 8, 22 * EMU))?.elementType).toBe('table');
    expect(mocks.render).toHaveBeenCalledOnce();
    expect((mocks.render.mock.calls[0]![4] as { measureTablesOnly?: boolean }).measureTablesOnly).toBe(true);
  });
});
