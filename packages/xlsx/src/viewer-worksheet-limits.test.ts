import { afterEach, describe, expect, it, vi } from 'vitest';
import { serializeWorkerError } from '@silurus/ooxml-core/worker';
import { OoxmlResourceLimitError } from '@silurus/ooxml-core';
import { XlsxSheetViewer, XlsxViewer } from './viewer.js';
import { XlsxWorkbook, loadXlsxSheetSource, type LoadOptions } from './workbook.js';
import { parseDelimitedWorksheet } from './delimited-text.js';
import { normalizeXlsxWorksheetPolicy } from './internal/worksheet-policy.js';
import type { DelimitedTextParseRequest } from './delimited-text-protocol.js';
import { installDom, makeContainer, makeEl } from './viewer-destroy-test-dom.js';

// #1917: the self-loading viewers build their own XlsxWorkbook load options.
// Any LoadOptions field omitted there is silently replaced by the workbook
// default, e.g. a 250,001-cell sheet failed at the default 250,000-cell limit
// even though `xlsxWorksheetLimits.maxCells` was raised on the viewer.

/** Runs the production delimited-text worker body in-process: it applies the
 *  worksheet policy carried by the request, exactly like delimited-text-worker.ts. */
class PolicyDelimitedTextWorker {
  static requests: DelimitedTextParseRequest[] = [];
  private readonly listeners = new Set<(event: MessageEvent) => void>();

  postMessage(message: unknown): void {
    const request = message as DelimitedTextParseRequest;
    if (request.type !== 'parseDelimitedText') return;
    PolicyDelimitedTextWorker.requests.push(request);
    let data: unknown;
    try {
      const policy = normalizeXlsxWorksheetPolicy({
        xlsxWorksheetLimits: request.worksheetPolicy?.worksheet,
      });
      const { workbook, worksheet } = parseDelimitedWorksheet(request.data, request.options, policy);
      const worksheetJson = new TextEncoder().encode(JSON.stringify(worksheet)).buffer as ArrayBuffer;
      data = { type: 'delimitedTextParsed', id: request.id, workbook, worksheetJson };
    } catch (error) {
      data = { type: 'error', id: request.id, ...serializeWorkerError(error) };
    }
    queueMicrotask(() => {
      for (const listener of this.listeners) listener({ data } as MessageEvent);
    });
  }

  addEventListener(type: string, listener: (event: MessageEvent) => void): void {
    if (type === 'message') this.listeners.add(listener);
  }

  removeEventListener(type: string, listener: (event: MessageEvent) => void): void {
    if (type === 'message') this.listeners.delete(listener);
  }

  terminate(): void {}
}

vi.mock('./delimited-text-worker-host', () => ({
  createDelimitedTextWorker: () => new PolicyDelimitedTextWorker() as unknown as Worker,
}));

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
  PolicyDelimitedTextWorker.requests = [];
});

const RAISED = { maxCells: 500_000 } as const;

function fakeWorkbook(mode: 'main' | 'worker'): XlsxWorkbook {
  return {
    mode,
    sheetNames: ['Sheet1'],
    tabColors: [null],
    destroy: vi.fn(),
    getWorksheet: vi.fn().mockResolvedValue(undefined),
  } as unknown as XlsxWorkbook;
}

type Surface = 'XlsxViewer' | 'XlsxSheetViewer';

function makeViewer(surface: Surface, options: LoadOptions & { onError?: (e: Error) => void }) {
  if (surface === 'XlsxViewer') {
    const viewer = new XlsxViewer(makeContainer() as unknown as HTMLElement, options);
    const showSheet = vi.spyOn(
      viewer as unknown as { showSheet(index: number): Promise<void> }, 'showSheet',
    ).mockResolvedValue(undefined);
    return { viewer, showSheet, load: (source: ArrayBuffer) => viewer.load(source) };
  }
  const viewer = new XlsxSheetViewer(makeEl('canvas') as unknown as HTMLCanvasElement, options);
  const engine = (viewer as unknown as { engine: { showSheet(index: number): Promise<void> } }).engine;
  const showSheet = vi.spyOn(engine, 'showSheet').mockResolvedValue(undefined);
  return { viewer, showSheet, load: (source: ArrayBuffer) => viewer.load(source) };
}

describe('self-loading viewers forward xlsxWorksheetLimits (#1917)', () => {
  for (const surface of ['XlsxViewer', 'XlsxSheetViewer'] as const) {
    for (const mode of ['main', 'worker'] as const) {
      it(`${surface} (${mode}) passes the configured limits to XlsxWorkbook.load on load and reload`, async () => {
        installDom();
        const load = vi.spyOn(XlsxWorkbook, 'load')
          .mockImplementation(async () => fakeWorkbook(mode));
        const limits = { maxCells: 500_000, maxRows: 300_000 };
        const { viewer, load: loadViewer } = makeViewer(surface, { mode, xlsxWorksheetLimits: limits });

        await loadViewer(new ArrayBuffer(1));
        await loadViewer(new ArrayBuffer(1));

        expect(load).toHaveBeenCalledTimes(2);
        for (const call of load.mock.calls) {
          expect(call[1]).toMatchObject({ mode, xlsxWorksheetLimits: limits });
        }
        viewer.destroy();
      });
    }
  }

  it('XlsxSheetViewer forwards the limits to a delimited-text source', async () => {
    installDom();
    const dispatch = vi.spyOn(XlsxWorkbook, loadXlsxSheetSource)
      .mockImplementation(async () => fakeWorkbook('main'));
    const viewer = new XlsxSheetViewer(makeEl('canvas') as unknown as HTMLCanvasElement, {
      xlsxWorksheetLimits: RAISED,
    });
    const engine = (viewer as unknown as { engine: { showSheet(index: number): Promise<void> } }).engine;
    vi.spyOn(engine, 'showSheet').mockResolvedValue(undefined);

    await viewer.load(new ArrayBuffer(1), { format: 'csv' });

    expect(dispatch.mock.lastCall?.[1]).toMatchObject({ xlsxWorksheetLimits: RAISED });
    expect(dispatch.mock.lastCall?.[2]).toEqual({ format: 'csv' });
    viewer.destroy();
  });

  it('forwards every workbook LoadOptions field, not only the limits', async () => {
    installDom();
    const load = vi.spyOn(XlsxWorkbook, 'load').mockImplementation(async () => fakeWorkbook('main'));
    // One distinct sentinel per LoadOptions key; the compile-time guard is
    // ForwardedViewerLoadOptions in viewer.ts.
    const sentinel = (name: string) => ({ sentinel: name });
    const options = {
      password: 'pw',
      fonts: [sentinel('fonts')],
      useGoogleFonts: false,
      cjkFallback: 'sc',
      maxZipEntryBytes: 1234,
      resourceLimits: { maxArchiveEntries: 7 },
      xlsxWorksheetLimits: RAISED,
      debug: true,
      onResourceMetrics: vi.fn(),
      workerTimeoutMs: 9876,
      wasmUrl: '/custom.wasm',
      math: sentinel('math'),
      threeD: sentinel('threeD'),
      regionMap: sentinel('regionMap'),
      chartEx: sentinel('chartEx'),
      tiff: sentinel('tiff'),
      mode: 'main',
      recalculateFormulas: true,
      formualizerWasm: '/custom-fml.wasm',
      legacyXls: { target: 'xlsx', claim: () => false, beginLoad: vi.fn() },
    } as unknown as Required<LoadOptions>;
    const expectedKeys = Object.keys(options).sort();

    for (const surface of ['XlsxViewer', 'XlsxSheetViewer'] as const) {
      const { viewer, load: loadViewer } = makeViewer(surface, options);
      await loadViewer(new ArrayBuffer(1));
      const forwarded = load.mock.lastCall?.[1] as Record<string, unknown>;
      expect(Object.keys(forwarded).sort()).toEqual(expectedKeys);
      for (const key of expectedKeys) {
        if (key === 'onResourceMetrics' && surface === 'XlsxSheetViewer') {
          // XlsxSheetViewer wraps the callback to retain its last metrics.
          expect(forwarded[key]).toBeTypeOf('function');
          continue;
        }
        expect(forwarded[key], `${surface}.${key}`).toBe(options[key as keyof LoadOptions]);
      }
      viewer.destroy();
    }
  });

  it('rejects invalid limits through the same workbook validation', async () => {
    installDom();
    const viewer = new XlsxViewer(makeContainer() as unknown as HTMLElement, {
      xlsxWorksheetLimits: { maxCells: 0 },
    });
    await expect(viewer.load(new ArrayBuffer(1))).rejects.toThrow(
      /xlsxWorksheetLimits\.maxCells must be a positive safe integer/,
    );
    viewer.destroy();
  });
});

describe('XlsxSheetViewer admits one cell over the default limit when raised (#1917)', () => {
  // 4,717 rows x 53 columns = 250,001 numeric cells: exactly one over the
  // default 250,000-cell worksheet limit and well inside the row limit.
  const ROWS = 4_717;
  const COLS = 53;
  const csv = () => new TextEncoder().encode(
    `${Array(COLS).fill('1').join(',')}\n`.repeat(ROWS),
  ).buffer as ArrayBuffer;

  async function settle(): Promise<void> {
    for (let i = 0; i < 5; i++) await new Promise((resolve) => setTimeout(resolve, 0));
  }

  it('fails at the default limit and succeeds with maxCells: 500000, with no async onError', async () => {
    installDom();
    const defaultError = vi.fn();
    const defaults = new XlsxSheetViewer(
      makeEl('canvas') as unknown as HTMLCanvasElement,
      { onError: defaultError },
    );
    const rejected = defaults.load(csv(), { format: 'csv' });
    await expect(rejected).rejects.toBeInstanceOf(OoxmlResourceLimitError);
    await expect(rejected).rejects.toThrow(/cells 250001 > 250000/);
    defaults.destroy();

    const onError = vi.fn();
    const raised = new XlsxSheetViewer(
      makeEl('canvas') as unknown as HTMLCanvasElement,
      { onError, xlsxWorksheetLimits: RAISED },
    );
    const engine = (raised as unknown as {
      engine: { showSheet(index: number): Promise<void>; workbook: XlsxWorkbook | null };
    }).engine;
    // Rendering 250k rows on the fake DOM is out of scope; the admission
    // decision is made by the parse request, which runs for real here.
    vi.spyOn(engine, 'showSheet').mockResolvedValue(undefined);
    await raised.load(csv(), { format: 'csv' });
    // load() may resolve before a later Viewer-managed failure reaches onError.
    await settle();

    expect(onError).not.toHaveBeenCalled();
    const [rejectedRequest, admittedRequest] = PolicyDelimitedTextWorker.requests;
    expect(rejectedRequest?.worksheetPolicy?.worksheet.maxCells).toBe(250_000);
    expect(admittedRequest?.worksheetPolicy?.worksheet.maxCells).toBe(500_000);
    expect(raised.sheetCount).toBe(1);
    raised.destroy();
  }, 30_000);
});
