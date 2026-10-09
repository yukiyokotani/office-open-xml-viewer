import { afterAll, beforeAll, describe, expect, it, vi } from 'vitest';
import { XlsxWorkbook } from './workbook.js';
import { withCfReportSink, type XlsxConditionalFormattingReport } from './cf-diagnostics.js';
import type { Worksheet } from './types.js';

// Exercises XlsxWorkbook.renderViewportToBitmap's real worker-mode host path
// (archive FIFO, wire decode, commit, sink, bitmap cleanup) against a fake
// bridge. It is NOT a real Worker; the browser suite must cover that.

const VIEWPORT = { row: 1, col: 1, rows: 10, cols: 5 };
const RECORD = { kind: 'unsupported', phase: 'expression', blockIndex: 0, ruleIndex: 1, row: 2, col: 3 } as const;
const OPTS = { width: 100, height: 50, dpr: 1 };

type Reply = { bitmap: ImageBitmap; conditionalFormatting: unknown };

function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((r) => { resolve = r; });
  return { promise, resolve };
}

function bitmap() {
  return { close: vi.fn() } as unknown as ImageBitmap & { close: ReturnType<typeof vi.fn> };
}

function workerModeWorkbook(replies: Array<Promise<Reply>>): XlsxWorkbook {
  const worksheet = {
    name: 'Sheet1', rows: [], colWidths: {}, rowHeights: {}, defaultColWidth: 8.43, defaultRowHeight: 15,
    mergeCells: [], freezeRows: 0, freezeCols: 0, conditionalFormats: [], images: [], charts: [],
  } as Worksheet;
  const workbook = Object.create(XlsxWorkbook.prototype) as XlsxWorkbook;
  let next = 0;
  Object.assign(workbook, {
    _mode: 'worker', generation: 1, resourceFailure: null, lastCfReport: null,
    parsedWorkbook: { workbook: { sheets: [{ name: 'Sheet1', sheetId: 1, rId: 'rId1' }] }, styles: {}, sharedStrings: [] },
    sheetCache: new Map([[0, worksheet]]), sheetLeases: new Map(), evictingSheets: new Map(), sheetLoads: new Map(),
    archiveOperationTail: Promise.resolve(),
    bridge: {
      request: async (build: (id: number) => unknown) => {
        const id = ++next;
        build(id);
        return { type: 'viewportRendered', id, ...(await replies[id - 1]) };
      },
    },
  });
  return workbook;
}

describe('XlsxWorkbook CF report commit (worker host path)', () => {
  beforeAll(() => {
    if (typeof OffscreenCanvas === 'undefined') {
      vi.stubGlobal('OffscreenCanvas', class { width = 1; height = 1; });
    }
  });
  afterAll(() => vi.unstubAllGlobals());

  it('commits a detached frozen report and delivers it to this call only', async () => {
    const surface = bitmap();
    const workbook = workerModeWorkbook([Promise.resolve({ bitmap: surface, conditionalFormatting: [RECORD] })]);
    expect(workbook.getLastConditionalFormattingDiagnostics()).toBeNull();
    const received: XlsxConditionalFormattingReport[] = [];
    const result = await workbook.renderViewportToBitmap(0, VIEWPORT, withCfReportSink(OPTS, (r) => received.push(r)));
    expect(result).toBe(surface);
    expect(surface.close).not.toHaveBeenCalled();
    const report = workbook.getLastConditionalFormattingDiagnostics();
    expect(report).toEqual({ sheetIndex: 0, viewport: VIEWPORT, diagnostics: [RECORD] });
    expect(received).toEqual([report]);
    expect(Object.isFrozen(report)).toBe(true);
    expect(Object.isFrozen(report!.diagnostics[0])).toBe(true);
    expect(structuredClone(report)).toEqual(report);
  });

  it('a successful supported frame replaces an older warning with []', async () => {
    const workbook = workerModeWorkbook([
      Promise.resolve({ bitmap: bitmap(), conditionalFormatting: [RECORD] }),
      Promise.resolve({ bitmap: bitmap(), conditionalFormatting: [] }),
    ]);
    await workbook.renderViewportToBitmap(0, VIEWPORT, OPTS);
    await workbook.renderViewportToBitmap(0, VIEWPORT, OPTS);
    expect(workbook.getLastConditionalFormattingDiagnostics()?.diagnostics).toEqual([]);
  });

  it('keeps the invocation viewport detached from later caller edits', async () => {
    const reply = deferred<Reply>();
    const workbook = workerModeWorkbook([reply.promise]);
    const viewport = { ...VIEWPORT };
    const pending = workbook.renderViewportToBitmap(0, viewport, OPTS);
    viewport.row = 50;
    reply.resolve({ bitmap: bitmap(), conditionalFormatting: [RECORD] });
    await pending;
    expect(workbook.getLastConditionalFormattingDiagnostics()?.viewport).toEqual(VIEWPORT);
  });

  it('a malformed batch rejects, closes the bitmap and keeps the committed report', async () => {
    const bad = bitmap();
    const workbook = workerModeWorkbook([
      Promise.resolve({ bitmap: bitmap(), conditionalFormatting: [RECORD] }),
      Promise.resolve({ bitmap: bad, conditionalFormatting: [{ ...RECORD, kind: 'bogus' }] }),
    ]);
    await workbook.renderViewportToBitmap(0, VIEWPORT, OPTS);
    const committed = workbook.getLastConditionalFormattingDiagnostics();
    const sink = vi.fn();
    await expect(workbook.renderViewportToBitmap(0, VIEWPORT, withCfReportSink(OPTS, sink))).rejects.toThrow(TypeError);
    expect(bad.close).toHaveBeenCalledTimes(1);
    expect(sink).not.toHaveBeenCalled();
    expect(workbook.getLastConditionalFormattingDiagnostics()).toBe(committed);
  });

  it('a failing operation-local sink propagates and closes its bitmap without committing a new report', async () => {
    const rejected = bitmap();
    const workbook = workerModeWorkbook([
      Promise.resolve({ bitmap: bitmap(), conditionalFormatting: [RECORD] }),
      Promise.resolve({ bitmap: rejected, conditionalFormatting: [] }),
    ]);
    await workbook.renderViewportToBitmap(0, VIEWPORT, OPTS);
    const previous = workbook.getLastConditionalFormattingDiagnostics();
    const failure = new Error('test sink failed');
    await expect(workbook.renderViewportToBitmap(0, VIEWPORT,
      withCfReportSink(OPTS, () => { throw failure; }))).rejects.toBe(failure);
    expect(rejected.close).toHaveBeenCalledOnce();
    expect(workbook.getLastConditionalFormattingDiagnostics()).toBe(previous);
  });

  it('concurrent calls: each sink gets its own frame; the getter is last completion', async () => {
    const first = deferred<Reply>();
    const second = deferred<Reply>();
    const workbook = workerModeWorkbook([first.promise, second.promise]);
    const a = vi.fn();
    const b = vi.fn();
    const pa = workbook.renderViewportToBitmap(0, VIEWPORT, withCfReportSink(OPTS, a));
    const pb = workbook.renderViewportToBitmap(0, { ...VIEWPORT, row: 50 }, withCfReportSink(OPTS, b));
    first.resolve({ bitmap: bitmap(), conditionalFormatting: [RECORD] });
    second.resolve({ bitmap: bitmap(), conditionalFormatting: [] });
    await Promise.all([pa, pb]);
    expect(a.mock.calls[0][0].diagnostics).toEqual([RECORD]);
    expect(b.mock.calls[0][0].diagnostics).toEqual([]);
    expect(b.mock.calls[0][0].viewport.row).toBe(50);
    expect(workbook.getLastConditionalFormattingDiagnostics()).toBe(b.mock.calls[0][0]);
  });
});
