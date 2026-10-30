import { afterEach, describe, expect, it, vi } from 'vitest';
import type { PullSessionCommand } from '@silurus/ooxml-core/worker';
import { XlsxWorkbook } from './workbook.js';
import { withViewerRenderContext, type RenderWorkerRequest } from './worker-protocol.js';
import { withXlsxRenderCommitGuard } from './render-orchestrator.js';
import type { ParsedWorkbook, Worksheet, WorkerRequest } from './types.js';

/**
 * #1974 — XLSX sheet failures are recorded only when main admits the terminal
 * model, and a function `errorPlaceholder` never crosses to the worker: when
 * main only holds a provisional preview, the worker asks back once and the
 * retry carries resolved text after main's own admission.
 */
type Message = WorkerRequest | RenderWorkerRequest | PullSessionCommand<number>;
const MESSAGE = 'xl/worksheets/sheet1.xml: unexpected end of stream';

const SHEET: Worksheet = {
  name: 'Budget',
  rows: [],
  colWidths: {},
  rowHeights: {},
  defaultColWidth: 8.43,
  defaultRowHeight: 15,
  mergeCells: [],
  freezeRows: 0,
  freezeCols: 0,
  conditionalFormats: [],
  images: [],
  charts: [],
};

function makeWorkbook(respond: (message: Message) => Promise<Record<string, unknown>>, mode: 'main' | 'worker' = 'worker') {
  vi.stubGlobal('OffscreenCanvas', class { constructor(readonly width: number, readonly height: number) {} });
  let nextId = 1;
  const messages: Message[] = [];
  const request = vi.fn((build: (id: number) => Message) => {
    const message = build(nextId++);
    messages.push(message);
    return respond(message);
  });
  const bridge = { request, transport: () => bridge, forgetOrphaned: vi.fn(), terminate: vi.fn() };
  const instance = Object.create(XlsxWorkbook.prototype) as Record<string, unknown>;
  Object.assign(instance, {
    _mode: mode,
    parsedWorkbook: { workbook: { sheets: [{ name: 'Budget' }] }, sharedStrings: [], styles: {} } as unknown as ParsedWorkbook,
    sheetCache: new Map(),
    sheetCacheUsage: new Map(),
    sheetLeases: new Map(),
    evictingSheets: new Map(),
    sheetLoads: new Map(),
    bridge,
    retainedSheetUsage: { rows: 0, cells: 0, ownedUtf8Bytes: 0, jsonBytes: 0 },
    resourceFailure: null,
    retainedFontSets: new Map(),
    rawParts: { clear: vi.fn() },
  });
  return { workbook: instance as unknown as XlsxWorkbook, messages, bridge };
}

function pullResponse(message: Message, options: { degraded?: boolean; loseTerminalAck?: boolean } = {}) {
  if ('type' in message) return { type: 'sheetSessionOpened', id: 'id' in message ? message.id : 0 };
  if (message.kind === 'pull') {
    const value = message.sequence === 0
      ? { kind: 'rows', rows: [] }
      : { kind: 'finished', worksheet: { ...SHEET, ...(options.degraded ? { parseError: MESSAGE } : {}) } };
    const payload = new TextEncoder().encode(JSON.stringify(value)).buffer;
    return { ...message, kind: 'chunk', byteLength: payload.byteLength, done: message.sequence === 1, payload };
  }
  if (options.loseTerminalAck && message.kind === 'ack' && message.sequence === 1) {
    throw new Error('terminal ACK reply lost');
  }
  return { ...message, kind: 'accepted', command: message.kind };
}

afterEach(() => {
  vi.unstubAllGlobals();
});

const failure = { kind: 'sheet', sheetIndex: 0, sheetName: 'Budget', message: MESSAGE };
const viewport = { row: 1, col: 1, rows: 1, cols: 1 };

describe('XlsxWorkbook part failures (#1974)', () => {
  it('records a degraded sheet once its terminal is admitted, deduplicated on re-pull', async () => {
    const { workbook } = makeWorkbook(async (message) => pullResponse(message, { degraded: true }));
    expect(workbook.partFailures).toEqual([]);
    await workbook.getWorksheet(0);
    expect(workbook.partFailures).toEqual([failure]);
    (workbook as unknown as { sheetCache: Map<number, Worksheet> }).sheetCache.clear();
    await workbook.getWorksheet(0);
    expect(workbook.partFailures).toEqual([failure]);
  });

  it('records nothing when the terminal ACK fails', async () => {
    const { workbook } = makeWorkbook(async (message) => pullResponse(message, { degraded: true, loseTerminalAck: true }));
    await expect(workbook.getWorksheet(0)).rejects.toThrow('terminal ACK reply lost');
    expect(workbook.partFailures).toEqual([]);
  });

  it('resolves the function on main after admission and never posts it', async () => {
    const bitmap = {} as ImageBitmap;
    const { workbook, messages } = makeWorkbook(async (message) => {
      if ('type' in message && message.type === 'renderViewport') return { type: 'viewportRendered', id: message.id, bitmap };
      return pullResponse(message, { degraded: true });
    });
    const resolver = vi.fn((f: { sheetName: string }) => ({ title: `${f.sheetName} is unavailable` }));
    await expect(workbook.renderViewportToBitmap(0, viewport, { width: 100, height: 80, errorPlaceholder: resolver })).resolves.toBe(bitmap);
    const render = messages.find((m) => 'type' in m && m.type === 'renderViewport') as Extract<RenderWorkerRequest, { type: 'renderViewport' }>;
    expect(render.opts).not.toHaveProperty('errorPlaceholder');
    expect(render.opts.errorPlaceholderText).toEqual({ state: 'resolved', title: 'Budget is unavailable', detail: undefined });
    expect(resolver).toHaveBeenCalledWith(failure);
    expect(() => structuredClone({ ...render, canvas: undefined })).not.toThrow();
  });

  it('never evaluates or posts the option for a healthy sheet', async () => {
    const bitmap = {} as ImageBitmap;
    const { workbook, messages } = makeWorkbook(async (message) => {
      if ('type' in message && message.type === 'renderViewport') return { type: 'viewportRendered', id: message.id, bitmap };
      return pullResponse(message);
    });
    const resolver = vi.fn(() => ({ title: 'x' }));
    await workbook.renderViewportToBitmap(0, viewport, { width: 100, height: 80, errorPlaceholder: resolver });
    const render = messages.find((m) => 'type' in m && m.type === 'renderViewport') as Extract<RenderWorkerRequest, { type: 'renderViewport' }>;
    expect(resolver).not.toHaveBeenCalled();
    expect(render.opts).not.toHaveProperty('errorPlaceholder');
    expect(render.opts).not.toHaveProperty('errorPlaceholderText');
  });

  function previewWorkbook(options: { loseTerminalAck?: boolean } = {}) {
    const bitmap = {} as ImageBitmap;
    const renders: Extract<RenderWorkerRequest, { type: 'renderViewport' }>[] = [];
    const harness = makeWorkbook(async (message) => {
      if ('type' in message && message.type === 'renderViewport') {
        renders.push(message);
        return message.opts.errorPlaceholderText?.state === 'resolve-on-main'
          ? { type: 'errorPlaceholderRequired', id: message.id, sheetIndex: 0 }
          : { type: 'viewportRendered', id: message.id, bitmap };
      }
      return pullResponse(message, { degraded: true, ...options });
    });
    // Main holds a provisional preview; the worker already holds the terminal.
    const provisional = { ...SHEET };
    (harness.workbook as unknown as { sheetPreviews: Map<number, unknown> }).sheetPreviews = new Map([
      [0, { worksheet: provisional, waitFor: async () => undefined }],
    ]);
    const opts = withViewerRenderContext({ width: 100, height: 80 }, 7, { worksheet: provisional });
    return { ...harness, bitmap, renders, opts };
  }

  it('asks back once from a provisional preview, then retries with resolved text', async () => {
    const { workbook, bitmap, renders, opts } = previewWorkbook();
    const resolver = vi.fn(() => ({ detail: null }));
    const rendered = workbook.renderViewportToBitmap(0, viewport, { ...opts, errorPlaceholder: resolver });
    await expect(rendered).resolves.toBe(bitmap);
    expect(renders.map((r) => r.opts.errorPlaceholderText?.state)).toEqual(['resolve-on-main', 'resolved']);
    expect(resolver).toHaveBeenCalledTimes(1);
    expect(workbook.partFailures).toEqual([failure]);
  });

  it('a resolver returning undefined is sent as built-in and never asked again', async () => {
    const { workbook, renders, opts } = previewWorkbook();
    await workbook.renderViewportToBitmap(0, viewport, { ...opts, errorPlaceholder: () => undefined });
    expect(renders.map((r) => r.opts.errorPlaceholderText?.state)).toEqual(['resolve-on-main', 'builtin']);
  });

  it('rejects without recording when admission fails between the reply and the retry', async () => {
    const { workbook, renders, opts } = previewWorkbook({ loseTerminalAck: true });
    await expect(workbook.renderViewportToBitmap(0, viewport, { ...opts, errorPlaceholder: () => ({ title: 't' }) }))
      .rejects.toThrow('terminal ACK reply lost');
    expect(renders).toHaveLength(1);
    expect(workbook.partFailures).toEqual([]);
  });

  it('drops the retry of a frame superseded during admission', async () => {
    const { workbook, renders, opts } = previewWorkbook();
    vi.stubGlobal('OffscreenCanvas', class {
      constructor(readonly width: number, readonly height: number) {}
      transferToImageBitmap() { return { superseded: true } as unknown as ImageBitmap; }
    });
    let current = true;
    const resolver = vi.fn(() => ({ title: 't' }));
    const guarded = withXlsxRenderCommitGuard({ ...opts, errorPlaceholder: resolver } as never, () => current);
    const pending = workbook.renderViewportToBitmap(0, viewport, guarded as never);
    current = false;
    await expect(pending).resolves.toEqual({ superseded: true });
    expect(renders).toHaveLength(1);
    expect(resolver).not.toHaveBeenCalled();
    expect(workbook.partFailures).toEqual([failure]);
  });

  it('sends static text without a round trip', async () => {
    const { workbook, renders, opts } = previewWorkbook();
    await workbook.renderViewportToBitmap(0, viewport, { ...opts, errorPlaceholder: { title: 'Static' } });
    expect(renders.map((r) => r.opts.errorPlaceholderText)).toEqual([{ state: 'resolved', title: 'Static', detail: undefined }]);
  });

  it('empties the list on destroy', async () => {
    const { workbook } = makeWorkbook(async (message) => pullResponse(message, { degraded: true }));
    await workbook.getWorksheet(0);
    try { workbook.destroy(); } catch { /* off-prototype teardown of unrelated resources */ }
    expect(workbook.partFailures).toEqual([]);
  });
});
