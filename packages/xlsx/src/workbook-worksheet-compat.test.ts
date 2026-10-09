import { afterEach, describe, expect, it, vi } from 'vitest';
import { XlsxWorkbook, acquireXlsxWorksheet } from './workbook.js';
import { OoxmlResourceLimitError } from '@silurus/ooxml-core';
import type { PullSessionCommand } from '@silurus/ooxml-core/worker';
import type { ParsedWorkbook, Worksheet, WorkerRequest } from './types.js';
import type { RenderWorkerRequest } from './worker-protocol.js';

const WORKSHEET: Worksheet = {
  name: 'Sheet1',
  rows: [{
    index: 1,
    height: null,
    cells: [{ col: 1, row: 1, value: { type: 'shared', si: 0 } }],
  }],
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

const PARSED_WORKBOOK = {
  workbook: { sheets: [{ name: 'Sheet1' }] },
  sharedStrings: [{ text: 'resolved' }],
  styles: {},
} as ParsedWorkbook;

interface WorkbookProbe {
  getWorksheet(index: number): Promise<Worksheet>;
  toMarkdown(): Promise<string>;
  renderViewportToBitmap(
    sheetIndex: number,
    viewport: { startRow: number; endRow: number; startCol: number; endCol: number },
    options: { width: number; height: number },
  ): Promise<ImageBitmap>;
}

function makeWorkbook(
  mode: 'main' | 'worker',
  respond: (
    request: WorkerRequest | RenderWorkerRequest | PullSessionCommand<number>,
  ) => Promise<Record<string, unknown>>,
  workerTimeoutMs?: number,
) {
  // Worker mode requires OffscreenCanvas; the caller creates the render surface.
  if (mode === 'worker') {
    vi.stubGlobal('OffscreenCanvas', class { constructor(readonly width: number, readonly height: number) {} });
  }
  let nextId = 1;
  const request = vi.fn(
    (
      build: (id: number) => WorkerRequest | RenderWorkerRequest | PullSessionCommand<number>,
      _transfer?: Transferable[],
      _options?: { timeoutMs?: number | false },
    ) =>
      respond(build(nextId++)),
  );
  const bridge = {
    request,
    transport: () => bridge,
    forgetOrphaned: vi.fn(),
    terminate: vi.fn(),
  };
  const instance = Object.create(XlsxWorkbook.prototype) as Record<string, unknown>;
  instance._mode = mode;
  instance.parsedWorkbook = structuredClone(PARSED_WORKBOOK);
  instance.sheetCache = new Map();
  instance.sheetCacheUsage = new Map();
  instance.sheetLeases = new Map();
  instance.evictingSheets = new Map();
  instance.sheetLoads = new Map();
  instance.bridge = bridge;
  instance.retainedSheetUsage = { rows: 0, cells: 0, ownedUtf8Bytes: 0, jsonBytes: 0 };
  instance.resourceFailure = null;
  instance.workerTimeoutMs = workerTimeoutMs;
  instance.retainedFontSets = new Map();
  instance.rawParts = { clear: vi.fn() };
  return { workbook: instance as unknown as WorkbookProbe, request, bridge };
}

function streamResponse(
  message: WorkerRequest | RenderWorkerRequest | PullSessionCommand<number>,
): Record<string, unknown> {
  if ('type' in message) {
    expect(message).toMatchObject({ type: 'openSheetSession', sheetIndex: 0, sheetName: 'Sheet1' });
    return { type: 'sheetSessionOpened', id: 'id' in message ? message.id : 0 };
  }
  if (message.kind === 'pull') {
    const value = message.sequence === 0
      ? { kind: 'rows', rows: WORKSHEET.rows }
      : { kind: 'finished', worksheet: { ...WORKSHEET, rows: [] } };
    const payload = new TextEncoder().encode(JSON.stringify(value)).buffer;
    return {
      ...message,
      kind: 'chunk',
      byteLength: payload.byteLength,
      done: message.sequence === 1,
      payload,
    };
  }
  return { ...message, kind: 'accepted', command: message.kind };
}

afterEach(() => {
  vi.unstubAllGlobals();
});

describe('XlsxWorkbook.getWorksheet compatibility materializer', () => {
  it('decodes main-mode bytes once, resolves shared strings, and preserves cache identity', async () => {
    const { workbook, request } = makeWorkbook('main', async (message) => streamResponse(message));

    const first = await workbook.getWorksheet(0);
    expect(first.rows[0].cells[0].value).toEqual({ type: 'text', text: 'resolved' });
    first.rows.push({ index: 2, height: null, cells: [] });
    const second = await workbook.getWorksheet(0);

    expect(second).toBe(first);
    expect(second.rows).toHaveLength(2);
    expect(request).toHaveBeenCalledTimes(5);
  });

  it('evicts an inactive sheet before terminal ACK and re-pulls an identical model', async () => {
    const messages: Array<WorkerRequest | RenderWorkerRequest | PullSessionCommand<number>> = [];
    const sheetBySession = new Map<number, number>();
    const { workbook, request } = makeWorkbook('main', async (message) => {
      messages.push(message);
      if ('type' in message) {
        if (message.type === 'openSheetSession') sheetBySession.set(message.sessionId, message.sheetIndex);
        return { type: 'sheetSessionOpened', id: 'id' in message ? message.id : 0 };
      }
      if (message.kind === 'pull') {
        const value = message.sequence === 0
          ? { kind: 'rows', rows: WORKSHEET.rows }
          : { kind: 'finished', worksheet: {
            ...WORKSHEET, name: `Sheet${(sheetBySession.get(message.sessionId) ?? 0) + 1}`, rows: [],
          } };
        const payload = new TextEncoder().encode(JSON.stringify(value)).buffer;
        return { ...message, kind: 'chunk', byteLength: payload.byteLength, done: message.sequence === 1, payload };
      }
      return { ...message, kind: 'accepted', command: message.kind };
    });
    const state = workbook as unknown as {
      parsedWorkbook: ParsedWorkbook;
      retainedSheetUsage: {
        rows: number;
        cells: number;
        ownedUtf8Bytes: number;
        jsonBytes: number;
      };
    };
    state.parsedWorkbook.workbook.sheets.push({ name: 'Sheet2' } as ParsedWorkbook['workbook']['sheets'][number]);
    state.retainedSheetUsage = {
      rows: 199_999,
      cells: 499_999,
      ownedUtf8Bytes: 0,
      jsonBytes: 0,
    };

    const first = await workbook.getWorksheet(0);
    const afterFirst = request.mock.calls.length;
    expect(await workbook.getWorksheet(0)).toBe(first);
    expect(request).toHaveBeenCalledTimes(afterFirst);
    expect(state.retainedSheetUsage).toMatchObject({ rows: 200_000, cells: 500_000 });
    const second = await workbook.getWorksheet(1);
    expect(second.name).toBe('Sheet2');
    expect(state.retainedSheetUsage).toMatchObject({ rows: 200_000, cells: 500_000 });
    expect((workbook as unknown as { sheetCache: Map<number, Worksheet> }).sheetCache.has(0)).toBe(false);
    const secondSession = messages.filter(
      (message): message is PullSessionCommand<number> =>
        'kind' in message && message.sessionId === 2,
    );
    expect(secondSession.some((message) => message.kind === 'ack' && message.sequence === 1)).toBe(true);
    const reacquired = await workbook.getWorksheet(0);
    expect(reacquired).toEqual(first);
    expect(reacquired).not.toBe(first);
  });

  it('keeps a leased sheet and retries admission after its lease ends', async () => {
    const { workbook, request } = makeWorkbook('main', async (message) => {
      if ('type' in message) return { type: 'sheetSessionOpened', id: 'id' in message ? message.id : 0 };
      if (message.kind === 'pull') {
        const payload = new TextEncoder().encode(JSON.stringify(message.sequence === 0
          ? { kind: 'rows', rows: WORKSHEET.rows }
          : { kind: 'finished', worksheet: { ...WORKSHEET, rows: [] } })).buffer;
        return { ...message, kind: 'chunk', byteLength: payload.byteLength,
          done: message.sequence === 1, payload };
      }
      return { ...message, kind: 'accepted', command: message.kind };
    });
    const state = workbook as unknown as {
      parsedWorkbook: ParsedWorkbook;
      retainedSheetUsage: { rows: number; cells: number; ownedUtf8Bytes: number; jsonBytes: number };
    };
    state.parsedWorkbook.workbook.sheets.push({ name: 'Sheet2' } as ParsedWorkbook['workbook']['sheets'][number]);
    state.retainedSheetUsage = { rows: 199_999, cells: 499_999, ownedUtf8Bytes: 0, jsonBytes: 0 };
    const lease = await acquireXlsxWorksheet(workbook as unknown as XlsxWorkbook, 0);
    await expect(workbook.getWorksheet(1)).rejects.toMatchObject({
      code: 'ooxml-resource-limit',
      details: { violation: { resource: 'worksheet-cache' } },
    });
    expect(request.mock.calls.some(([build]) => {
      const message = build(0);
      return 'kind' in message && message.kind === 'ack' && message.sessionId === 2 && message.sequence === 1;
    })).toBe(false);
    lease.release();
    await expect(workbook.getWorksheet(0)).resolves.toBe(lease.worksheet);
    await expect(workbook.getWorksheet(1)).resolves.toBeDefined();
  });

  it('keeps eviction accounting exact when the incoming terminal ACK fails', async () => {
    const { workbook } = makeWorkbook('main', async (message) => {
      if ('type' in message) return { type: 'sheetSessionOpened', id: 'id' in message ? message.id : 0 };
      if (message.kind === 'pull') {
        const payload = new TextEncoder().encode(JSON.stringify(message.sequence === 0
          ? { kind: 'rows', rows: WORKSHEET.rows }
          : { kind: 'finished', worksheet: { ...WORKSHEET, rows: [] } })).buffer;
        return { ...message, kind: 'chunk', byteLength: payload.byteLength,
          done: message.sequence === 1, payload };
      }
      if (message.kind === 'ack' && message.sessionId === 2 && message.sequence === 1) {
        throw new Error('terminal ACK failed');
      }
      return { ...message, kind: 'accepted', command: message.kind };
    });
    const state = workbook as unknown as {
      parsedWorkbook: ParsedWorkbook;
      retainedSheetUsage: { rows: number; cells: number; ownedUtf8Bytes: number; jsonBytes: number };
      sheetCache: Map<number, Worksheet>;
    };
    state.parsedWorkbook.workbook.sheets.push({ name: 'Sheet2' } as ParsedWorkbook['workbook']['sheets'][number]);
    state.retainedSheetUsage = { rows: 199_999, cells: 499_999, ownedUtf8Bytes: 0, jsonBytes: 0 };
    await workbook.getWorksheet(0);
    await expect(workbook.getWorksheet(1)).rejects.toThrow('terminal ACK failed');
    expect(state.sheetCache.size).toBe(0);
    expect(state.retainedSheetUsage).toMatchObject({ rows: 199_999, cells: 499_999 });
    await expect(workbook.getWorksheet(1)).resolves.toBeDefined();
  });

  it('keeps one LRU order and evicts the worker copy before admitting a replacement', async () => {
    const workerSheets = new Set<number>();
    const opened = new Map<number, number>();
    const messages: string[] = [];
    const { workbook } = makeWorkbook('worker', async (message) => {
      if ('type' in message) {
        if (message.type === 'evictWorksheets') {
          for (const index of message.sheetIndices) workerSheets.delete(index);
          messages.push(`evict:${message.sheetIndices.join(',')}`);
          return { type: 'worksheetsEvicted', id: message.id };
        }
        if (message.type === 'openSheetSession') {
          opened.set(message.sessionId, message.sheetIndex);
          return { type: 'sheetSessionOpened', id: message.id };
        }
        throw new Error('unexpected worker request');
      }
      if (message.kind === 'pull') {
        const sheet = opened.get(message.sessionId) ?? 0;
        const payload = new TextEncoder().encode(JSON.stringify(message.sequence === 0
          ? { kind: 'rows', rows: [
            ...WORKSHEET.rows,
            { index: 2, height: null, cells: [] },
          ] }
          : { kind: 'finished', worksheet: { ...WORKSHEET, name: `Sheet${sheet + 1}`, rows: [] } })).buffer;
        return { ...message, kind: 'chunk', byteLength: payload.byteLength,
          done: message.sequence === 1, payload };
      }
      if (message.kind === 'ack' && message.sequence === 1) {
        const sheet = opened.get(message.sessionId) ?? 0;
        workerSheets.add(sheet);
        messages.push(`ack:${sheet}`);
      }
      return { ...message, kind: 'accepted', command: message.kind };
    });
    const state = workbook as unknown as {
      parsedWorkbook: ParsedWorkbook;
      retainedSheetUsage: { rows: number; cells: number; ownedUtf8Bytes: number; jsonBytes: number };
      sheetCache: Map<number, Worksheet>;
    };
    state.parsedWorkbook.workbook.sheets.push(
      { name: 'Sheet2' } as ParsedWorkbook['workbook']['sheets'][number],
      { name: 'Sheet3' } as ParsedWorkbook['workbook']['sheets'][number],
    );
    state.retainedSheetUsage = { rows: 199_996, cells: 499_998, ownedUtf8Bytes: 0, jsonBytes: 0 };
    await workbook.getWorksheet(0);
    await workbook.getWorksheet(1);
    await workbook.getWorksheet(0); // touch first sheet; second is now LRU
    await workbook.getWorksheet(2);
    expect(messages.slice(-2)).toEqual(['evict:1', 'ack:2']);
    expect([...state.sheetCache.keys()]).toEqual([0, 2]);
    expect(state.retainedSheetUsage.rows).toBe(200_000);
    expect([...workerSheets].sort()).toEqual([0, 2]);
  });

  it('waits for a pending worker eviction before leasing or rendering its victim', async () => {
    let completeEviction!: () => void;
    let evictionStarted!: () => void;
    const blockedEviction = new Promise<void>((resolve) => { completeEviction = resolve; });
    const evictionPending = new Promise<void>((resolve) => { evictionStarted = resolve; });
    const opened = new Map<number, number>();
    const events: string[] = [];
    const bitmap = {} as ImageBitmap;
    const { workbook } = makeWorkbook('worker', async (message) => {
      if ('type' in message) {
        if (message.type === 'evictWorksheets') {
          events.push('evict-start');
          evictionStarted();
          await blockedEviction;
          events.push('evict-done');
          return { type: 'worksheetsEvicted', id: message.id };
        }
        if (message.type === 'openSheetSession') {
          opened.set(message.sessionId, message.sheetIndex);
          events.push(`open:${message.sheetIndex}`);
          return { type: 'sheetSessionOpened', id: message.id };
        }
        if (message.type === 'renderViewport') {
          events.push(`render:${message.sheetIndex}`);
          return { type: 'viewportRendered', id: message.id, bitmap };
        }
        throw new Error('unexpected worker request');
      }
      if (message.kind === 'pull') {
        const sheet = opened.get(message.sessionId) ?? 0;
        const payload = new TextEncoder().encode(JSON.stringify(message.sequence === 0
          ? { kind: 'rows', rows: WORKSHEET.rows }
          : { kind: 'finished', worksheet: { ...WORKSHEET, name: `Sheet${sheet + 1}`, rows: [] } })).buffer;
        return { ...message, kind: 'chunk', byteLength: payload.byteLength,
          done: message.sequence === 1, payload };
      }
      return { ...message, kind: 'accepted', command: message.kind };
    });
    const state = workbook as unknown as {
      parsedWorkbook: ParsedWorkbook;
      retainedSheetUsage: { rows: number; cells: number; ownedUtf8Bytes: number; jsonBytes: number };
    };
    state.parsedWorkbook.workbook.sheets.push({ name: 'Sheet2' } as ParsedWorkbook['workbook']['sheets'][number]);
    state.retainedSheetUsage = { rows: 199_999, cells: 499_999, ownedUtf8Bytes: 0, jsonBytes: 0 };
    const first = await workbook.getWorksheet(0);
    const incoming = workbook.getWorksheet(1);
    await evictionPending;

    let acquired = false;
    const reacquiring = acquireXlsxWorksheet(workbook as unknown as XlsxWorkbook, 0).then((lease) => {
      acquired = true;
      return lease;
    });
    const rendering = workbook.renderViewportToBitmap(
      0, { startRow: 1, endRow: 1, startCol: 1, endCol: 1 }, { width: 100, height: 80 },
    );
    await Promise.resolve();
    expect(acquired).toBe(false);
    expect(events).not.toContain('render:0');

    completeEviction();
    await incoming;
    const lease = await reacquiring;
    expect(lease.worksheet).toEqual(first);
    expect(lease.worksheet).not.toBe(first);
    await expect(rendering).resolves.toBe(bitmap);
    expect(events.indexOf('evict-done')).toBeLessThan(events.lastIndexOf('open:0'));
    expect(events.indexOf('evict-done')).toBeLessThan(events.indexOf('render:0'));
    lease.release();
  });

  it('closes both caches when a worker eviction reply is lost', async () => {
    const { workbook, bridge } = makeWorkbook('worker', async (message) => {
      if ('type' in message) {
        if (message.type === 'evictWorksheets') throw new Error('eviction reply lost');
        return { type: 'sheetSessionOpened', id: 'id' in message ? message.id : 0 };
      }
      if (message.kind === 'pull') {
        const payload = new TextEncoder().encode(JSON.stringify(message.sequence === 0
          ? { kind: 'rows', rows: WORKSHEET.rows }
          : { kind: 'finished', worksheet: { ...WORKSHEET, rows: [] } })).buffer;
        return { ...message, kind: 'chunk', byteLength: payload.byteLength,
          done: message.sequence === 1, payload };
      }
      return { ...message, kind: 'accepted', command: message.kind };
    });
    const state = workbook as unknown as {
      parsedWorkbook: ParsedWorkbook | null;
      retainedSheetUsage: { rows: number; cells: number; ownedUtf8Bytes: number; jsonBytes: number };
      sheetCache: Map<number, Worksheet>;
    };
    if (!state.parsedWorkbook) throw new Error('missing test workbook');
    state.parsedWorkbook.workbook.sheets.push({ name: 'Sheet2' } as ParsedWorkbook['workbook']['sheets'][number]);
    state.retainedSheetUsage = { rows: 199_999, cells: 499_999, ownedUtf8Bytes: 0, jsonBytes: 0 };
    await workbook.getWorksheet(0);

    await expect(workbook.getWorksheet(1)).rejects.toThrow('eviction reply lost');
    expect(bridge.terminate).toHaveBeenCalledOnce();
    expect(state.parsedWorkbook).toBeNull();
    expect(state.sheetCache.size).toBe(0);
  });

  it('closes both caches when a worker terminal ACK reply is lost', async () => {
    const { workbook, bridge } = makeWorkbook('worker', async (message) => {
      if ('type' in message) return { type: 'sheetSessionOpened', id: 'id' in message ? message.id : 0 };
      if (message.kind === 'pull') {
        const payload = new TextEncoder().encode(JSON.stringify(message.sequence === 0
          ? { kind: 'rows', rows: WORKSHEET.rows }
          : { kind: 'finished', worksheet: { ...WORKSHEET, rows: [] } })).buffer;
        return { ...message, kind: 'chunk', byteLength: payload.byteLength,
          done: message.sequence === 1, payload };
      }
      if (message.kind === 'ack' && message.sequence === 1) {
        throw new Error('terminal ACK reply lost');
      }
      return { ...message, kind: 'accepted', command: message.kind };
    });
    const state = workbook as unknown as { sheetCache: Map<number, Worksheet> };

    await expect(workbook.getWorksheet(0)).rejects.toThrow('terminal ACK reply lost');
    expect(bridge.terminate).toHaveBeenCalledOnce();
    expect(state.sheetCache.size).toBe(0);
  });

  it('rejects a single oversized worksheet before cache admission', async () => {
    const oversizedRows = Array.from({ length: 100_001 }, (_, index) => ({
      index: index + 1, height: null, cells: [],
    }));
    const { workbook } = makeWorkbook('main', async (message) => {
      if ('type' in message) return { type: 'sheetSessionOpened', id: 'id' in message ? message.id : 0 };
      if (message.kind === 'pull') {
        const payload = new TextEncoder().encode(JSON.stringify({
          kind: 'rows', rows: oversizedRows,
        })).buffer;
        return { ...message, kind: 'chunk', byteLength: payload.byteLength, done: false, payload };
      }
      return { ...message, kind: 'accepted', command: message.kind };
    });
    await expect(workbook.getWorksheet(0)).rejects.toMatchObject({
      code: 'ooxml-resource-limit',
      details: { violation: { resource: 'worksheet-model', metric: 'rows' } },
    });
    expect((workbook as unknown as { sheetCache: Map<number, Worksheet> }).sheetCache.size).toBe(0);
  });

  it('deduplicates concurrent worker-mode materialization into one request and object', async () => {
    let releaseOpen: (() => void) | undefined;
    const openBarrier = new Promise<void>((resolve) => {
      releaseOpen = resolve;
    });
    const { workbook, request } = makeWorkbook('worker', async (message) => {
      if ('type' in message) await openBarrier;
      return streamResponse(message);
    });

    const first = workbook.getWorksheet(0);
    const second = workbook.getWorksheet(0);
    await Promise.resolve();
    expect(request).toHaveBeenCalledTimes(1);
    releaseOpen?.();

    const [left, right] = await Promise.all([first, second]);
    expect(left).toBe(right);
    expect(left.rows[0].cells[0].value).toEqual({ type: 'text', text: 'resolved' });
  });

  it('rejects an out-of-range sheet before opening a worker operation', async () => {
    const { workbook, request } = makeWorkbook('main', async () => {
      throw new Error('must not be called');
    });

    await expect(workbook.getWorksheet(1)).rejects.toThrow('Sheet index 1 out of range');
    expect(request).not.toHaveBeenCalled();
  });

  it('does not retain a rejected in-flight materialization', async () => {
    let attempt = 0;
    const { workbook, request } = makeWorkbook('main', async (message) => {
      attempt += 1;
      if (attempt === 1) throw new Error('injected sheet failure');
      return streamResponse(message);
    });

    await expect(workbook.getWorksheet(0)).rejects.toThrow('injected sheet failure');
    await expect(workbook.getWorksheet(0)).resolves.toMatchObject({ name: 'Sheet1' });
    // Failed open + correlated cancel, followed by a full successful stream.
    expect(request).toHaveBeenCalledTimes(7);
  });

  it('cancels a failed open so a later ordinary archive operation proceeds', async () => {
    const messages: Array<WorkerRequest | RenderWorkerRequest | PullSessionCommand<number>> = [];
    const { workbook } = makeWorkbook('main', async (message) => {
      messages.push(message);
      if ('type' in message && message.type === 'openSheetSession') {
        throw new Error('open response timed out');
      }
      if ('type' in message && message.type === 'toMarkdown') {
        return { type: 'markdownRendered', id: message.id, markdown: 'ok' };
      }
      return streamResponse(message);
    });

    await expect(workbook.getWorksheet(0)).rejects.toThrow('open response timed out');
    await expect(workbook.toMarkdown()).resolves.toBe('ok');
    expect(messages.map((message) => 'type' in message ? message.type : message.kind)).toEqual([
      'openSheetSession',
      'cancel',
      'toMarkdown',
    ]);
  });

  it('applies the configured timeout to open, every pull, and every acknowledgement', async () => {
    const { workbook, request } = makeWorkbook(
      'main',
      async (message) => streamResponse(message),
      250,
    );

    await expect(workbook.getWorksheet(0)).resolves.toMatchObject({ name: 'Sheet1' });
    expect(request.mock.calls.map((call) => call[2]?.timeoutMs)).toEqual([
      250, // open
      250, // rows pull
      250, // rows ACK
      250, // terminal pull
      250, // terminal ACK
    ]);
  });

  it('cancels with the non-expiring lifecycle path after a timed-out worksheet pull', async () => {
    const messages: Array<WorkerRequest | RenderWorkerRequest | PullSessionCommand<number>> = [];
    const { workbook, request } = makeWorkbook('main', async (message) => {
      messages.push(message);
      if (!('type' in message) && message.kind === 'pull') {
        throw new Error('worker request timed out after 25ms');
      }
      return streamResponse(message);
    }, 25);

    await expect(workbook.getWorksheet(0)).rejects.toThrow('timed out after 25ms');
    expect(messages.map((message) => 'type' in message ? message.type : message.kind)).toEqual([
      'openSheetSession',
      'pull',
      'cancel',
    ]);
    expect(request.mock.calls.map((call) => call[2]?.timeoutMs)).toEqual([25, 25, false]);
  });

  it('cancels and poisons every later sibling operation before retaining an over-limit row chunk', async () => {
    const oversizedRows = Array.from({ length: 100_001 }, (_, index) => ({
      index: index + 1,
      height: null,
      cells: [],
    }));
    const messages: Array<WorkerRequest | RenderWorkerRequest | PullSessionCommand<number>> = [];
    const { workbook } = makeWorkbook('main', async (message) => {
      messages.push(message);
      if ('type' in message) return streamResponse(message);
      if (message.kind === 'pull') {
        const payload = new TextEncoder().encode(JSON.stringify({ kind: 'rows', rows: oversizedRows })).buffer;
        return { ...message, kind: 'chunk', byteLength: payload.byteLength, done: false, payload };
      }
      return { ...message, kind: 'accepted', command: message.kind };
    });

    let first: unknown;
    try {
      await workbook.getWorksheet(0);
    } catch (error) {
      first = error;
    }
    expect(first).toBeInstanceOf(OoxmlResourceLimitError);
    expect((first as OoxmlResourceLimitError).details.violation).toMatchObject({
      resource: 'worksheet-model', metric: 'rows', observed: 100_001,
    });
    const beforeSibling = messages.length;
    await expect(workbook.toMarkdown()).rejects.toBe(first);
    expect(messages).toHaveLength(beforeSibling);
    expect(messages.some((message) => !('type' in message) && message.kind === 'cancel')).toBe(true);
  });

  it('latches a renderer resource failure across later image and markdown operations', async () => {
    const fatal = new OoxmlResourceLimitError('renderer index limit', {
      stage: 'layout',
      violation: {
        format: 'xlsx', operation: 'render-viewport', resource: 'renderer-index', metric: 'entries',
        limit: 250_000, observed: 250_001, configurable: false,
        usage: { archiveEntryCount: 1, declaredInflatedBytes: 2, distinctInflatedBytes: 3, operationInflatedBytes: 4 },
      },
    });
    const { workbook, request } = makeWorkbook('worker', async (message) => {
      if ('type' in message && message.type === 'renderViewport') throw fatal;
      return streamResponse(message);
    });
    (workbook as unknown as { sheetCache: Map<number, Worksheet> }).sheetCache.set(0, WORKSHEET);

    await expect(workbook.renderViewportToBitmap(
      0,
      { startRow: 0, endRow: 1, startCol: 0, endCol: 1 },
      { width: 100, height: 100 },
    )).rejects.toBe(fatal);
    const requestCount = request.mock.calls.length;
    await expect((workbook as unknown as XlsxWorkbook).getImage('xl/media/a.png', 'image/png'))
      .rejects.toBe(fatal);
    await expect(workbook.toMarkdown()).rejects.toBe(fatal);
    expect(request).toHaveBeenCalledTimes(requestCount);
  });

  it('keeps an uncached worker render atomic ahead of later markdown', async () => {
    const order: string[] = [];
    const bitmap = {} as ImageBitmap;
    const { workbook } = makeWorkbook('worker', async (message) => {
      order.push('type' in message ? message.type : `${message.kind}:${'sequence' in message ? message.sequence : ''}`);
      if ('type' in message && message.type === 'renderViewport') {
        return { type: 'viewportRendered', id: message.id, bitmap };
      }
      if ('type' in message && message.type === 'toMarkdown') {
        return { type: 'markdownRendered', id: message.id, markdown: 'after' };
      }
      return streamResponse(message);
    });

    const rendered = workbook.renderViewportToBitmap(
      0,
      { startRow: 1, endRow: 1, startCol: 1, endCol: 1 },
      { width: 100, height: 80 },
    );
    const markdown = workbook.toMarkdown();
    await expect(rendered).resolves.toBe(bitmap);
    await expect(markdown).resolves.toBe('after');
    expect(order.indexOf('renderViewport')).toBeLessThan(order.indexOf('toMarkdown'));
  });

  it('settles open cleanup when a worker error makes the bridge permanently unusable', async () => {
    const { workbook, request } = makeWorkbook('main', async () => {
      throw new Error('Worker error: worker is unusable');
    });

    await expect(workbook.getWorksheet(0)).rejects.toThrow('Worker error: worker is unusable');
    // Open plus the shared session's immediate cancel attempt; neither hangs.
    expect(request).toHaveBeenCalledTimes(2);
  });
});
