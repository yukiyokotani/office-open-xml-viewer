import { describe, expect, it, vi } from 'vitest';
import { XlsxWorkbook } from './workbook.js';
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
  instance.sheetLoads = new Map();
  instance.bridge = bridge;
  instance.retainedSheetUsage = { rows: 0, cells: 0, ownedUtf8Bytes: 0, jsonBytes: 0 };
  instance.resourceFailure = null;
  instance.workerTimeoutMs = workerTimeoutMs;
  return { workbook: instance as unknown as WorkbookProbe, request };
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

  it('commits workbook cache totals after ACK and never double-charges a cached sheet', async () => {
    const messages: Array<WorkerRequest | RenderWorkerRequest | PullSessionCommand<number>> = [];
    const { workbook, request } = makeWorkbook('main', async (message) => {
      messages.push(message);
      if ('type' in message) {
        return { type: 'sheetSessionOpened', id: 'id' in message ? message.id : 0 };
      }
      if (message.kind === 'pull') {
        const value = message.sequence === 0
          ? { kind: 'rows', rows: WORKSHEET.rows }
          : { kind: 'finished', worksheet: { ...WORKSHEET, name: `Sheet${message.sessionId}`, rows: [] } };
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
    const committedUsage = { ...state.retainedSheetUsage };

    await expect(workbook.getWorksheet(1)).rejects.toBeInstanceOf(OoxmlResourceLimitError);
    expect(state.retainedSheetUsage).toEqual(committedUsage);
    const secondSession = messages.filter(
      (message): message is PullSessionCommand<number> =>
        'kind' in message && message.sessionId === 2,
    );
    expect(secondSession.some((message) => message.kind === 'ack' && message.sequence === 1)).toBe(false);
    expect(secondSession.some((message) => message.kind === 'cancel')).toBe(true);
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
