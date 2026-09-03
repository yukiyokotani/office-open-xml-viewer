import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  XlsxWorkbook,
  loadXlsxSheetSource,
} from './workbook.js';
import { parseDelimitedWorksheet } from './delimited-text.js';
import type { DelimitedTextParseRequest } from './delimited-text-protocol.js';

class DelimitedTextWorkerProbe {
  static instances: DelimitedTextWorkerProbe[] = [];
  readonly messages: unknown[] = [];
  readonly transfers: Transferable[][] = [];
  terminated = false;
  private readonly messageListeners = new Set<(event: MessageEvent) => void>();

  constructor() {
    DelimitedTextWorkerProbe.instances.push(this);
  }

  postMessage(message: unknown, transfer: Transferable[] = []): void {
    this.messages.push(message);
    this.transfers.push(transfer);
    const request = message as DelimitedTextParseRequest;
    if (request.type !== 'parseDelimitedText') return;
    const { workbook, worksheet } = parseDelimitedWorksheet(request.data, request.options);
    const worksheetJson = new TextEncoder().encode(JSON.stringify(worksheet)).buffer as ArrayBuffer;
    queueMicrotask(() => {
      for (const listener of this.messageListeners) {
        listener({
          data: {
            type: 'delimitedTextParsed',
            id: request.id,
            workbook,
            worksheetJson,
          },
        } as MessageEvent);
      }
    });
  }

  addEventListener(type: 'message', listener: (event: MessageEvent) => void): void;
  addEventListener(type: 'messageerror', listener: (event: MessageEvent) => void): void;
  addEventListener(type: 'error', listener: (event: ErrorEvent) => void): void;
  addEventListener(
    type: 'message' | 'messageerror' | 'error',
    listener: ((event: MessageEvent) => void) | ((event: ErrorEvent) => void),
  ): void {
    if (type === 'message') {
      this.messageListeners.add(listener as (event: MessageEvent) => void);
    }
  }

  removeEventListener(type: 'message', listener: (event: MessageEvent) => void): void;
  removeEventListener(type: 'messageerror', listener: (event: MessageEvent) => void): void;
  removeEventListener(type: 'error', listener: (event: ErrorEvent) => void): void;
  removeEventListener(
    type: 'message' | 'messageerror' | 'error',
    listener: ((event: MessageEvent) => void) | ((event: ErrorEvent) => void),
  ): void {
    if (type === 'message') {
      this.messageListeners.delete(listener as (event: MessageEvent) => void);
    }
  }

  terminate(): void {
    this.terminated = true;
  }
}

vi.mock('./delimited-text-worker-host', () => ({
  createDelimitedTextWorker: () => new DelimitedTextWorkerProbe() as unknown as Worker,
}));

afterEach(() => {
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
  DelimitedTextWorkerProbe.instances = [];
});

describe('XlsxSheetViewer delimited-text workbook bridge', () => {
  it('loads CSV off the main thread without initializing XLSX WASM', async () => {
    const source = new TextEncoder().encode('id,name\n00123,Ada').buffer as ArrayBuffer;
    const workbook = await XlsxWorkbook[loadXlsxSheetSource](source, {}, {
      format: 'csv',
      sheetName: 'Import',
    });

    const worker = DelimitedTextWorkerProbe.instances[0]!;
    expect(worker.messages).toHaveLength(1);
    expect(worker.messages[0]).toMatchObject({
      type: 'parseDelimitedText',
      options: { delimiter: ',', encoding: 'utf-8', sheetName: 'Import' },
    });
    expect(worker.messages[0]).not.toMatchObject({ type: 'init' });
    expect(worker.transfers[0]).toHaveLength(1);
    expect((worker.messages[0] as DelimitedTextParseRequest).data).not.toBe(source);
    expect(source.byteLength).toBe(17);
    expect(worker.terminated).toBe(true);

    expect(workbook.sheetNames).toEqual(['Import']);
    const worksheet = await workbook.getWorksheet(0);
    expect(worksheet.rows[1]?.cells.map((cell) => cell.value)).toEqual([
      { type: 'text', text: '00123' },
      { type: 'text', text: 'Ada' },
    ]);
    expect(workbook.mode).toBe('main');
    await expect(workbook.getResourceMetrics()).resolves.toMatchObject({
      format: 'xlsx',
      mode: 'main',
      status: 'ok',
      sourceBytes: source.byteLength,
    });

    workbook.destroy();
  });

  it('fetches a generic delimited URL with the same string-source contract as XLSX', async () => {
    const bytes = new TextEncoder().encode('left|right').buffer as ArrayBuffer;
    const fetch = vi.fn().mockResolvedValue({
      ok: true,
      arrayBuffer: () => Promise.resolve(bytes),
    });
    vi.stubGlobal('fetch', fetch);

    const workbook = await XlsxWorkbook[loadXlsxSheetSource]('/data/report.dat', {}, {
      format: 'delimited-text',
      delimiter: '|',
    });

    expect(fetch).toHaveBeenCalledWith('/data/report.dat');
    expect((await workbook.getWorksheet(0)).rows[0]?.cells).toHaveLength(2);
    workbook.destroy();
  });

  it('keeps the public XlsxWorkbook loader XLSX-only', async () => {
    const load = vi.spyOn(XlsxWorkbook, 'load').mockResolvedValue({} as XlsxWorkbook);
    const source = new ArrayBuffer(0);

    await XlsxWorkbook[loadXlsxSheetSource](source, {});

    expect(load).toHaveBeenCalledWith(source, {});
    expect(DelimitedTextWorkerProbe.instances).toHaveLength(0);
  });
});
