import { beforeEach, describe, expect, it, vi } from 'vitest';
import type { ModelSourceModuleDescriptor } from '@silurus/ooxml-core';
import { XLSX_HOST_LAYOUT_REQUEST, XLSX_HOST_LAYOUT_RESULT } from './internal/host-layout.js';

const state = vi.hoisted(() => ({
  ensureReady: vi.fn(async () => undefined),
  setWasmInput: vi.fn(),
  openSource: vi.fn(),
  ooxmlConstruct: vi.fn(),
  computeMdw: vi.fn(() => 9),
}));

vi.mock('@silurus/ooxml-core', async (load) => {
  const actual = await load<typeof import('@silurus/ooxml-core')>();
  class Host<T extends { free(): void }> {
    archive: T | null = null;
    setWasmInput = state.setWasmInput;
    ensureReady = state.ensureReady;
    run<R>(operation: () => R): R { return operation(); }
    setArchive(value: T): void { this.archive?.free(); this.archive = value; }
    disposeArchive(): void { this.archive?.free(); this.archive = null; }
  }
  return {
    ...actual,
    WasmParserHost: Host,
    openModelSourceModule: (...args: unknown[]) => state.openSource(...args),
  };
});

class OoxmlArchive {
  constructor() { state.ooxmlConstruct(); }
  free = vi.fn();
  assert_healthy() {}
  parse() { return new TextEncoder().encode(JSON.stringify(bootstrap)); }
  resource_usage() { return new Uint8Array(48); }
  open_sheet_cursor() {}
  pull_sheet_cursor() { return new Uint8Array(); }
  sheet_cursor_pull_finished() { return false; }
  sheet_cursor_resource_usage() { throw new Error('worksheet cursor usage is unavailable'); }
  acknowledge_sheet_cursor_terminal() {}
  cancel_sheet_cursor() {}
  close_sheet_cursor() {}
  extract_image() { return new Uint8Array([8]); }
  to_markdown() { return ''; }
}

vi.mock('./wasm/xlsx_parser.js', () => ({
  default: vi.fn(async () => undefined),
  reinit: vi.fn(async () => undefined),
  XlsxArchive: OoxmlArchive,
}));

// Parse assertions do not render; the render worker only needs computeMdw.
vi.mock('./renderer.js', () => ({ computeMdw: state.computeMdw }));
vi.mock('./render-orchestrator.js', () => ({}));
vi.mock('./delimited-text.js', () => ({}));

const bootstrap = { workbook: { sheets: [] }, styles: {}, sharedStrings: [] };
const CALIBRI_11 = { family: 'Calibri', sizePt: 11, bold: false, italic: false };
const descriptor: ModelSourceModuleDescriptor = {
  protocol: 'ooxml-model-source-module/v1',
  target: 'xlsx',
  moduleUrl: 'https://example.test/source.mjs',
  config: {},
};
const resourcePolicy = { maxArchiveEntryBytes: 1, maxTotalInflatedBytes: 1, maxArchiveEntries: 1 };

/** A source archive without resource_usage / to_markdown that asks for host layout. */
function sourceArchive(font: typeof CALIBRI_11 | undefined) {
  return {
    ...(font ? { host_layout_request: vi.fn(() => new TextEncoder().encode(JSON.stringify(font))) } : {}),
    configure_host_layout: vi.fn(),
    parse: vi.fn(() => new TextEncoder().encode(JSON.stringify(bootstrap))),
    assert_healthy: vi.fn(),
    open_sheet_cursor: vi.fn(),
    pull_sheet_cursor: vi.fn(() => new Uint8Array()),
    sheet_cursor_pull_finished: vi.fn(() => false),
    acknowledge_sheet_cursor_terminal: vi.fn(),
    cancel_sheet_cursor: vi.fn(),
    close_sheet_cursor: vi.fn(),
    extract_image: vi.fn(() => new Uint8Array([1, 2, 3])),
  };
}

/** A worker global whose page answers host layout requests with `width`. */
function workerScope(width: number) {
  const listeners = new Set<EventListener>();
  const posted: unknown[] = [];
  const scope = {
    onmessage: null as ((event: MessageEvent) => Promise<void>) | null,
    addEventListener: (_type: string, listener: EventListener) => listeners.add(listener),
    removeEventListener: (_type: string, listener: EventListener) => listeners.delete(listener),
    postMessage: (message: unknown) => {
      posted.push(message);
      const request = message as { type?: string; requestId?: number };
      if (request.type === XLSX_HOST_LAYOUT_REQUEST) {
        queueMicrotask(() => {
          const event = {
            data: { type: XLSX_HOST_LAYOUT_RESULT, requestId: request.requestId, maximumDigitWidth: width },
          } as MessageEvent;
          for (const listener of [...listeners]) listener(event);
        });
      }
    },
  };
  return { scope, posted, listeners };
}

const ofType = (posted: unknown[], type: string) =>
  posted.filter((message) => (message as { type?: string }).type === type) as Record<string, unknown>[];

describe('XLSX workers with a model source', () => {
  beforeEach(() => {
    vi.resetModules();
    vi.clearAllMocks();
    state.openSource.mockReset();
  });

  it('parse worker asks the page to measure, configures the reply and reports layoutMetrics', async () => {
    const { scope, posted, listeners } = workerScope(7);
    vi.stubGlobal('self', scope);
    const first = sourceArchive(CALIBRI_11);
    const second = sourceArchive(undefined);
    const firstClose = vi.fn();
    const secondClose = vi.fn();
    state.openSource
      .mockResolvedValueOnce({ archive: first, viewDefaults: {}, close: firstClose })
      .mockResolvedValueOnce({ archive: second, viewDefaults: {}, close: secondClose });
    await import('./worker.js');
    const dispatch = scope.onmessage!;

    await dispatch({ data: { type: 'init', wasmUrl: 'https://example.test/xlsx.wasm' } } as MessageEvent);
    await dispatch({ data: { type: 'parse', id: 1, data: new ArrayBuffer(3), resourcePolicy, source: descriptor } } as MessageEvent);
    expect(state.ensureReady).not.toHaveBeenCalled();
    expect(state.setWasmInput).not.toHaveBeenCalled();
    expect(ofType(posted, XLSX_HOST_LAYOUT_REQUEST)).toEqual([expect.objectContaining({ font: CALIBRI_11 })]);
    expect(first.configure_host_layout).toHaveBeenCalledExactlyOnceWith(7);
    expect(first.configure_host_layout.mock.invocationCallOrder[0])
      .toBeLessThan(first.parse.mock.invocationCallOrder[0]!);
    expect(ofType(posted, 'parsed')).toEqual([
      expect.objectContaining({ id: 1, layoutMetrics: { maximumDigitWidth: 7 }, usage: undefined }),
    ]);
    expect(listeners.size).toBe(0);

    await dispatch({ data: { type: 'openSheetSession', id: 2, sheetIndex: 0, sheetName: 'S', sessionId: 1, operationId: 1, generation: 1 } } as MessageEvent);
    expect(first.open_sheet_cursor).toHaveBeenCalledWith(0, 'S');
    await dispatch({ data: {
      protocol: 'ooxml-pull-v1', kind: 'close', requestId: 1,
      sessionId: 1, operationId: 1, generation: 1,
    } } as MessageEvent);
    await dispatch({ data: { type: 'extractImage', id: 3, path: 'xl/media/1' } } as MessageEvent);
    expect(first.extract_image).toHaveBeenCalledOnce();
    await dispatch({ data: { type: 'toMarkdown', id: 4 } } as MessageEvent);
    expect(posted).toContainEqual(expect.objectContaining({
      type: 'error', id: 4, message: 'Markdown conversion is unsupported for this source',
    }));

    // Reparse closes the old source; a source that asks nothing gets no layoutMetrics.
    await dispatch({ data: { type: 'parse', id: 5, data: new ArrayBuffer(3), resourcePolicy, source: descriptor } } as MessageEvent);
    expect(firstClose).toHaveBeenCalledOnce();
    expect(second.configure_host_layout).not.toHaveBeenCalled();
    expect(ofType(posted, 'parsed').at(-1)).not.toHaveProperty('layoutMetrics');

    await dispatch({ data: { type: 'parse', id: 6, data: new ArrayBuffer(3), resourcePolicy } } as MessageEvent);
    expect(secondClose).toHaveBeenCalledOnce();
    expect(state.ensureReady).toHaveBeenCalled();
    expect(state.ooxmlConstruct).toHaveBeenCalledOnce();
    expect(ofType(posted, XLSX_HOST_LAYOUT_REQUEST)).toHaveLength(1);
  });

  it('render worker measures locally with computeMdw and keeps the OOXML runtime idle', async () => {
    const { scope, posted } = workerScope(99);
    vi.stubGlobal('self', scope);
    const first = sourceArchive(CALIBRI_11);
    const second = sourceArchive(undefined);
    const firstClose = vi.fn();
    const secondClose = vi.fn();
    state.openSource
      .mockResolvedValueOnce({ archive: first, viewDefaults: {}, close: firstClose })
      .mockResolvedValueOnce({ archive: second, viewDefaults: {}, close: secondClose });
    await import('./render-worker.js');
    const dispatch = scope.onmessage!;

    await dispatch({ data: { type: 'init', wasmUrl: 'https://example.test/xlsx.wasm' } } as MessageEvent);
    await dispatch({ data: {
      type: 'parse', id: 11, data: new ArrayBuffer(3), resourcePolicy, source: descriptor,
    } } as MessageEvent);
    expect(state.ensureReady).not.toHaveBeenCalled();
    expect(state.setWasmInput).not.toHaveBeenCalled();
    expect(state.computeMdw).toHaveBeenCalledExactlyOnceWith('Calibri', 11, undefined, false, 400, 'normal');
    expect(first.configure_host_layout).toHaveBeenCalledExactlyOnceWith(9);
    expect(ofType(posted, XLSX_HOST_LAYOUT_REQUEST)).toEqual([]);
    expect(ofType(posted, 'parsed')).toEqual([expect.objectContaining({
      id: 11, workbook: expect.objectContaining({ layoutMetrics: { maximumDigitWidth: 9 } }),
    })]);

    await dispatch({ data: {
      type: 'parse', id: 12, data: new ArrayBuffer(3), resourcePolicy, source: descriptor,
    } } as MessageEvent);
    expect(firstClose).toHaveBeenCalledOnce();
    expect(second.parse).toHaveBeenCalledOnce();

    await dispatch({ data: { type: 'parse', id: 13, data: new ArrayBuffer(3), resourcePolicy } } as MessageEvent);
    expect(secondClose).toHaveBeenCalledOnce();
    expect(state.ensureReady).toHaveBeenCalled();
    expect(state.ooxmlConstruct).toHaveBeenCalledOnce();
  });
});
