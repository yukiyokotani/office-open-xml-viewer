import { beforeEach, describe, expect, it, vi } from 'vitest';

const state = vi.hoisted(() => ({
  ensureReady: vi.fn(async () => undefined),
  setWasmInput: vi.fn(),
  directOpen: vi.fn(),
  directClose: vi.fn(),
  ooxmlConstruct: vi.fn(),
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
  return { ...actual, WasmParserHost: Host };
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

const bootstrap = { workbook: { sheets: [] }, styles: {}, sharedStrings: [] };

function directArchive() {
  return {
    free: vi.fn(),
    measurement_request: vi.fn(() => new TextEncoder().encode('{"required":false,"font":null}')),
    configure_mdw: vi.fn(),
    parse: vi.fn(() => new TextEncoder().encode(JSON.stringify(bootstrap))),
    assert_healthy: vi.fn(),
    resource_usage: vi.fn(() => { throw new Error('xlsx resource usage is unavailable'); }),
    open_sheet_cursor: vi.fn(),
    pull_sheet_cursor: vi.fn(() => new Uint8Array()),
    sheet_cursor_pull_finished: vi.fn(() => false),
    sheet_cursor_resource_usage: vi.fn(() => { throw new Error('worksheet cursor usage is unavailable'); }),
    acknowledge_sheet_cursor_terminal: vi.fn(),
    cancel_sheet_cursor: vi.fn(),
    close_sheet_cursor: vi.fn(),
    extract_image: vi.fn(() => new Uint8Array([1, 2, 3])),
    to_markdown: vi.fn(() => { throw new Error('unsupported'); }),
    close_workbook_session: vi.fn(),
  };
}

vi.mock('@silurus/ooxml-legacy-converter/internal/direct-xls-engine', async (load) => ({
  ...await load<typeof import('@silurus/ooxml-legacy-converter/internal/direct-xls-engine')>(),
  openLegacyXlsSource: state.directOpen,
}));

describe('XLSX parse worker direct source dispatch', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    state.directOpen.mockReset();
  });

  it('does not initialize OOXML for direct parse and closes on reparse', async () => {
    const posted = vi.fn();
    const scope = { postMessage: posted, onmessage: null as ((event: MessageEvent) => Promise<void>) | null };
    vi.stubGlobal('self', scope);
    const first = directArchive();
    const second = directArchive();
    state.directOpen
      .mockResolvedValueOnce({ archive: first, sourceByteLength: 3, closeArchive: state.directClose })
      .mockResolvedValueOnce({ archive: second, sourceByteLength: 3, closeArchive: state.directClose });
    await import('./worker.js');
    const dispatch = scope.onmessage!;
    const source = {
      protocol: 'ooxml-legacy-xls-source/v1', builtin: 'xls',
      wasmUrl: 'https://example.test/direct.wasm',
    } as const;
    const policy = { maxArchiveEntryBytes: 1, maxTotalInflatedBytes: 1, maxArchiveEntries: 1 };

    await dispatch({ data: { type: 'init', wasmUrl: 'https://example.test/xlsx.wasm' } } as MessageEvent);
    await dispatch({ data: { type: 'parse', id: 1, data: new ArrayBuffer(3), resourcePolicy: policy, source } } as MessageEvent);
    expect(state.ensureReady).not.toHaveBeenCalled();
    expect(state.setWasmInput).not.toHaveBeenCalled();
    expect(first.parse).toHaveBeenCalledOnce();

    await dispatch({ data: { type: 'openSheetSession', id: 2, sheetIndex: 0, sheetName: 'S', sessionId: 1, operationId: 1, generation: 1 } } as MessageEvent);
    expect(first.open_sheet_cursor).toHaveBeenCalledWith(0, 'S');
    await dispatch({ data: {
      protocol: 'ooxml-pull-v1', kind: 'close', requestId: 1,
      sessionId: 1, operationId: 1, generation: 1,
    } } as MessageEvent);
    await dispatch({ data: { type: 'extractImage', id: 3, path: 'legacy-xls/image/1' } } as MessageEvent);
    expect(first.extract_image).toHaveBeenCalledOnce();

    await dispatch({ data: { type: 'parse', id: 4, data: new ArrayBuffer(3), resourcePolicy: policy, source } } as MessageEvent);
    expect(state.directClose).toHaveBeenCalledOnce();
    expect(second.parse).toHaveBeenCalledOnce();

    await dispatch({ data: { type: 'parse', id: 5, data: new ArrayBuffer(3), resourcePolicy: policy } } as MessageEvent);
    expect(state.directClose).toHaveBeenCalledTimes(2);
    expect(state.ensureReady).toHaveBeenCalledOnce();
    expect(state.ooxmlConstruct).toHaveBeenCalledOnce();
  });

  it('keeps render-worker direct parse off the OOXML runtime and cleans up on replacement', async () => {
    const posted = vi.fn();
    const scope = { postMessage: posted, onmessage: null as ((event: MessageEvent) => Promise<void>) | null };
    vi.stubGlobal('self', scope);
    const first = directArchive();
    const second = directArchive();
    state.directOpen
      .mockResolvedValueOnce({ archive: first, sourceByteLength: 3, closeArchive: state.directClose })
      .mockResolvedValueOnce({ archive: second, sourceByteLength: 3, closeArchive: state.directClose });
    await import('./render-worker.js');
    const dispatch = scope.onmessage!;
    const source = {
      protocol: 'ooxml-legacy-xls-source/v1', builtin: 'xls',
      wasmUrl: 'https://example.test/direct.wasm',
    } as const;
    const resourcePolicy = {
      maxArchiveEntryBytes: 1, maxTotalInflatedBytes: 1, maxArchiveEntries: 1,
    };

    await dispatch({ data: { type: 'init', wasmUrl: 'https://example.test/xlsx.wasm' } } as MessageEvent);
    await dispatch({ data: {
      type: 'parse', id: 11, data: new ArrayBuffer(3), resourcePolicy, source,
    } } as MessageEvent);
    expect(state.ensureReady).not.toHaveBeenCalled();
    expect(state.setWasmInput).not.toHaveBeenCalled();
    expect(first.parse).toHaveBeenCalledOnce();

    await dispatch({ data: {
      type: 'parse', id: 12, data: new ArrayBuffer(3), resourcePolicy, source,
    } } as MessageEvent);
    expect(state.directClose).toHaveBeenCalledOnce();
    expect(second.parse).toHaveBeenCalledOnce();

    await dispatch({ data: {
      type: 'parse', id: 13, data: new ArrayBuffer(3), resourcePolicy,
    } } as MessageEvent);
    expect(state.directClose).toHaveBeenCalledTimes(2);
    expect(state.ensureReady).toHaveBeenCalledOnce();
    expect(state.ooxmlConstruct).toHaveBeenCalledOnce();
  });
});
