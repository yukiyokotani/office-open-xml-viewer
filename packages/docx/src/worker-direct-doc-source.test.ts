import { describe, expect, it, vi } from 'vitest';
import type { WorkerBridgeTransport } from '@silurus/ooxml-core';
import type { PullSessionResponse } from '@silurus/ooxml-core/worker';
import { materializeDocumentPullSession } from './document-pull-client.js';

const state = vi.hoisted(() => ({
  init: vi.fn(async () => undefined), ensureReady: vi.fn(async () => undefined),
  setWasmInput: vi.fn(), ooxmlConstruct: vi.fn(), directOpen: vi.fn(),
}));

vi.mock('@silurus/ooxml-core', async (load) => {
  const actual = await load<typeof import('@silurus/ooxml-core')>();
  class Host<T extends { free(): void }> {
    archive: T | null = null;
    setWasmInput(value: unknown): void { state.setWasmInput(value); }
    async ensureReady(): Promise<void> { state.ensureReady(); await state.init(); }
    run<R>(operation: () => R): R { return operation(); }
    setArchive(value: T): void { this.disposeArchive(); this.archive = value; }
    disposeArchive(): void { this.archive?.free(); this.archive = null; }
  }
  return { ...actual, WasmParserHost: Host };
});

vi.mock('./wasm/docx_parser.js', () => ({
  default: state.init,
  reinit: vi.fn(async () => undefined),
  DocxArchive: class extends CursorArchive {
    resource_usage = vi.fn(() => new Uint8Array(48));
    to_markdown = vi.fn(() => 'ooxml markdown');
    constructor() { super(); state.ooxmlConstruct(); }
  },
}));

vi.mock('@silurus/ooxml-legacy-converter/internal/direct-doc-engine', () => ({
  openLegacyDocSource: (...args: unknown[]) => state.directOpen(...args),
}));

describe('DOCX parse worker direct source dispatch', () => {
  it('streams native units, retains images, cleans replacement, fails closed, and preserves OOXML', async () => {
    const first = new CursorArchive();
    const second = new CursorArchive();
    const third = new CursorArchive();
    const broken = new CursorArchive();
    broken.open_document_cursor.mockImplementation(() => { throw new Error('cursor open failed'); });
    const firstClose = vi.fn();
    const secondClose = vi.fn();
    const thirdClose = vi.fn();
    const brokenClose = vi.fn();
    let resolveSecond!: (value: unknown) => void;
    const secondPending = new Promise((resolve) => { resolveSecond = resolve; });
    state.directOpen
      .mockResolvedValueOnce({ archive: first, sourceByteLength: 3, closeArchive: firstClose })
      .mockReturnValueOnce(secondPending)
      .mockResolvedValueOnce({ archive: third, sourceByteLength: 3, closeArchive: thirdClose })
      .mockResolvedValueOnce({ archive: broken, sourceByteLength: 3, closeArchive: brokenClose })
      .mockRejectedValueOnce(new Error('native failed'))
      .mockRejectedValueOnce(new TypeError('unsupported legacy DOC source builtin'));
    const harness = workerHarness();
    vi.stubGlobal('self', harness.scope);
    await import('./worker.js');
    const dispatch = harness.scope.onmessage!;
    const policy = { maxArchiveEntryBytes: 1, maxTotalInflatedBytes: 1, maxArchiveEntries: 1 };
    const source = { protocol: 'ooxml-legacy-doc-source/v1', builtin: 'doc',
      wasmUrl: 'https://example.test/doc.wasm' } as const;

    await dispatch({ data: { type: 'init', wasmUrl: 'https://example.test/docx.wasm' } } as MessageEvent);
    await dispatch({ data: { type: 'parse', id: 1, data: new ArrayBuffer(3), resourcePolicy: policy, source } } as MessageEvent);
    expect(state.ensureReady).not.toHaveBeenCalled();
    expect(state.setWasmInput).not.toHaveBeenCalled();
    expect(state.ooxmlConstruct).not.toHaveBeenCalled();
    const opened = harness.posts.find((message) => (message as { id?: number }).id === 1) as
      { sessionId: number; operationId: number; generation: number };
    const model = await materializeDocumentPullSession(harness.transport(dispatch), opened);
    expect(model.body.flatMap((element) => element.type === 'paragraph'
      ? element.runs.filter((run) => run.type === 'text').map((run) => run.text) : []))
      .toEqual(['first', 'second']);

    await dispatch({ data: { type: 'extractImage', id: 2, path: 'legacy-doc/image/1' } } as MessageEvent);
    expect(first.extract_image).toHaveBeenCalledWith('legacy-doc/image/1');
    expect(firstClose).not.toHaveBeenCalled();
    await dispatch({ data: { type: 'resourceUsage', id: 3 } } as MessageEvent);
    await dispatch({ data: { type: 'toMarkdown', id: 4 } } as MessageEvent);
    expect(harness.posts.filter((message) => (message as { id?: number }).id === 3)[0])
      .toMatchObject({ type: 'error' });
    expect(harness.posts.filter((message) => (message as { id?: number }).id === 4)[0])
      .toMatchObject({ type: 'error' });

    const superseded = dispatch({ data: { type: 'parse', id: 5, data: new ArrayBuffer(3), resourcePolicy: policy, source } } as MessageEvent);
    await vi.waitFor(() => expect(state.directOpen).toHaveBeenCalledTimes(2));
    expect(firstClose).toHaveBeenCalledTimes(1);
    await dispatch({ data: { type: 'parse', id: 6, data: new ArrayBuffer(3), resourcePolicy: policy, source } } as MessageEvent);
    resolveSecond({ archive: second, sourceByteLength: 3, closeArchive: secondClose });
    await superseded;
    expect(secondClose).toHaveBeenCalledTimes(1);
    expect(thirdClose).not.toHaveBeenCalled();
    expect(harness.posts.find((message) => (message as { id?: number }).id === 5))
      .toMatchObject({ type: 'error', message: expect.stringContaining('superseded') });
    expect(state.ooxmlConstruct).not.toHaveBeenCalled();
    expect(harness.posts.find((message) => (message as { id?: number }).id === 6))
      .toMatchObject({ type: 'documentSessionOpened' });

    await dispatch({ data: { type: 'parse', id: 7, data: new ArrayBuffer(3), resourcePolicy: policy, source } } as MessageEvent);
    expect(thirdClose).toHaveBeenCalledTimes(1);
    expect(harness.posts.find((message) => (message as { id?: number }).id === 7))
      .toMatchObject({ type: 'error' });
    expect(brokenClose).toHaveBeenCalledTimes(1);

    await dispatch({ data: { type: 'parse', id: 8, data: new ArrayBuffer(3), resourcePolicy: policy, source } } as MessageEvent);
    expect(harness.posts.find((message) => (message as { id?: number }).id === 8))
      .toMatchObject({ type: 'error' });

    await dispatch({ data: { type: 'parse', id: 9, data: new ArrayBuffer(3), resourcePolicy: policy,
      source: { ...source, builtin: 'xls' } } } as MessageEvent);
    expect(state.ooxmlConstruct).not.toHaveBeenCalled();
    expect(harness.posts.find((message) => (message as { id?: number }).id === 9))
      .toMatchObject({ type: 'error' });

    await dispatch({ data: { type: 'parse', id: 10, data: new ArrayBuffer(3), resourcePolicy: policy } } as MessageEvent);
    expect(state.setWasmInput).toHaveBeenCalledOnce();
    expect(state.ensureReady).toHaveBeenCalledOnce();
    expect(state.ooxmlConstruct).toHaveBeenCalledOnce();
    expect(harness.posts.find((message) => (message as { id?: number }).id === 10))
      .toMatchObject({ type: 'documentSessionOpened' });
  });
});

class CursorArchive {
  private index = 0;
  private delivered = false;
  private done = false;
  readonly free = vi.fn();
  readonly extract_image = vi.fn(() => new Uint8Array([9]));
  open_document_cursor = vi.fn();
  pull_document_chunk(sequence: number): Uint8Array {
    if (sequence !== this.index || this.delivered) throw new Error('bad sequence');
    const units = [body('first'), body('second'), { kind: 'complete', document: { body: [] } }];
    this.delivered = true; this.done = sequence === 2;
    return new TextEncoder().encode(JSON.stringify(units[sequence]));
  }
  document_chunk_done(): boolean { return this.done; }
  acknowledge_document_chunk(sequence: number): void {
    if (!this.delivered || sequence !== this.index) throw new Error('bad acknowledgement');
    this.delivered = false; this.index += 1;
  }
  cancel_document_cursor = vi.fn();
  close_document_session = vi.fn();
  assert_healthy = vi.fn();
}

function body(text: string) {
  return { kind: 'body', body: [{ type: 'paragraph', alignment: 'left', indentLeft: 0,
    indentRight: 0, indentFirst: 0, spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [], runs: [{ type: 'text', text, bold: false, italic: false,
      underline: false, strikethrough: false, fontSize: 10, color: null, fontFamily: null,
      isLink: false, background: null }] }] };
}

function workerHarness() {
  const posts: unknown[] = [];
  const waiters = new Map<number, (value: PullSessionResponse<ArrayBuffer, number>) => void>();
  const scope = {
    onmessage: null as ((event: MessageEvent) => Promise<void>) | null,
    postMessage(message: unknown) {
      posts.push(message);
      const response = message as PullSessionResponse<ArrayBuffer, number>;
      if ('requestId' in response) waiters.get(response.requestId)?.(response);
    },
  };
  let requestId = 1;
  return { scope, posts, transport(dispatch: (event: MessageEvent) => Promise<void>) {
    return {
      request(build) {
        const id = requestId++;
        return new Promise<PullSessionResponse<ArrayBuffer, number>>((resolve, reject) => {
          waiters.set(id, resolve);
          void dispatch({ data: build(id) } as MessageEvent).catch(reject);
        });
      },
      forgetOrphaned() {}, terminate() {},
    } satisfies WorkerBridgeTransport<PullSessionResponse<ArrayBuffer, number>>;
  } };
}
