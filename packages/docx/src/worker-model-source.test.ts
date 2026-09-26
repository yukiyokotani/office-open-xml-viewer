import { describe, expect, it, vi } from 'vitest';
import type { ModelSourceModuleDescriptor, WorkerBridgeTransport } from '@silurus/ooxml-core';
import type { PullSessionResponse } from '@silurus/ooxml-core/worker';
import { materializeDocumentPullSession } from './document-pull-client.js';

const state = vi.hoisted(() => ({
  init: vi.fn(async () => undefined), ensureReady: vi.fn(async () => undefined),
  setWasmInput: vi.fn(), ooxmlConstruct: vi.fn(), openSource: vi.fn(),
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
  return {
    ...actual,
    WasmParserHost: Host,
    openModelSourceModule: (...args: unknown[]) => state.openSource(...args),
  };
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

const source: ModelSourceModuleDescriptor = {
  protocol: 'ooxml-model-source-module/v1',
  target: 'docx',
  moduleUrl: 'https://example.test/source.mjs',
  config: {},
};

describe('DOCX parse worker with a model source', () => {
  it('streams source units without the OOXML runtime, supersedes cleanly, fails closed and keeps OOXML working', async () => {
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
    state.openSource
      .mockResolvedValueOnce({ archive: first, viewDefaults: { showTrackedChanges: true }, close: firstClose })
      .mockReturnValueOnce(secondPending)
      .mockResolvedValueOnce({ archive: third, viewDefaults: {}, close: thirdClose })
      .mockResolvedValueOnce({ archive: broken, viewDefaults: {}, close: brokenClose })
      .mockRejectedValueOnce(new Error('source open failed'));
    const harness = workerHarness();
    vi.stubGlobal('self', harness.scope);
    await import('./worker.js');
    const dispatch = harness.scope.onmessage!;
    const policy = { maxArchiveEntryBytes: 1, maxTotalInflatedBytes: 1, maxArchiveEntries: 1 };
    const parse = (id: number, extra: object = { source }) => dispatch({
      data: { type: 'parse', id, data: new ArrayBuffer(3), resourcePolicy: policy, ...extra },
    } as MessageEvent);
    const response = (id: number) => harness.posts.find((message) => (message as { id?: number }).id === id);

    await dispatch({ data: { type: 'init', wasmUrl: 'https://example.test/docx.wasm' } } as MessageEvent);
    await parse(1);
    expect(state.ensureReady).not.toHaveBeenCalled();
    expect(state.setWasmInput).not.toHaveBeenCalled();
    expect(state.ooxmlConstruct).not.toHaveBeenCalled();
    const opened = response(1) as { sessionId: number; operationId: number; generation: number };
    expect(opened).toMatchObject({ type: 'documentSessionOpened', viewDefaults: { showTrackedChanges: true } });
    const model = await materializeDocumentPullSession(harness.transport(dispatch), opened);
    expect(model.body.flatMap((element) => element.type === 'paragraph'
      ? element.runs.filter((run) => run.type === 'text').map((run) => run.text) : []))
      .toEqual(['first', 'second']);

    // Media stay readable after the cursor's terminal ACK; missing optional
    // capabilities degrade rather than falling back to an OOXML archive.
    await dispatch({ data: { type: 'extractImage', id: 2, path: 'media/1' } } as MessageEvent);
    expect(first.extract_image).toHaveBeenCalledWith('media/1');
    expect(response(2)).toMatchObject({ type: 'imageExtracted' });
    await dispatch({ data: { type: 'resourceUsage', id: 3 } } as MessageEvent);
    await dispatch({ data: { type: 'toMarkdown', id: 4 } } as MessageEvent);
    expect(response(3)).toEqual({ type: 'resourceUsage', id: 3, usage: undefined });
    expect(response(4)).toMatchObject({
      type: 'error', message: 'Markdown conversion is unsupported for this source',
    });
    expect(firstClose).not.toHaveBeenCalled();

    // A parse superseded while its source is still opening closes that source
    // and never touches the newer generation's source.
    const superseded = parse(5);
    await vi.waitFor(() => expect(state.openSource).toHaveBeenCalledTimes(2));
    expect(firstClose).toHaveBeenCalledTimes(1);
    await parse(6);
    resolveSecond({ archive: second, viewDefaults: {}, close: secondClose });
    await superseded;
    expect(secondClose).toHaveBeenCalledTimes(1);
    expect(thirdClose).not.toHaveBeenCalled();
    expect(response(5)).toMatchObject({ type: 'error', message: expect.stringContaining('superseded') });
    expect(response(6)).toMatchObject({ type: 'documentSessionOpened' });
    expect(response(6)).not.toHaveProperty('viewDefaults.showTrackedChanges');

    // Cursor-open failure closes the source; an open failure does not fall back.
    await parse(7);
    expect(thirdClose).toHaveBeenCalledTimes(1);
    expect(brokenClose).toHaveBeenCalledTimes(1);
    expect(response(7)).toMatchObject({ type: 'error', message: 'cursor open failed' });
    await parse(8);
    expect(response(8)).toMatchObject({ type: 'error', message: 'source open failed' });
    expect(state.ooxmlConstruct).not.toHaveBeenCalled();

    // Without a source the ordinary OOXML path initializes and answers without viewDefaults.
    await parse(10, {});
    expect(state.setWasmInput).toHaveBeenCalledOnce();
    expect(state.ensureReady).toHaveBeenCalledOnce();
    expect(state.ooxmlConstruct).toHaveBeenCalledOnce();
    expect(response(10)).toMatchObject({ type: 'documentSessionOpened' });
    expect(response(10)).not.toHaveProperty('viewDefaults');
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
