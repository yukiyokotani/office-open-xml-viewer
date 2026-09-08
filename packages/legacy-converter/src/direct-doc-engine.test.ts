import { describe, expect, it, vi } from 'vitest';
import type { LegacyDocDirectSourceDescriptor } from '@silurus/ooxml-core/internal/legacy-doc-source';
import {
  createLegacyDocSourceEngine,
  openLegacyDocSource,
  type LegacyDocNativeDocument,
} from './direct-doc-engine.js';

const defaultGlue = vi.hoisted(() => ({ loaded: vi.fn() }));
vi.mock('./wasm-direct-doc/legacy_office_converter.js', () => {
  defaultGlue.loaded();
  return { default: vi.fn(), LegacyDocDocument: class {} };
});
import { DocumentPullWorker } from '../../docx/src/document-pull-worker.js';
import { createLocalDocumentPullTransport } from '../../docx/src/document-pull-worker.js';
import { materializeDocumentPullSession } from '../../docx/src/document-pull-client.js';

const descriptor: LegacyDocDirectSourceDescriptor = {
  protocol: 'ooxml-legacy-doc-source/v1', builtin: 'doc',
  wasmUrl: 'https://example.test/direct-doc.wasm',
};

describe('direct DOC source engine', () => {
  it('keeps the production generated-glue loader lazy on pre-admission failure', async () => {
    const abort = new AbortController(); abort.abort();
    await expect(openLegacyDocSource(new Uint8Array(), descriptor, abort.signal))
      .rejects.toMatchObject({ name: 'AbortError' });
    expect(defaultGlue.loaded).not.toHaveBeenCalled();
  });

  it('initializes once, pins the URL, and passes the optional model budget', async () => {
    const documents: FakeDocument[] = [];
    const glue = {
      default: vi.fn(async () => undefined),
      LegacyDocDocument: class extends FakeDocument {
        constructor(bytes: Uint8Array, budget?: number) { super(bytes, budget); documents.push(this); }
      },
    };
    const resolve = vi.fn(async () => new Uint8Array([0]));
    const engine = createLegacyDocSourceEngine(async () => glue, resolve, 4096);
    const first = await engine.open(new Uint8Array([1]), descriptor);
    const second = await engine.open(new Uint8Array([2]), descriptor);
    expect(glue.default).toHaveBeenCalledTimes(1);
    expect(resolve).toHaveBeenCalledTimes(1);
    expect(documents.map(({ budget }) => budget)).toEqual([4096, 4096]);
    await expect(engine.open(new Uint8Array(), { ...descriptor,
      wasmUrl: 'https://example.test/other.wasm' })).rejects.toThrow('pinned');
    first.closeArchive(); second.closeArchive();
  });

  it('uses the existing DocumentPullWorker identity, pull, ACK and terminal protocol', async () => {
    const document = new FakeDocument(new Uint8Array([1]));
    const source = await engineFor(document).open(new Uint8Array([1]), descriptor);
    const worker = new DocumentPullWorker(() => source.archive);
    const identity = { sessionId: 17, operationId: 23, generation: 29 };
    worker.open(identity);
    await expect(materializeDocumentPullSession(
      createLocalDocumentPullTransport(worker), identity,
    )).resolves.toMatchObject({ body: [] });
    expect(() => source.archive.acknowledge_document_chunk(0, 24, 29)).toThrow('stale identity');
    expect(document.close).not.toHaveBeenCalled();
    source.archive.assert_healthy();
    expect(source.archive.extract_image('image/1')).toEqual(new Uint8Array([9]));
    source.archive.cancel_document_cursor();
    expect(document.calls).toEqual([
      'open:23:29', 'pull:0:23:29', 'done', 'ack:0:23:29',
      'healthy', 'image:image/1', 'cancel',
    ]);
    source.closeArchive();
    expect(document.close).toHaveBeenCalledTimes(1);
    expect(document.released).toHaveBeenCalledTimes(1);
  });

  it('rejects validation, byte budget, and pre-abort before loading', async () => {
    const load = vi.fn();
    const engine = createLegacyDocSourceEngine(load, vi.fn());
    await expect(engine.open(new Uint8Array(), { ...descriptor, builtin: 'xls' } as never)).rejects.toThrow();
    await expect(engine.open({ byteLength: 256 * 1024 * 1024 + 1 } as Uint8Array, descriptor)).rejects.toThrow('byte budget');
    const abort = new AbortController(); abort.abort();
    await expect(engine.open(new Uint8Array(), descriptor, abort.signal)).rejects.toMatchObject({ name: 'AbortError' });
    expect(load).not.toHaveBeenCalled();
  });

  it.each([0, -1, 1.5, Number.NaN, 256 * 1024 * 1024 + 1])(
    'rejects invalid model budget %s before loading glue',
    (budget) => {
      const load = vi.fn();
      expect(() => createLegacyDocSourceEngine(load, vi.fn(), budget)).toThrow(RangeError);
      expect(load).not.toHaveBeenCalled();
    },
  );

  it('cleans an archive aborted during construction and closes/frees exactly once', async () => {
    const controller = new AbortController();
    const document = new FakeDocument(new Uint8Array());
    const glue = { default: vi.fn(async () => undefined), LegacyDocDocument: class extends FakeDocument {
      constructor(bytes: Uint8Array) { super(bytes); controller.abort(); return document; }
    } };
    await expect(createLegacyDocSourceEngine(async () => glue, async () => new Uint8Array())
      .open(new Uint8Array(), descriptor, controller.signal)).rejects.toMatchObject({ name: 'AbortError' });
    expect(document.close).toHaveBeenCalledTimes(1);
    expect(document.released).toHaveBeenCalledTimes(1);
  });

  it('keeps initialization failure sticky', async () => {
    const glue = { default: vi.fn(async () => { throw new Error('init failed'); }),
      LegacyDocDocument: FakeDocument };
    const engine = createLegacyDocSourceEngine(async () => glue, async () => new Uint8Array());
    await expect(engine.open(new Uint8Array(), descriptor)).rejects.toThrow('init failed');
    await expect(engine.open(new Uint8Array(), descriptor)).rejects.toThrow('init failed');
    expect(glue.default).toHaveBeenCalledTimes(1);
  });
});

function engineFor(document: FakeDocument) {
  const glue = { default: vi.fn(async () => undefined), LegacyDocDocument: class extends FakeDocument {
    constructor(bytes: Uint8Array) { super(bytes); return document; }
  } };
  return createLegacyDocSourceEngine(async () => glue, async () => new Uint8Array());
}

class FakeDocument implements LegacyDocNativeDocument {
  readonly calls: string[] = [];
  readonly close = vi.fn(); readonly released = vi.fn();
  private identity: { operation: number; generation: number } | undefined;
  constructor(readonly bytes: Uint8Array, readonly budget?: number) {}
  free(): void { this.released(); }
  open_document_cursor(operation: number, generation: number): void {
    this.identity = { operation, generation };
    this.calls.push(`open:${operation}:${generation}`);
  }
  pull_document_chunk(sequence: number, operation: number, generation: number): Uint8Array {
    this.assertIdentity(operation, generation);
    this.calls.push(`pull:${sequence}:${operation}:${generation}`);
    return new TextEncoder().encode(JSON.stringify({ kind: 'complete', document: { body: [] } }));
  }
  document_chunk_done(): boolean { this.calls.push('done'); return true; }
  acknowledge_document_chunk(sequence: number, operation: number, generation: number): void {
    this.assertIdentity(operation, generation);
    this.calls.push(`ack:${sequence}:${operation}:${generation}`);
  }
  cancel_document_cursor(): void { this.calls.push('cancel'); }
  close_document_session(): void { this.close(); }
  assert_healthy(): void { this.calls.push('healthy'); }
  extract_image(key: string): Uint8Array { this.calls.push(`image:${key}`); return new Uint8Array([9]); }
  private assertIdentity(operation: number, generation: number): void {
    if (this.identity?.operation !== operation || this.identity.generation !== generation) {
      throw new Error('stale identity');
    }
  }
}
