import { describe, expect, it, vi } from 'vitest';
import {
  createLegacyDocSourceEngine,
  legacyDocViewDefaults,
  MAX_LEGACY_DOC_SOURCE_BYTES,
  type LegacyDocGlue,
  type LegacyDocNativeDocument,
} from './direct-doc-engine.js';

const productionGlue = vi.hoisted(() => ({ loaded: vi.fn() }));
vi.mock('./wasm-direct-doc/legacy_doc_direct.js', () => {
  productionGlue.loaded();
  return { default: vi.fn(), LegacyDocDocument: class {} };
});

const wasmUrl = 'https://example.test/direct-doc.wasm';

class FakeDocument {
  readonly close = vi.fn();
  readonly released = vi.fn();
  constructor(readonly bytes: Uint8Array, readonly budget?: number) {}
  open_document_cursor(): void {}
  pull_document_chunk(): Uint8Array { return new Uint8Array(); }
  document_chunk_done(): boolean { return true; }
  acknowledge_document_chunk(): void {}
  cancel_document_cursor(): void {}
  assert_healthy(): void {}
  extract_image(): Uint8Array { return new Uint8Array(); }
  free(): void { this.released(); }
  close_document_session(): void { this.close(); }
}
const glueFor = (construct: (bytes: Uint8Array, budget?: number) => FakeDocument, init = vi.fn(async () => undefined)) => ({
  default: init,
  LegacyDocDocument: function (bytes: Uint8Array, budget?: number) { return construct(bytes, budget); },
}) as unknown as LegacyDocGlue & { default: typeof init };

describe('direct DOC source engine', () => {
  it('keeps the production source module from loading generated glue on pre-admission failure', async () => {
    const { openModelSource } = await import('./legacy-doc-source-module.js');
    const abort = new AbortController(); abort.abort();
    await expect(openModelSource(new Uint8Array(), { wasmUrl, maxInputBytes: 16 }, abort.signal))
      .rejects.toMatchObject({ name: 'AbortError' });
    await expect(openModelSource(new Uint8Array(32), { wasmUrl, maxInputBytes: 16 })).rejects.toThrow(RangeError);
    expect(productionGlue.loaded).not.toHaveBeenCalled();
  });

  it('initializes once, pins the WASM URL, and passes the optional model budget', async () => {
    const documents: FakeDocument[] = [];
    const glue = glueFor((bytes, budget) => { const document = new FakeDocument(bytes, budget); documents.push(document); return document; });
    const resolve = vi.fn(async () => new Uint8Array([0]));
    const engine = createLegacyDocSourceEngine(async () => glue, resolve, 4096);
    const first = await engine.open(new Uint8Array([1]), wasmUrl);
    const second = await engine.open(new Uint8Array([2]), wasmUrl);
    expect(glue.default).toHaveBeenCalledTimes(1);
    expect(resolve).toHaveBeenCalledTimes(1);
    expect(documents.map(({ bytes, budget }) => [bytes[0], budget])).toEqual([[1, 4096], [2, 4096]]);
    await expect(engine.open(new Uint8Array(), 'https://example.test/other.wasm')).rejects.toThrow('pinned');
    first.closeArchive(); second.closeArchive();
    expect(documents.map(d => [d.close.mock.calls.length, d.released.mock.calls.length])).toEqual([[1, 1], [1, 1]]);
  });

  it('keeps explicit reading placement local to each concurrent native open', async () => {
    const argumentsSeen: Array<readonly [number, number | undefined, boolean | undefined]> = [];
    const documents: FakeDocument[] = [];
    const init = vi.fn(async () => undefined);
    const glue = { default: init, LegacyDocDocument: function(bytes: Uint8Array, budget?: number, reading?: boolean) {
      argumentsSeen.push([bytes[0], budget, reading]); const document = new FakeDocument(bytes, budget); documents.push(document); return document;
    } } as unknown as LegacyDocGlue;
    const engine = createLegacyDocSourceEngine(async () => glue, async () => new Uint8Array(), 4096);
    const [reading, strict] = await Promise.all([engine.open(new Uint8Array([1]), wasmUrl, undefined, 'relocateForReading'), engine.open(new Uint8Array([2]), wasmUrl)]);
    expect(argumentsSeen).toEqual([[1, 4096, true], [2, 4096, undefined]]);
    expect(init).toHaveBeenCalledTimes(1);
    reading.closeArchive(); strict.closeArchive();
    expect(documents.map(document => [document.close.mock.calls.length, document.released.mock.calls.length])).toEqual([[1, 1], [1, 1]]);
  });

  it('rejects an invalid URL, the byte budget, and pre-abort before loading', async () => {
    const load = vi.fn();
    const engine = createLegacyDocSourceEngine(load, vi.fn());
    await expect(engine.open(new Uint8Array(), 'direct-doc.wasm')).rejects.toThrow(TypeError);
    await expect(engine.open({ byteLength: MAX_LEGACY_DOC_SOURCE_BYTES + 1 } as Uint8Array, wasmUrl)).rejects.toThrow('byte budget');
    const abort = new AbortController(); abort.abort();
    await expect(engine.open(new Uint8Array(), wasmUrl, abort.signal)).rejects.toMatchObject({ name: 'AbortError' });
    expect(load).not.toHaveBeenCalled();
  });

  it.each([0, 1.5, Number.NaN, MAX_LEGACY_DOC_SOURCE_BYTES + 1])('rejects invalid model budget %s before loading glue', budget => {
    const load = vi.fn();
    expect(() => createLegacyDocSourceEngine(load, vi.fn(), budget)).toThrow(RangeError);
    expect(load).not.toHaveBeenCalled();
  });

  it('closes and frees exactly once a document aborted during construction', async () => {
    const controller = new AbortController();
    const document = new FakeDocument(new Uint8Array());
    const glue = glueFor(() => { controller.abort(); return document; });
    await expect(createLegacyDocSourceEngine(async () => glue, async () => new Uint8Array())
      .open(new Uint8Array(), wasmUrl, controller.signal)).rejects.toMatchObject({ name: 'AbortError' });
    expect(document.close).toHaveBeenCalledTimes(1);
    expect(document.released).toHaveBeenCalledTimes(1);
  });

  it('keeps initialization failure sticky', async () => {
    const glue = glueFor(bytes => new FakeDocument(bytes), vi.fn(async () => { throw new Error('init failed'); }));
    const engine = createLegacyDocSourceEngine(async () => glue, async () => new Uint8Array());
    await expect(engine.open(new Uint8Array(), wasmUrl)).rejects.toThrow('init failed');
    await expect(engine.open(new Uint8Array(), wasmUrl)).rejects.toThrow('init failed');
    expect(glue.default).toHaveBeenCalledTimes(1);
  });
});

describe('explicit DOC word-breaking reading construction', () => {
  it('selects the fourth native argument and keeps relocation strict on every open', async () => {
    const calls: unknown[][] = [];
    class Archive extends FakeDocument {
      constructor(bytes: Uint8Array, budget?: number, relocation?: boolean, reading?: boolean) {
        super(bytes, budget); calls.push([budget, relocation, reading]);
      }
    }
    const glue = { default: async () => undefined, LegacyDocDocument: Archive } satisfies LegacyDocGlue;
    const engine = createLegacyDocSourceEngine(async () => glue, async () => new Uint8Array(), 4096);
    const a = await engine.open(new Uint8Array(), wasmUrl);
    const b = await engine.open(new Uint8Array(), wasmUrl, undefined, 'simplifyForReading');
    const c = await engine.open(new Uint8Array(), wasmUrl, undefined, 'strict');
    expect(calls).toEqual([[4096, undefined, undefined], [4096, false, true], [4096, undefined, undefined]]);
    a.closeArchive(); b.closeArchive(); c.closeArchive();
  });
});

describe('legacyDocViewDefaults', () => {
  const document = (inPrint?: boolean, marks?: boolean) => ({
    ...(inPrint === undefined ? {} : { revision_markup_in_print: () => inPrint }),
    ...(marks === undefined ? {} : { has_revision_marks: () => marks }),
  }) as Pick<LegacyDocNativeDocument, 'revision_markup_in_print' | 'has_revision_marks'>;

  it('shows tracked changes only when the DOC prints revision markup and carries revision marks', () => {
    // MS-DOC 2.7.2 fRMPrint: Word's PDF output is the display target.
    expect(legacyDocViewDefaults(document(true, true))).toEqual({ showTrackedChanges: true });
    for (const [inPrint, marks] of [[true, false], [false, true], [undefined, true], [true, undefined]] as const) {
      expect(legacyDocViewDefaults(document(inPrint, marks))).toEqual({});
    }
  });
});

describe('DOC picture-bullet constructor capability slot', () => {
  it('uses only the fifth argument for stored-size reading, preserving strict construction', async () => {
    const calls: unknown[][] = [];
    class Document implements LegacyDocNativeDocument {
      constructor(...args: [Uint8Array, number?, boolean?, boolean?, boolean?]) { calls.push(args); }
      free(): void {}
      open_document_cursor(): void {}
      pull_document_chunk(): Uint8Array { return new Uint8Array(); }
      document_chunk_done(): boolean { return true; }
      acknowledge_document_chunk(): void {}
      cancel_document_cursor(): void {}
      close_document_session(): void {}
      assert_healthy(): void {}
      extract_image(): Uint8Array { return new Uint8Array(); }
    }
    const glue: LegacyDocGlue = { default: async () => undefined, LegacyDocDocument: Document };
    const strict = createLegacyDocSourceEngine(async () => glue, async () => new Uint8Array(), 4096);
    const reading = createLegacyDocSourceEngine(async () => glue, async () => new Uint8Array(), 4096, 'storedSizeForReading');
    (await strict.open(new Uint8Array([1]), wasmUrl)).closeArchive();
    (await reading.open(new Uint8Array([2]), wasmUrl)).closeArchive();
    expect(calls[0]).toEqual([new Uint8Array([1]), 4096]);
    expect(calls[1]).toEqual([new Uint8Array([2]), 4096, false, false, true]);
  });
});
