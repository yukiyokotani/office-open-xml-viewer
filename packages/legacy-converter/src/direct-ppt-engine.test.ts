// PPT wiring of the shared direct-source runtime: the generated glue's
// LegacyPptPresentation, its native close, the 256 MiB source budget and the
// initialization/abort/cleanup ownership rules. URL validation, pinning and
// trap containment are tested once in direct-source-runtime.test.ts.
import { describe, expect, it, vi } from 'vitest';
import {
  createLegacyPptSourceEngine,
  MAX_LEGACY_PPT_SOURCE_BYTES,
  type LegacyPptGlue,
  type LegacyPptNativeArchive,
} from './direct-ppt-engine.js';

const wasmUrl = 'https://example.test/direct-ppt.wasm';

class FakeArchive implements LegacyPptNativeArchive {
  readonly close = vi.fn();
  readonly released = vi.fn();
  constructor(readonly bytes: Uint8Array) {}
  free(): void { this.released(); }
  close_presentation_session(): void { this.close(); }
  presentation_bootstrap(): Uint8Array { return new Uint8Array(); }
  pull_slide(): Uint8Array { return new Uint8Array(); }
  acknowledge_slide(): void {}
  cancel_slide(): void {}
  assert_healthy(): void {}
  extract_image(): Uint8Array { return new Uint8Array(); }
  slide_cursor_resource_usage(): Uint8Array { throw new Error('unavailable'); }
}

function glueWith(Presentation: new (bytes: Uint8Array) => LegacyPptNativeArchive, init = vi.fn(async () => undefined)) {
  return { default: init, LegacyPptPresentation: Presentation } as unknown as LegacyPptGlue & { default: typeof init };
}

describe('direct PPT source engine', () => {
  it('shares one initialization, owns distinct presentations and closes each natively once before freeing', async () => {
    const archives: FakeArchive[] = [];
    const glue = glueWith(class extends FakeArchive {
      constructor(bytes: Uint8Array) { super(bytes); archives.push(this); }
    });
    const engine = createLegacyPptSourceEngine(async () => glue, async () => new Uint8Array([0]));
    const [first, second] = await Promise.all([
      engine.open(new Uint8Array([1]), wasmUrl),
      engine.open(new Uint8Array([2, 3]), wasmUrl),
    ]);
    expect(glue.default).toHaveBeenCalledTimes(1);
    expect(archives.map(archive => [...archive.bytes])).toEqual([[1], [2, 3]]);
    expect(second.sourceByteLength).toBe(2);
    first.closeArchive();
    first.closeArchive();
    expect(archives[0]!.close).toHaveBeenCalledTimes(1);
    expect(archives[0]!.released).toHaveBeenCalledTimes(1);
    expect(archives[0]!.close.mock.invocationCallOrder[0]).toBeLessThan(archives[0]!.released.mock.invocationCallOrder[0]!);
    expect(() => first.archive.assert_healthy()).toThrow(/closed/);
    // Closing one presentation leaves its sibling usable.
    expect(() => second.archive.assert_healthy()).not.toThrow();
    expect(archives[1]!.released).not.toHaveBeenCalled();
    second.closeArchive();
  });

  it('rejects the source byte budget before loading the glue', async () => {
    const load = vi.fn();
    const engine = createLegacyPptSourceEngine(load, vi.fn());
    await expect(engine.open({ byteLength: MAX_LEGACY_PPT_SOURCE_BYTES + 1 } as Uint8Array, wasmUrl)).rejects.toThrow('byte budget');
    expect(MAX_LEGACY_PPT_SOURCE_BYTES).toBe(256 * 1024 * 1024);
    expect(load).not.toHaveBeenCalled();
  });

  it('aborts one pending owner promptly while the shared initialization serves another', async () => {
    let finish!: () => void;
    const pending = new Promise<undefined>((resolve) => { finish = () => resolve(undefined); });
    const archives: FakeArchive[] = [];
    const glue = glueWith(class extends FakeArchive {
      constructor(bytes: Uint8Array) { super(bytes); archives.push(this); }
    }, vi.fn(() => pending));
    const loadGlue = vi.fn(async () => glue);
    const resolveWasm = vi.fn(async () => new Uint8Array([0]));
    const engine = createLegacyPptSourceEngine(loadGlue, resolveWasm);
    const controller = new AbortController();
    const aborted = engine.open(new Uint8Array([1]), wasmUrl, controller.signal);
    const retained = engine.open(new Uint8Array([2]), wasmUrl);
    controller.abort();
    await expect(aborted).rejects.toMatchObject({ name: 'AbortError' });
    expect(archives).toHaveLength(0);
    finish();
    const source = await retained;
    expect([loadGlue, resolveWasm, glue.default].map(fn => fn.mock.calls.length)).toEqual([1, 1, 1]);
    expect(archives.map(archive => [...archive.bytes])).toEqual([[2]]);
    source.closeArchive();
  });

  it('keeps an initialization failure sticky without constructing', async () => {
    const construct = vi.fn();
    const glue = glueWith(class extends FakeArchive {
      constructor(bytes: Uint8Array) { super(bytes); construct(); }
    }, vi.fn(async () => { throw new Error('initialization failed'); }));
    const engine = createLegacyPptSourceEngine(async () => glue, async () => new Uint8Array([0]));
    await expect(engine.open(new Uint8Array([1]), wasmUrl)).rejects.toThrow('initialization failed');
    await expect(engine.open(new Uint8Array([2]), wasmUrl)).rejects.toThrow('initialization failed');
    expect(glue.default).toHaveBeenCalledTimes(1);
    expect(construct).not.toHaveBeenCalled();
  });

  it.each([false, true])('releases a presentation aborted during construction and keeps AbortError (cleanup throws=%s)', async (cleanupThrows) => {
    const controller = new AbortController();
    const archive = new FakeArchive(new Uint8Array());
    if (cleanupThrows) {
      archive.close.mockImplementation(() => { throw new Error('close failed'); });
      archive.released.mockImplementation(() => { throw new Error('free failed'); });
    }
    const glue = glueWith(class extends FakeArchive {
      constructor(bytes: Uint8Array) { super(bytes); controller.abort(); return archive; }
    });
    const engine = createLegacyPptSourceEngine(async () => glue, async () => new Uint8Array([0]));
    await expect(engine.open(new Uint8Array(), wasmUrl, controller.signal)).rejects.toMatchObject({ name: 'AbortError' });
    expect(archive.close).toHaveBeenCalledTimes(1);
    expect(archive.released).toHaveBeenCalledTimes(1);
  });

  it('frees after a native close failure, reports the close failure and never retries', async () => {
    const archive = new FakeArchive(new Uint8Array());
    archive.close.mockImplementation(() => { throw undefined; });
    archive.released.mockImplementation(() => { throw new Error('free failed'); });
    const glue = glueWith(class extends FakeArchive {
      constructor(bytes: Uint8Array) { super(bytes); return archive; }
    });
    const source = await createLegacyPptSourceEngine(async () => glue, async () => new Uint8Array([0])).open(new Uint8Array(), wasmUrl);
    let thrown: unknown = 'not thrown';
    try { source.closeArchive(); } catch (error) { thrown = error; }
    // The primary (close) failure is reported even when it is `undefined`.
    expect(thrown).toBeUndefined();
    expect(archive.close).toHaveBeenCalledTimes(1);
    expect(archive.released).toHaveBeenCalledTimes(1);
    expect(() => source.closeArchive()).not.toThrow();
    expect(archive.close).toHaveBeenCalledTimes(1);
  });
});
