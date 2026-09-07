import { describe, expect, it, vi } from 'vitest';
import {
  createLegacyPptSourceEngine,
  type LegacyPptNativeArchive,
} from './direct-ppt-engine.js';
import type { LegacyPptDirectSourceDescriptor } from '@silurus/ooxml-core/internal/legacy-ppt-source';

const descriptor: LegacyPptDirectSourceDescriptor = {
  protocol: 'ooxml-legacy-ppt-source/v1',
  builtin: 'ppt',
  wasmUrl: 'https://example.test/direct-ppt.wasm',
};

describe('direct PPT source engine', () => {
  it('shares initialization but creates and owns distinct source sessions', async () => {
    const archives: FakeArchive[] = [];
    const glue = {
      default: vi.fn(async () => undefined),
      LegacyPptPresentation: class extends FakeArchive {
        constructor(bytes: Uint8Array) {
          super(bytes);
          archives.push(this);
        }
      },
    };
    const engine = createLegacyPptSourceEngine(async () => glue, async () => new Uint8Array([0]));
    const [first, second] = await Promise.all([
      engine.open(new Uint8Array([1]), descriptor),
      engine.open(new Uint8Array([2]), descriptor),
    ]);
    expect(glue.default).toHaveBeenCalledTimes(1);
    expect(archives).toHaveLength(2);
    expect(first.archive).not.toBe(second.archive);
    first.closeArchive();
    expect(() => second.archive.assert_healthy()).not.toThrow();
    expect(archives[1]!.released).not.toHaveBeenCalled();
    second.closeArchive();
  });

  it('aborts one pending owner promptly while shared initialization serves another', async () => {
    let finishInitialization!: () => void;
    const initialization = new Promise<void>((resolve) => { finishInitialization = resolve; });
    const archives: FakeArchive[] = [];
    const glue = {
      default: vi.fn(() => initialization),
      LegacyPptPresentation: class extends FakeArchive {
        constructor(bytes: Uint8Array) {
          super(bytes);
          archives.push(this);
        }
      },
    };
    const loadGlue = vi.fn(async () => glue);
    const resolveWasm = vi.fn(async () => new Uint8Array([0]));
    const engine = createLegacyPptSourceEngine(loadGlue, resolveWasm);
    const controller = new AbortController();

    const aborted = engine.open(new Uint8Array([1]), descriptor, controller.signal);
    const retained = engine.open(new Uint8Array([2]), descriptor);
    await expect(engine.open(new Uint8Array([3]), {
      ...descriptor,
      wasmUrl: 'https://example.test/other.wasm',
    })).rejects.toThrow('pinned');
    controller.abort();
    await expect(aborted).rejects.toMatchObject({ name: 'AbortError' });
    expect(archives).toHaveLength(0);

    finishInitialization();
    const source = await retained;
    expect(loadGlue).toHaveBeenCalledTimes(1);
    expect(resolveWasm).toHaveBeenCalledTimes(1);
    expect(glue.default).toHaveBeenCalledTimes(1);
    expect(archives).toHaveLength(1);
    expect(source.sourceByteLength).toBe(1);
    await expect(engine.open(new Uint8Array([4]), {
      ...descriptor,
      wasmUrl: 'https://example.test/other.wasm',
    })).rejects.toThrow('pinned');

    source.closeArchive();
    source.closeArchive();
    expect(archives[0]!.close).toHaveBeenCalledTimes(1);
    expect(archives[0]!.released).toHaveBeenCalledTimes(1);
  });

  it('keeps initialization failure sticky', async () => {
    let attempt = 0;
    const glue = {
      default: vi.fn(async () => {
        attempt += 1;
        throw new Error('initialization failed');
      }),
      LegacyPptPresentation: class extends FakeArchive {
        constructor(bytes: Uint8Array) {
          super(bytes);
          throw new Error('must not construct');
        }
      },
    };
    const engine = createLegacyPptSourceEngine(async () => glue, async () => new Uint8Array([0]));
    await expect(engine.open(new Uint8Array([1]), descriptor)).rejects.toThrow('initialization failed');
    await expect(engine.open(new Uint8Array([2]), descriptor)).rejects.toThrow('initialization failed');
    expect(attempt).toBe(1);
  });

  it('frees a source aborted synchronously after construction', async () => {
    const controller = new AbortController();
    const archive = new FakeArchive(new Uint8Array([1]));
    const glue = {
      default: vi.fn(async () => undefined),
      LegacyPptPresentation: class extends FakeArchive {
        constructor(bytes: Uint8Array) {
          super(bytes);
          controller.abort();
          return archive;
        }
      },
    };
    const engine = createLegacyPptSourceEngine(async () => glue, async () => new Uint8Array([0]));
    await expect(engine.open(new Uint8Array([2]), descriptor, controller.signal))
      .rejects.toMatchObject({ name: 'AbortError' });
    expect(archive.close).toHaveBeenCalledTimes(1);
    expect(archive.released).toHaveBeenCalledTimes(1);
  });

  it('rejects over-budget bytes and pre-aborted calls before loading glue', async () => {
    const loadGlue = vi.fn();
    const engine = createLegacyPptSourceEngine(loadGlue, vi.fn());
    const oversized = { byteLength: 256 * 1024 * 1024 + 1 } as Uint8Array;
    await expect(engine.open(oversized, descriptor)).rejects.toThrow('byte budget');
    const controller = new AbortController();
    controller.abort();
    await expect(engine.open(new Uint8Array(), descriptor, controller.signal))
      .rejects.toMatchObject({ name: 'AbortError' });
    expect(loadGlue).not.toHaveBeenCalled();
  });

  it('frees after close failure, preserves the primary value, and never retries close', async () => {
    const archive = new FakeArchive(new Uint8Array());
    archive.close.mockImplementation(() => { throw undefined; });
    archive.released.mockImplementation(() => { throw new Error('free failed'); });
    const glue = {
      default: vi.fn(async () => undefined),
      LegacyPptPresentation: class extends FakeArchive {
        constructor(bytes: Uint8Array) {
          super(bytes);
          return archive;
        }
      },
    };
    const source = await createLegacyPptSourceEngine(
      async () => glue,
      async () => new Uint8Array([0]),
    ).open(new Uint8Array(), descriptor);
    let thrown: unknown = 'not thrown';
    try { source.closeArchive(); } catch (error) { thrown = error; }
    expect(thrown).toBeUndefined();
    expect(archive.close).toHaveBeenCalledTimes(1);
    expect(archive.released).toHaveBeenCalledTimes(1);
    expect(() => source.closeArchive()).not.toThrow();
    expect(archive.close).toHaveBeenCalledTimes(1);
  });

  it('preserves AbortError when constructor-abort cleanup throws', async () => {
    const controller = new AbortController();
    const archive = new FakeArchive(new Uint8Array());
    archive.close.mockImplementation(() => { throw new Error('close failed'); });
    archive.released.mockImplementation(() => { throw new Error('free failed'); });
    const glue = {
      default: vi.fn(async () => undefined),
      LegacyPptPresentation: class extends FakeArchive {
        constructor(bytes: Uint8Array) {
          super(bytes);
          controller.abort();
          return archive;
        }
      },
    };
    const engine = createLegacyPptSourceEngine(async () => glue, async () => new Uint8Array([0]));
    await expect(engine.open(new Uint8Array(), descriptor, controller.signal))
      .rejects.toMatchObject({ name: 'AbortError' });
    expect(archive.close).toHaveBeenCalledTimes(1);
    expect(archive.released).toHaveBeenCalledTimes(1);
  });
});

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
  extract_media(): Uint8Array { throw new Error('unsupported'); }
  extract_font(): Uint8Array { throw new Error('unsupported'); }
  slide_cursor_resource_usage(): Uint8Array { throw new Error('unavailable'); }
}
