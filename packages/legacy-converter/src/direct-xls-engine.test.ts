// XLS wiring of the shared direct-source runtime: the generated glue's
// LegacyXlsWorkbook, its native close, the 256 MiB source budget and the
// initialization/abort ownership rules. URL validation and trap containment
// are tested once in direct-source-runtime.test.ts; the Normal-font host
// layout is the XLSX host's (packages/xlsx host-layout.ts), not the engine's.
import { describe, expect, it, vi } from 'vitest';
import {
  createLegacyXlsSourceEngine,
  MAX_LEGACY_XLS_SOURCE_BYTES,
  type LegacyXlsGlue,
  type LegacyXlsNativeArchive,
} from './direct-xls-engine.js';

const wasmUrl = 'https://example.test/direct-xls.wasm';

class FakeArchive implements LegacyXlsNativeArchive {
  readonly released = vi.fn(); readonly closeNative = vi.fn(); readonly configure = vi.fn();
  constructor(readonly bytes: Uint8Array) {}
  free(): void { this.released(); }
  close_workbook_session(): void { this.closeNative(); }
  host_layout_request(): Uint8Array { return new TextEncoder().encode('null'); }
  configure_host_layout(mdw?: number): void { this.configure(mdw); }
  parse(): Uint8Array { return new Uint8Array(); }
  open_sheet_cursor(): void {} pull_sheet_cursor(): Uint8Array { return new Uint8Array(); }
  sheet_cursor_pull_finished(): boolean { return false; }
  acknowledge_sheet_cursor_terminal(): void {} cancel_sheet_cursor(): void {} close_sheet_cursor(): void {}
  sheet_cursor_resource_usage(): Uint8Array { throw new Error('unavailable'); }
  extract_image(): Uint8Array { return new Uint8Array(); }
  assert_healthy(): void {}
}

function glueWith(Workbook: new (bytes: Uint8Array) => LegacyXlsNativeArchive, init = vi.fn(async () => undefined)) {
  return { default: init, LegacyXlsWorkbook: Workbook } as unknown as LegacyXlsGlue & { default: typeof init };
}

describe('direct XLS source engine', () => {
  it('shares one initialization and owns distinct archives it never configures', async () => {
    const archives: FakeArchive[] = [];
    const glue = glueWith(class extends FakeArchive {
      constructor(bytes: Uint8Array) { super(bytes); archives.push(this); }
    });
    const engine = createLegacyXlsSourceEngine(async () => glue, async () => new Uint8Array());
    const [first, second] = await Promise.all([
      engine.open(new Uint8Array([1]), wasmUrl),
      engine.open(new Uint8Array([2]), wasmUrl),
    ]);
    expect(glue.default).toHaveBeenCalledTimes(1);
    expect(archives.map(archive => [...archive.bytes])).toEqual([[1], [2]]);
    expect(first.sourceByteLength).toBe(1);
    // Host layout is answered by the XLSX host through the archive.
    expect(archives[0]!.configure).not.toHaveBeenCalled();
    first.closeArchive(); first.closeArchive(); second.closeArchive();
    expect(archives[0]!.closeNative).toHaveBeenCalledTimes(1);
    expect(archives[0]!.released).toHaveBeenCalledTimes(1);
    expect(archives[0]!.closeNative.mock.invocationCallOrder[0])
      .toBeLessThan(archives[0]!.released.mock.invocationCallOrder[0]!);
    expect(() => first.archive.parse()).toThrow(/closed/);
  });

  it('rejects the source byte budget before initialization', async () => {
    const load = vi.fn();
    const engine = createLegacyXlsSourceEngine(load, vi.fn());
    await expect(engine.open({ byteLength: MAX_LEGACY_XLS_SOURCE_BYTES + 1 } as Uint8Array, wasmUrl))
      .rejects.toThrow('byte budget');
    expect(MAX_LEGACY_XLS_SOURCE_BYTES).toBe(256 * 1024 * 1024);
    expect(load).not.toHaveBeenCalled();
  });

  it('cancels pending initialization and cleans an archive aborted after construction', async () => {
    let finish!: () => void;
    const pending = new Promise<undefined>((resolve) => { finish = () => resolve(undefined); });
    const controller = new AbortController();
    const archive = new FakeArchive(new Uint8Array());
    const glue = glueWith(class extends FakeArchive {
      constructor(bytes: Uint8Array) { super(bytes); controller.abort(); return archive; }
    }, vi.fn(() => pending));
    const engine = createLegacyXlsSourceEngine(async () => glue, async () => new Uint8Array());
    const firstController = new AbortController();
    const first = engine.open(new Uint8Array(), wasmUrl, firstController.signal);
    firstController.abort();
    await expect(first).rejects.toMatchObject({ name: 'AbortError' });
    finish();
    await expect(engine.open(new Uint8Array(), wasmUrl, controller.signal))
      .rejects.toMatchObject({ name: 'AbortError' });
    expect(archive.closeNative).toHaveBeenCalledTimes(1);
    expect(archive.released).toHaveBeenCalledTimes(1);
  });

  it('keeps an initialization failure sticky without retrying', async () => {
    const glue = glueWith(FakeArchive, vi.fn(async () => { throw new Error('init failed'); }));
    const engine = createLegacyXlsSourceEngine(async () => glue, async () => new Uint8Array());
    await expect(engine.open(new Uint8Array(), wasmUrl)).rejects.toThrow('init failed');
    await expect(engine.open(new Uint8Array(), wasmUrl)).rejects.toThrow('init failed');
    expect(glue.default).toHaveBeenCalledTimes(1);
  });
});
