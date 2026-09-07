import { describe, expect, it, vi } from 'vitest';
import {
  createLegacyXlsSourceEngine,
  type LegacyXlsNativeArchive,
} from './direct-xls-engine.js';
import type { LegacyXlsDirectSourceDescriptor } from '@silurus/ooxml-core/internal/legacy-xls-source';

const descriptor: LegacyXlsDirectSourceDescriptor = {
  protocol: 'ooxml-legacy-xls-source/v1', builtin: 'xls',
  wasmUrl: 'https://example.test/direct-xls.wasm',
};

describe('direct XLS source engine', () => {
  it('shares sticky initialization and owns distinct unconfigured archives', async () => {
    const archives: FakeArchive[] = [];
    const glue = {
      default: vi.fn(async () => undefined),
      LegacyXlsWorkbook: class extends FakeArchive {
        constructor(bytes: Uint8Array) { super(bytes); archives.push(this); }
      },
    };
    const engine = createLegacyXlsSourceEngine(async () => glue, async () => new Uint8Array());
    const [first, second] = await Promise.all([
      engine.open(new Uint8Array([1]), descriptor),
      engine.open(new Uint8Array([2]), descriptor),
    ]);
    expect(glue.default).toHaveBeenCalledTimes(1);
    expect(archives).toHaveLength(2);
    expect(archives[0]!.configure).not.toHaveBeenCalled();
    first.closeArchive(); first.closeArchive(); second.closeArchive();
    expect(archives[0]!.closeNative).toHaveBeenCalledTimes(1);
    expect(archives[0]!.released).toHaveBeenCalledTimes(1);
  });

  it('rejects validation, limits, and pre-abort before initialization', async () => {
    const load = vi.fn();
    const engine = createLegacyXlsSourceEngine(load, vi.fn());
    await expect(engine.open(new Uint8Array(), { ...descriptor, builtin: 'ppt' } as never))
      .rejects.toThrow();
    await expect(engine.open({ byteLength: 256 * 1024 * 1024 + 1 } as Uint8Array, descriptor))
      .rejects.toThrow('byte budget');
    const abort = new AbortController(); abort.abort();
    await expect(engine.open(new Uint8Array(), descriptor, abort.signal))
      .rejects.toMatchObject({ name: 'AbortError' });
    expect(load).not.toHaveBeenCalled();
  });

  it('cancels pending initialization and cleans an archive aborted after construction', async () => {
    let finish!: () => void;
    const pending = new Promise<void>((resolve) => { finish = resolve; });
    const controller = new AbortController();
    const archive = new FakeArchive(new Uint8Array());
    const glue = {
      default: vi.fn(() => pending),
      LegacyXlsWorkbook: class extends FakeArchive {
        constructor(bytes: Uint8Array) { super(bytes); controller.abort(); return archive; }
      },
    };
    const engine = createLegacyXlsSourceEngine(async () => glue, async () => new Uint8Array());
    const firstController = new AbortController();
    const first = engine.open(new Uint8Array(), descriptor, firstController.signal);
    firstController.abort();
    await expect(first).rejects.toMatchObject({ name: 'AbortError' });
    finish();
    await expect(engine.open(new Uint8Array(), descriptor, controller.signal))
      .rejects.toMatchObject({ name: 'AbortError' });
    expect(archive.closeNative).toHaveBeenCalledTimes(1);
    expect(archive.released).toHaveBeenCalledTimes(1);
  });

  it('pins URL and keeps initialization failure sticky', async () => {
    const glue = { default: vi.fn(async () => { throw new Error('init failed'); }), LegacyXlsWorkbook: FakeArchive };
    const engine = createLegacyXlsSourceEngine(async () => glue, async () => new Uint8Array());
    await expect(engine.open(new Uint8Array(), descriptor)).rejects.toThrow('init failed');
    await expect(engine.open(new Uint8Array(), descriptor)).rejects.toThrow('init failed');
    await expect(engine.open(new Uint8Array(), { ...descriptor, wasmUrl: 'https://example.test/b.wasm' }))
      .rejects.toThrow('pinned');
    expect(glue.default).toHaveBeenCalledTimes(1);
  });
});

class FakeArchive implements LegacyXlsNativeArchive {
  readonly released = vi.fn(); readonly closeNative = vi.fn(); readonly configure = vi.fn();
  constructor(readonly bytes: Uint8Array) {}
  free(): void { this.released(); }
  close_workbook_session(): void { this.closeNative(); }
  measurement_request(): Uint8Array { return new Uint8Array(); }
  configure_mdw(mdw?: number): void { this.configure(mdw); }
  parse(): Uint8Array { return new Uint8Array(); }
  open_sheet_cursor(): void {} pull_sheet_cursor(): Uint8Array { return new Uint8Array(); }
  sheet_cursor_pull_finished(): boolean { return false; }
  acknowledge_sheet_cursor_terminal(): void {} cancel_sheet_cursor(): void {} close_sheet_cursor(): void {}
  resource_usage(): Uint8Array { throw new Error('unavailable'); }
  sheet_cursor_resource_usage(): Uint8Array { throw new Error('unavailable'); }
  extract_image(): Uint8Array { return new Uint8Array(); }
  to_markdown(): string { throw new Error('unsupported'); }
  assert_healthy(): void {}
}
