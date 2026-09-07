import { describe, expect, it, vi } from 'vitest';
import { TerminalResourceOwner } from '@silurus/ooxml-core/internal/canvas-viewer-mechanics';
import type { XlsxWorkbook } from './workbook.js';
import { settleLegacyXlsLoad } from './legacy-xls-load.js';

const sourceOptions = {
  xls: { source: {
    protocol: 'ooxml-legacy-xls-source/v1' as const,
    builtin: 'xls' as const,
    wasmUrl: 'https://example.test/direct.wasm',
  } },
};

function ownedWorkbook(): XlsxWorkbook {
  let retained: () => void = () => undefined;
  return {
    _retainLegacyXlsSignalCleanup(cleanup: () => void) { retained = cleanup; },
    destroy() {
      const cleanup = retained;
      retained = () => undefined;
      cleanup();
    },
  } as unknown as XlsxWorkbook;
}

describe('settleLegacyXlsLoad', () => {
  it('releases direct-source signal composition when loading fails', async () => {
    const failure = new Error('load failed');
    const cleanup = vi.fn();
    await expect(settleLegacyXlsLoad(Promise.reject(failure), {
      options: sourceOptions, cleanup,
    })).rejects.toBe(failure);
    expect(cleanup).toHaveBeenCalledOnce();
  });

  it('retains successful direct-source cleanup but settles converter cleanup immediately', async () => {
    const cleanup = vi.fn();
    const retain = vi.fn();
    const workbook = { _retainLegacyXlsSignalCleanup: retain } as unknown as XlsxWorkbook;
    await expect(settleLegacyXlsLoad(Promise.resolve(workbook), {
      options: sourceOptions, cleanup,
    })).resolves.toBe(workbook);
    expect(cleanup).not.toHaveBeenCalled();
    expect(retain).toHaveBeenCalledWith(cleanup);

    await settleLegacyXlsLoad(Promise.resolve(workbook), {
      options: { xls: { converter: { convert: vi.fn() } } }, cleanup,
    });
    expect(cleanup).toHaveBeenCalledOnce();
  });

  it('releases composed listeners once for replacement, stale completion, and destroy', async () => {
    const owner = new TerminalResourceOwner<XlsxWorkbook>('test');
    const firstCleanup = vi.fn();
    const staleCleanup = vi.fn();
    const winnerCleanup = vi.fn();
    await owner.replace(() => settleLegacyXlsLoad(Promise.resolve(ownedWorkbook()), {
      options: sourceOptions, cleanup: firstCleanup,
    }));

    let resolveStale!: (workbook: XlsxWorkbook) => void;
    const stale = owner.replace(() => settleLegacyXlsLoad(
      new Promise<XlsxWorkbook>((resolve) => { resolveStale = resolve; }),
      { options: sourceOptions, cleanup: staleCleanup },
    ));
    await owner.replace(() => settleLegacyXlsLoad(Promise.resolve(ownedWorkbook()), {
      options: sourceOptions, cleanup: winnerCleanup,
    }));
    expect(firstCleanup).toHaveBeenCalledOnce();
    resolveStale(ownedWorkbook());
    await stale;
    expect(staleCleanup).toHaveBeenCalledOnce();
    expect(winnerCleanup).not.toHaveBeenCalled();
    owner.close();
    owner.close();
    expect(winnerCleanup).toHaveBeenCalledOnce();
  });
});
