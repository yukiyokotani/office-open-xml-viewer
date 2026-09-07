import { describe, expect, it, vi } from 'vitest';
import { XlsxWorkbook } from './workbook.js';

describe('XlsxWorkbook direct XLS signal ownership', () => {
  it('destroys a successfully loaded session when its retained source signal aborts', async () => {
    const controller = new AbortController();
    const workbook = Object.create(XlsxWorkbook.prototype) as XlsxWorkbook;
    const destroy = vi.spyOn(workbook, 'destroy').mockImplementation(() => undefined);
    const pending = Promise.resolve('loaded');
    const bound = (workbook as unknown as {
      bindLegacyXlsSignal<T>(pending: Promise<T>, signal: AbortSignal): Promise<T>;
    }).bindLegacyXlsSignal(pending, controller.signal);
    await expect(bound).resolves.toBe('loaded');
    controller.abort();
    expect(destroy).toHaveBeenCalledOnce();
  });

  it('runs retained cleanup immediately for an already-destroyed workbook and only once otherwise', () => {
    const destroyed = Object.create(XlsxWorkbook.prototype) as XlsxWorkbook;
    (destroyed as unknown as Record<string, unknown>).destroyed = true;
    const immediate = vi.fn();
    destroyed._retainLegacyXlsSignalCleanup(immediate);
    expect(immediate).toHaveBeenCalledOnce();

    const workbook = Object.create(XlsxWorkbook.prototype) as XlsxWorkbook;
    Object.assign(workbook as unknown as Record<string, unknown>, {
      sheetCache: new Map(),
      sheetLoads: new Map(),
      retainedFontSets: new Map(),
      googleFontNames: [],
      rawParts: { clear: vi.fn() },
      queuedImageLoads: new Map(),
      _fetchImage: vi.fn(),
    });
    const cleanup = vi.fn();
    workbook._retainLegacyXlsSignalCleanup(cleanup);
    workbook.destroy();
    workbook.destroy();
    expect(cleanup).toHaveBeenCalledOnce();
  });
});
