// The Node XLSX workbook session over the direct XLS reader: native row
// streaming, host layout through the session canvas, cancellation and native
// ownership. Row streaming never acquires the OOXML parser WASM.
import { describe, expect, it, vi } from 'vitest';
import { buildXlsFixture } from '../test-fixtures.js';
import { buildXlsPicturesFixture } from '../xls-pictures-fixture.js';
import { testXlsSource } from '../test-sources.js';
import { GridGeometry, openXlsxWorkbook } from './node-facade.js';
import { digitWidthFactory } from './xls-host-layout-fixture.js';

const { loadOoxml } = vi.hoisted(() => ({
  loadOoxml: vi.fn(() => { throw new Error('OOXML WASM must not load for direct XLS'); }),
}));
vi.mock('../../../node/src/wasm-loader.ts', async (importOriginal) => ({
  ...await importOriginal<Record<string, unknown>>(),
  createLazyWasmModule: () => loadOoxml,
}));

type Session = Awaited<ReturnType<typeof openXlsxWorkbook>>;
const open = (bytes: Uint8Array, options: Omit<Parameters<typeof openXlsxWorkbook>[1] & object, 'modelSources'> = {}) =>
  openXlsxWorkbook(bytes, { ...options, modelSources: [testXlsSource()] });

async function rows(workbook: Session) {
  const result = [];
  for await (const chunk of workbook.worksheetRows(0)) {
    if (chunk.kind === 'rows') result.push(...chunk.rows);
  }
  return result;
}

/** Count native close/free calls on the generated glue for one operation. */
async function nativeReleases(operation: () => Promise<unknown>) {
  const { LegacyXlsWorkbook } = await import('../wasm-direct-xls/legacy_xls_direct.js');
  const free = vi.spyOn(LegacyXlsWorkbook.prototype, 'free');
  const close = vi.spyOn(LegacyXlsWorkbook.prototype, 'close_workbook_session');
  try {
    await operation();
    return { close: close.mock.calls.length, free: free.mock.calls.length };
  } finally { free.mockRestore(); close.mockRestore(); }
}

describe('Node direct XLS session', () => {
  it('streams BIFF rows through the XLSX row contract repeatably and closes idempotently', async () => {
    const workbook = await open(buildXlsFixture());
    try {
      expect(workbook.sheetNames).toEqual(['表計算']);
      const first = await rows(workbook);
      expect(first.flatMap(row => row.cells).map(cell => [cell.row, cell.col, cell.value])).toEqual([
        [1, 1, { type: 'number', number: 42.5 }], [2, 2, { type: 'text', text: '日本語' }],
      ]);
      expect(await rows(workbook)).toEqual(first);
      // The direct archive reports no OOXML resource-usage snapshot.
      expect(workbook.resourceUsage).toBeUndefined();
    } finally { await workbook.close(); }
    await workbook.close();
    await expect(workbook.worksheetRows(0).next()).rejects.toThrow(/closed/);
    expect(loadOoxml).not.toHaveBeenCalled();
  });

  it('allows a new native cursor after the consumer breaks early', async () => {
    const workbook = await open(buildXlsFixture());
    try {
      for await (const chunk of workbook.worksheetRows(0)) {
        expect(chunk.kind).toBe('rows');
        break;
      }
      expect(await rows(workbook)).toHaveLength(2);
    } finally { await workbook.close(); }
  });

  it('keeps the session cancellation signal effective after opening', async () => {
    const abort = new AbortController();
    const workbook = await open(buildXlsFixture(), { signal: abort.signal });
    abort.abort();
    await expect(rows(workbook)).rejects.toMatchObject({ name: 'AbortError' });
    await workbook.close();
    expect(loadOoxml).not.toHaveBeenCalled();
  });

  it.each([7, 9])('keeps measured width %i for native pictures and the worksheet grid geometry', async (width) => {
    const factory = digitWidthFactory(width);
    const workbook = await open(buildXlsPicturesFixture(), { factory });
    try {
      expect(workbook.workbookIndex.layoutMetrics).toEqual({ maximumDigitWidth: width });
      for (let repeat = 0; repeat < 2; repeat++) {
        for await (const chunk of workbook.worksheetRows(0)) {
          if (chunk.kind !== 'finished') continue;
          // Authored STANDARDWIDTH is 10 characters; the anchor is at 512/1024
          // of its first column ([MS-XLS] 2.5.193 OfficeArtClientAnchorSheet.dxL).
          expect(chunk.worksheet.images).toEqual([expect.objectContaining({ fromColOff: 10 * width / 2 * 9525 })]);
          const geometry = GridGeometry.forWorksheetMeasured(chunk.worksheet, () => {
            throw new Error('source geometry must not be remeasured');
          });
          expect(geometry.maximumDigitWidth).toBe(width);
        }
      }
      // One host measurement per session, of the Normal font only.
      expect(factory.fonts).toHaveLength(1);
      expect(loadOoxml).not.toHaveBeenCalled();
    } finally { await workbook.close(); }
  });

  it.each([
    ['the measurement throws', () => ({ factory: digitWidthFactory(() => { throw new Error('measurement failed'); }) }), 'measurement failed'],
    ['the session is aborted during measurement', () => {
      const abort = new AbortController();
      return { signal: abort.signal, factory: digitWidthFactory(() => { abort.abort(); return 7; }) };
    }, /aborted/i],
  ] as const)('releases the native owner exactly once when %s', async (_label, options, error) => {
    const released = await nativeReleases(async () => {
      await expect(open(buildXlsPicturesFixture(), options())).rejects.toThrow(error);
    });
    expect(released).toEqual({ close: 1, free: 1 });
  });
});
