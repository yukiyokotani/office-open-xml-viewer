import { describe, expect, it, vi } from 'vitest';
import type { LegacyXlsDirectSourceDescriptor } from '@silurus/ooxml-core/internal/legacy-xls-source';
import { buildXlsFixture } from '../../legacy-converter/src/test-fixtures.ts';
import { openXlsxWorkbook, type XlsxWorkbookSession } from './xlsx.ts';

const { loadOoxml } = vi.hoisted(() => ({
  loadOoxml: vi.fn(() => { throw new Error('OOXML WASM must not load for direct XLS'); }),
}));
vi.mock('./wasm-loader.ts', async (importOriginal) => ({
  ...await importOriginal<typeof import('./wasm-loader.ts')>(),
  createLazyWasmModule: () => loadOoxml,
}));

const source: LegacyXlsDirectSourceDescriptor = {
  protocol: 'ooxml-legacy-xls-source/v1', builtin: 'xls',
  wasmUrl: new URL('../../legacy-converter/src/wasm-direct-xls/legacy_office_converter_bg.wasm', import.meta.url).href,
};

function fixture() {
  const bytes = buildXlsFixture();
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const sectorSize = 2 ** view.getUint16(30, true);
  const directoryOffset = (view.getUint32(48, true) + 1) * sectorSize;
  // This legacy fixture builder leaves streams unlinked. The direct reader
  // requires root ownership: its sole Workbook entry is directory ID 1.
  view.setUint32(directoryOffset + 76, 1, true);
  return bytes;
}

async function rows(workbook: XlsxWorkbookSession) {
  const result = [];
  for await (const chunk of workbook.worksheetRows(0)) {
    if (chunk.kind === 'rows') result.push(...chunk.rows);
  }
  return result;
}

describe('Node direct XLS native session', () => {
  it('streams real BIFF through the XLSX row contract without loading OOXML WASM', async () => {
    const workbook = await openXlsxWorkbook(fixture(), { legacyConversion: { xls: { source } } });
    try {
      expect(workbook.sheetNames).toEqual(['表計算']);
      const first = await rows(workbook);
      expect(first.flatMap(row => row.cells).map(cell => cell.value)).toEqual([
        { type: 'number', number: 42.5 }, { type: 'text', text: '日本語' },
      ]);
      expect(await rows(workbook)).toEqual(first);
      expect(workbook.resourceUsage).toBeUndefined();
      expect(loadOoxml).not.toHaveBeenCalled();
    } finally { await workbook.close(); }
    await workbook.close();
    await expect(workbook.worksheetRows(0).next()).rejects.toThrow(/closed/);
  });

  it('allows a new native cursor after the consumer breaks early', async () => {
    const workbook = await openXlsxWorkbook(fixture(), { legacyConversion: { xls: { source } } });
    try {
      for await (const chunk of workbook.worksheetRows(0)) {
        expect(chunk.kind).toBe('rows');
        break;
      }
      expect(await rows(workbook)).toHaveLength(2);
    } finally { await workbook.close(); }
  });

  it('keeps the source cancellation signal effective after opening', async () => {
    const abort = new AbortController();
    const workbook = await openXlsxWorkbook(fixture(), {
      legacyConversion: { xls: { source, signal: abort.signal } },
    });
    abort.abort();
    await expect(rows(workbook)).rejects.toMatchObject({ name: 'AbortError' });
    await workbook.close();
    expect(loadOoxml).not.toHaveBeenCalled();
  });
});
