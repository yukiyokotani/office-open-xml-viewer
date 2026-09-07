import { describe, expect, it, vi } from 'vitest';
import { readFile, readdir } from 'node:fs/promises';
import { join } from 'node:path';
import type { LegacyXlsDirectSourceDescriptor } from '@silurus/ooxml-core/internal/legacy-xls-source';
import { buildXlsFixture } from '../../legacy-converter/src/test-fixtures.ts';
import { buildXlsPicturesFixture } from '../../legacy-converter/src/xls-pictures-fixture.ts';
import { GridGeometry } from '@silurus/ooxml-xlsx/internal/session';
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
  return linkSoleWorkbook(buildXlsFixture());
}

function linkSoleWorkbook(bytes: Uint8Array) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const sectorSize = 2 ** view.getUint16(30, true);
  const directoryOffset = (view.getUint32(48, true) + 1) * sectorSize;
  // This legacy fixture builder leaves streams unlinked. The direct reader
  // requires root ownership: its sole Workbook entry is directory ID 1.
  view.setUint32(directoryOffset + 76, 1, true);
  for (const entry of [0, 1]) {
    view.setUint32(directoryOffset + entry * 128 + 68, 0xffffffff, true);
    view.setUint32(directoryOffset + entry * 128 + 72, 0xffffffff, true);
  }
  view.setUint32(directoryOffset + 128 + 76, 0xffffffff, true);
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
  it('releases the native owner when the measurement provider rejects', async () => {
    const { LegacyXlsWorkbook } = await import('../../legacy-converter/src/wasm-direct-xls/legacy_office_converter.js');
    const free = vi.spyOn(LegacyXlsWorkbook.prototype, 'free');
    const close = vi.spyOn(LegacyXlsWorkbook.prototype, 'close_workbook_session');
    try {
      await expect(openXlsxWorkbook(linkSoleWorkbook(buildXlsPicturesFixture()), {
        legacyConversion: { xls: { source } },
        measureLegacyXlsNormalFont: () => { throw new Error('measurement failed'); },
      })).rejects.toThrow('measurement failed');
      expect(close).toHaveBeenCalledOnce();
      expect(free).toHaveBeenCalledOnce();
    } finally { free.mockRestore(); close.mockRestore(); }
  });

  it.each([7, 9])('retains measured width %i for native pictures and renderer geometry', async (width) => {
    const measure = vi.fn(() => width);
    const workbook = await openXlsxWorkbook(linkSoleWorkbook(buildXlsPicturesFixture()), {
      legacyConversion: { xls: { source } }, measureLegacyXlsNormalFont: measure,
    });
    try {
      expect(workbook.maximumDigitWidth).toBe(width);
      for (let repeat = 0; repeat < 2; repeat++) {
        for await (const chunk of workbook.worksheetRows(0)) {
          if (chunk.kind !== 'finished') continue;
          expect(chunk.worksheet.images).toHaveLength(1);
          // Authored STANDARDWIDTH is 10 characters; the anchor is at
          // 512/1024 of its first column ([MS-XLS] 2.5.193,
          // OfficeArtClientAnchorSheet.dxL).
          expect(chunk.worksheet.images?.[0].fromColOff).toBe(10 * width / 2 * 9525);
          const geometry = GridGeometry.forWorksheetMeasured(chunk.worksheet, () => {
            throw new Error('source geometry must not be remeasured');
          });
          expect(geometry.maximumDigitWidth).toBe(width);
        }
      }
      expect(measure).toHaveBeenCalledOnce();
      expect(measure).toHaveBeenCalledWith({ family: 'Arial', sizePoints: 11, bold: false, italic: false }, expect.any(AbortSignal));
      expect(loadOoxml).not.toHaveBeenCalled();
    } finally { await workbook.close(); }
  });

  it('does not invent font metrics when no provider is available', async () => {
    const workbook = await openXlsxWorkbook(linkSoleWorkbook(buildXlsPicturesFixture()), { legacyConversion: { xls: { source } } });
    try {
      expect(workbook.maximumDigitWidth).toBeUndefined();
      for await (const chunk of workbook.worksheetRows(0)) {
        if (chunk.kind === 'finished') expect(chunk.worksheet.images ?? []).toHaveLength(0);
      }
    } finally { await workbook.close(); }
  });

  it.skipIf(!process.env.LEGACY_XLS_SESSION_CORPUS)('streams every local corpus workbook without OOXML acquisition', async () => {
    const directory = process.env.LEGACY_XLS_SESSION_CORPUS as string;
    const files = (await readdir(directory)).filter(name => name.toLowerCase().endsWith('.xls')).sort();
    expect(files.length).toBeGreaterThan(0);
    let sheetCount = 0;
    let rowCount = 0;
    for (const file of files) {
      const workbook = await openXlsxWorkbook(await readFile(join(directory, file)), {
        legacyConversion: { xls: { source } },
      });
      try {
        for (let index = workbook.sheetCount - 1; index >= 0; index--) {
          let finished = false;
          for await (const chunk of workbook.worksheetRows(index)) {
            if (chunk.kind === 'rows') rowCount += chunk.rows.length;
            else finished = true;
          }
          expect(finished).toBe(true);
          sheetCount++;
        }
      } finally { await workbook.close(); }
    }
    expect(loadOoxml).not.toHaveBeenCalled();
    // Aggregate counts only: never publish local corpus filenames or contents.
    console.info({ workbooks: files.length, sheets: sheetCount, rows: rowCount });
  }, 120_000);

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
