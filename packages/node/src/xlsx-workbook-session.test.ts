import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterAll, beforeAll, describe, expect, it, vi } from 'vitest';
import type { Row } from '@silurus/ooxml-xlsx';
// @ts-ignore — wasm-pack generated JavaScript is local build output.
import * as xlsxWasm from '../../xlsx/src/wasm/xlsx_parser.js';
import { generateSyntheticXlsx } from '../scripts/generate-synthetic-xlsx.mjs';
import { openXlsxWorkbook } from './xlsx.ts';

let directory = '';
let bytes: Buffer;

beforeAll(async () => {
  directory = await mkdtemp(join(tmpdir(), 'ooxml-xlsx-session-'));
  const fixture = join(directory, 'workbook.xlsx');
  await generateSyntheticXlsx(fixture, { rows: 257, columns: 8 });
  bytes = await readFile(fixture);
});

afterAll(async () => {
  if (directory) await rm(directory, { recursive: true, force: true });
});

describe('Node bounded XLSX workbook session', () => {
  it('keeps the degraded-container diagnostic when archive usage is unavailable', async () => {
    const workbook = await openXlsxWorkbook(new Uint8Array([0x50, 0x4b, 0x03, 0x04]));
    try {
      expect(workbook.resourceUsage).toBeUndefined();
      expect(workbook.workbookIndex.workbook.parseError)
        .toContain('(zip container): ZIP central directory preflight failed');

      let terminalError: string | undefined;
      for await (const chunk of workbook.worksheetRows(0)) {
        if (chunk.kind === 'finished') terminalError = chunk.worksheet.parseError;
      }
      expect(terminalError).toContain('(zip container): ZIP central directory preflight failed');
    } finally {
      await workbook.close();
    }
  });

  it('opens one workbook and streams worksheets sequentially from the retained archive', async () => {
    const free = vi.spyOn(archivePrototype(), 'free');
    try {
      const workbook = await openXlsxWorkbook(bytes);
      expect(workbook.sheetCount).toBe(1);
      expect(workbook.sheetNames).toEqual(['Synthetic']);

      const first = await collectRows(workbook.worksheetRows(0));
      const second = await collectRows(workbook.worksheetRows(0));
      expect(first).toHaveLength(257);
      expect(second).toEqual(first);
      expect(free).not.toHaveBeenCalled();

      await workbook.close();
      await workbook.close();
      expect(free).toHaveBeenCalledOnce();
    } finally {
      free.mockRestore();
    }
  });

  it('rejects concurrent worksheet streams instead of silently queueing ownership', async () => {
    const workbook = await openXlsxWorkbook(bytes);
    try {
      const first = workbook.worksheetRows(0);
      expect((await first.next()).value?.kind).toBe('rows');
      await expect(workbook.worksheetRows(0).next()).rejects.toThrow(/already active/);
      await first.return();
      await expect(collectRows(workbook.worksheetRows(0))).resolves.toHaveLength(257);
    } finally {
      await workbook.close();
    }
  });

  it('keeps invalid indexes local to the operation and rejects use after close', async () => {
    const workbook = await openXlsxWorkbook(bytes);
    await expect(workbook.worksheetRows(1).next()).rejects.toThrow(/out of range/);
    await expect(collectRows(workbook.worksheetRows(0))).resolves.toHaveLength(257);
    await workbook.close();
    await expect(workbook.worksheetRows(0).next()).rejects.toThrow(/closed/);
  });

  it('cancels an active worksheet stream when workbook ownership closes', async () => {
    const free = vi.spyOn(archivePrototype(), 'free');
    try {
      const workbook = await openXlsxWorkbook(bytes);
      const rows = workbook.worksheetRows(0);
      expect((await rows.next()).value?.kind).toBe('rows');

      await workbook.close();
      expect(free).toHaveBeenCalledOnce();
      await expect(rows.next()).rejects.toThrow(/closed/);
      await workbook.close();
      expect(free).toHaveBeenCalledOnce();
    } finally {
      free.mockRestore();
    }
  });

  it('reports one workbook-scoped metrics result when ownership closes', async () => {
    const onResourceMetrics = vi.fn();
    const workbook = await openXlsxWorkbook(bytes, { onResourceMetrics });
    await collectRows(workbook.worksheetRows(0));
    expect(onResourceMetrics).not.toHaveBeenCalled();
    await workbook.close();
    expect(onResourceMetrics).toHaveBeenCalledOnce();
    expect(onResourceMetrics).toHaveBeenCalledWith(expect.objectContaining({
      format: 'xlsx',
      scope: 'session',
      status: 'ok',
      outcome: expect.objectContaining({ worksheets: 1, rows: 257 }),
    }));
  });

  it('applies the configurable archive entry-count policy before workbook parsing', async () => {
    await expect(openXlsxWorkbook(bytes, {
      resourceLimits: { maxArchiveEntries: 1 },
    })).rejects.toMatchObject({
      name: 'OoxmlResourceLimitError',
      code: 'ooxml-resource-limit',
      details: expect.objectContaining({
        violation: expect.objectContaining({
          metric: 'entry-count',
          configurable: true,
          limit: 1,
        }),
      }),
    });
  });

  it('invalidates sibling sessions after a trap and recovers on a fresh generation', async () => {
    const openCursor = vi.spyOn(archivePrototype(), 'open_sheet_cursor')
      .mockImplementationOnce(() => { throw new RangeError('synthetic trap'); });
    const first = await openXlsxWorkbook(bytes);
    const sibling = await openXlsxWorkbook(bytes);
    try {
      await expect(first.worksheetRows(0).next())
        .rejects.toMatchObject({ name: 'WasmTrapError', code: 'parser-crashed' });
      await expect(sibling.worksheetRows(0).next())
        .rejects.toMatchObject({ name: 'WasmTrapError', code: 'parser-crashed' });
    } finally {
      await first.close().catch(() => undefined);
      await sibling.close().catch(() => undefined);
      openCursor.mockRestore();
    }

    const recovered = await openXlsxWorkbook(bytes);
    await expect(collectRows(recovered.worksheetRows(0))).resolves.toHaveLength(257);
    await recovered.close();
  });
});

async function collectRows(
  chunks: AsyncIterable<{ kind: string; rows?: Row[] }>,
): Promise<Row[]> {
  const rows: Row[] = [];
  for await (const chunk of chunks) {
    if (chunk.kind === 'rows' && chunk.rows) rows.push(...chunk.rows);
  }
  return rows;
}

type ArchivePrototype = {
  free(): void;
  open_sheet_cursor(sheetIndex: number, name: string): void;
};

function archivePrototype(): ArchivePrototype {
  return (xlsxWasm as unknown as { XlsxArchive: { prototype: ArchivePrototype } })
    .XlsxArchive.prototype;
}
