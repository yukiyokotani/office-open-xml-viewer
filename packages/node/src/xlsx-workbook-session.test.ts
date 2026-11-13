import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterAll, beforeAll, describe, expect, it, vi } from 'vitest';
import type { Row } from '@silurus/ooxml-xlsx';
// @ts-ignore — wasm-pack generated JavaScript is local build output.
import * as xlsxWasm from '../../xlsx/src/wasm/xlsx_parser.js';
import { generateSyntheticXlsx } from '../scripts/generate-synthetic-xlsx.mjs';
import { OoxmlError } from '@silurus/ooxml-core';
import { nonOoxmlInputs } from '@silurus/ooxml-core/testing';
import { materializeXlsxWorkbook, openXlsxWorkbook } from './xlsx.ts';
import { storedZip } from './test-ooxml-package.ts';

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
  it.each(nonOoxmlInputs('xlsx'))('rejects %s with OoxmlError not-ooxml', async (_name, input) => {
    for (const open of [() => openXlsxWorkbook(input), () => materializeXlsxWorkbook(input)]) {
      const error = await open().then(
        () => { throw new Error('resolved a non-OOXML workbook'); },
        (rejection: unknown) => rejection,
      );
      expect(error).toBeInstanceOf(OoxmlError);
      expect(error).toMatchObject({ code: 'not-ooxml' });
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

  it('extracts picture parts through the session-owned bounded cache and rejects after close', async () => {
    const extract = vi.spyOn(archivePrototype(), 'extract_image');
    try {
      const session = await openXlsxWorkbook(pictureWorkbook());
      const first = await session.getImage('xl/media/image1.png', 'image/png');
      expect(await session.getImage('xl/media/image1.png', 'image/png')).toBe(first);
      expect(first.type).toBe('image/png');
      expect(new Uint8Array(await first.arrayBuffer())).toEqual(PICTURE);
      expect((await session.getImage('xl/media/image1.png')).type).toBe('application/octet-stream');
      expect(extract).toHaveBeenCalledTimes(1);
      expect(session.resourceUsage?.distinctInflatedBytes).toBeGreaterThanOrEqual(PICTURE.byteLength);
      await session.close();
      await expect(session.getImage('xl/media/image1.png', 'image/png')).rejects.toThrow(/closed/);
      expect(extract).toHaveBeenCalledTimes(1);
    } finally {
      extract.mockRestore();
    }
  });

  it('rejects a non-picture or missing path without failing the session', async () => {
    const onResourceMetrics = vi.fn();
    const picture = new Uint8Array(65_536);
    const session = await openXlsxWorkbook(pictureWorkbook(picture), { onResourceMetrics });
    await session.getImage('xl/media/image1.png', 'image/png');
    // XML parts are not pictures, whatever their path spelling.
    await expect(session.getImage('xl/workbook.xml')).rejects.toThrow(TypeError);
    await expect(session.getImage('/XL/Workbook.xml')).rejects.toThrow(TypeError);
    await expect(session.getImage('xl/media/missing.png', 'image/png')).rejects.toBeDefined();
    expect(onResourceMetrics).not.toHaveBeenCalled();
    expect((await session.getImage('xl/media/image1.png', 'image/png')).size).toBe(picture.byteLength);
    expect(await session.worksheetRows(0).next()).toBeDefined();
    await session.close();
    // The session closes normally and keeps the earlier picture read's counters.
    expect(onResourceMetrics).toHaveBeenCalledTimes(1);
    expect(onResourceMetrics).toHaveBeenCalledWith(expect.objectContaining({
      status: 'ok',
      usage: expect.objectContaining({
        distinctInflatedBytes: expect.toSatisfy((bytes: number) => bytes >= picture.byteLength),
      }),
    }));
  });

  it('fails the session when a picture read exceeds the archive entry budget', async () => {
    const session = await openXlsxWorkbook(pictureWorkbook(new Uint8Array(4096)), {
      resourceLimits: { maxArchiveEntryBytes: 2048 },
    });
    try {
      const failure = await session.getImage('xl/media/image1.png', 'image/png').then(
        () => { throw new Error('resolved an over-budget picture'); },
        (error: unknown) => error,
      );
      expect(failure).toMatchObject({ name: 'OoxmlResourceLimitError', code: 'ooxml-resource-limit' });
      await expect(session.getImage('xl/media/image1.png', 'image/png')).rejects.toBe(failure);
      await expect(session.worksheetRows(0).next()).rejects.toBe(failure);
    } finally {
      await session.close();
    }
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
  extract_image(path: string): Uint8Array;
};

/** PNG signature plus a few bytes; extraction never decodes it. */
const PICTURE = Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3, 4]);

/** A one-sheet workbook that carries one media part. */
function pictureWorkbook(picture: Uint8Array = PICTURE): Uint8Array {
  const ns = 'http://schemas.openxmlformats.org';
  return storedZip({
    '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="${ns}/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`,
    '_rels/.rels': `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${ns}/package/2006/relationships"><Relationship Id="rId1" Type="${ns}/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    'xl/workbook.xml': `<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="${ns}/spreadsheetml/2006/main" xmlns:r="${ns}/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${ns}/package/2006/relationships"><Relationship Id="rId1" Type="${ns}/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${ns}/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
    'xl/styles.xml': `<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="${ns}/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellXfs count="1"><xf/></cellXfs></styleSheet>`,
    'xl/worksheets/sheet1.xml': `<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="${ns}/spreadsheetml/2006/main"><sheetData/></worksheet>`,
    'xl/media/image1.png': picture,
  });
}

function archivePrototype(): ArchivePrototype {
  return (xlsxWasm as unknown as { XlsxArchive: { prototype: ArchivePrototype } })
    .XlsxArchive.prototype;
}
