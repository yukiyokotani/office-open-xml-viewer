import { describe, expect, it } from 'vitest';
import {
  parseDelimitedWorksheet,
  resolveDelimitedTextOptions,
} from './delimited-text.js';
import { MAX_WORKSHEET_COL } from './internal/grid-geometry.js';
import { XLSX_MAX_MATERIALIZED_ROWS } from './worksheet-resource-limits.js';

const encode = (value: string): ArrayBuffer => new TextEncoder().encode(value).buffer;

describe('delimited worksheet preview input', () => {
  it('parses quoted CSV fields without inferring cell types', () => {
    const { workbook, worksheet } = parseDelimitedWorksheet(
      encode('id,description\r\n00123,"first line\nsecond, line"\r\n"""quoted""",=1+1\r\n'),
      resolveDelimitedTextOptions({ format: 'csv' }),
    );

    expect(workbook.workbook.sheets).toEqual([
      { name: 'Sheet1', sheetId: 1, rId: 'rId1' },
    ]);
    expect(worksheet.rows).toEqual([
      {
        index: 1,
        height: null,
        cells: [
          { row: 1, col: 1, value: { type: 'text', text: 'id' } },
          { row: 1, col: 2, value: { type: 'text', text: 'description' } },
        ],
      },
      {
        index: 2,
        height: null,
        cells: [
          { row: 2, col: 1, value: { type: 'text', text: '00123' } },
          { row: 2, col: 2, value: { type: 'text', text: 'first line\nsecond, line' } },
        ],
      },
      {
        index: 3,
        height: null,
        cells: [
          { row: 3, col: 1, value: { type: 'text', text: '"quoted"' } },
          { row: 3, col: 2, value: { type: 'text', text: '=1+1' } },
        ],
      },
    ]);
    expect(workbook.styles.fonts[0]).toMatchObject({ name: 'Calibri', size: 11 });
    expect(worksheet).toMatchObject({
      name: 'Sheet1',
      defaultColWidth: 8.43,
      defaultRowHeight: 15,
      defaultFontFamily: 'Calibri',
      defaultFontSize: 11,
    });
  });

  it('uses TSV and generic delimited-text separators', () => {
    const tsv = parseDelimitedWorksheet(
      encode('left\tright'),
      resolveDelimitedTextOptions({ format: 'tsv', sheetName: 'Data' }),
    );
    expect(tsv.worksheet.name).toBe('Data');
    expect(tsv.worksheet.rows[0]?.cells.map((cell) => cell.value)).toEqual([
      { type: 'text', text: 'left' },
      { type: 'text', text: 'right' },
    ]);

    const semicolon = parseDelimitedWorksheet(
      encode('left;right'),
      resolveDelimitedTextOptions({ format: 'delimited-text', delimiter: ';' }),
    );
    expect(semicolon.worksheet.rows[0]?.cells).toHaveLength(2);
  });

  it('keeps interior blank rows and fields but ignores one terminal record break', () => {
    const { worksheet } = parseDelimitedWorksheet(
      encode('a,,c\n\nlast,\n'),
      resolveDelimitedTextOptions({ format: 'csv' }),
    );

    expect(worksheet.rows).toHaveLength(3);
    expect(worksheet.rows[0]?.cells.map((cell) => cell.col)).toEqual([1, 3]);
    expect(worksheet.rows[1]).toEqual({ index: 2, height: null, cells: [] });
    expect(worksheet.rows[2]?.cells).toEqual([
      { row: 3, col: 1, value: { type: 'text', text: 'last' } },
    ]);
  });

  it('normalizes CR and CRLF line breaks inside quoted fields to renderer hard breaks', () => {
    const { worksheet } = parseDelimitedWorksheet(
      encode('id,notes\r\n1,"first\rsecond"\r\n2,"third\r\nfourth"\r\n'),
      resolveDelimitedTextOptions({ format: 'csv' }),
    );

    expect(worksheet.rows[1]?.cells[1]?.value).toEqual({
      type: 'text',
      text: 'first\nsecond',
    });
    expect(worksheet.rows[2]?.cells[1]?.value).toEqual({
      type: 'text',
      text: 'third\nfourth',
    });
  });

  it('supports an explicit browser TextDecoder encoding', () => {
    const bytes = new Uint8Array([0x63, 0x61, 0x66, 0xe9]).buffer;
    const { worksheet } = parseDelimitedWorksheet(
      bytes,
      resolveDelimitedTextOptions({ format: 'csv', encoding: 'windows-1252' }),
    );
    expect(worksheet.rows[0]?.cells[0]?.value).toEqual({ type: 'text', text: 'café' });
  });

  it('keeps large individual fields intact', () => {
    const value = 'x'.repeat(10_000);
    const { worksheet } = parseDelimitedWorksheet(
      encode(`"${value}"|tail`),
      resolveDelimitedTextOptions({ format: 'delimited-text', delimiter: '|' }),
    );
    expect(worksheet.rows[0]?.cells.map((cell) => cell.value)).toEqual([
      { type: 'text', text: value },
      { type: 'text', text: 'tail' },
    ]);
  });

  it('rejects ambiguous or malformed input explicitly', () => {
    expect(() => resolveDelimitedTextOptions({
      format: 'delimited-text',
    } as unknown as Parameters<typeof resolveDelimitedTextOptions>[0]))
      .toThrow("delimiter is required for format 'delimited-text'");
    expect(() => resolveDelimitedTextOptions({ format: 'csv', delimiter: '||' }))
      .toThrow('delimiter must be exactly one character');
    expect(() => resolveDelimitedTextOptions({ format: 'csv', delimiter: '"' }))
      .toThrow('delimiter cannot be a quote or record separator');
    expect(() => parseDelimitedWorksheet(
      encode('a,"unterminated'),
      resolveDelimitedTextOptions({ format: 'csv' }),
    )).toThrow('Unterminated quoted field');
  });

  it('enforces worksheet row and column bounds while parsing', () => {
    const csv = resolveDelimitedTextOptions({ format: 'csv' });
    expect(() => parseDelimitedWorksheet(
      encode('\n'.repeat(XLSX_MAX_MATERIALIZED_ROWS + 1)),
      csv,
    )).toThrow(`rows ${XLSX_MAX_MATERIALIZED_ROWS + 1} > ${XLSX_MAX_MATERIALIZED_ROWS}`);
    expect(() => parseDelimitedWorksheet(encode(','.repeat(MAX_WORKSHEET_COL)), csv))
      .toThrow(`more than ${MAX_WORKSHEET_COL} columns`);
  });
});
