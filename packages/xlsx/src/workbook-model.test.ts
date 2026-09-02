import { describe, expect, it, vi } from 'vitest';
import type { ParsedWorkbook, Styles, Worksheet } from './types.js';

const { renderWorksheetViewport } = vi.hoisted(() => ({
  renderWorksheetViewport: vi.fn(async () => undefined),
}));

vi.mock('./render-orchestrator.js', () => ({ renderWorksheetViewport }));

import { XlsxWorkbook } from './workbook.js';

const styles: Styles = {
  fonts: [{
    bold: false,
    italic: false,
    underline: false,
    strike: false,
    size: 11,
    color: '#000000',
    name: 'Arial',
  }],
  fills: [{ patternType: 'none', fgColor: null, bgColor: null }],
  borders: [{ left: null, right: null, top: null, bottom: null }],
  cellXfs: [{
    fontId: 0,
    fillId: 0,
    borderId: 0,
    numFmtId: 0,
    alignH: null,
    alignV: null,
    wrapText: false,
  }],
  numFmts: [],
  dxfs: [],
};

function worksheet(name = 'CSV'): Worksheet {
  return {
    name,
    rows: [{
      index: 1,
      height: null,
      cells: [{ row: 1, col: 1, value: { type: 'text', text: '00123' } }],
    }],
    colWidths: {},
    rowHeights: {},
    defaultColWidth: 8.43,
    defaultRowHeight: 15,
    mergeCells: [],
    freezeRows: 0,
    freezeCols: 0,
    conditionalFormats: [],
    images: [],
    charts: [],
  };
}

function model(name = 'CSV'): ParsedWorkbook {
  return {
    workbook: {
      sheets: [{ name, sheetId: 1, rId: 'rId1' }],
    },
    styles,
    sharedStrings: [],
  };
}

describe('XlsxWorkbook.fromModel', () => {
  it('installs worksheet models without loading an OOXML archive', async () => {
    const sheet = worksheet();
    const workbook = XlsxWorkbook.fromModel(model(), [sheet]);

    expect(workbook.mode).toBe('main');
    expect(workbook.sheetNames).toEqual(['CSV']);
    expect(await workbook.getWorksheet(0)).toBe(sheet);
    expect(workbook.cellText(sheet, sheet.rows[0]!.cells[0]!)).toBe('00123');
    const canvas = {} as HTMLCanvasElement;
    const viewport = { row: 1, col: 1, rows: 10, cols: 10 };
    await workbook.renderViewport(canvas, 0, viewport);
    expect(renderWorksheetViewport).toHaveBeenCalledWith(
      expect.objectContaining({ ws: sheet, styles }),
      canvas,
      viewport,
      expect.objectContaining({ fetchImage: expect.any(Function) }),
    );
    await expect(workbook.getResourceMetrics()).resolves.toMatchObject({
      format: 'xlsx',
      mode: 'main',
      status: 'ok',
      outcome: { sheets: 1 },
    });

    await expect(workbook.toMarkdown()).rejects.toThrow(
      'This operation requires an active archive-backed workbook',
    );
    await expect(workbook.getImage('xl/media/image1.png', 'image/png')).rejects.toThrow(
      'This operation requires an active archive-backed workbook',
    );

    expect(() => workbook.destroy()).not.toThrow();
  });

  it('rejects worksheet metadata that does not match the supplied models', () => {
    expect(() => XlsxWorkbook.fromModel(model(), [])).toThrow(
      'Workbook metadata has 1 sheet but 0 worksheet models were supplied',
    );
    expect(() => XlsxWorkbook.fromModel(model(), [worksheet('Other')])).toThrow(
      'Worksheet 0 is named "Other" but workbook metadata names it "CSV"',
    );
  });

  it('resolves shared strings before retaining and rendering worksheet models', async () => {
    const sheet = worksheet();
    sheet.rows[0]!.cells[0]!.value = { type: 'shared', si: 0 };
    const parsed = model();
    parsed.sharedStrings = [{ text: 'resolved' }];

    const workbook = XlsxWorkbook.fromModel(parsed, [sheet]);

    expect(workbook.cellText(sheet, sheet.rows[0]!.cells[0]!)).toBe('resolved');
    expect(sheet.rows[0]!.cells[0]!.value).toEqual({ type: 'text', text: 'resolved' });
  });
});
