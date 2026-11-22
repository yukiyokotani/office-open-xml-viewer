/// <reference types="node" />

import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import type { DocxDocumentModel } from './types.js';
import { normalizeInternalDocumentModel } from './parser-model.js';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { storeZip } from './conformance/generate.js';

// WORD_TABLE_CELL_SPACING_GEOMETRY for an AutoFit table with preferred
// widths (§17.4.45 tblCellSpacing, §17.4.71 tcW, §17.4.63 tblW). Word's
// controls: A4, margins left 40 / right 90 pt (465.3 pt of text width), tables
// with tcW 72 + 108 pt (or 60 + 72 + 108 pt) and single 0.5 pt borders between
// two paragraphs.
// - tblW auto: the cell boxes keep their preferred widths and the gaps widen
//   the table (192 pt at 2 pt spacing, 210 pt at 5 pt), up to the text width;
// - a preferred table width, dxa or pct, or a table that would pass the text
//   width: the gaps are carved out of the cells proportionally, as for a
//   fixed table.
// The same holds for the upright block table of a tbRl section.

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const encoder = new TextEncoder();
const xml = (body: string) => encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${body}`);
const paragraph = (text: string) => `<w:p><w:r><w:rPr><w:sz w:val="21"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;

const BORDERS = '<w:tblBorders>'
  + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
    .map((edge) => `<w:${edge} w:val="single" w:sz="4"/>`).join('')
  + '</w:tblBorders>';
const spacing = (pt: number) => (pt > 0 ? `<w:tblCellSpacing w:w="${pt * 20}" w:type="dxa"/>` : '');
const cell = (width: number, text: string, properties = '') =>
  `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${properties}</w:tcPr>${paragraph(text)}</w:tc>`;

/** A table of rows of cells with these dxa grid widths. */
function tableXml(tableProperties: string, grid: readonly number[], rows: readonly string[]): string {
  return `<w:tbl><w:tblPr>${tableProperties}${BORDERS}</w:tblPr><w:tblGrid>`
    + grid.map((width) => `<w:gridCol w:w="${width}"/>`).join('')
    + `</w:tblGrid>${rows.map((row) => `<w:tr>${row}</w:tr>`).join('')}</w:tbl>`;
}

/** The two-column 72 + 108 pt table of the controls. */
const twoColumns = (tableProperties: string, widths: readonly [number, number] = [1440, 2160]) =>
  tableXml(tableProperties, widths, ['1', '2'].map((row) =>
    cell(widths[0], `A${row}`) + cell(widths[1], `B${row}`)));

function docx(table: string, direction: string | null, columns = 1, mode: number | null = null): Uint8Array {
  const document = `<w:document xmlns:w="${W}"><w:body>${paragraph('before')}${table}${paragraph('after')}`
    + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
    + '<w:pgMar w:top="1200" w:right="1800" w:bottom="1600" w:left="800" w:header="720" w:footer="720" w:gutter="0"/>'
    + (columns > 1 ? `<w:cols w:num="${columns}" w:space="400"/>` : '')
    + (direction ? `<w:textDirection w:val="${direction}"/>` : '')
    + '</w:sectPr></w:body></w:document>';
  return storeZip(new Map([
    ['[Content_Types].xml', xml('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
      + '</Types>')],
    ['word/_rels/document.xml.rels', xml('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + `<Relationship Id="rId1" Type="${REL}/settings" Target="settings.xml"/></Relationships>`)],
    ['word/settings.xml', xml(`<w:settings xmlns:w="${W}"><w:compat>`
      + (mode === null ? '' : `<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${mode}"/>`)
      + '</w:compat></w:settings>')],
    ['_rels/.rels', xml('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + `<Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`)],
    ['word/document.xml', xml(document)],
  ]));
}

function measureContext(): CanvasRenderingContext2D {
  let font = '10px serif';
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    fontKerning: 'normal',
    measureText: (text: string) => {
      const px = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      return {
        width: [...text].length * px,
        fontBoundingBoxAscent: px * 0.8,
        fontBoundingBoxDescent: px * 0.2,
        actualBoundingBoxAscent: px * 0.8,
        actualBoundingBoxDescent: px * 0.2,
      } as TextMetrics;
    },
  } as unknown as CanvasRenderingContext2D;
}

/** The laid-out table: its left edge, overall width, the widths of the first
 * row's cell boxes and its horizontal rules as "y:x0-x1". */
function tableOf(table: string, direction: string | null = null, columns = 1) {
  const archive = new DocxArchive(docx(table, direction, columns));
  let model: DocxDocumentModel;
  try {
    model = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document;
  } finally {
    archive.free();
  }
  const layout = layoutDocument(
    model,
    createLayoutServices(model, { measureContext: measureContext() }),
    { currentDateMs: 0 },
  );
  const laidOut = layout.pages[0]!.layers.body.find((node) => node.kind === 'table');
  if (laidOut?.kind !== 'table') throw new Error('expected a table');
  return {
    left: laidOut.flowBounds.xPt,
    table: laidOut.columnWidthsPt.reduce((sum, width) => sum + width, 0),
    cells: laidOut.rows[0]!.cells.map((box) => box.flowBounds.widthPt),
    cellHeights: laidOut.rows.map((row) => row.cells.map((box) => box.flowBounds.heightPt)),
    horizontalRules: laidOut.borders.filter((border) => border.from.yPt === border.to.yPt),
  };
}

/** Cell boxes of the table nested in the first cell of `table`: their widths
 * and the first one's offset from the outer cell's content box. */
function nestedCells(table: string, mode: number | null = 15): { widths: number[]; firstCellXPt: number } {
  const archive = new DocxArchive(docx(table, null, 1, mode));
  let model: DocxDocumentModel;
  try {
    model = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document;
  } finally {
    archive.free();
  }
  const layout = layoutDocument(
    model,
    createLayoutServices(model, { measureContext: measureContext() }),
    { currentDateMs: 0 },
  );
  const outer = layout.pages[0]!.layers.body.find((node) => node.kind === 'table');
  if (outer?.kind !== 'table') throw new Error('expected a table');
  const inner = outer.rows[0]!.cells[0]!.blocks.map((block) => block.layout)
    .find((block) => block.kind === 'table');
  if (inner?.kind !== 'table') throw new Error('expected a nested table');
  return {
    widths: inner.rows[0]!.cells.map((box) => Math.round(box.flowBounds.widthPt * 1e6) / 1e6),
    firstCellXPt: inner.rows[0]!.cells[0]!.flowBounds.xPt,
  };
}

/** The table fragments of every page, in page order. */
function tablePages(table: string, mode: number | null = null) {
  const archive = new DocxArchive(docx(table, null, 1, mode));
  let model: DocxDocumentModel;
  try {
    model = normalizeInternalDocumentModel(JSON.parse(new TextDecoder().decode(archive.parse()))).document;
  } finally {
    archive.free();
  }
  const layout = layoutDocument(
    model,
    createLayoutServices(model, { measureContext: measureContext() }),
    { currentDateMs: 0 },
  );
  return layout.pages.flatMap((page) => page.layers.body.flatMap((node) => node.kind === 'table' ? [node] : []));
}

const tableWidths = (spacingPt: number, tableWidth: string, direction: string | null = null) => {
  const { table, cells } = tableOf(twoColumns(tableWidth + spacing(spacingPt)), direction);
  return { table, cells };
};

const AUTO = '<w:tblW w:w="0" w:type="auto"/>';
const DXA_180 = '<w:tblW w:w="3600" w:type="dxa"/>';

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});

describe('cell spacing in an AutoFit table with preferred cell widths', () => {
  it.each([
    [2, 192],
    [5, 210],
  ])('adds %d pt spacing outside the cell boxes when the table width is automatic', (spacingPt, widthPt) => {
    const widths = tableWidths(spacingPt, AUTO);
    expect(widths.table).toBeCloseTo(widthPt, 6);
    expect(widths.cells[0]).toBeCloseTo(72, 6);
    expect(widths.cells[1]).toBeCloseTo(108, 6);
  });

  it('carves the spacing out of the cells proportionally under a preferred table width', () => {
    const widths = tableWidths(2, DXA_180);
    expect(widths.table).toBeCloseTo(180, 6);
    expect(widths.cells[0]).toBeCloseTo(67.2, 6);
    expect(widths.cells[1]).toBeCloseTo(100.8, 6);
  });

  it('leaves a table without spacing at its grid', () => {
    expect(tableWidths(0, AUTO)).toEqual({ table: 180, cells: [72, 108] });
  });

  it('widens a three-column table by a gap at every boundary', () => {
    const three = tableXml(AUTO + spacing(2), [1200, 1440, 2160], ['1', '2'].map((row) =>
      cell(1200, `A${row}`) + cell(1440, `B${row}`) + cell(2160, `C${row}`)));
    const widths = tableOf(three);
    expect(widths.table).toBeCloseTo(240 + 4 * 4, 6);
    expect(widths.cells.map((width) => Math.round(width * 1e6) / 1e6)).toEqual([60, 72, 108]);
  });

  it('carves the spacing out of a percentage table width', () => {
    // 50% of the 465.3 pt text width, less three 4 pt gaps, shared 2:3.
    const widths = tableWidths(2, '<w:tblW w:w="2500" w:type="pct"/>');
    expect(widths.table).toBeCloseTo(232.65, 6);
    expect(widths.cells[0]).toBeCloseTo((232.65 - 12) * 0.4, 6);
    expect(widths.cells[1]).toBeCloseTo((232.65 - 12) * 0.6, 6);
  });

  it('carves the spacing out of a table that would pass the text width', () => {
    // 220 + 240 pt cells and 12 pt of gaps are 472 pt in 465.3 pt. The
    // table's border box, half a 0.5 pt rule past the grid on either side,
    // is held to the text width.
    const widths = tableOf(twoColumns(AUTO + spacing(2), [4400, 4800]));
    expect(widths.table).toBeCloseTo(465.3 - 0.5, 6);
    // The cells give way as in an unspaced table that wide: from 220 and
    // 240 pt toward their content (two characters and two 0.5 pt rules).
    const fraction = (465.3 - 0.5 - 12 - 2 * 22) / (460 - 2 * 22);
    expect(widths.cells[0]).toBeCloseTo(22 + (220 - 22) * fraction, 6);
    expect(widths.cells[1]).toBeCloseTo(22 + (240 - 22) * fraction, 6);
  });

  it('keeps the interior gap inside a spanning cell', () => {
    const spanned = tableXml(AUTO + spacing(2), [1440, 2160], [
      cell(1440, 'A1') + cell(2160, 'B1'),
      cell(3600, 'AB2', '<w:gridSpan w:val="2"/>'),
    ]);
    expect(tableOf(spanned).table).toBeCloseTo(192, 6);
  });

  it('draws a vertically merged cell as one box across its rows', () => {
    const merged = tableXml(AUTO + spacing(2), [1440, 2160], [
      cell(1440, 'A1', '<w:vMerge w:val="restart"/>') + cell(2160, 'B1'),
      cell(1440, '', '<w:vMerge/>') + cell(2160, 'B2'),
      cell(1440, '', '<w:vMerge/>') + cell(2160, 'B3'),
    ]);
    const { table, left, horizontalRules } = tableOf(merged);
    expect(table).toBeCloseTo(192, 6);
    // Column A's cell box spans x 4–76 from the table's left edge: only its
    // top and bottom rules lie there, against two per row in column B.
    const within = (from: number, to: number) => horizontalRules.filter((rule) =>
      rule.from.xPt >= left + from - 1 && rule.to.xPt <= left + to + 1).length;
    expect(within(4, 76)).toBe(2);
    expect(within(80, 188)).toBe(6);
  });

  it('measures the same with the spacing set on the rows', () => {
    const onRows = tableXml(AUTO, [1440, 2160], ['1', '2'].map((row) =>
      `<w:trPr>${spacing(2)}</w:trPr>` + cell(1440, `A${row}`) + cell(2160, `B${row}`)));
    const widths = tableOf(onRows);
    expect(widths.table).toBeCloseTo(192, 6);
    expect(widths.cells[0]).toBeCloseTo(72, 6);
  });

  it.each([
    // [outer cell, its side margins, nested tblW, width of each nested cell box]
    [144, 0, AUTO, 30],
    [144, 0, '<w:tblW w:w="5000" w:type="pct"/>', (144 - 0.5 - 0.5 - 12) / 2],
    [72, 0, AUTO, (72 - 0.5 - 0.5 - 12) / 2],
    [72, 5.4, AUTO, (72 - 10.8 - 0.5 - 12) / 2],
  ])('gives a spaced table nested in a %d pt cell with %d pt margins the same two forms', (outerPt, marginPt, nestedWidth, boxPt) => {
    // A 30 + 30 pt table with 2 pt spacing wants 72 pt. It keeps that in a
    // wider cell, takes the cell's width when asked for 100%, and is held to
    // a 72 pt cell by its own border box. The cell's 0.5 pt side rules each
    // take half their width from a cell without margins.
    const nested = tableXml(nestedWidth + spacing(2), [600, 600], [cell(600, 'n1') + cell(600, 'n2')]);
    const margins = `<w:tblCellMar><w:left w:w="${marginPt * 20}" w:type="dxa"/><w:right w:w="${marginPt * 20}" w:type="dxa"/></w:tblCellMar>`;
    const outer = `<w:tbl><w:tblPr>${AUTO}${BORDERS}${margins}</w:tblPr><w:tblGrid>`
      + `<w:gridCol w:w="${outerPt * 20}"/><w:gridCol w:w="2160"/></w:tblGrid><w:tr>`
      + `<w:tc><w:tcPr><w:tcW w:w="${outerPt * 20}" w:type="dxa"/></w:tcPr>${nested}${paragraph('A1')}</w:tc>`
      + `${cell(2160, 'B1')}</w:tr></w:tbl>`;
    const { widths, firstCellXPt } = nestedCells(outer);
    expect(widths).toEqual([boxPt, boxPt].map((width) => Math.round(width * 1e6) / 1e6));
    // From the cell's content box: the half rule its margin leaves
    // uncovered, half the table's own rule (mode 15), then the gap.
    expect(firstCellXPt).toBeCloseTo(Math.max(0, 0.25 - marginPt) + 0.25 + 4, 6);
  });

  it('draws the box of a vertically merged cell on every page it continues on', () => {
    // 60 rows of one line each pass the first page; the first column is one
    // merged cell. Each page draws the left and right rule of its piece.
    const rows = Array.from({ length: 60 }, (_unused, row) =>
      cell(1440, row === 0 ? 'A' : '', `<w:vMerge${row === 0 ? ' w:val="restart"' : ''}/>`) + cell(2160, `B${row}`));
    const pages = tablePages(tableXml(AUTO + spacing(2), [1440, 2160], rows));
    expect(pages.length).toBeGreaterThan(1);
    for (const page of pages) {
      const firstColumnRules = page.borders.filter((border) =>
        border.from.xPt === border.to.xPt && border.from.xPt < page.flowBounds.xPt + 4 + 72 + 1
        && border.from.xPt > page.flowBounds.xPt + 2);
      expect(firstColumnRules.map((border) => Math.round((border.from.xPt - page.flowBounds.xPt) * 100) / 100).sort())
        .toEqual([4.25, 75.75]);
      // One box: the rules run the height of the piece between its gaps.
      for (const rule of firstColumnRules) {
        expect(Math.abs(rule.to.yPt - rule.from.yPt)).toBeGreaterThan(page.flowBounds.heightPt - 12);
      }
    }
  });

  it('keeps the growth of a cell without a preferred width when the table fits', () => {
    // The second cell has no tcW and holds 21 characters (210 pt with its
    // margins aside). Spacing adds the gaps; the content-sized box also
    // holds its own two 0.5 pt rules.
    const text = 'aaaa bbbb cccc dddd e';
    const mixed = (extra: string) => tableXml(AUTO + extra, [1440, 2160], [
      cell(1440, 'A1') + `<w:tc><w:tcPr/>${paragraph(text)}</w:tc>`,
    ]);
    const plain = tableOf(mixed(''));
    const spaced = tableOf(mixed(spacing(2)));
    expect(spaced.cells[0]).toBeCloseTo(plain.cells[0]!, 6);
    expect(plain.cells[1]).toBeGreaterThan(200);
    expect(spaced.cells[1]).toBeCloseTo(plain.cells[1]! + 1, 6);
    expect(spaced.table).toBeCloseTo(plain.table + 12 + 1, 6);
  });

  it('holds a table in a two-column section to its column', () => {
    // Two 222.65 pt columns 20 pt apart. 72 + 108 pt cells keep their
    // widths; 220 + 240 pt cells are resolved in the column.
    expect(tableOf(twoColumns(AUTO + spacing(2)), null, 2).table).toBeCloseTo(192, 6);
    const held = tableOf(twoColumns(AUTO + spacing(2), [4400, 4800]), null, 2);
    expect(held.table).toBeCloseTo(222.65 - 0.5, 6);
  });

  it.each([
    // [columns, tblInd, cells, tblW, table width]
    [2, 40, [1440, 2160], AUTO, 222.65 - 40 - 0.5],
    [2, 20, [1440, 2160], AUTO, 192],
    [2, 20, [4400, 4800], AUTO, 222.65 - 20 - 0.5],
    [2, 40, [1440, 2160], DXA_180, 180],
    [2, 40, [1440, 2160], '<w:tblW w:w="5000" w:type="pct"/>', 222.65 - 0.5],
    [1, 40, [1440, 2160], AUTO, 192],
    [1, 40, [4400, 4800], AUTO, 465.3 - 40 - 0.5],
    [1, 40, [1440, 2160], DXA_180, 180],
    [1, 40, [1440, 2160], '<w:tblW w:w="5000" w:type="pct"/>', 465.3 - 0.5],
    // A negative indent leaves the table's end at the column's end.
    [1, -40, [4400, 4800], AUTO, 472],
    [2, -40, [4400, 4800], AUTO, 222.65 + 40 - 0.5],
    [2, -40, [1440, 2160], '<w:tblW w:w="5000" w:type="pct"/>', 222.65 - 0.5],
  ] as const)('holds an indented automatic table to what the indent leaves of its column (%d columns, %d pt, %j)', (columnCount, indentPt, widths, tableWidth, expected) => {
    // An automatic table ends with its column; a preferred width is kept
    // and moves with the indent.
    const indented = tableOf(twoColumns(`${tableWidth}<w:tblInd w:w="${indentPt * 20}" w:type="dxa"/>${spacing(2)}`,
      widths as unknown as [number, number]), null, columnCount);
    expect(indented.table).toBeCloseTo(expected, 6);
  });

  it('resolves a held automatic table like the same table in the band the gaps leave', () => {
    // 72 pt beside a cell without a preferred width holding 60 characters:
    // the table is held to the text width, the 72 pt cell keeps its width
    // and the other takes the rest.
    const text = Array.from({ length: 12 }, () => 'aaaa').join(' ');
    const held = tableOf(tableXml(AUTO + spacing(2), [1440, 2160], [
      cell(1440, 'A1') + `<w:tc><w:tcPr/>${paragraph(text)}</w:tc>`,
    ]));
    expect(held.table).toBeCloseTo(465.3 - 0.5, 6);
    expect(held.cells[0]).toBeCloseTo(72, 6);
    expect(held.cells[1]).toBeCloseTo(465.3 - 0.5 - 12 - 72, 6);
  });

  it('gives a vertically merged cell room for its own rules', () => {
    // One 10.5 pt line in a cell merged over two rows, 6 pt rules: the box
    // holds both rules and the line.
    const sz48 = BORDERS.replaceAll('w:sz="4"', 'w:sz="48"');
    const merged = `<w:tbl><w:tblPr>${AUTO}${spacing(2)}${sz48}</w:tblPr><w:tblGrid><w:gridCol w:w="1440"/></w:tblGrid>`
      + `<w:tr>${cell(1440, 'A', '<w:vMerge w:val="restart"/>')}</w:tr>`
      + `<w:tr><w:tc><w:tcPr><w:tcW w:w="1440" w:type="dxa"/><w:vMerge/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>`;
    const laidOut = tablePages(merged)[0]!;
    const rules = laidOut.borders.filter((border) => border.from.yPt === border.to.yPt
      && border.from.xPt > laidOut.flowBounds.xPt + 1).map((border) => border.from.yPt).sort((a, b) => a - b);
    expect(rules.length).toBe(2);
    // Between the inner edges of the two 6 pt rules.
    expect(rules[1]! - rules[0]! - 6).toBeGreaterThanOrEqual(10.5);
    // The cell's own bounds, which its fill uses, end with its bottom rule.
    const box = laidOut.rows[0]!.cells[0]!.flowBounds;
    expect(box.yPt).toBeCloseTo(rules[0]! - 3, 6);
    expect(box.yPt + box.heightPt).toBeCloseTo(rules[1]! + 3, 6);
  });

  it('holds the preferred-width form to its column in mode 14 as well', () => {
    // 5.4 pt cell margins: the legacy outer-margin allowance is no part of
    // this ceiling.
    const margins = '<w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>';
    const wide = `<w:tbl><w:tblPr>${AUTO}${spacing(2)}${BORDERS}${margins}</w:tblPr><w:tblGrid>`
      + `<w:gridCol w:w="4400"/><w:gridCol w:w="4800"/></w:tblGrid><w:tr>${cell(4400, 'A1')}${cell(4800, 'B1')}</w:tr></w:tbl>`;
    for (const mode of [14, 15]) {
      const page = tablePages(wide, mode)[0]!;
      expect(page.columnWidthsPt.reduce((sum, width) => sum + width, 0)).toBeCloseTo(465.3 - 0.5, 6);
    }
  });

  it('keeps a repeated header merge apart from a body merge continuing below it', () => {
    const header = (row: number) => `<w:tr><w:trPr><w:tblHeader/></w:trPr>`
      + cell(1440, row === 0 ? 'H' : '', `<w:vMerge${row === 0 ? ' w:val="restart"' : ''}/>`) + cell(2160, `H${row}`) + '</w:tr>';
    const body = Array.from({ length: 70 }, (_unused, row) => '<w:tr>'
      + cell(1440, row === 0 ? 'A' : '', `<w:vMerge${row === 0 ? ' w:val="restart"' : ''}/>`) + cell(2160, `B${row}`) + '</w:tr>').join('');
    const pages = tablePages(`<w:tbl><w:tblPr>${AUTO}${spacing(2)}${BORDERS}</w:tblPr><w:tblGrid>`
      + `<w:gridCol w:w="1440"/><w:gridCol w:w="2160"/></w:tblGrid>${header(0)}${header(1)}${body}</w:tbl>`);
    expect(pages.length).toBeGreaterThan(1);
    for (const page of pages) {
      // The first column holds two boxes: the header merge and the body merge.
      const leftRules = page.borders.filter((border) => border.from.xPt === border.to.xPt
        && Math.abs(border.from.xPt - page.flowBounds.xPt - 4.25) < 0.01);
      expect(leftRules.length).toBe(2);
      // The header cell itself (its fill) ends with the two header rows.
      const headerCell = page.rows[0]!.cells[0]!;
      const headerRows = page.rows[0]!.advancePt + page.rows[1]!.advancePt;
      expect(headerCell.flowBounds.heightPt).toBeLessThanOrEqual(headerRows);
    }
    // On a later page the body merge has no restart of its own: its first
    // row there paints it, not the header's restart.
    const column = pages[1]!.rows.map((row) => row.cells[0]! as {
      verticalMerge: string; visualMergeOwnership?: string; flowBounds: { yPt: number; heightPt: number } });
    expect(column[2]!.verticalMerge).toBe('continue');
    expect(column[2]!.visualMergeOwnership).toBe('continuation');
    // It paints the whole piece, gaps between its rows included; no other
    // cell of the column paints.
    const last = pages[1]!.rows.at(-1)!;
    expect(column[2]!.flowBounds.yPt + column[2]!.flowBounds.heightPt)
      .toBeGreaterThan(last.flowBounds.yPt + last.flowBounds.heightPt - 6);
    expect(column.slice(3).every((continuation) => continuation.visualMergeOwnership === undefined)).toBe(true);
    expect(column[1]!.visualMergeOwnership).toBeUndefined();
  });

  it('keeps a nested table inside the rule a vertical merge beside it starts with', () => {
    // The merged first column opens with a 6 pt right rule; its second row
    // has no override. The nested table of the second row starts inside it.
    const noMargins = '<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>';
    const nested = tableXml('<w:tblW w:w="5000" w:type="pct"/>' + spacing(2), [600, 600], [cell(600, 'n1') + cell(600, 'n2')]);
    const outer = `<w:tbl><w:tblPr>${AUTO}${BORDERS}${noMargins}</w:tblPr><w:tblGrid>`
      + '<w:gridCol w:w="1440"/><w:gridCol w:w="2880"/></w:tblGrid>'
      + `<w:tr>${cell(1440, 'A', '<w:vMerge w:val="restart"/><w:tcBorders><w:right w:val="single" w:sz="48"/></w:tcBorders>')}${cell(2880, 'B1')}</w:tr>`
      + `<w:tr>${cell(1440, '', '<w:vMerge/>')}<w:tc><w:tcPr><w:tcW w:w="2880" w:type="dxa"/></w:tcPr>${nested}${paragraph('B2')}</w:tc></w:tr></w:tbl>`;
    const inner = tablePages(outer, 15)[0]!.rows[1]!.cells.at(-1)!.blocks.map((block) => block.layout)
      .find((block) => block.kind === 'table');
    if (inner?.kind !== 'table') throw new Error('expected a nested table');
    const boxes = inner.rows[0]!.cells.map((box) => box.flowBounds);
    // Half the 6 pt rule on the left, half the 0.5 pt table rule on the right.
    expect(boxes[0]!.xPt).toBeCloseTo(3 + 0.25 + 4, 6);
    expect(boxes[0]!.widthPt).toBeCloseTo((144 - 3 - 0.25 - 0.5 - 12) / 2, 6);
  });

  it('keeps a nested table outside mode 15, or right-aligned, on its previous geometry', () => {
    const noMargins = '<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>';
    const outer = (nestedProperties: string) => `<w:tbl><w:tblPr>${AUTO}${BORDERS}${noMargins}</w:tblPr><w:tblGrid>`
      + '<w:gridCol w:w="1440"/><w:gridCol w:w="2160"/></w:tblGrid><w:tr>'
      + `<w:tc><w:tcPr><w:tcW w:w="1440" w:type="dxa"/></w:tcPr>${tableXml(nestedProperties, [600, 600], [cell(600, 'n1') + cell(600, 'n2')])}${paragraph('A1')}</w:tc>`
      + `${cell(2160, 'B1')}</w:tr></w:tbl>`;
    const full = '<w:tblW w:w="5000" w:type="pct"/>';
    const inRule = nestedCells(outer(full + spacing(2))).widths;
    expect(inRule).toEqual([29.5, 29.5]);
    // The rule's boxes are not used: the previous geometry gives other widths.
    expect(nestedCells(outer(full + spacing(2)), null).widths).not.toEqual(inRule);
    expect(nestedCells(outer(full + '<w:jc w:val="right"/>' + spacing(2))).widths).not.toEqual(inRule);
    expect(nestedCells(outer(full + '<w:bidiVisual/>' + spacing(2))).widths).not.toEqual(inRule);
  });

  it('keeps table spacing mixed with row spacing on its previous geometry', () => {
    // tblPr spacing and the same spacing on one row only: outside the rule,
    // the table keeps its 180 pt.
    const mixed = tableXml(AUTO + spacing(2), [1440, 2160], [
      cell(1440, 'A1') + cell(2160, 'B1'),
      `<w:trPr>${spacing(2)}</w:trPr>` + cell(1440, 'A2') + cell(2160, 'B2'),
    ]);
    expect(tableOf(mixed).table).toBeCloseTo(180, 6);
  });

  it('carries the cell that opens a merge down a long table', () => {
    // Four rows, first column merged with a 6 pt right rule, a nested spaced
    // table in the last row: it starts inside that rule.
    const noMargins = '<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>';
    const nested = tableXml('<w:tblW w:w="5000" w:type="pct"/>' + spacing(2), [600, 600], [cell(600, 'n1') + cell(600, 'n2')]);
    const rows = Array.from({ length: 4 }, (_unused, row) => `<w:tr>`
      + cell(1440, row === 0 ? 'A' : '', row === 0
        ? '<w:vMerge w:val="restart"/><w:tcBorders><w:right w:val="single" w:sz="48"/></w:tcBorders>' : '<w:vMerge/>')
      + (row === 3 ? `<w:tc><w:tcPr><w:tcW w:w="2880" w:type="dxa"/></w:tcPr>${nested}${paragraph('B')}</w:tc>` : cell(2880, `B${row}`))
      + '</w:tr>').join('');
    const outer = `<w:tbl><w:tblPr>${AUTO}${BORDERS}${noMargins}</w:tblPr><w:tblGrid>`
      + `<w:gridCol w:w="1440"/><w:gridCol w:w="2880"/></w:tblGrid>${rows}</w:tbl>`;
    const inner = tablePages(outer, 15)[0]!.rows[3]!.cells.at(-1)!.blocks.map((block) => block.layout)
      .find((block) => block.kind === 'table');
    if (inner?.kind !== 'table') throw new Error('expected a nested table');
    expect(inner.rows[0]!.cells[0]!.flowBounds.xPt).toBeCloseTo(3 + 0.25 + 4, 6);
  });

  it('reads the row exception of the row that opens the merge', () => {
    // The opening row's tblPrEx gives a 6 pt insideV; the next row has none.
    const noMargins = '<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>';
    const nested = tableXml('<w:tblW w:w="5000" w:type="pct"/>' + spacing(2), [600, 600], [cell(600, 'n1') + cell(600, 'n2')]);
    const exception = '<w:tblPrEx><w:tblBorders><w:insideV w:val="single" w:sz="48"/></w:tblBorders></w:tblPrEx>';
    const outer = `<w:tbl><w:tblPr>${AUTO}${BORDERS}${noMargins}</w:tblPr><w:tblGrid>`
      + '<w:gridCol w:w="1440"/><w:gridCol w:w="2880"/></w:tblGrid>'
      + `<w:tr>${exception}${cell(1440, 'A', '<w:vMerge w:val="restart"/>')}${cell(2880, 'B1')}</w:tr>`
      + `<w:tr>${cell(1440, '', '<w:vMerge/>')}<w:tc><w:tcPr><w:tcW w:w="2880" w:type="dxa"/></w:tcPr>${nested}${paragraph('B2')}</w:tc></w:tr></w:tbl>`;
    const inner = tablePages(outer, 15)[0]!.rows[1]!.cells.at(-1)!.blocks.map((block) => block.layout)
      .find((block) => block.kind === 'table');
    if (inner?.kind !== 'table') throw new Error('expected a nested table');
    expect(inner.rows[0]!.cells[0]!.flowBounds.xPt).toBeCloseTo(3 + 0.25 + 4, 6);
  });

  describe('a merge continuing below repeated header rows', () => {
    const headerText = Array.from({ length: 12 }, () => 'hhhhhh').join(' ');
    const paged = (spacingPt: number, headerContent: string) => {
      const header = (row: number) => '<w:tr><w:trPr><w:tblHeader/></w:trPr>'
        + cell(1440, row === 0 ? headerContent : '', `<w:vMerge${row === 0 ? ' w:val="restart"' : ''}/>`) + cell(2160, `H${row}`) + '</w:tr>';
      const body = Array.from({ length: 70 }, (_unused, row) => '<w:tr>'
        + cell(1440, row === 0 ? 'A' : '', `<w:vMerge${row === 0 ? ' w:val="restart"' : ''}/>`) + cell(2160, `B${row}`) + '</w:tr>').join('');
      return tablePages(`<w:tbl><w:tblPr>${AUTO}${spacing(spacingPt)}${BORDERS}</w:tblPr><w:tblGrid>`
        + `<w:gridCol w:w="1440"/><w:gridCol w:w="2160"/></w:tblGrid>${header(0)}${header(1)}${body}</w:tbl>`);
    };

    it.each([0, 2])('sizes the repeated header rows for the header merge alone (%d pt spacing)', (spacingPt) => {
      // The merged header cell holds text of several lines; its two rows
      // grow to hold it on every page, apart from the body merge below.
      const pages = paged(spacingPt, headerText);
      expect(pages.length).toBeGreaterThan(1);
      const heights = pages.map((page) => page.rows[0]!.cells[0]!.flowBounds.heightPt);
      expect(heights[0]).toBeGreaterThan(40);
      expect(heights[1]).toBeCloseTo(heights[0]!, 6);
      for (const page of pages) {
        const headerCell = page.rows[0]!.cells[0]!;
        expect(headerCell.flowBounds.heightPt)
          .toBeLessThanOrEqual(page.rows[0]!.advancePt + page.rows[1]!.advancePt + 1e-6);
      }
    });

    it('draws the collapsed rules of the body merge on the later page', () => {
      const page = paged(0, 'H')[1]!;
      const bottom = page.flowBounds.yPt + page.flowBounds.heightPt;
      // The left rule runs the height of the piece.
      const left = page.borders.filter((border) => border.from.xPt === border.to.xPt
        && Math.abs(border.from.xPt - page.flowBounds.xPt) < 0.5);
      const covered = left.reduce((sum, border) => sum + Math.abs(border.to.yPt - border.from.yPt), 0);
      expect(covered).toBeGreaterThan(page.flowBounds.heightPt - 2);
      // The bottom rule runs the width of both columns.
      const bottomRule = page.borders.filter((border) => border.from.yPt === border.to.yPt
        && Math.abs(border.from.yPt - bottom) < 1);
      const width = bottomRule.reduce((sum, border) => sum + Math.abs(border.to.xPt - border.from.xPt), 0);
      expect(width).toBeGreaterThan(180 - 1);
    });
  });

  it('gives a continued merge the shading and side rules of the cell that opens it', () => {
    // Shading on the restart only; the continuations carry an 8 pt right
    // rule the merge never shows. Later pages keep the restart's.
    const rows = Array.from({ length: 140 }, (_unused, row) => cell(1440, row === 0 ? 'A' : '', row === 0
      ? '<w:vMerge w:val="restart"/><w:shd w:val="clear" w:fill="FF0000"/>'
      : '<w:vMerge/><w:tcBorders><w:right w:val="single" w:sz="64"/></w:tcBorders>') + cell(2160, `B${row}`));
    for (const spacingPt of [0, 2]) {
      const pages = tablePages(tableXml(AUTO + spacing(spacingPt), [1440, 2160], rows));
      expect(pages.length).toBeGreaterThan(1);
      for (const page of pages) {
        const owner = page.rows[0]!.cells[0]! as { background?: { color: string } };
        expect(owner.background?.color.toLowerCase()).toContain('ff0000');
        expect(page.borders.every((border) => border.widthPt === 0.5)).toBe(true);
      }
    }
  });

  it.each([
    // [what the merge's rows author, painted rule between the columns]
    ['an 8 pt right rule on the continuations, which the merge never shows',
      '', '<w:tcBorders><w:right w:val="single" w:sz="64"/></w:tcBorders>', '', 0.5],
    ['a 6 pt right rule on the cell that opens the merge',
      '<w:tcBorders><w:right w:val="single" w:sz="48"/></w:tcBorders>', '', '', 6],
    ['a 6 pt insideV in the row exception of the row that opens the merge',
      '', '', '<w:tblPrEx><w:tblBorders><w:insideV w:val="single" w:sz="48"/></w:tblBorders></w:tblPrEx>', 6],
  ])('keeps a nested table clear of the rule painted beside a merge on a later page: %s', (_name, restart, continuation, exception, rulePt) => {
    // Mode 15, no cell margins, rows that do not split, the first column
    // merged over every row; the last row holds a 100% spaced table.
    const noMargins = '<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>';
    const nested = tableXml('<w:tblW w:w="5000" w:type="pct"/>' + spacing(2), [600, 600], [cell(600, 'n1') + cell(600, 'n2')]);
    const rowCount = 140;
    const rows = Array.from({ length: rowCount }, (_unused, row) => `<w:tr>${row === 0 ? exception : ''}<w:trPr><w:cantSplit/></w:trPr>`
      + cell(1440, row === 0 ? 'A' : '', row === 0 ? `<w:vMerge w:val="restart"/>${restart}` : `<w:vMerge/>${continuation}`)
      + (row === rowCount - 1
        ? `<w:tc><w:tcPr><w:tcW w:w="2880" w:type="dxa"/></w:tcPr>${nested}${paragraph('B')}</w:tc>` : cell(2880, `B${row}`))
      + '</w:tr>').join('');
    const pages = tablePages(`<w:tbl><w:tblPr>${AUTO}${BORDERS}${noMargins}</w:tblPr><w:tblGrid>`
      + `<w:gridCol w:w="1440"/><w:gridCol w:w="2880"/></w:tblGrid>${rows}</w:tbl>`, 15);
    expect(pages.length).toBeGreaterThan(1);
    const page = pages.at(-1)!;
    const host = page.rows.at(-1)!.cells.at(-1)!;
    // The rule painted between the columns beside the last row.
    const shared = page.borders.filter((border) => border.from.xPt === border.to.xPt
      && Math.abs(border.from.xPt - host.flowBounds.xPt) < 0.01
      && Math.max(border.from.yPt, border.to.yPt) >= host.flowBounds.yPt + host.flowBounds.heightPt - 1);
    expect(shared.length).toBeGreaterThan(0);
    expect(Math.max(...shared.map((border) => border.widthPt))).toBe(rulePt);
    const inner = host.blocks.map((block) => block.layout).find((block) => block.kind === 'table');
    if (inner?.kind !== 'table') throw new Error('expected a nested table');
    // The nested table's border box starts where the painted rule ends.
    const firstBox = inner.rows[0]!.cells[0]!.flowBounds.xPt;
    expect(firstBox - 4 - 0.25).toBeCloseTo(rulePt / 2, 6);
  });

  it('leaves the bottom rule of a continued merge to its last row', () => {
    // Only the opening row's exception has a 6 pt bottom rule. A last page
    // holding one row of the merge ends with the table's 0.5 pt rule.
    const exception = '<w:tblPrEx><w:tblBorders><w:bottom w:val="single" w:sz="48"/></w:tblBorders></w:tblPrEx>';
    const table = (rowCount: number) => tableXml(AUTO, [1440, 2160], Array.from({ length: rowCount }, (_unused, row) =>
      (row === 0 ? exception : '') + cell(1440, row === 0 ? 'A' : '', `<w:vMerge${row === 0 ? ' w:val="restart"' : ''}/>`) + cell(2160, `B${row}`)));
    let lastPage: ReturnType<typeof tablePages>[number] | undefined;
    for (let rowCount = 60; rowCount < 90 && lastPage === undefined; rowCount += 1) {
      const pages = tablePages(table(rowCount), 15);
      if (pages.length > 1 && pages.at(-1)!.rows.length === 1) lastPage = pages.at(-1);
    }
    if (!lastPage) throw new Error('expected a last page of one row');
    const bottom = lastPage.flowBounds.yPt + lastPage.flowBounds.heightPt;
    const bottomRules = lastPage.borders.filter((border) => border.from.yPt === border.to.yPt
      && Math.abs(border.from.yPt - bottom) < 1);
    expect(bottomRules.length).toBeGreaterThan(0);
    expect(bottomRules.every((border) => border.widthPt === 0.5)).toBe(true);
  });

  it('fits a nested table between wide side rules of its cell', () => {
    // A 36 pt cell with 6 pt collapsed side rules leaves 30 pt: two boxes of
    // 8.75 pt with their gaps and the table's own half rules.
    const sz48 = BORDERS.replaceAll('w:sz="4"', 'w:sz="48"');
    const noMargins = '<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>';
    const nested = tableXml('<w:tblW w:w="5000" w:type="pct"/>' + spacing(2), [360, 360], [cell(360, '') + cell(360, '')]);
    const outer = `<w:tbl><w:tblPr>${AUTO}${sz48}${noMargins}</w:tblPr><w:tblGrid>`
      + '<w:gridCol w:w="720"/><w:gridCol w:w="2160"/></w:tblGrid><w:tr>'
      + `<w:tc><w:tcPr><w:tcW w:w="720" w:type="dxa"/></w:tcPr>${nested}${paragraph('')}</w:tc>`
      + `${cell(2160, 'B1')}</w:tr></w:tbl>`;
    const { widths, firstCellXPt } = nestedCells(outer);
    expect(widths).toEqual([8.75, 8.75]);
    expect(firstCellXPt).toBeCloseTo(3 + 0.25 + 4, 6);
  });

  it('keeps a nested table in a merged cell clear of the widest rule beside the merge', () => {
    // The first column is one cell over eight rows and holds the nested
    // table. The cells beside it share a 0.5 pt rule in the first row and
    // carry a 6 pt left rule below it.
    const noMargins = '<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>';
    const nested = tableXml('<w:tblW w:w="5000" w:type="pct"/>' + spacing(2), [600, 600], [cell(600, 'n1') + cell(600, 'n2')]);
    const rows = Array.from({ length: 8 }, (_unused, row) => '<w:tr>'
      + (row === 0
        ? `<w:tc><w:tcPr><w:tcW w:w="2880" w:type="dxa"/><w:vMerge w:val="restart"/></w:tcPr>${nested}${paragraph('A')}</w:tc>`
        : cell(2880, '', '<w:vMerge/>'))
      + cell(1440, `B${row}`, row === 0 ? '' : '<w:tcBorders><w:left w:val="single" w:sz="48"/></w:tcBorders>')
      + '</w:tr>').join('');
    const outer = `<w:tbl><w:tblPr>${AUTO}${BORDERS}${noMargins}</w:tblPr><w:tblGrid>`
      + `<w:gridCol w:w="2880"/><w:gridCol w:w="1440"/></w:tblGrid>${rows}</w:tbl>`;
    const { widths, firstCellXPt } = nestedCells(outer);
    // Half the table's 0.5 pt rule on the left, half the 6 pt rule on the right.
    expect(widths).toEqual([(144 - 0.25 - 3 - 0.5 - 12) / 2, (144 - 0.25 - 3 - 0.5 - 12) / 2]);
    expect(firstCellXPt).toBeCloseTo(0.25 + 0.25 + 4, 6);
  });

  it('mirrors the outer rules of a right-to-left spaced table', () => {
    // start is red and end is blue: in a bidiVisual table start is on the right.
    const borders = '<w:tblBorders><w:top w:val="single" w:sz="4"/><w:start w:val="single" w:sz="4" w:color="FF0000"/>'
      + '<w:bottom w:val="single" w:sz="4"/><w:end w:val="single" w:sz="4" w:color="0000FF"/>'
      + '<w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders>';
    const table = (bidi: string) => `<w:tbl><w:tblPr>${bidi}${AUTO}${spacing(2)}${borders}</w:tblPr><w:tblGrid>`
      + `<w:gridCol w:w="1440"/><w:gridCol w:w="2160"/></w:tblGrid><w:tr>${cell(1440, 'A1')}${cell(2160, 'B1')}</w:tr></w:tbl>`;
    const frame = (bidi: string) => {
      const page = tablePages(table(bidi), 15)[0]!;
      const vertical = page.borders.filter((border) => border.from.xPt === border.to.xPt)
        .sort((a, b) => a.from.xPt - b.from.xPt);
      return [vertical[0]!.color.toLowerCase(), vertical.at(-1)!.color.toLowerCase()];
    };
    expect(frame('')).toEqual(['#ff0000', '#0000ff']);
    expect(frame('<w:bidiVisual/>')).toEqual(['#0000ff', '#ff0000']);
  });

  it('keeps the upright table of a two-column vertical section on its previous geometry', () => {
    expect(tableOf(twoColumns(AUTO + spacing(2)), 'tbRl', 2).table).toBeCloseTo(180, 6);
    expect(tableOf(twoColumns(AUTO + spacing(2)), 'tbRl', 1).table).toBeCloseTo(192, 6);
  });

  it('paints a continued merge as one box on the later page', () => {
    // No header: the first row of the second page opens the merge there.
    const rows = Array.from({ length: 60 }, (_unused, row) =>
      cell(1440, row === 0 ? 'A' : '', `<w:vMerge${row === 0 ? ' w:val="restart"' : ''}/><w:shd w:val="clear" w:fill="FF0000"/>`) + cell(2160, `B${row}`));
    const page = tablePages(tableXml(AUTO + spacing(2), [1440, 2160], rows))[1]!;
    const column = page.rows.map((row) => row.cells[0]! as {
      visualMergeOwnership?: string; flowBounds: { yPt: number; heightPt: number } });
    expect(column[0]!.visualMergeOwnership).toBe('continuation');
    expect(column.slice(1).every((continuation) => continuation.visualMergeOwnership === undefined)).toBe(true);
    // Its bounds are the border box drawn for it.
    const rules = page.borders.filter((border) => border.from.yPt === border.to.yPt
      && Math.abs(border.from.xPt - page.flowBounds.xPt - 4) < 0.01).map((border) => border.from.yPt).sort((a, b) => a - b);
    expect(rules.length).toBe(2);
    expect(column[0]!.flowBounds.yPt).toBeCloseTo(rules[0]! - 0.25, 6);
    expect(column[0]!.flowBounds.yPt + column[0]!.flowBounds.heightPt).toBeCloseTo(rules[1]! + 0.25, 6);
  });

  it('leaves spacing that would empty the cells outside the rule', () => {
    // Mode 14, 76.5 pt spacing, 10 pt indent: the gaps take the band.
    const margins = '<w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>';
    const wide = `<w:tbl><w:tblPr>${AUTO}<w:tblCellSpacing w:w="1530" w:type="dxa"/><w:tblInd w:w="200" w:type="dxa"/>${BORDERS}${margins}</w:tblPr><w:tblGrid>`
      + `<w:gridCol w:w="1440"/><w:gridCol w:w="2160"/></w:tblGrid><w:tr>${cell(1440, 'A1')}${cell(2160, 'B1')}</w:tr></w:tbl>`;
    const page = tablePages(wide, 14)[0]!;
    for (const box of page.rows[0]!.cells) expect(box.flowBounds.widthPt).toBeGreaterThan(1);
    for (const border of page.borders) {
      expect(border.to.xPt).toBeGreaterThanOrEqual(border.from.xPt);
    }
  });

  it('ends a negatively indented table with its column in every mode', () => {
    // 220 + 240 pt cells and their gaps are 472 pt: from an origin 20 pt
    // left of the text they end before the column does.
    const indented = `<w:tbl><w:tblPr>${AUTO}${spacing(2)}<w:tblInd w:w="-400" w:type="dxa"/>${BORDERS}</w:tblPr><w:tblGrid>`
      + `<w:gridCol w:w="4400"/><w:gridCol w:w="4800"/></w:tblGrid><w:tr>${cell(4400, 'A1')}${cell(4800, 'B1')}</w:tr></w:tbl>`;
    for (const mode of [12, 14, 15]) {
      const page = tablePages(indented, mode)[0]!;
      expect(page.columnWidthsPt.reduce((sum, width) => sum + width, 0)).toBeCloseTo(472, 6);
    }
  });

  it('applies both forms to the upright block table of a tbRl section', () => {
    expect(tableWidths(5, AUTO, 'tbRl').table).toBeCloseTo(210, 6);
    const preferred = tableWidths(2, DXA_180, 'tbRl');
    expect(preferred.table).toBeCloseTo(180, 6);
    expect(preferred.cells[0]).toBeCloseTo(67.2, 6);
  });
});
