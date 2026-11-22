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

function docx(table: string, direction: string | null, columns = 1): Uint8Array {
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
      + '</Types>')],
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
function nestedCells(table: string): { widths: number[]; firstCellXPt: number } {
  const archive = new DocxArchive(docx(table, null));
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
function tablePages(table: string) {
  const archive = new DocxArchive(docx(table, null));
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
    expect(widths.cells[0]).toBeCloseTo((465.3 - 0.5 - 12) * 220 / 460, 6);
    expect(widths.cells[1]).toBeCloseTo((465.3 - 0.5 - 12) * 240 / 460, 6);
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
    // uncovered, then the gap.
    expect(firstCellXPt).toBeCloseTo(Math.max(0, 0.25 - marginPt) + 4, 6);
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
    // With an indent the table keeps the geometry it has without this rule.
    const indented = tableOf(twoColumns(AUTO + '<w:tblInd w:w="800" w:type="dxa"/>' + spacing(2), [4400, 4800]), null, 2);
    expect(indented.table).toBeCloseTo(222.65, 6);
  });

  it('applies both forms to the upright block table of a tbRl section', () => {
    expect(tableWidths(5, AUTO, 'tbRl').table).toBeCloseTo(210, 6);
    const preferred = tableWidths(2, DXA_180, 'tbRl');
    expect(preferred.table).toBeCloseTo(180, 6);
    expect(preferred.cells[0]).toBeCloseTo(67.2, 6);
  });
});
