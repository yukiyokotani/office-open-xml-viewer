/// <reference types="node" />

import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import type { DocxDocumentModel } from './types.js';
import { normalizeInternalDocumentModel } from './parser-model.js';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { storeZip } from './conformance/generate.js';

// WORD_AUTOFIT_CONTENT_RESOLUTION. The controls are A4 pages with 465.3 pt
// of text width in compatibility mode 15 and one AutoFit table of 10.5 pt
// text, 5.25 pt per character, with omitted cell margins. Word adds half a
// point on either side of a cell's text there, which is not modelled: the
// expected widths are the resolution of Word's controls for text alone, so
// a width that content decides is Word's less 1 pt per cell of content.

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const encoder = new TextEncoder();
const xml = (body: string) => encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${body}`);
const paragraph = (text: string) => `<w:p><w:r><w:rPr><w:sz w:val="21"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;

const BORDERS = '<w:tblBorders>'
  + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
    .map((edge) => `<w:${edge} w:val="single" w:sz="4"/>`).join('')
  + '</w:tblBorders>';
const AUTO = '<w:tblW w:w="0" w:type="auto"/>';
const dxa = (pt: number) => `<w:tblW w:w="${pt * 20}" w:type="dxa"/>`;
const pct = (percent: number) => `<w:tblW w:w="${percent * 50}" w:type="pct"/>`;

/** Unbreakable text `pt - 1` wide: the word that is `pt` wide in Word's cell. */
const word = (pt: number) => 'W'.repeat(Math.round((pt - 1) / 5.25));
const TWO = 'A1'; // 10.5 pt
const SHORT = 'aaa bbb ccc ddd'; // 78.75 pt on one line, 15.75 pt at its widest word
const LONG = Array.from({ length: 60 }, () => 'word').join(' '); // 1569.75 pt, 21 pt
const HALF = Array.from({ length: 30 }, () => 'word').join(' '); // 782.25 pt, 21 pt

type Width = number | `${number}%` | null;
function cell(text: string, width: Width = null, properties = '', inner = ''): string {
  const tcW = width === null ? ''
    : typeof width === 'number' ? `<w:tcW w:w="${width * 20}" w:type="dxa"/>`
      : `<w:tcW w:w="${parseFloat(width) * 50}" w:type="pct"/>`;
  return `<w:tc><w:tcPr>${tcW}${properties}</w:tcPr>${inner}${paragraph(text)}</w:tc>`;
}
function table(tableWidth: string, grid: readonly number[], rows: readonly string[], layout = ''): string {
  return `<w:tbl><w:tblPr>${tableWidth}${layout}${BORDERS}</w:tblPr><w:tblGrid>`
    + grid.map((width) => `<w:gridCol w:w="${width * 20}"/>`).join('')
    + `</w:tblGrid>${rows.map((row) => `<w:tr>${row}</w:tr>`).join('')}</w:tbl>`;
}
/** One row of cells over a 72 + 108 pt grid (the grid takes no part). */
const row = (tableWidth: string, ...cells: readonly [string, Width?][]) =>
  table(tableWidth, cells.length === 3 ? [60, 72, 108] : [72, 108], [cells.map(([text, width]) => cell(text, width ?? null)).join('')]);

function docx(body: string, mode: number | null): Uint8Array {
  const document = `<w:document xmlns:w="${W}"><w:body>${paragraph('before')}${body}${paragraph('after')}`
    + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
    + '<w:pgMar w:top="1200" w:right="1800" w:bottom="1600" w:left="800" w:header="720" w:footer="720" w:gutter="0"/>'
    + '</w:sectPr></w:body></w:document>';
  const settings = `<w:settings xmlns:w="${W}"><w:compat>`
    + (mode === null ? '' : `<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="${mode}"/>`)
    + '</w:compat></w:settings>';
  return storeZip(new Map([
    ['[Content_Types].xml', xml('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
      + '</Types>')],
    ['_rels/.rels', xml('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + `<Relationship Id="rId1" Type="${REL}/officeDocument" Target="word/document.xml"/></Relationships>`)],
    ['word/_rels/document.xml.rels', xml('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + `<Relationship Id="rId1" Type="${REL}/settings" Target="settings.xml"/></Relationships>`)],
    ['word/document.xml', xml(document)],
    ['word/settings.xml', xml(settings)],
  ]));
}

/** Half-width Latin characters, as in the monospaced face of the controls. */
function measureContext(): CanvasRenderingContext2D {
  let font = '10px serif';
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    fontKerning: 'normal',
    measureText(text: string) {
      const px = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      return {
        width: [...text].length * px / 2,
        fontBoundingBoxAscent: px * 0.8,
        fontBoundingBoxDescent: px * 0.2,
        actualBoundingBoxAscent: px * 0.8,
        actualBoundingBoxDescent: px * 0.2,
      } as TextMetrics;
    },
  } as unknown as CanvasRenderingContext2D;
}

/** Column widths of the table and of the table nested in its first cell. */
function columns(body: string, mode: number | null = 15): { outer: number[]; nested: number[] } {
  const archive = new DocxArchive(docx(body, mode));
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
  const inner = laidOut.rows[0]!.cells[0]!.blocks.map((block) => block.layout).find((block) => block.kind === 'table');
  return {
    outer: [...laidOut.columnWidthsPt],
    nested: inner?.kind === 'table' ? [...inner.columnWidthsPt] : [],
  };
}

/** Word's rule positions are on a 0.12 pt raster. */
function expectWidths(actual: readonly number[], expected: readonly number[]): void {
  expect(actual.length).toBe(expected.length);
  actual.forEach((width, index) => {
    expect(Math.abs(width - expected[index]!), `${actual.join(' / ')} against ${expected.join(' / ')}`).toBeLessThan(0.2);
  });
}

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});

describe('AutoFit column widths from content and preferred widths', () => {
  it.each<[string, string, number[]]>([
    ['the saved grid takes no part', table(AUTO, [72, 108], [cell(TWO, 100) + cell(TWO, 100)]), [100, 100]],
    ['a cell without a width holds its content', row(AUTO, [TWO, 72], [TWO]), [72, 10.5]],
    ['a cell without a width holds its content on one line', row(AUTO, [TWO, 72], [SHORT]), [72, 78.75]],
    ['a width of type auto is no width', table(AUTO, [72, 108], [cell(TWO, 72) + `<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr>${paragraph(TWO)}</w:tc>`]), [72, 10.5]],
    ['three columns, the middle one without a width', row(AUTO, [TWO, 72], [TWO], [TWO, 108]), [72, 10.5, 108]],
    ['content wider than its preferred width widens the table', row(AUTO, [word(106), 72], [TWO, 108]), [105, 108]],
    ['breakable text keeps its preferred width', row(AUTO, [LONG, 72], [TWO, 108]), [72, 108]],
    ['no widths at all', row(AUTO, [SHORT], [TWO]), [78.75, 10.5]],
    ['the largest preferred width of a column counts',
      table(AUTO, [72, 108], [cell(TWO, 72) + cell(TWO, 108), cell(TWO, 108) + cell(TWO, 72)]), [108, 108]],
    ['a preferred width in one row only',
      table(AUTO, [72, 108], [cell(TWO, 72) + cell(TWO, 108), cell(TWO) + cell(TWO)]), [72, 108]],
  ])('sizes an automatic table: %s', (_name, body, expected) => {
    expectWidths(columns(body).outer, expected);
  });

  it.each<[string, string, number[]]>([
    ['150 pt', row(dxa(150), [TWO, 72], [TWO, 108]), [60.4, 89.6]],
    ['120 pt', row(dxa(120), [TWO, 72], [TWO, 108]), [48.79, 71.21]],
    ['90 pt', row(dxa(90), [TWO, 72], [TWO, 108]), [37.19, 52.81]],
    ['30 pt', row(dxa(30), [TWO, 72], [TWO, 108]), [13.98, 16.02]],
    ['150 pt, a 64 pt word in the second cell', row(dxa(150), [TWO, 72], [word(64), 108]), [54.7, 95.3]],
    ['120 pt, four short words in the first cell', row(dxa(120), [SHORT, 72], [TWO, 108]), [50.05, 69.95]],
    ['150 pt, three columns', row(dxa(150), [TWO, 60], [TWO, 72], [TWO, 108]), [38.63, 45.45, 65.91]],
    ['180 pt, a 106 pt word in a 72 pt cell', row(dxa(180), [word(106), 72], [TWO, 108]), [105, 75]],
    ['150 pt, a 64 pt word in a 72 pt cell', row(dxa(150), [word(64), 72], [TWO, 108]), [69.46, 80.54]],
    ['180 pt, no widths, long text', row(dxa(180), [LONG], [TWO]), [169.5, 10.5]],
    ['180 pt, no widths, long text twice', row(dxa(180), [LONG], [LONG]), [90, 90]],
  ])('shares a table below its maxima between minimum and maximum: %s', (_name, body, expected) => {
    expectWidths(columns(body).outer, expected);
  });

  it('holds a table to the text width the same way', () => {
    // 465.3 pt here; Word's outer rules are 464.9 pt apart.
    const interpolate = (minimums: number[], maximums: number[]) => {
      const fraction = (465.3 - minimums.reduce((a, b) => a + b)) / maximums.reduce((sum, maximum, index) => sum + maximum - minimums[index]!, 0);
      return minimums.map((minimum, index) => minimum + fraction * (maximums[index]! - minimum));
    };
    expectWidths(columns(row(AUTO, [TWO, 300], [TWO, 300])).outer, [232.65, 232.65]);
    expectWidths(columns(row(AUTO, [SHORT, 200], [TWO, 400])).outer, interpolate([15.75, 10.5], [200, 400]));
    expectWidths(columns(row(AUTO, [LONG], [SHORT])).outer, interpolate([21, 15.75], [1569.75, 78.75]));
    expectWidths(columns(row(AUTO, [LONG], [HALF])).outer, interpolate([21, 21], [1569.75, 782.25]));
    // A column without a width gives way before one with a width.
    expectWidths(columns(row(AUTO, [TWO, 72], [LONG])).outer, [72, 465.3 - 72]);
  });

  it.each<[string, string, number[]]>([
    ['every column has a width: all grow', row(dxa(240), [TWO, 72], [TWO, 108]), [96, 144]],
    ['every column has a width, one holds a wide word', row(dxa(300), [word(106), 72], [TWO, 108]), [147.89, 152.11]],
    ['a column without a width takes the room', row(dxa(180), [TWO, 72], [TWO]), [72, 108]],
    ['a column without a width takes the room, also with long text', row(dxa(300), [TWO, 72], [LONG]), [72, 228]],
    ['no widths: equal content', row(dxa(180), [TWO], [TWO]), [90, 90]],
    ['no widths: by one-line content', row(dxa(180), [SHORT], [TWO]), [158.82, 21.18]],
    ['no widths: a wide word', row(dxa(180), [word(106)], [TWO]), [163.64, 16.36]],
    ['no widths: three columns', row(dxa(300), [SHORT], [TWO], [TWO]), [236.84, 31.58, 31.58]],
  ])('shares the room of a table above its maxima: %s', (_name, body, expected) => {
    expectWidths(columns(body).outer, expected);
  });

  it('never goes below the content', () => {
    expectWidths(columns(row(dxa(60), [word(64)], [TWO])).outer, [63, 10.5]);
    expectWidths(columns(row(dxa(180), [word(106), 72], [word(158.5), 108])).outer, [105, 157.5]);
  });

  it.each<[string, string, number[]]>([
    ['two letters', row(AUTO, [TWO, '40%'], [TWO, '60%']), [10.5, 15.75]],
    ['four short words', row(AUTO, [SHORT, '40%'], [TWO, '60%']), [78.75, 118.12]],
    ['a 106 pt word', row(AUTO, [word(106), '40%'], [TWO, '60%']), [105, 157.5]],
    ['long text takes the text width', row(AUTO, [LONG, '40%'], [TWO, '60%']), [186.1, 279.2]],
    ['20% and 80%', row(AUTO, [TWO, '20%'], [TWO, '80%']), [10.5, 42]],
    ['30% and 30% share the table equally', row(AUTO, [TWO, '30%'], [TWO, '30%']), [17.5, 17.5]],
    ['30% and 30% under 180 pt', row(dxa(180), [TWO, '30%'], [TWO, '30%']), [90, 90]],
    ['40% beside a cell without a width', row(AUTO, [TWO, '40%'], [TWO]), [10.5, 15.75]],
    ['40% beside 108 pt', row(AUTO, [TWO, '40%'], [TWO, 108]), [72, 108]],
    ['under 180 pt', row(dxa(180), [TWO, '40%'], [TWO, '60%']), [72, 108]],
    ['under 50%', row(pct(50), [TWO, '40%'], [TWO, '60%']), [93.06, 139.59]],
  ])('sizes percentage cells as shares of the table: %s', (_name, body, expected) => {
    expectWidths(columns(body).outer, expected);
  });

  it.each<[string, string, number[]]>([
    ['a wide word over 72 + 108 pt',
      table(AUTO, [72, 108], [cell(word(237.25), 180, '<w:gridSpan w:val="2"/>'), cell(TWO, 72) + cell(TWO, 108)]), [96.97, 139.28]],
    ['the same under 180 pt, spanning row second',
      table(dxa(180), [72, 108], [cell(TWO, 72) + cell(TWO, 108), cell(word(237.25), 180, '<w:gridSpan w:val="2"/>')]), [96.97, 139.28]],
    ['a wide word over cells without widths',
      table(AUTO, [72, 108], [cell(word(237.25), null, '<w:gridSpan w:val="2"/>'), cell(TWO) + cell(TWO)]), [118.12, 118.12]],
    ['a wide word over unequal cells without widths',
      table(AUTO, [72, 108], [cell(word(237.25), null, '<w:gridSpan w:val="2"/>'), cell(SHORT) + cell(TWO)]), [193.3, 42.95]],
    ['a wide word over two of three columns',
      table(AUTO, [60, 72, 108], [cell(word(158.5), 132, '<w:gridSpan w:val="2"/>') + cell(TWO, 108), cell(TWO, 60) + cell(TWO, 72) + cell(TWO, 108)]),
      [72.57, 84.93, 108]],
    ['a preferred width beyond its columns goes to the last',
      table(AUTO, [72, 108], [cell(TWO, 240, '<w:gridSpan w:val="2"/>'), cell(TWO, 72) + cell(TWO, 108)]), [72, 168]],
    ['a preferred width within its columns changes nothing',
      table(AUTO, [72, 108], [cell(TWO, 120, '<w:gridSpan w:val="2"/>'), cell(TWO, 72) + cell(TWO, 108)]), [72, 108]],
    ['a vertically merged cell counts in its first row',
      table(AUTO, [72, 108], [cell(word(106), 72, '<w:vMerge w:val="restart"/>') + cell(TWO, 108), cell('', 72, '<w:vMerge/>') + cell(TWO, 108)]), [105, 108]],
  ])('resolves merged cells: %s', (_name, body, expected) => {
    expectWidths(columns(body).outer, expected);
  });

  describe('nested tables', () => {
    const nested = (tableWidth: string, cells: readonly [string, Width?][], layout = '') =>
      table(tableWidth, [50, 50], [cells.map(([text, width]) => cell(text, width ?? null)).join('')], layout);
    const outer = (tableWidth: string, first: Width, second: Width, inner: string) =>
      table(tableWidth, [72, 108], [cell(TWO, first, '', inner) + cell(TWO, second)]);

    it('shrinks a nested table with preferred cell widths into its cell', () => {
      const result = columns(outer(AUTO, 72, 108, nested(AUTO, [['n1', 50], ['n2', 50]])));
      expectWidths(result.outer, [72, 108]);
      expectWidths(result.nested, [36, 36]);
    });

    it('widens the cell by nested content that cannot shrink', () => {
      const inner = nested(AUTO, [[word(64)], ['n2']]);
      const result = columns(outer(AUTO, 72, 108, inner));
      expectWidths(result.nested, [63, 10.5]);
      expectWidths(result.outer, [73.5, 108]);
      // Under a preferred table width the other cell gives way.
      expectWidths(columns(outer(dxa(180), 72, 108, inner)).outer, [73.5, 106.5]);
    });

    it('gives a cell without a width the unconstrained width of its nested table', () => {
      const inner = nested(AUTO, [['n1', 50], ['n2', 50]]);
      const result = columns(outer(dxa(180), null, null, inner));
      expectWidths(result.nested, [50, 50]);
      expectWidths(result.outer, [180 * 100 / 110.5, 180 * 10.5 / 110.5]);
    });

    it('counts a 100% nested table by its content', () => {
      const inner = nested(pct(100), [['n1'], ['n2']]);
      expectWidths(columns(outer(AUTO, null, null, inner)).outer, [21, 10.5]);
      // In a cell with a width it fills the cell.
      expectWidths(columns(outer(AUTO, 72, 108, inner)).nested, [36, 36]);
    });

    it('holds a fixed nested table at its width', () => {
      const inner = nested(dxa(100), [['n1', 50], ['n2', 50]], '<w:tblLayout w:type="fixed"/>');
      const result = columns(outer(AUTO, 72, 108, inner));
      expectWidths(result.nested, [50, 50]);
      expectWidths(result.outer, [100, 108]);
    });

    it('sizes nested percentage cells from their content', () => {
      const inner = nested(AUTO, [['n1', '40%'], ['n2', '60%']]);
      const result = columns(table(AUTO, [144, 108], [cell(TWO, 144, '', inner) + cell(TWO, 108)]));
      expectWidths(result.nested, [10.5, 15.75]);
      expectWidths(result.outer, [144, 108]);
    });
  });

  describe('outside the rule', () => {
    const below = row(dxa(150), [TWO, 72], [TWO, 108]);

    it('keeps the proportional reduction outside compatibility mode 15', () => {
      expectWidths(columns(below, 14).outer, [60, 90]);
      expectWidths(columns(below, null).outer, [60, 90]);
    });

    it('keeps fixed tables on their widths', () => {
      expectWidths(columns(table(dxa(150), [72, 108], [cell(TWO, 72) + cell(TWO, 108)], '<w:tblLayout w:type="fixed"/>')).outer, [60, 90]);
    });

    it('keeps the saved grid of an indented or centred table', () => {
      const mixed = (extra: string) => `<w:tbl><w:tblPr>${AUTO}${extra}${BORDERS}</w:tblPr><w:tblGrid>`
        + `<w:gridCol w:w="1440"/><w:gridCol w:w="2160"/></w:tblGrid><w:tr>${cell(TWO, 72)}${cell(TWO)}</w:tr></w:tbl>`;
      expectWidths(columns(mixed('')).outer, [72, 10.5]);
      expect(columns(mixed('<w:jc w:val="center"/>')).outer[1]).toBeGreaterThan(100);
      expect(columns(mixed('<w:tblInd w:w="400" w:type="dxa"/>')).outer[1]).toBeGreaterThan(100);
    });

    it('keeps the saved grid of a table with cell margins', () => {
      const margins = '<w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>';
      const mixed = `<w:tbl><w:tblPr>${AUTO}${BORDERS}${margins}</w:tblPr><w:tblGrid>`
        + `<w:gridCol w:w="1440"/><w:gridCol w:w="2160"/></w:tblGrid><w:tr>${cell(TWO, 72)}${cell(TWO)}</w:tr></w:tbl>`;
      expect(columns(mixed).outer[1]).toBeGreaterThan(100);
    });
  });
});
