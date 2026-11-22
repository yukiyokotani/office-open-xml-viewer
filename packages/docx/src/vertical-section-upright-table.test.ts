/// <reference types="node" />

import { readFile } from 'node:fs/promises';
import { beforeAll, describe, expect, it } from 'vitest';
import init, { DocxArchive } from './wasm/docx_parser.js';
import type { DocxDocumentModel } from './types.js';
import { normalizeInternalDocumentModel } from './parser-model.js';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { storeZip } from './conformance/generate.js';

// `word-vertical-section-upright-block-table`: a block table of a tbRl or
// tbLrV section (ECMA-376 §17.6.20) is an upright, zero-advance block at the
// column top and the current block position, and every line strip whose block
// interval meets the table's (closed interval) starts below it, including the
// strip laid out just before the table.
//
// Geometry of the Word controls: A4, margins left 40, right 90, top 60,
// bottom 80 pt, an 18 pt line grid, 10.5 pt text and a 2x2 single-bordered
// table 72 + 108 pt wide. The table's height comes from the test metrics, so
// every expectation below is stated against the laid-out table box.

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PAGE_W = 11906 / 20;
const MARGIN = { left: 40, right: 90, top: 60 };
const PITCH = 18;
/** Physical y of the column end: page height less the 80 pt bottom margin. */
const COLUMN_END = 16838 / 20 - 80;
const TABLE_W = 180;
const encoder = new TextEncoder();
const xml = (body: string) => encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${body}`);

const paragraph = (text: string) =>
  `<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr><w:r><w:rPr><w:sz w:val="21"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;
const P = paragraph('あいうえお');
const cell = (width: number, text: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>${paragraph(text)}</w:tc>`;
const table = (rows: number, rowHeightPt?: number) => '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>'
  + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
    .map((edge) => `<w:${edge} w:val="single" w:sz="4"/>`).join('')
  + '</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="1440"/><w:gridCol w:w="2160"/></w:tblGrid>'
  + Array.from({ length: rows }, (_, row) =>
    `<w:tr>${rowHeightPt === undefined ? '' : `<w:trPr><w:trHeight w:val="${rowHeightPt * 20}" w:hRule="atLeast"/></w:trPr>`}`
    + `${cell(1440, `A${row + 1}`)}${cell(2160, `B${row + 1}`)}</w:tr>`).join('')
  + '</w:tbl>';
const T = table(2);

type Direction = 'tbRl' | 'tbLrV';

function docx(direction: Direction, body: readonly string[], columns = 1): Uint8Array {
  const document = `<w:document xmlns:w="${W}"><w:body>${body.join('')}`
    + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
    + '<w:pgMar w:top="1200" w:right="1800" w:bottom="1600" w:left="800" w:header="720" w:footer="720" w:gutter="0"/>'
    + (columns > 1 ? `<w:cols w:num="${columns}" w:space="425"/>` : '')
    + `<w:textDirection w:val="${direction}"/><w:docGrid w:type="lines" w:linePitch="360"/>`
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

interface Page {
  /** Upright table boxes on the physical page, in source order. */
  readonly tables: Array<Readonly<{ left: number; top: number; bottom: number }>>;
  /** Line strips per body paragraph, in source order: the strip's distance
   * from the first strip of the page along the block axis, and the physical
   * y at which its inline flow starts. */
  readonly strips: Array<Array<Readonly<{ block: number; inlineStart: number }>>>;
}

function layOut(direction: Direction, body: readonly string[]): Page {
  const pages = layOutPages(direction, body);
  expect(pages).toHaveLength(1);
  return pages[0]!;
}

function layOutPages(direction: Direction, body: readonly string[], columns = 1): Page[] {
  const archive = new DocxArchive(docx(direction, body, columns));
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
  // The section-logical block origin: the right margin in tbRl (lines advance
  // leftward from it), the left margin in tbLrV.
  const blockOrigin = direction === 'tbRl' ? MARGIN.right : MARGIN.left;
  // Each column's strips restart at the block origin.
  return layout.pages.map((laidOut) => {
  const page: Page = { tables: [], strips: [] };
  for (const node of laidOut.layers.body) {
    if (node.kind === 'table') {
      page.tables.push({
        left: node.flowBounds.xPt,
        top: node.flowBounds.yPt,
        bottom: node.flowBounds.yPt + node.advancePt,
      });
    } else if (node.kind === 'paragraph') {
      page.strips.push(node.lines.map((line) => ({
        block: line.bounds.yPt - blockOrigin,
        inlineStart: line.bounds.xPt,
      })));
    }
  }
  return page;
  });
}

/** Width of the tables' border rules (w:sz 4). */
const RULE = 0.5;
/** Physical left grid edge of a table whose box starts at block position
 * `block`: the outer edge of its near rule is there, the grid half a rule on. */
const tableLeft = (direction: Direction, block: number) => direction === 'tbRl'
  ? PAGE_W - MARGIN.right - block - RULE / 2 - TABLE_W
  : MARGIN.left + block + RULE / 2;

beforeAll(async () => {
  await init({ module_or_path: await readFile(new URL('./wasm/docx_parser_bg.wasm', import.meta.url)) });
});

describe('upright block table of a vertical section', () => {
  it.each(['tbRl', 'tbLrV'] as const)(
    'sits at the column top and the block position of a %s section without advancing the flow',
    (direction) => {
      const { tables, strips } = layOut(direction, [P, T, P, P]);
      expect(tables).toHaveLength(1);
      const table = tables[0]!;
      expect(table.top).toBeCloseTo(MARGIN.top, 6);
      expect(table.bottom).toBeGreaterThan(table.top);
      expect(table.left).toBeCloseTo(tableLeft(direction, PITCH), 6);
      // The paragraphs after the table take the next strips, and all three
      // strips meet the table, the one before it at its closed near edge.
      expect(strips.map((lines) => lines.map((line) => line.block))).toEqual([[0], [PITCH], [2 * PITCH]]);
      for (const [line] of strips) expect(line!.inlineStart).toBeCloseTo(table.bottom, 6);
    },
  );

  it('wraps, of the strips laid out before the table, only the one that ends where it starts', () => {
    const { tables, strips } = layOut('tbLrV', [P, P, P, T, P]);
    const table = tables[0]!;
    expect(table.left).toBeCloseTo(tableLeft('tbLrV', 3 * PITCH), 6);
    expect(strips.map(([line]) => line!.inlineStart)).toEqual([
      MARGIN.top,
      MARGIN.top,
      table.bottom,
      table.bottom,
    ]);
  });

  it('places a table that meets an earlier table below it, with the strips meeting both below the lower one', () => {
    const { tables, strips } = layOut('tbRl', [P, T, P, T, P]);
    expect(tables).toHaveLength(2);
    const [first, second] = tables as [Page['tables'][number], Page['tables'][number]];
    expect(first.top).toBeCloseTo(MARGIN.top, 6);
    expect(first.left).toBeCloseTo(tableLeft('tbRl', PITCH), 6);
    expect(second.top).toBeCloseTo(first.bottom, 6);
    expect(second.left).toBeCloseTo(tableLeft('tbRl', 2 * PITCH), 6);
    expect(strips.map((lines) => lines.map((line) => line.block))).toEqual([[0], [PITCH], [2 * PITCH]]);
    // The first strip meets only the first table; the second ends where the
    // second table starts and lies inside the first.
    expect(strips.map(([line]) => line!.inlineStart)).toEqual([first.bottom, second.bottom, second.bottom]);
  });

  it('meets the strip that starts at its far edge and not the one after', () => {
    const { tables, strips } = layOut('tbRl', [T, ...Array.from({ length: 12 }, () => P)]);
    const starts = strips.map(([line]) => line!.inlineStart);
    // Strips 0–9 lie inside the 180 pt table and strip 10 starts at its far edge.
    expect(starts.slice(0, 11)).toEqual(Array.from({ length: 11 }, () => tables[0]!.bottom));
    expect(starts[11]).toBe(MARGIN.top);
  });

  it('moves a table taller than the column to a fresh page and continues its rows on the next', () => {
    const pages = layOutPages('tbRl', [P, table(80), P]);
    expect(pages).toHaveLength(3);
    expect(pages[0]!.tables).toHaveLength(0);
    expect(pages[0]!.strips).toEqual([[{ block: 0, inlineStart: MARGIN.top }]]);
    // Both fragments start at the column top and the flow start of their page.
    for (const page of pages.slice(1)) {
      expect(page.tables).toHaveLength(1);
      expect(page.tables[0]!.top).toBeCloseTo(MARGIN.top, 6);
      expect(page.tables[0]!.left).toBeCloseTo(tableLeft('tbRl', 0), 6);
    }
    expect(pages[1]!.tables[0]!.bottom).toBeLessThanOrEqual(COLUMN_END);
    expect(pages[1]!.tables[0]!.bottom).toBeGreaterThan(COLUMN_END - 2 * PITCH);
    expect(pages[1]!.strips).toEqual([]);
    // The paragraph after the table takes the first strip beside its last fragment.
    expect(pages[2]!.strips).toEqual([[{ block: 0, inlineStart: pages[2]!.tables[0]!.bottom }]]);
  });

  it('places a table that meets only a stacked table at the first free position from the column top', () => {
    // X covers strips 1–10, Y (five strips later) stacks under it, and Z
    // (six strips after Y) meets Y but not X.
    const strips = (count: number) => Array.from({ length: count }, () => P);
    const { tables } = layOut('tbRl', [P, T, ...strips(5), T, ...strips(6), T, P]);
    const [x, y, z] = tables as [Page['tables'][number], Page['tables'][number], Page['tables'][number]];
    expect(y.top).toBeCloseTo(x.bottom, 6);
    expect(z.left).toBeCloseTo(tableLeft('tbRl', 12 * PITCH), 6);
    expect(z.top).toBeCloseTo(MARGIN.top, 6);
  });

  it('moves a table that does not fit below the table it meets past that table, at the column top', () => {
    // 45 + 25 rows are taller than the band; each fits it alone.
    const { tables, strips } = layOut('tbRl', [P, table(45), P, table(25), P]);
    const [x, y] = tables as [Page['tables'][number], Page['tables'][number]];
    expect(x.left).toBeCloseTo(tableLeft('tbRl', PITCH), 6);
    expect(y.top).toBeCloseTo(MARGIN.top, 6);
    // X's box ends at block 198 plus its two half rules; Y's box clears it by
    // the measured 14.28 pt.
    expect(y.left).toBeCloseTo(tableLeft('tbRl', PITCH + TABLE_W + RULE + 14.28), 6);
    // The flow is not advanced: the paragraphs keep consecutive strips.
    expect(strips.map((lines) => lines.map((line) => line.block))).toEqual([[0], [PITCH], [2 * PITCH]]);
  });

  it.each([
    ['a table indent', '<w:tblInd w:w="200" w:type="dxa"/>'],
    ['center alignment', '<w:jc w:val="center"/>'],
    ['right alignment', '<w:jc w:val="right"/>'],
  ])('leaves a table with %s at the block position', (_name, property) => {
    const { tables } = layOut('tbRl', [P, T.replace('<w:tblBorders>', `${property}<w:tblBorders>`), P]);
    expect(tables[0]!.left).toBeCloseTo(tableLeft('tbRl', PITCH), 6);
  });

  it('places the outer edge of the near border rule at the block position', () => {
    const thick = T.replaceAll('w:sz="4"', 'w:sz="24"');
    const { tables, strips } = layOut('tbRl', [P, thick, P]);
    // 3 pt rules: the grid starts 1.5 pt past the block position.
    expect(tables[0]!.left).toBeCloseTo(PAGE_W - MARGIN.right - PITCH - 1.5 - TABLE_W, 6);
    expect(strips.map(([line]) => line!.inlineStart)).toEqual([tables[0]!.bottom, tables[0]!.bottom]);
  });

  it('resumes text with no room beside a table past its far edge, at the column top', () => {
    // 63 rows leave less than one character of room below the table.
    const { tables, strips } = layOut('tbRl', [table(63), P, P]);
    expect(COLUMN_END - tables[0]!.bottom).toBeLessThan(10.5);
    expect(strips.map(([line]) => line!.inlineStart)).toEqual([MARGIN.top, MARGIN.top]);
    // The measured 7.14 pt past the box's far edge, then one strip further.
    expect(strips[0]![0]!.block).toBeCloseTo(TABLE_W + RULE + 7.14, 6);
    expect(strips[1]![0]!.block).toBeCloseTo(TABLE_W + RULE + 7.14 + PITCH, 6);
  });

  it.each(['tbRl', 'tbLrV'] as const)('stacks a later table below the rows a %s table continues with on the next page', (direction) => {
    const pages = layOutPages(direction, [table(80), P, T, P]);
    expect(pages).toHaveLength(2);
    const [rest, later] = pages[1]!.tables as [Page['tables'][number], Page['tables'][number]];
    expect(rest.top).toBeCloseTo(MARGIN.top, 6);
    expect(later.top).toBeCloseTo(rest.bottom, 6);
    expect(later.left).toBeCloseTo(tableLeft(direction, PITCH), 6);
    // Both strips meet the continued rows and the later table.
    expect(pages[1]!.strips).toEqual([
      [{ block: 0, inlineStart: later.bottom }],
      [{ block: PITCH, inlineStart: later.bottom }],
    ]);
  });

  it('keeps rows whole: a row that fits a fresh page is not cut at the page end', () => {
    // Two 450 pt rows in a 701.9 pt column.
    const pages = layOutPages('tbRl', [table(2, 450)]);
    expect(pages).toHaveLength(2);
    for (const page of pages) {
      expect(page.tables).toHaveLength(1);
      expect(page.tables[0]!.top).toBeCloseTo(MARGIN.top, 6);
      expect(page.tables[0]!.bottom - page.tables[0]!.top).toBeGreaterThanOrEqual(450);
      expect(page.tables[0]!.bottom).toBeLessThanOrEqual(COLUMN_END);
    }
  });

  it('moves a keepNext paragraph with a table that does not fit the block extent left', () => {
    // 19 strips leave 123.3 pt: the keepNext strip and the 180 pt table need 198.
    const keepNext = P.replace('<w:pPr>', '<w:pPr><w:keepNext/>');
    const pages = layOutPages('tbRl', [...Array.from({ length: 19 }, () => P), keepNext, T, P]);
    expect(pages).toHaveLength(2);
    expect(pages[0]!.strips).toHaveLength(19);
    expect(pages[1]!.tables[0]!.left).toBeCloseTo(tableLeft('tbRl', PITCH), 6);
    expect(pages[1]!.strips.map(([line]) => line!.block)).toEqual([0, PITCH]);
  });

  describe('in a two-column section', () => {
    it('settles tables whose strips cannot wrap beside them', () => {
      const pages = layOutPages('tbRl', [P, table(30), P, table(30), P], 2);
      expect(pages.flatMap((page) => page.tables)).toHaveLength(2);
      expect(pages.flatMap((page) => page.strips)).toHaveLength(3);
    });

    it('leaves a table wider than the block extent left where the flow reached it', () => {
      const pages = layOutPages('tbRl', [...Array.from({ length: 20 }, () => P), T, P], 2);
      expect(pages).toHaveLength(1);
      expect(pages[0]!.tables[0]!.left).toBeCloseTo(tableLeft('tbRl', 20 * PITCH), 6);
      expect(pages[0]!.tables[0]!.top).toBeCloseTo(MARGIN.top, 6);
    });

    it('measures a tall table against the page band, not its column', () => {
      // 40 rows are taller than one column and shorter than the page band:
      // the table stays whole where the flow reached it.
      const pages = layOutPages('tbRl', [P, table(40), P], 2);
      expect(pages).toHaveLength(1);
      const [tall] = pages[0]!.tables;
      expect(tall!.top).toBeCloseTo(MARGIN.top, 6);
      expect(tall!.bottom).toBeGreaterThan((MARGIN.top + COLUMN_END) / 2);
      expect(tall!.bottom).toBeLessThanOrEqual(COLUMN_END);
    });

    it.each(['tbRl', 'tbLrV'] as const)('keeps a %s table taller than the page whole and raises it to the page top', (direction) => {
      const pages = layOutPages(direction, [P, table(80), P], 2);
      expect(pages).toHaveLength(1);
      expect(pages[0]!.tables).toHaveLength(1);
      expect(pages[0]!.tables[0]!.left).toBeCloseTo(tableLeft(direction, PITCH), 6);
      expect(pages[0]!.tables[0]!.top).toBe(0);
    });
  });

  it('settles two tables nearly as tall as the column, with no room for the strips before them', () => {
    // The strip before each table cannot wrap beside it; the upright block
    // plan pins each table once and ends.
    const pages = layOutPages('tbRl', [P, table(63), P, table(63), P]);
    expect(pages.flatMap((page) => page.tables).length).toBeGreaterThanOrEqual(2);
    expect(pages.flatMap((page) => page.strips)).toHaveLength(3);
  });

  it('keeps the table where the full-length strip before it left the flow', () => {
    // `word-vertical-section-upright-table-rewrapped-strip`: 66 characters of
    // 10.5 pt fill the 701.9 pt column in one strip, but not beside the table.
    const { tables, strips } = layOut('tbRl', [paragraph('長'.repeat(66)), T, P]);
    const table = tables[0]!;
    expect(table.left).toBeCloseTo(tableLeft('tbRl', PITCH), 6);
    expect(strips.map((lines) => lines.map((line) => line.block))).toEqual([[0, PITCH], [2 * PITCH]]);
    for (const line of strips.flat()) expect(line.inlineStart).toBeCloseTo(table.bottom, 6);
  });
});
