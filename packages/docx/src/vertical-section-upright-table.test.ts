import { beforeAll, describe, expect, it } from 'vitest';
import {
  COLUMN_END,
  MARGIN,
  P,
  PAGE_W,
  PITCH,
  T,
  TABLE_W,
  layOutPages,
  loadParser,
  paragraph,
  table,
  type Direction,
  type Page,
} from './test-support/vertical-section-upright.test-support.js';

// `word-vertical-section-upright-block-table`: a block table of a tbRl or
// tbLrV section (ECMA-376 §17.6.20) is an upright, zero-advance block at the
// column top and the current block position, and every line strip whose block
// interval meets the table's (closed interval) starts below it, including the
// strip laid out just before the table.
//
// The fixtures (test-support/vertical-section-upright.test-support.ts) use the
// geometry of the Word controls. A table's height comes from the test
// metrics, so every expectation below is stated against the laid-out box.

function layOut(direction: Direction, body: readonly string[]): Page {
  const pages = layOutPages(direction, body);
  expect(pages).toHaveLength(1);
  return pages[0]!;
}

/** Width of the tables' border rules (w:sz 4). */
const RULE = 0.5;
/** Physical left grid edge of a table whose box starts at block position
 * `block`: the outer edge of its near rule is there, the grid half a rule on. */
const tableLeft = (direction: Direction, block: number) => direction === 'tbRl'
  ? PAGE_W - MARGIN.right - block - RULE / 2 - TABLE_W
  : MARGIN.left + block + RULE / 2;

beforeAll(loadParser);

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

  it('moves a keepNext paragraph with a table that clearing an earlier table moves to the next page', () => {
    // A 300 pt wide table with a 500 pt row, fifteen strips, then a table
    // with a 350 pt row: its 180 pt box fits the 195.3 pt left, but it meets
    // the first table, does not fit below it and has no room past it on the
    // page. Look-ahead places it as placement does, so the strip kept with it
    // goes along.
    const first = table(1, 500).replaceAll('"1440"', '"2400"').replaceAll('"2160"', '"3600"');
    const keepNext = P.replace('<w:pPr>', '<w:pPr><w:keepNext/>');
    const pages = layOutPages('tbRl', [first, ...Array.from({ length: 14 }, () => P), keepNext, table(1, 350), P]);
    expect(pages).toHaveLength(2);
    expect(pages[0]!.tables).toHaveLength(1);
    expect(pages[0]!.strips).toHaveLength(14);
    expect(pages[1]!.tables).toHaveLength(1);
    expect(pages[1]!.strips.map(([line]) => line!.block)).toEqual([0, PITCH]);
  });

  it('moves a keepNext paragraph first in the flow with a table that clearing moves to the next page', () => {
    // The page starts with a table, which adds no flow advance, so the kept
    // strip has the whole block extent of a fresh page before it. The second
    // table still starts the next page to clear the first, and takes none of
    // that page's block extent: the strip fits there with it.
    const first = table(1, 400).replaceAll('"1440"', '"2400"').replaceAll('"2160"', '"3600"');
    const keepNext = P.replace('<w:pPr>', '<w:pPr><w:keepNext/>');
    const pages = layOutPages('tbRl', [first, keepNext, table(1, 400), P]);
    expect(pages).toHaveLength(2);
    expect(pages[0]!.strips).toHaveLength(0);
    expect(pages[1]!.tables[0]!.left).toBeCloseTo(tableLeft('tbRl', PITCH), 6);
    expect(pages[1]!.strips.map(([line]) => line!.block)).toEqual([0, PITCH]);
  });

  it('wraps a line that a negative indent takes into the first column around a table there', () => {
    // A 350 pt table of column 1 reaches past the column's end. The second
    // column's paragraph starts 30 pt before its band, inside the table's
    // inline extent, and wraps at the table's bottom edge.
    const columnBreak = '<w:p><w:r><w:br w:type="column"/></w:r></w:p>';
    const outdented = P.replace('<w:pPr>', '<w:pPr><w:ind w:left="-600"/>');
    const pages = layOutPages('tbRl', [table(1, 350), columnBreak, outdented], 2);
    const placed = pages[0]!.tables[0]!;
    const line = pages[0]!.strips.at(-1)![0]!;
    expect(line.inlineStart).toBeCloseTo(placed.bottom, 6);
  });

  it('wraps a line that the negative indent of its numbering level takes into the first column', () => {
    // The same with the indent on the list level (§17.9.6), not on the
    // paragraph: the level's -30 pt start indent, with no hanging indent,
    // takes the numbered line into the table's inline extent.
    const columnBreak = '<w:p><w:r><w:br w:type="column"/></w:r></w:p>';
    const level = '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1"/>'
      + '<w:suff w:val="nothing"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="-600" w:hanging="0"/></w:pPr></w:lvl>';
    const numbered = P.replace('<w:pPr>', '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>');
    const body = [table(1, 350), columnBreak, numbered];
    const pages = layOutPages('tbRl', body, 2, 360, level);
    const placed = pages[0]!.tables[0]!;
    // The line wraps at the table's bottom edge; its text follows the
    // one-digit number (10.5 pt in the test metrics).
    expect(pages[0]!.strips.at(-1)![0]!.inlineStart).toBeCloseTo(placed.bottom + 10.5, 6);
    // Without the level's indent the line starts at the top of column 2,
    // past the table.
    const plain = layOutPages('tbRl', body, 2, 360, level.replace('w:left="-600"', 'w:left="0"'));
    expect(plain[0]!.strips.at(-1)![0]!.inlineStart).toBeGreaterThan(placed.bottom + 20);
  });

  it.each(['tbRl', 'tbLrV'] as const)('wraps the strip before a %s table on a line grid of a fractional pitch', (direction) => {
    // A 15.15 pt pitch: the strip's end and the table's start are the same
    // block position computed two ways, equal only to within rounding.
    const [page] = layOutPages(direction, [P, T, P], 1, 303);
    const placed = page!.tables[0]!;
    expect(page!.strips.map(([line]) => line!.inlineStart)).toEqual([placed.bottom, placed.bottom]);
  });

  it('moves a table whose border rules pass the block extent left to the next page', () => {
    // Fifteen strips leave 195.3 pt. The 195 pt grid fits; its box with two
    // 3 pt side rules is 198 pt wide and does not.
    const wide = T.replaceAll('w:sz="4"', 'w:sz="24"').replaceAll('"2160"', '"2460"');
    const pages = layOutPages('tbRl', [...Array.from({ length: 15 }, () => P), wide, P]);
    expect(pages).toHaveLength(2);
    expect(pages[0]!.tables).toHaveLength(0);
    expect(pages[1]!.tables).toHaveLength(1);
  });

  it('moves a keepNext paragraph with a table whose border rules pass the block extent left', () => {
    // After fifteen strips 195.3 pt are left: the 195 pt grid fits, the
    // 198 pt box does not, so the fifteenth strip goes with the table.
    const wide = T.replaceAll('w:sz="4"', 'w:sz="24"').replaceAll('"2160"', '"2460"');
    const keepNext = P.replace('<w:pPr>', '<w:pPr><w:keepNext/>');
    const pages = layOutPages('tbRl', [...Array.from({ length: 14 }, () => P), keepNext, wide, P]);
    expect(pages).toHaveLength(2);
    expect(pages[0]!.strips).toHaveLength(14);
    expect(pages[1]!.tables).toHaveLength(1);
    expect(pages[1]!.strips[0]![0]!.block).toBe(0);
  });

  it('takes the box from the outermost painted rules under cell spacing', () => {
    // No table borders, 6 pt cell rules and 2 pt cell spacing; one 696 pt row
    // leaves no room for text below the table.
    const cellRules = '<w:tcBorders>' + ['top', 'left', 'bottom', 'right']
      .map((edge) => `<w:${edge} w:val="single" w:sz="48"/>`).join('') + '</w:tcBorders>';
    const spaced = table(1, 696).replace(/<w:tblBorders>.*?<\/w:tblBorders>/, '<w:tblCellSpacing w:w="40" w:type="dxa"/>')
      .replaceAll('</w:tcPr>', `${cellRules}</w:tcPr>`);
    const { tables, strips } = layOut('tbRl', [spaced, P]);
    const [placed] = tables;
    expect(COLUMN_END - placed!.bottom).toBeLessThan(10.5);
    // The near rule's outer edge is at the block position, and the text
    // resumes 7.14 pt past the outer edge of the far rule.
    expect(placed!.paintedRight).toBeCloseTo(PAGE_W - MARGIN.right, 6);
    expect(strips[0]![0]!.inlineStart).toBe(MARGIN.top);
    expect(strips[0]![0]!.block).toBeCloseTo(placed!.paintedRight - placed!.paintedLeft + 7.14, 6);
  });

  it('keeps a continued fragment at the offset of the whole table\'s border rules', () => {
    // Forty 20 pt rows without table borders; only the first row's first cell
    // has a 3 pt left rule. The rows continued on page 2 keep its offset.
    const row = (index: number) => '<w:tr><w:trPr><w:trHeight w:val="400" w:hRule="atLeast"/></w:trPr>'
      + '<w:tc><w:tcPr><w:tcW w:w="3600" w:type="dxa"/>'
      + (index === 0 ? '<w:tcBorders><w:left w:val="single" w:sz="24"/></w:tcBorders>' : '')
      + `</w:tcPr>${paragraph(`R${index}`)}</w:tc></w:tr>`;
    const long = '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3600"/></w:tblGrid>'
      + Array.from({ length: 40 }, (_, index) => row(index)).join('') + '</w:tbl>';
    const pages = layOutPages('tbLrV', [long]);
    expect(pages).toHaveLength(2);
    expect(pages[0]!.tables[0]!.left).toBeCloseTo(MARGIN.left + 1.5, 6);
    expect(pages[1]!.tables[0]!.left).toBeCloseTo(MARGIN.left + 1.5, 6);
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

    it('moves a tall table that meets an earlier one past it, then raises it', () => {
      // The 75-row table would pass the page edge below the 2-row table it
      // meets: it clears that table along the block axis and is raised.
      const pages = layOutPages('tbRl', [P, T, P, table(75), P], 2);
      expect(pages).toHaveLength(1);
      const [short, tall] = pages[0]!.tables as [Page['tables'][number], Page['tables'][number]];
      expect(short.top).toBeCloseTo(MARGIN.top, 6);
      expect(tall.left).toBeCloseTo(tableLeft('tbRl', PITCH + TABLE_W + RULE + 14.28), 6);
      expect(tall.top).toBeLessThan(MARGIN.top);
      expect(tall.bottom).toBeCloseTo(16838 / 20, 6);
    });

    it('keeps a table that directly follows another below it', () => {
      // Two tables of one 800 pt row each with nothing between them run on as
      // one block: the second is neither moved aside nor raised over the
      // first, and the block is raised by its whole height.
      const pages = layOutPages('tbRl', [table(1, 800), table(1, 800)], 2);
      expect(pages).toHaveLength(1);
      const [first, second] = pages[0]!.tables as [Page['tables'][number], Page['tables'][number]];
      expect(second.left).toBeCloseTo(first.left, 6);
      expect(second.top).toBeCloseTo(first.bottom, 6);
      // Together they are taller than the page, so the pair starts at its top.
      expect(first.top).toBe(0);
    });

    it('keeps a tall table below a short one it directly follows', () => {
      // The 20 pt table is raised with the 830 pt one that runs on from it;
      // the tall one starts where the short one ends.
      const pages = layOutPages('tbRl', [table(1, 20), table(1, 830)], 2);
      const [short, tall] = pages[0]!.tables as [Page['tables'][number], Page['tables'][number]];
      expect(short.top).toBe(0);
      expect(tall.left).toBeCloseTo(short.left, 6);
      expect(tall.top).toBeCloseTo(short.bottom, 6);
    });

    it('stacks a long run of directly following tables in order', () => {
      const pages = layOutPages('tbRl', Array.from({ length: 48 }, () => T), 2);
      const tables = pages.flatMap((page) => page.tables);
      expect(tables).toHaveLength(48);
      tables.slice(1).forEach((next, index) => {
        expect(next.top).toBeCloseTo(tables[index]!.bottom, 6);
      });
    });

    it('does not raise a table of the second column into a table of the first', () => {
      // Measured: the 800 pt table of column 2 is raised above its column
      // and would cross the 20 pt table of column 1 at the same block
      // position: it clears that table along the block axis first.
      const columnBreak = '<w:p><w:r><w:br w:type="column"/></w:r></w:p>';
      const pages = layOutPages('tbRl', [table(1, 20), columnBreak, table(1, 800)], 2);
      const [short, tall] = pages[0]!.tables as [Page['tables'][number], Page['tables'][number]];
      expect(tall.top).toBeLessThan(short.bottom);
      expect(tall.left).toBeCloseTo(tableLeft('tbRl', TABLE_W + RULE + 14.28), 6);
    });

    it('clears a table of the second column that lies under a raised table of the first', () => {
      // Measured: the 800 pt table of column 1 is raised over both columns.
      // The empty strip the column-break paragraph leaves in column 2 has no
      // room beside it and resumes 7.14 pt past it; the 20 pt table follows
      // one strip further on, at the top of column 2, and the paragraph after
      // it starts at its bottom edge.
      const columnBreak = '<w:p><w:r><w:br w:type="column"/></w:r></w:p>';
      const pages = layOutPages('tbRl', [table(1, 800), columnBreak, table(1, 20), P], 2);
      expect(pages).toHaveLength(1);
      const [tall, short] = pages[0]!.tables as [Page['tables'][number], Page['tables'][number]];
      expect(tall.left).toBeCloseTo(tableLeft('tbRl', 0), 6);
      expect(tall.bottom).toBeCloseTo(16838 / 20, 6);
      expect(short.left).toBeCloseTo(tableLeft('tbRl', TABLE_W + RULE + 7.14 + PITCH), 6);
      expect(short.top).toBeGreaterThan((MARGIN.top + COLUMN_END) / 2);
      const after = pages[0]!.strips.at(-1)![0]!;
      expect(after.block).toBeCloseTo(TABLE_W + RULE + 7.14 + PITCH, 6);
      expect(after.inlineStart).toBeCloseTo(short.bottom, 6);
    });

    it.each(['tbRl', 'tbLrV'] as const)('clears and raises a %s table that meets a table of its own second column', (direction) => {
      // Measured: both tables lie in column 2, whose first strip is the empty
      // one of the column-break paragraph. The 800 pt table clears the 20 pt
      // one by 14.28 pt along the block axis and is raised to the page edge;
      // the paragraphs beside the 20 pt table start at its bottom edge.
      const columnBreak = '<w:p><w:r><w:br w:type="column"/></w:r></w:p>';
      const pages = layOutPages(direction, [P, columnBreak, table(1, 20), P, table(1, 800), P], 2);
      expect(pages).toHaveLength(1);
      const [short, tall] = pages[0]!.tables as [Page['tables'][number], Page['tables'][number]];
      expect(short.left).toBeCloseTo(tableLeft(direction, PITCH), 6);
      expect(tall.left).toBeCloseTo(tableLeft(direction, PITCH + TABLE_W + RULE + 14.28), 6);
      expect(tall.bottom).toBeCloseTo(16838 / 20, 6);
      const [between, after] = pages[0]!.strips.slice(-2) as [Page['strips'][number], Page['strips'][number]];
      expect(between[0]!.inlineStart).toBeCloseTo(short.bottom, 6);
      expect(after[0]!.inlineStart).toBeCloseTo(short.bottom, 6);
      expect([between[0]!.block, after[0]!.block]).toEqual([PITCH, 2 * PITCH]);
    });

    it('keeps a table below the one it directly follows when that one was moved aside', () => {
      // The 800 pt table clears the 20 pt table before it along the block
      // axis; the table that directly follows runs on from it there.
      const pages = layOutPages('tbRl', [table(1, 20), P, table(1, 800), table(1, 20)], 2);
      const [, tall, follower] = pages[0]!.tables as [unknown, Page['tables'][number], Page['tables'][number]];
      expect(tall.left).toBeCloseTo(tableLeft('tbRl', TABLE_W + RULE + 14.28), 6);
      expect(follower.left).toBeCloseTo(tall.left, 6);
      expect(follower.top).toBeCloseTo(tall.bottom, 6);
    });

    it('does not move a keepNext paragraph and a tall table to the second column', () => {
      const keepNext = P.replace('<w:pPr>', '<w:pPr><w:keepNext/>');
      const pages = layOutPages('tbLrV', [P, keepNext, table(1, 800)], 2);
      expect(pages).toHaveLength(1);
      expect(pages[0]!.tables[0]!.left).toBeCloseTo(tableLeft('tbLrV', 2 * PITCH), 6);
      // Both paragraphs stay in the first column, whose strips start at the top margin.
      expect(pages[0]!.strips.flat().map((line) => line.inlineStart)).toEqual([MARGIN.top, MARGIN.top]);
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

  describe('placement in one forward pass', () => {
    /** Pairs of tables of one page whose painted boxes share an area. */
    const overlapping = (pages: readonly Page[]) => pages.flatMap((page, pageIndex) =>
      page.tables.flatMap((first, at) => page.tables.slice(at + 1).flatMap((second) => {
        const shared = (startA: number, endA: number, startB: number, endB: number) =>
          Math.min(endA, endB) - Math.max(startA, startB) > 1e-6;
        return shared(first.paintedLeft, first.paintedRight, second.paintedLeft, second.paintedRight)
          && shared(first.top, first.bottom, second.top, second.bottom)
          ? [{ pageIndex, first, second }] : [];
      })));
    const strips = (count: number) => Array.from({ length: count }, () => P);

    it('places a table, a paragraph and two tables that run on in two columns', () => {
      // A table placed later never changes the text before the table it runs
      // on from, so nothing is placed twice: the 20 pt and 750 pt tables are
      // one block, which does not fit below the first table and clears it
      // along the block axis, at the column top.
      const pages = layOutPages('tbRl', [table(1, 20), P, table(1, 20), table(1, 750)], 2);
      expect(pages).toHaveLength(1);
      const [first, leader, follower] = pages[0]!.tables as [
        Page['tables'][number], Page['tables'][number], Page['tables'][number],
      ];
      expect(overlapping(pages)).toEqual([]);
      expect(follower.top).toBeCloseTo(leader.bottom, 6);
      expect(follower.left).toBeCloseTo(leader.left, 6);
      expect(first.left).toBeCloseTo(tableLeft('tbRl', 0), 6);
      expect(leader.top).toBeCloseTo(MARGIN.top, 6);
      expect(leader.left).toBeCloseTo(tableLeft('tbRl', TABLE_W + RULE + 14.28), 6);
    });

    it.each([1, 2])('keeps a table that runs on clear of an earlier table in %i column(s)', (columns) => {
      // X at the column top, Y stacked under it, then Z and a table that
      // directly follows Z, which both meet Y but not X.
      const pages = layOutPages('tbRl', [P, T, ...strips(5), T, ...strips(6), T, T], columns);
      expect(pages).toHaveLength(1);
      expect(pages[0]!.tables).toHaveLength(4);
      expect(overlapping(pages)).toEqual([]);
      const [x, y, z, follower] = pages[0]!.tables as [
        Page['tables'][number], Page['tables'][number], Page['tables'][number], Page['tables'][number],
      ];
      expect(y.top).toBeCloseTo(x.bottom, 6);
      // Z and the table after it are laid out as one table: it does not fit
      // above Y and is placed below it.
      expect(follower.left).toBeCloseTo(z.left, 6);
      expect(z.top).toBeCloseTo(y.bottom, 6);
      expect(follower.top).toBeCloseTo(z.bottom, 6);
    });

    it('lays out generated table and paragraph sequences without overlapping tables', () => {
      // A fixed pseudo-random stream (mulberry32) of short and full-length
      // paragraphs, column breaks and tables from two rows to taller than the
      // page, in one and two columns of both directions. Every sequence has
      // to lay out (pagination ends) with no two table boxes of a page
      // sharing an area, and with every table present.
      let seed = 0x2326;
      const random = () => {
        seed = (seed + 0x6d2b79f5) | 0;
        let value = Math.imul(seed ^ (seed >>> 15), 1 | seed);
        value = (value + Math.imul(value ^ (value >>> 7), 61 | value)) ^ value;
        return ((value ^ (value >>> 14)) >>> 0) / 4294967296;
      };
      const pick = <Item>(items: readonly Item[]) => items[Math.floor(random() * items.length)]!;
      const thick = T.replaceAll('w:sz="4"', 'w:sz="24"');
      const tables = [T, T, thick, table(1, 20), table(30), table(45), table(63), table(1, 800), table(80)];
      const paragraphs = [P, P, P, paragraph('長'.repeat(30)), paragraph('長'.repeat(66)),
        P.replace('<w:pPr>', '<w:pPr><w:keepNext/>')];
      const columnBreak = '<w:p><w:r><w:br w:type="column"/></w:r></w:p>';
      for (const direction of ['tbRl', 'tbLrV'] as const) {
        for (const columns of [1, 2]) {
          for (let sample = 0; sample < 10; sample += 1) {
            const body: string[] = [];
            let tableCount = 0;
            const length = 3 + Math.floor(random() * 10);
            for (let at = 0; at < length; at += 1) {
              const kind = random();
              if (kind < 0.45) {
                body.push(pick(tables));
                tableCount += 1;
              } else if (kind < 0.5 && columns > 1) {
                body.push(columnBreak);
              } else {
                for (let run = 1 + Math.floor(random() * 6); run > 0; run -= 1) body.push(pick(paragraphs));
              }
            }
            const context = `${direction} ${columns} column(s) sample ${sample}`;
            const pages = layOutPages(direction, body, columns);
            expect(overlapping(pages), context).toEqual([]);
            // A table continued on later pages has several fragments.
            expect(pages.flatMap((page) => page.tables).length, context).toBeGreaterThanOrEqual(tableCount);
          }
        }
      }
    }, 120_000);
  });

  describe('text laid out before a table that is raised over it', () => {
    const columnBreak = '<w:p><w:r><w:br w:type="column"/></w:r></w:p>';
    const strips = (count: number) => Array.from({ length: count }, () => P);
    /** Block position at which text resumes past a table whose box starts at `block`. */
    const past = (block: number) => block + TABLE_W + RULE + 7.14;

    it.each(['tbRl', 'tbLrV'] as const)('flows the first column of a %s page around a table raised from the second', (direction) => {
      // Measured: the 800 pt table of column 2 is raised over both columns,
      // one strip into the flow (the empty strip of the column-break
      // paragraph). The twelve strips of column 1 have no room beside it and
      // resume past it, as the text of column 2 does.
      const pages = layOutPages(direction, [...strips(12), columnBreak, table(1, 800), P], 2);
      expect(pages).toHaveLength(1);
      const [page] = pages;
      expect(page!.tables[0]!.left).toBeCloseTo(tableLeft(direction, PITCH), 6);
      expect(page!.tables[0]!.bottom).toBeCloseTo(16838 / 20, 6);
      const first = page!.strips.slice(0, 12).map(([line]) => line!);
      first.forEach((line, at) => {
        expect(line.block).toBeCloseTo(past(PITCH) + at * PITCH, 6);
        expect(line.inlineStart).toBe(MARGIN.top);
      });
      // The paragraph after the table: one strip after the empty one.
      expect(page!.strips.at(-1)![0]!.block).toBeCloseTo(past(PITCH) + PITCH, 6);
    });

    it('keeps the table where the flow reaches it when the text before it is laid out without it', () => {
      // Measured: 27 strips fill column 1 (25) and start column 2, so the
      // table lies two strips into column 2. Raised, it takes the room of
      // both columns from there: column 1 keeps one strip before it and
      // thirteen past it, and column 2 holds the rest the same way.
      const pages = layOutPages('tbRl', [...strips(27), table(1, 800), P], 2);
      expect(pages).toHaveLength(1);
      const [page] = pages;
      expect(page!.tables[0]!.left).toBeCloseTo(tableLeft('tbRl', 2 * PITCH), 6);
      const lines = page!.strips.map(([line]) => line!);
      expect(lines).toHaveLength(28);
      const columnTwoTop = lines[14]!.inlineStart;
      expect(columnTwoTop).toBeGreaterThan((MARGIN.top + COLUMN_END) / 2);
      const expected = (at: number) => at === 0 ? 0 : past(2 * PITCH) + (at - 1) * PITCH;
      lines.forEach((line, at) => {
        const inColumn = at < 14 ? at : at - 14;
        expect(line.block, `strip ${at}`).toBeCloseTo(expected(inColumn), 6);
        expect(line.inlineStart, `strip ${at}`).toBe(at < 14 ? MARGIN.top : columnTwoTop);
      });
    });

    it('leaves a table in the first column when the strip before it moves to the second', () => {
      // Measured: 24 strips and a full-length strip fill column 1; the table
      // after them is wider than the block extent left and stays, overflowing.
      // Beside it the full-length strip needs a second line that column 1 has
      // no room for, so the paragraph moves to column 2 and the table stays.
      const full = paragraph('長'.repeat(32));
      const pages = layOutPages('tbLrV', [...strips(24), full, T, P], 2);
      expect(pages).toHaveLength(1);
      const [page] = pages;
      expect(page!.tables[0]!.left).toBeCloseTo(tableLeft('tbLrV', 25 * PITCH), 6);
      expect(page!.tables[0]!.top).toBeCloseTo(MARGIN.top, 6);
      const [moved, after] = page!.strips.slice(-2) as [Page['strips'][number], Page['strips'][number]];
      expect(moved[0]!.block).toBe(0);
      expect(moved[0]!.inlineStart).toBeGreaterThan((MARGIN.top + COLUMN_END) / 2);
      expect(after[0]!.block).toBe(moved.length * PITCH);
    });

    it('starts the next page with a table too wide for the block extent left in the last column', () => {
      // Measured: the same at the end of column 2. The table does not stay
      // there overflowing: it starts page 2, and the strip before it is not
      // wrapped.
      const full = paragraph('長'.repeat(32));
      const pages = layOutPages('tbRl', [P, columnBreak, ...strips(23), full, T, P], 2);
      expect(pages).toHaveLength(2);
      expect(pages[0]!.tables).toHaveLength(0);
      expect(pages[0]!.strips.at(-1)).toHaveLength(1);
      expect(pages[1]!.tables[0]!.left).toBeCloseTo(tableLeft('tbRl', 0), 6);
      expect(pages[1]!.strips[0]![0]!.inlineStart).toBeCloseTo(pages[1]!.tables[0]!.bottom, 6);
    });
  });

  describe('tables that directly follow one another', () => {
    const narrow = (rows: number) => table(rows).replaceAll('"1440"', '"720"').replaceAll('"2160"', '"1080"');

    it.each([
      ['a wide table then a narrow one', (): string[] => [T, narrow(2)], 1],
      ['a narrow table then a wide one', (): string[] => [narrow(2), T], 0],
    ] as const)('lays out %s as one table, left edges aligned', (_name, tables, narrowAt) => {
      // Measured in tbRl and tbLrV: the second table starts where the first
      // ends, both at the left edge of the wider one, and every strip beside
      // either starts below both.
      for (const direction of ['tbRl', 'tbLrV'] as const) {
        const [page] = layOutPages(direction, [P, ...tables(), P, P], 2);
        const [first, second] = page!.tables as [Page['tables'][number], Page['tables'][number]];
        expect(second.top).toBeCloseTo(first.bottom, 6);
        expect(second.left).toBeCloseTo(first.left, 6);
        expect(page!.tables[narrowAt]!.paintedRight - page!.tables[narrowAt]!.paintedLeft).toBeCloseTo(90 + RULE, 6);
        expect(page!.tables[1 - narrowAt]!.left).toBeCloseTo(tableLeft(direction, PITCH), 6);
        for (const [line] of page!.strips) expect(line!.inlineStart).toBeCloseTo(second.bottom, 6);
      }
    });

    it.each(['tbRl', 'tbLrV'] as const)('paginates %s tables taller than the column together as one table', (direction) => {
      // Measured: a paragraph, 40 rows and 25 rows. Together they do not fit
      // the column, so both start the next page; the second is cut between
      // rows at the band end and continues on a third page, where the last
      // paragraph starts at its bottom edge.
      const pages = layOutPages(direction, [P, table(40), table(25), P]);
      expect(pages).toHaveLength(3);
      expect(pages[0]!.tables).toHaveLength(0);
      const [first, second] = pages[1]!.tables as [Page['tables'][number], Page['tables'][number]];
      expect(first.top).toBeCloseTo(MARGIN.top, 6);
      expect(second.top).toBeCloseTo(first.bottom, 6);
      expect(second.bottom).toBeLessThanOrEqual(COLUMN_END);
      expect(second.bottom).toBeGreaterThan(COLUMN_END - 2 * PITCH);
      const [rest] = pages[2]!.tables;
      expect(rest!.top).toBeCloseTo(MARGIN.top, 6);
      expect(pages[2]!.strips).toEqual([[{ block: 0, inlineStart: rest!.bottom }]]);
    });

    it('continues below the last rows of the table before it', () => {
      // Measured: 40, 25 and 2 rows. The third follows the second's last
      // rows on the page they end on.
      const pages = layOutPages('tbRl', [table(40), table(25), T, P]);
      expect(pages).toHaveLength(2);
      const [rest, last] = pages[1]!.tables as [Page['tables'][number], Page['tables'][number]];
      expect(last.top).toBeCloseTo(rest.bottom, 6);
      expect(last.left).toBeCloseTo(rest.left, 6);
      expect(pages[1]!.strips[0]![0]!.inlineStart).toBeCloseTo(last.bottom, 6);
    });
  });

  /** A square-wrapped 80 x 40 pt rectangle anchored at page position (x, 66 pt). */
  const rectangleAt = (x: number) =>
    '<w:r><w:drawing><wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
    + ' distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
    + '<wp:simplePos x="0" y="0"/>'
    + `<wp:positionH relativeFrom="page"><wp:posOffset>${x * 12700}</wp:posOffset></wp:positionH>`
    + `<wp:positionV relativeFrom="page"><wp:posOffset>${66 * 12700}</wp:posOffset></wp:positionV>`
    + `<wp:extent cx="${80 * 12700}" cy="${40 * 12700}"/><wp:wrapSquare wrapText="bothSides"/>`
    + '<wp:docPr id="1" name="Rect"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
    + '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
    + '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:cNvSpPr/><wps:spPr>'
    + `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${80 * 12700}" cy="${40 * 12700}"/></a:xfrm>`
    + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:bodyPr/></wps:wsp>'
    + '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>';

  it.each(['tbRl', 'tbLrV'] as const)('stacks a %s table below a wrapped drawing anchored before it', (direction) => {
    // Measured: a square-wrapped 80 x 40 pt rectangle at a page position over
    // the place the table would take, anchored by the paragraph before it.
    // The table is placed directly below the rectangle; the strip before the
    // table holds its text above the table.
    const drawing = rectangleAt(direction === 'tbRl' ? 400 : 150);
    const [page] = layOutPages(direction, [P.replace('</w:p>', `${drawing}</w:p>`), T, P]);
    const placed = page!.tables[0]!;
    expect(placed.left).toBeCloseTo(tableLeft(direction, PITCH), 6);
    expect(placed.top).toBeCloseTo(66 + 40, 6);
    expect(page!.strips[0]![0]!.inlineStart).toBe(MARGIN.top);
    // The strip after the table lies beside the rectangle in tbRl and starts
    // below the table; in tbLrV the rectangle is further along the block
    // axis and the strip uses the room above the table.
    expect(page!.strips[1]![0]!.inlineStart)
      .toBeCloseTo(direction === 'tbRl' ? placed.bottom : MARGIN.top, 6);
  });

  it.each(['tbRl', 'tbLrV'] as const)('moves a wrapped drawing anchored after a %s table below the table', (direction) => {
    // Measured: the same rectangle anchored by the paragraph after the
    // table. The table keeps the column top and the rectangle lies directly
    // below it; in tbRl the paragraphs beside the rectangle start below it.
    const drawing = rectangleAt(direction === 'tbRl' ? 400 : 150);
    const [page] = layOutPages(direction, [P, T, P.replace('</w:p>', `${drawing}</w:p>`), P]);
    const placed = page!.tables[0]!;
    expect(placed.top).toBeCloseTo(MARGIN.top, 6);
    expect(placed.left).toBeCloseTo(tableLeft(direction, PITCH), 6);
    expect(page!.drawings).toHaveLength(1);
    expect(page!.drawings[0]!.top).toBeCloseTo(placed.bottom, 6);
    expect(page!.drawings[0]!.left).toBeCloseTo(direction === 'tbRl' ? 400 : 150, 6);
    const starts = page!.strips.map(([line]) => line!.inlineStart);
    expect(starts[0]).toBeCloseTo(placed.bottom, 6);
    for (const start of starts.slice(1)) {
      expect(start).toBeCloseTo(direction === 'tbRl' ? placed.bottom + 40 : placed.bottom, 6);
    }
  });
});
