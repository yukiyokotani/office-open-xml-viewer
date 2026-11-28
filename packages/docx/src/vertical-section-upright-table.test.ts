import { beforeAll, describe, expect, it } from 'vitest';
import {
  anchoredRectangle,
  anchoring,
  collisions,
  COLUMN_END,
  ENDNOTE_REFERENCE,
  FOOTNOTE_REFERENCE,
  MARGIN,
  PAGE_H,
  P,
  PAGE_W,
  PITCH,
  T,
  TABLE_W,
  layOutPages,
  loadParser,
  paragraph,
  romanPageTable,
  shared,
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

  it.each(['tbRl', 'tbLrV'] as const)(
    'holds text below a table only in at least 18 pt of room before the column end (%s)',
    (direction) => {
      // A table from the column top that ends 17 pt above the column's end
      // leaves a 10.5 pt character room there, but the strips that meet it
      // have no text below it and resume past it at the column top. In
      // 19 pt of room each holds one character below the table.
      const layOutWithRoom = (roomPt: number) => {
        const pages = layOutPages(direction, [table(1, COLUMN_END - MARGIN.top - 2 * RULE - roomPt), P, P]);
        expect(pages).toHaveLength(1);
        expect(COLUMN_END - pages[0]!.tables[0]!.bottom).toBeCloseTo(roomPt, 6);
        return pages[0]!;
      };
      const tight = layOutWithRoom(17);
      expect(tight.strips.flat().map((strip) => strip.inlineStart)).toEqual([MARGIN.top, MARGIN.top]);
      expect(tight.strips[0]![0]!.block).toBeCloseTo(TABLE_W + RULE + 7.14, 6);
      const roomy = layOutWithRoom(19);
      expect(roomy.strips[0]![0]!.block).toBe(0);
      for (const strip of roomy.strips[0]!) expect(strip.inlineStart).toBeCloseTo(roomy.tables[0]!.bottom, 6);
      expect(roomy.strips[0]).toHaveLength(5);
    },
  );

  it.each(['tbRl', 'tbLrV'] as const)(
    'starts the next page with a table whose box would reach the footnote it refers to (%s)',
    (direction) => {
      // Fifteen strips and a 180 pt table fit the page's block extent. The
      // table's first cell refers to a four-paragraph footnote, whose band
      // at the block end of the page the table's box would cross: the table
      // adds no flow advance, so it is its box that has to end before the
      // band. The table and its note start the next page.
      const noted = table(2, 40).replace('</w:p>', `${FOOTNOTE_REFERENCE}</w:p>`);
      const stories = { footnote: paragraph('note').repeat(4) };
      const body = (tableXml: string) => [...Array.from({ length: 15 }, () => P), tableXml, P];
      const plain = layOutPages(direction, body(table(2, 40)), 1, 360, undefined, undefined, stories);
      expect(plain.map((page) => page.tables.length)).toEqual([1]);
      const pages = layOutPages(direction, body(noted), 1, 360, undefined, undefined, stories);
      expect(pages.map((page) => [page.strips.length, page.tables.length])).toEqual([[15, 0], [1, 1]]);
      expect(pages[1]!.tables[0]!.top).toBeCloseTo(MARGIN.top, 6);
    },
  );

  describe('footnotes beside upright tables', () => {
    // A four-paragraph footnote takes its band at the block end of the page.
    // In one column no table box shares block extent with it unless nothing
    // can move.
    const stories = { footnote: paragraph('note').repeat(4) };
    const noted = (xml: string) => xml.replace('</w:p>', `${FOOTNOTE_REFERENCE}</w:p>`);
    const strips = (count: number) => Array.from({ length: count }, () => P);
    /** A one-row table `widthPt` wide (two equal columns). */
    const wide = (widthPt: number, rowPt: number) => table(1, rowPt)
      .replace(/w:w="1440"/g, `w:w="${widthPt * 10}"`).replace(/w:w="2160"/g, `w:w="${widthPt * 10}"`);
    /** How far the page's tables reach along the block axis, from the first strip. */
    const tablesEnd = (direction: Direction, page: Page) => Math.max(...page.tables.map((placed) =>
      direction === 'tbRl' ? PAGE_W - MARGIN.right - placed.paintedLeft : placed.paintedRight - MARGIN.left));
    const layOutNoted = (direction: Direction, body: readonly string[], columns = 1, more: object = {}) =>
      layOutPages(direction, body, columns, 360, undefined, undefined, { ...stories, ...more });

    it.each(['tbRl', 'tbLrV'] as const)('starts the next page with a table moved past another into its own note (%s)', (direction) => {
      // The 300 pt-wide table meets the 520 pt-tall one, does not fit below
      // it and is placed past it along the block axis, where its box would
      // reach the band of the note it refers to. It starts the next page.
      const pages = layOutNoted(direction, [wide(90, 520), P, noted(wide(300, 300))]);
      expect(pages.map((page) => page.tables.length)).toEqual([1, 1]);
      expect(pages[0]!.noteStart).toBeNull();
      expect(tablesEnd(direction, pages[1]!)).toBeLessThanOrEqual(pages[1]!.noteStart! + 1e-6);
    });

    it.each(['tbRl', 'tbLrV'] as const)('moves a run of tables whose second table refers to a note (%s)', (direction) => {
      // Two directly following tables are one block. The note the second
      // refers to counts when the first is placed: both start the next page.
      const second = noted(table(2, 40)).replace(/w:sz="4"/g, 'w:sz="8"');
      const pages = layOutNoted(direction, [...strips(15), table(2, 40), second]);
      expect(pages.map((page) => [page.strips.length, page.tables.length])).toEqual([[15, 0], [0, 2]]);
      expect(tablesEnd(direction, pages[1]!)).toBeLessThanOrEqual(pages[1]!.noteStart! + 1e-6);
    });

    it.each(['tbRl', 'tbLrV'] as const)('moves a paragraph whose note would reach a table placed before it (%s)', (direction) => {
      // The table adds no flow advance, so the flow position after it does
      // not say how far the page is occupied: its box does. The paragraph
      // after it, whose note band would reach the box, starts the next page.
      const pages = layOutNoted(direction, [...strips(15), table(2, 40), noted(P)]);
      expect(pages.map((page) => [page.strips.length, page.tables.length])).toEqual([[15, 1], [1, 0]]);
      expect(pages[0]!.noteStart).toBeNull();
    });

    it.each(['tbRl', 'tbLrV'] as const)('moves a keepNext paragraph with a table whose own note it has no room for (%s)', (direction) => {
      const kept = P.replace('<w:pPr>', '<w:pPr><w:keepNext/>');
      const pages = layOutNoted(direction, [...strips(14), kept, noted(table(2, 40))]);
      expect(pages.map((page) => [page.strips.length, page.tables.length])).toEqual([[14, 0], [1, 1]]);
    });

    it.each(['tbRl', 'tbLrV'] as const)('continues the note of a table that is first on its page past the table (%s)', (direction) => {
      // A 420 pt-wide table first on its page cannot move. Its note takes
      // the room its box leaves and continues on the next page.
      const first = noted(table(2, 40).replace(/w:w="1440"/g, 'w:w="4200"').replace(/w:w="2160"/g, 'w:w="4200"'));
      const pages = layOutNoted(direction, [first, P], 1, { continueFootnotes: true });
      expect(pages).toHaveLength(2);
      expect(pages[0]!.tables).toHaveLength(1);
      expect(tablesEnd(direction, pages[0]!)).toBeLessThanOrEqual(pages[0]!.noteStart! + 1e-6);
      expect(pages[1]!.noteStart).not.toBeNull();
    });

    it.each(['tbRl', 'tbLrV'] as const)('keeps a page-start run whose second table refers to a continued note (%s)', (direction) => {
      // The 420 pt-wide table is first on its page and cannot move, nor can
      // the table placed with it. The note the second table refers to takes
      // the room the run's box leaves and continues on the next page.
      const first = table(2, 40).replace(/w:w="1440"/g, 'w:w="4200"').replace(/w:w="2160"/g, 'w:w="4200"');
      const second = noted(table(2, 41)).replace(/w:sz="4"/g, 'w:sz="8"');
      const pages = layOutNoted(direction, [first, second, P], 1, { continueFootnotes: true });
      expect(pages.map((page) => page.tables.length)).toEqual([2, 0]);
      expect(tablesEnd(direction, pages[0]!)).toBeLessThanOrEqual(pages[0]!.noteStart! + 1e-6);
      expect(pages[1]!.noteStart).not.toBeNull();
    });

    it.each(['tbRl', 'tbLrV'] as const)('starts document-end notes past a final table (%s)', (direction) => {
      // The table adds no flow advance: the endnotes after the final empty
      // paragraph start past its box, not at the flow position.
      const endnoted = table(2, 40).replace('</w:p>', `${ENDNOTE_REFERENCE}</w:p>`);
      const pages = layOutPages(direction, [P, endnoted, '<w:p/>'], 1, 360, undefined, undefined,
        { endnote: paragraph('note').repeat(2) });
      expect(pages).toHaveLength(1);
      expect(pages[0]!.noteStart).not.toBeNull();
      expect(tablesEnd(direction, pages[0]!)).toBeLessThanOrEqual(pages[0]!.noteStart! + 1e-6);
    });

    it.each(['tbRl', 'tbLrV'] as const)('leaves a table of two columns where the flow reaches it, also over its note (%s)', (direction) => {
      const pages = layOutNoted(direction, [...strips(15), noted(table(2, 40)), P], 2);
      expect(pages).toHaveLength(1);
      expect(pages[0]!.tables).toHaveLength(1);
      expect(tablesEnd(direction, pages[0]!)).toBeGreaterThan(pages[0]!.noteStart!);
    });
  });

  it.each(['tbRl', 'tbLrV'] as const)(
    'moves a keepNext paragraph with a table that its footnote leaves no room for (%s)',
    (direction) => {
      // Fourteen strips, a keepNext paragraph and a table fit the page. The
      // kept paragraph refers to a four-paragraph footnote, whose reserve
      // at the block end of the page leaves the 180 pt table no room there.
      // The look-ahead places the table in the block extent left after the
      // note, as the placement does, so the paragraph moves with it.
      const kept = P.replace('<w:pPr>', '<w:pPr><w:keepNext/>');
      const body = (reference: string) => [
        ...Array.from({ length: 14 }, () => P), kept.replace('</w:p>', `${reference}</w:p>`), T, P,
      ];
      const stories = { footnote: paragraph('note').repeat(4) };
      const without = layOutPages(direction, body(''), 1, 360, undefined, undefined, stories);
      expect(without.map((page) => [page.strips.length, page.tables.length])).toEqual([[16, 1]]);
      const pages = layOutPages(direction, body(FOOTNOTE_REFERENCE), 1, 360, undefined, undefined, stories);
      expect(pages.map((page) => [page.strips.length, page.tables.length])).toEqual([[14, 0], [2, 1]]);
    },
  );

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

    it('measures every table a new plan admits on the page that places it', () => {
      // Eight pages of text, then seven tables with nothing between them.
      // The first (700 pt) fills page 9. On page 10 the roman PAGE fields of
      // two of the others are one character shorter than on page 9 ("X" for
      // "IX"), which makes room for the table of arabic PAGE fields behind
      // them, and that one is a character longer ("10" for "9"): it has to
      // be measured on page 10 before it is placed, or it runs over the
      // table after it.
      const pageBreak = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
      const pages = layOutPages('tbRl', [
        ...Array.from({ length: 8 }, () => [P, pageBreak]).flat(),
        table(1, 700), table(1, 200), table(1, 100),
        romanPageTable(10, 1), romanPageTable(15, 1), romanPageTable(10, 1, 'Arabic'), table(1, 30),
      ]);
      expect(pages.length).toBeGreaterThanOrEqual(10);
      expect(pages.slice(8).flatMap((page) => page.tables).length).toBeGreaterThanOrEqual(7);
      expect(overlapping(pages)).toEqual([]);
      for (const page of pages) {
        page.tables.forEach((placed, at) => {
          expect(placed.bottom).toBeLessThanOrEqual(COLUMN_END + 1e-6);
          if (at > 0) expect(placed.top).toBeCloseTo(page.tables[at - 1]!.bottom, 6);
        });
      }
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

    /** Block position of the far edge of a table whose box starts at `block`. */
    const blockStartOf = (direction: Direction, placed: Page['tables'][number]) => direction === 'tbRl'
      ? PAGE_W - MARGIN.right - placed.paintedRight : placed.paintedLeft - MARGIN.left;

    it.each(['tbRl', 'tbLrV'] as const)(
      'moves to the next page a second %s table whose rise would move the text of column 1 off the page',
      (direction) => {
        // Twelve paragraphs, a column break, an 800 pt table X, two
        // paragraphs and an 800 pt table Y. X is raised over column 1,
        // whose strips resume past it; Y, raised there too, would leave
        // column 1 no room for them. Y gives way: it starts page 2, and
        // page 1 keeps X and its text.
        const pages = layOutPages(direction, [
          ...strips(12), columnBreak, table(1, 800), P, P, table(1, 800), P,
        ], 2);
        expect(pages).toHaveLength(2);
        const [first, second] = pages as [Page, Page];
        expect(first.tables).toHaveLength(1);
        expect(first.tables[0]!.left).toBeCloseTo(tableLeft(direction, PITCH), 6);
        first.strips.slice(0, 12).forEach(([line], at) => {
          expect(line!.block).toBeCloseTo(past(PITCH) + at * PITCH, 6);
          expect(line!.inlineStart).toBe(MARGIN.top);
        });
        // The two paragraphs of column 2, past X after the empty strip.
        const columnTwo = first.strips.slice(-2).map(([line]) => line!);
        expect(columnTwo.map((line) => line.block)).toEqual([past(PITCH) + PITCH, past(PITCH) + 2 * PITCH]
          .map((block) => expect.closeTo(block, 6)));
        for (const line of columnTwo) expect(line.inlineStart).toBeGreaterThan((MARGIN.top + COLUMN_END) / 2);
        expect(second.tables).toHaveLength(1);
        expect(second.tables[0]!.left).toBeCloseTo(tableLeft(direction, 0), 6);
        expect(second.strips.map(([line]) => line!.block)).toEqual([expect.closeTo(past(0), 6)]);
      },
    );

    it.each([['tbRl', 800], ['tbLrV', 800], ['tbRl', 500]] as const)(
      'lays the page out without a %s table whose rise moves the text before it off the page (second table %i pt)',
      (direction, secondPt) => {
        // Thirty paragraphs (25 fill column 1), an 800 pt table X, three
        // paragraphs and a table Y. Laid out without X the flow reaches
        // it five strips into column 2; raised from there it would move the
        // text of column 1 past it, off the page. Page 1 holds the thirty
        // paragraphs without X, which starts page 2, where the paragraphs
        // after it resume past it and Y follows the third of them.
        const pages = layOutPages(direction, [
          ...strips(30), table(1, 800), P, P, P, table(1, secondPt), P,
        ], 2);
        expect(pages).toHaveLength(2);
        const [first, second] = pages as [Page, Page];
        expect(first.tables).toHaveLength(0);
        expect(first.strips).toHaveLength(30);
        for (const [line] of first.strips) {
          expect([MARGIN.top, first.strips[25]![0]!.inlineStart]).toContain(line!.inlineStart);
        }
        const [x, y] = second.tables as [Page['tables'][number], Page['tables'][number]];
        expect(x.left).toBeCloseTo(tableLeft(direction, 0), 6);
        expect(blockStartOf(direction, y)).toBeCloseTo(past(0) + 3 * PITCH, 6);
        // An 800 pt table is raised to end at the bottom page edge; a 500 pt
        // one stays at the column top.
        if (secondPt === 800) expect(y.bottom).toBeCloseTo(16838 / 20, 6);
        else expect(y.top).toBeCloseTo(MARGIN.top, 6);
        // The third paragraph ends where Y starts and resumes past it.
        expect(second.strips.map(([line]) => line!.block)).toEqual([
          past(0), past(0) + PITCH, past(past(0) + 3 * PITCH), past(past(0) + 3 * PITCH) + PITCH,
        ].map((block) => expect.closeTo(block, 6)));
      },
    );

    it('keeps a table raised clear of the text of column 1 and moves a later one too wide for the page', () => {
      // The same with X of a 500 pt row, raised over the room of
      // column 1 below its short strips, and Y of an 800 pt row: X stays five
      // strips into column 2 and Y, wider than the block extent left after
      // the paragraphs past X, starts page 2.
      const pages = layOutPages('tbRl', [...strips(30), table(1, 500), P, P, P, table(1, 800), P], 2);
      expect(pages).toHaveLength(2);
      const [first, second] = pages as [Page, Page];
      expect(first.tables).toHaveLength(1);
      expect(blockStartOf('tbRl', first.tables[0]!)).toBeCloseTo(5 * PITCH, 6);
      for (const [line] of first.strips.slice(0, 25)) expect(line!.inlineStart).toBe(MARGIN.top);
      expect(first.strips.slice(29).map(([line]) => line!.block)).toEqual([0, 1, 2, 3]
        .map((at) => expect.closeTo(past(5 * PITCH) + at * PITCH, 6)));
      expect(second.tables).toHaveLength(1);
      expect(second.tables[0]!.left).toBeCloseTo(tableLeft('tbRl', 0), 6);
    });

    it('pins a table raised over column 1 and a table with a drawing over it on one page', () => {
      // Twelve paragraphs, a column break, a 2-row table X at the
      // top of column 2, a paragraph anchoring a rectangle over X, two
      // paragraphs and an 800 pt table Y. The rectangle lies below X. Y,
      // which meets X, is placed 14.28 pt past it and raised; the strip of
      // column 1 that ends inside Y resumes past it, the others keep the
      // column top.
      const over = P.replace('</w:p>', `${anchoredRectangle(400, 430, 80, 40)}</w:p>`);
      const pages = layOutPages('tbRl', [...strips(12), columnBreak, T, over, P, P, table(1, 800), P], 2);
      expect(pages).toHaveLength(1);
      const [page] = pages;
      const [x, y] = page!.tables as [Page['tables'][number], Page['tables'][number]];
      expect(blockStartOf('tbRl', x)).toBeCloseTo(PITCH, 6);
      expect(page!.drawings[0]!.top).toBeCloseTo(x.bottom, 6);
      expect(blockStartOf('tbRl', y)).toBeCloseTo(PITCH + TABLE_W + RULE + 14.28, 6);
      expect(y.bottom).toBeCloseTo(16838 / 20, 6);
      const columnOne = page!.strips.slice(0, 12).map(([line]) => line!);
      columnOne.slice(0, 11).forEach((line, at) => {
        expect(line.block).toBeCloseTo(at * PITCH, 6);
        expect(line.inlineStart).toBe(MARGIN.top);
      });
      expect(columnOne[11]!.block).toBeCloseTo(past(PITCH + TABLE_W + RULE + 14.28), 6);
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

    it('flows the first column around a raised table whose own strip moved to the next column', () => {
      // Three columns, 40 strips and a full-length strip before the table:
      // laid out without it they fill column 1 (25) and reach 16 strips into
      // column 2, where the table is raised over the page. Beside it the
      // full-length strip has no room in column 2 and moves to column 3,
      // which leaves the table placed ahead of the flow in another column.
      // It still lies over column 1, whose text has to flow around it too:
      // every column then holds the strips before the table's block range.
      const full = paragraph('長'.repeat(20));
      const pages = layOutPages('tbLrV', [...strips(40), full, table(1, 800), P], 3);
      expect(pages).toHaveLength(1);
      const [page] = pages;
      const [raised] = page!.tables;
      expect(raised!.bottom).toBeCloseTo(16838 / 20, 6);
      const blockStart = raised!.left - MARGIN.left - RULE / 2;
      expect(blockStart).toBeCloseTo(16 * PITCH, 6);
      const lines = page!.strips.flat();
      expect(lines).toHaveLength(42);
      // The table spans the page: no strip lies in its exclusion, and none
      // ends where it starts.
      for (const line of lines) expect(line.block + PITCH, `strip at block ${line.block}`).toBeLessThan(blockStart + 1e-6);
      expect(lines.filter((line) => line.inlineStart === MARGIN.top)).toHaveLength(15);
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

    it('keeps the shared left edge of a run on the page its rows continue on', () => {
      // A 90 pt wide table of 45 rows of 18 pt, taller than the band, directly
      // before a 180 pt wide table: on page 2 the wide table follows the
      // narrow one's last rows at the same left edge.
      const pages = layOutPages('tbRl', [narrow(45).replaceAll('<w:tr>', '<w:tr><w:trPr><w:trHeight w:val="360" w:hRule="atLeast"/></w:trPr>'), T]);
      expect(pages).toHaveLength(2);
      const [rest, wide] = pages[1]!.tables as [Page['tables'][number], Page['tables'][number]];
      expect(wide.top).toBeCloseTo(rest.bottom, 6);
      expect(rest.paintedLeft).toBeCloseTo(wide.paintedLeft, 6);
      expect(pages[0]!.tables[0]!.paintedLeft).toBeCloseTo(wide.paintedLeft, 6);
      expect(wide.left).toBeCloseTo(tableLeft('tbRl', 0), 6);
    });

    it('takes a run again on a page where its tables measure differently', () => {
      // One 12 pt wide column of cells holding a PAGE field in roman
      // numerals: "I" is one line, "II" two, so every row is twice as tall
      // on page 2. Three 30-row tables fill page 1 and continue on page 2,
      // where the 15-row table after them no longer fits whole below the
      // continued rows, although it did by the heights taken on page 1.
      const field = '<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>'
        + '<w:fldSimple w:instr=" PAGE \\* Roman "><w:r><w:rPr><w:sz w:val="21"/></w:rPr><w:t>I</w:t></w:r></w:fldSimple></w:p>';
      const numbered = (rows: number) => '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/></w:tblPr>'
        + '<w:tblGrid><w:gridCol w:w="240"/></w:tblGrid>'
        + Array.from({ length: rows }, () => '<w:tr><w:tc><w:tcPr><w:tcW w:w="240" w:type="dxa"/>'
          + '<w:tcMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tcMar></w:tcPr>'
          + `${field}</w:tc></w:tr>`).join('') + '</w:tbl>';
      const pages = layOutPages('tbRl', [numbered(30), numbered(30), numbered(30), numbered(15)]);
      expect(pages.length).toBeGreaterThanOrEqual(3);
      const second = pages[1]!.tables;
      // The rows continued on page 2 are twice as tall as on page 1.
      const firstPageRow = (pages[0]!.tables[0]!.bottom - pages[0]!.tables[0]!.top) / 30;
      expect((second[0]!.bottom - second[0]!.top) / firstPageRow).toBeGreaterThan(40);
      for (const page of pages) {
        page.tables.forEach((placed, at) => {
          expect(placed.bottom).toBeLessThanOrEqual(COLUMN_END + 1e-6);
          if (at > 0) expect(placed.top).toBeCloseTo(page.tables[at - 1]!.bottom, 6);
        });
      }
    });

    describe('that the parser groups into one logical table', () => {
      // A default table style gives directly following tables one style
      // identity, so they reach layout as one logical table (§17.4.37).
      const styles = '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/></w:style>'
        + '<w:style w:type="table" w:styleId="Other"><w:name w:val="Other"/></w:style>';
      const layOutStyled = (direction: Direction, body: readonly string[]) =>
        layOutPages(direction, body, 1, 360, undefined, styles);

      it.each(['tbRl', 'tbLrV'] as const)('places a %s logical table upright, without a flow advance', (direction) => {
        const single = layOutStyled(direction, [P, T, P]);
        const pages = layOutStyled(direction, [P, T, T, P, P]);
        expect(pages).toHaveLength(1);
        // One table of four rows.
        const [placed] = pages[0]!.tables;
        expect(pages[0]!.tables).toHaveLength(1);
        const oneTable = single[0]!.tables[0]!;
        expect(placed!.top).toBeCloseTo(MARGIN.top, 6);
        expect(placed!.left).toBeCloseTo(tableLeft(direction, PITCH), 6);
        expect(placed!.bottom - placed!.top).toBeGreaterThan(1.9 * (oneTable.bottom - oneTable.top));
        expect(pages[0]!.strips.map((lines) => lines.map((line) => line.block)))
          .toEqual([[0], [PITCH], [2 * PITCH]]);
        for (const [line] of pages[0]!.strips) expect(line!.inlineStart).toBeCloseTo(placed!.bottom, 6);
      });

      it('places the segments of a logical table with a cell-owned frame row upright', () => {
        // The second table's first cell starts with a framed paragraph, so
        // the logical table is projected into owner segments (a host for
        // that row, ordinary segments around it). They are upright blocks
        // that add no flow advance, like those of a single such table.
        const framed = T.replace('<w:p><w:pPr>', '<w:p><w:pPr><w:framePr w:w="2000" w:hSpace="141"'
          + ' w:wrap="around" w:vAnchor="page" w:hAnchor="page" w:x="4000" w:y="6000"/>');
        const [page] = layOutStyled('tbRl', [P, T, framed, P, P]);
        const [before, host, after] = page!.tables as [
          Page['tables'][number], Page['tables'][number], Page['tables'][number],
        ];
        expect(page!.tables).toHaveLength(3);
        expect(before.top).toBeCloseTo(MARGIN.top, 6);
        expect(before.left).toBeCloseTo(tableLeft('tbRl', PITCH), 6);
        expect(host.top).toBeCloseTo(300, 6);
        expect(after.top).toBeCloseTo(before.bottom, 6);
        expect(after.left).toBeCloseTo(before.left, 6);
        expect(page!.strips.map(([line]) => line!.block)).toEqual([0, PITCH, 2 * PITCH]);
        for (const [line] of page!.strips.slice(1)) expect(line!.inlineStart).toBeCloseTo(after.bottom, 6);
      });

      it('moves an ordinary segment whose border rules pass the block extent left to the next page', () => {
        // Fifteen strips leave 195.3 pt. The logical table's 195 pt grid
        // fits; its first segment's box with two 3 pt side rules is 198 pt
        // wide and does not, like that of a table without owner segments.
        const wide = T.replaceAll('w:sz="4"', 'w:sz="24"').replaceAll('"2160"', '"2460"');
        const framed = wide.replace('<w:p><w:pPr>', '<w:p><w:pPr><w:framePr w:w="2000" w:hSpace="141"'
          + ' w:wrap="around" w:vAnchor="page" w:hAnchor="page" w:x="4000" w:y="6000"/>');
        const pages = layOutStyled('tbRl', [...Array.from({ length: 15 }, () => P), wide, framed, P]);
        expect(pages).toHaveLength(2);
        expect(pages[0]!.tables).toHaveLength(0);
        expect(pages[0]!.strips).toHaveLength(15);
        expect(pages[1]!.tables.length).toBeGreaterThan(0);
        expect(pages[1]!.tables[0]!.paintedRight).toBeCloseTo(PAGE_W - MARGIN.right, 6);
      });

      it('runs a logical table on from a table of another style', () => {
        const other = T.replace('<w:tblPr>', '<w:tblPr><w:tblStyle w:val="Other"/>');
        const [page] = layOutStyled('tbRl', [P, other, T, T, P]);
        const [first, grouped] = page!.tables as [Page['tables'][number], Page['tables'][number]];
        expect(page!.tables).toHaveLength(2);
        expect(first.top).toBeCloseTo(MARGIN.top, 6);
        expect(grouped.top).toBeCloseTo(first.bottom, 6);
        expect(grouped.left).toBeCloseTo(first.left, 6);
        // Text flows around the box the tables are painted in.
        for (const [line] of page!.strips) expect(line!.inlineStart).toBeCloseTo(grouped.bottom, 6);
        expect(page!.strips.map(([line]) => line!.block)).toEqual([0, PITCH]);
      });
    });

    it('fills the band exactly with a run whose total rounds past it', () => {
      // Three borderless one-row tables of exactly 4001, 4002 and 6035
      // twips: 701.9 pt, the band. Summed in body order they fit it; the
      // run's recorded total, summed the other way, rounds just past it.
      const exact = (twips: number) => '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>'
        + '<w:tblGrid><w:gridCol w:w="3600"/></w:tblGrid>'
        + `<w:tr><w:trPr><w:trHeight w:val="${twips}" w:hRule="exact"/></w:trPr>`
        + `<w:tc><w:tcPr><w:tcW w:w="3600" w:type="dxa"/></w:tcPr>${paragraph('A')}</w:tc></w:tr></w:tbl>`;
      const pages = layOutPages('tbRl', [exact(4001), exact(4002), exact(6035)]);
      expect(pages).toHaveLength(1);
      const [first, second, third] = pages[0]!.tables as [
        Page['tables'][number], Page['tables'][number], Page['tables'][number],
      ];
      expect(pages[0]!.tables).toHaveLength(3);
      expect(first.top).toBeCloseTo(MARGIN.top, 6);
      expect(second.top).toBeCloseTo(first.bottom, 6);
      expect(third.top).toBeCloseTo(second.bottom, 6);
      expect(third.bottom).toBeCloseTo(COLUMN_END, 6);
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

  it('moves a wrapped drawing anchored after a table below its ordinary row beside a cell-owned frame row', () => {
    // The table's first cell opens with a framed paragraph: that row is a
    // host at its page position and the other row an upright block at the
    // column top, which the rectangle anchored after the table would cover.
    const framed = T.replace('<w:p><w:pPr>', '<w:p><w:pPr><w:framePr w:w="2000" w:hSpace="141"'
      + ' w:wrap="around" w:vAnchor="page" w:hAnchor="page" w:x="4000" w:y="6000"/>');
    const [page] = layOutPages('tbRl', [P, framed, P.replace('</w:p>', `${rectangleAt(400)}</w:p>`)]);
    const ordinary = page!.tables.find((placed) => Math.abs(placed.top - MARGIN.top) < 1e-6)!;
    expect(ordinary).toBeDefined();
    expect(page!.drawings).toHaveLength(1);
    expect(page!.drawings[0]!.top).toBeCloseTo(ordinary.bottom, 6);
    expect(page!.drawings[0]!.left).toBeCloseTo(400, 6);
  });

  describe('drawings anchored after a table split around a cell-owned frame row', () => {
    const framePr = '<w:framePr w:w="2000" w:hSpace="141" w:wrap="around" w:vAnchor="page" w:hAnchor="page" w:x="4000" w:y="6000"/>';
    /** Three rows: an ordinary one, a host at page (200, 300), an ordinary one. */
    const split = (() => {
      const rows = table(3).split('<w:tr>');
      rows[2] = rows[2]!.replace('<w:p><w:pPr>', `<w:p><w:pPr>${framePr}`);
      return rows.join('<w:tr>');
    })();
    /** Block range of a drawing `width` wide whose physical left is `left`, in tbRl. */
    const blockRange = (left: number, width: number) =>
      [PAGE_W - MARGIN.right - left - width, PAGE_W - MARGIN.right - left] as const;

    it('moves the drawing over each ordinary row, and its text exclusion, below the rows', () => {
      // Both ordinary rows stack at the column top, each covered by one of
      // two rectangles the paragraph after the table anchors. Both lie below
      // the rows, and the strips in each one's block range start below it:
      // their exclusions moved with them.
      const upper = anchoredRectangle(440, 62, 30, 6);
      const lower = anchoredRectangle(360, 74, 30, 6);
      const [page] = layOutPages('tbRl', [
        P, split, P.replace('</w:p>', `${upper}${lower}</w:p>`), ...Array.from({ length: 8 }, () => P),
      ]);
      const ordinary = page!.tables.filter((placed) => placed.top < 200);
      expect(ordinary).toHaveLength(2);
      const rowsEnd = Math.max(...ordinary.map((placed) => placed.bottom));
      expect(page!.drawings).toHaveLength(2);
      for (const drawing of page!.drawings) expect(drawing.top).toBeCloseTo(rowsEnd, 6);
      const lines = page!.strips.slice(1).flat();
      for (const left of [440, 360]) {
        const [start, end] = blockRange(left, 30);
        const beside = lines.filter((line) => line.block < end && line.block + PITCH > start);
        expect(beside.length, `strips beside the rectangle at ${left}`).toBeGreaterThan(0);
        for (const line of beside) expect(line.inlineStart).toBeCloseTo(rowsEnd + 6, 6);
      }
    });

    it('keeps a pinned row where the drawing was placed below it when the text before it changes', () => {
      // The rectangle is authored over the first strip as well as the row:
      // at that place it shortens the full-length strip before the table,
      // which then takes a second strip, so the row starts two strips in.
      // Once the rectangle lies below the row the strip fits again, but the
      // row is pinned: it stays where the rectangle was placed below it.
      const full = paragraph('長'.repeat(66));
      const wide = anchoredRectangle(440, 62, 60, 6);
      const [page] = layOutPages('tbRl', [full, split, P.replace('</w:p>', `${wide}</w:p>`), P]);
      const ordinary = page!.tables.filter((placed) => placed.top < 200);
      const rowsEnd = Math.max(...ordinary.map((placed) => placed.bottom));
      expect(page!.drawings).toHaveLength(1);
      expect(page!.drawings[0]!.top).toBeCloseTo(rowsEnd, 6);
      // The drawing covers no row as painted.
      for (const placed of page!.tables) {
        const apart = page!.drawings[0]!.top >= placed.bottom - 1e-6
          || page!.drawings[0]!.top + 6 <= placed.top + 1e-6
          || 440 >= placed.paintedRight - 1e-6 || 500 <= placed.paintedLeft + 1e-6;
        expect(apart).toBe(true);
      }
    });
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
  it.each(['tbRl', 'tbLrV'] as const)('moves two drawings that one paragraph after two %s tables anchors below them', (direction) => {
    // A paragraph, table X, eleven paragraphs, table Y, then one
    // paragraph anchoring a rectangle over Y and one over X. Each rectangle
    // lies directly below the table it covers, and the strips beside it
    // start below it.
    const [overX, overY] = direction === 'tbRl' ? [400, 150] : [150, 350];
    const anchoring = P.replace('</w:p>', `${anchoredRectangle(overY, 66, 80, 40)}${anchoredRectangle(overX, 66, 80, 40)}</w:p>`);
    const [page] = layOutPages(direction, [P, T, ...Array.from({ length: 11 }, () => P), T, anchoring, P]);
    const [x, y] = page!.tables as [Page['tables'][number], Page['tables'][number]];
    expect(x.top).toBeCloseTo(MARGIN.top, 6);
    expect(y.top).toBeCloseTo(MARGIN.top, 6);
    const lying = [...page!.drawings].sort((left, right) => left.left - right.left);
    for (const drawing of lying) expect(drawing.top).toBeCloseTo(x.bottom, 6);
    expect(lying.map((drawing) => drawing.left)).toEqual([Math.min(overX, overY), Math.max(overX, overY)]
      .map((left) => expect.closeTo(left, 6)));
    // The strips across the page from the rectangle over X start below it.
    const physicalLeft = (block: number) => direction === 'tbRl' ? PAGE_W - MARGIN.right - block - PITCH : MARGIN.left + block;
    for (const [line] of page!.strips) {
      const left = physicalLeft(line!.block);
      const besideX = left < overX + 80 && left + PITCH > overX;
      expect(line!.inlineStart).toBeCloseTo(besideX ? x.bottom + 40 : x.bottom, 6);
    }
  });

  describe('a table after a drawing moved below an earlier table', () => {
    /** A paragraph anchoring a 60 x 20 pt square-wrapped rectangle at page
     * (400, 64), over the place of the table before it. */
    const shifted = P.replace('</w:p>', `${anchoredRectangle(400, 64, 60, 20)}</w:p>`);
    /** Whether the rectangle drawn at `drawing` and the table box are apart. */
    const apart = (drawing: Page['drawings'][number], placed: Page['tables'][number]) =>
      drawing.top >= placed.bottom - 1e-6 || drawing.top + 20 <= placed.top + 1e-6
      || drawing.left >= placed.paintedRight - 1e-6 || drawing.left + 60 <= placed.paintedLeft + 1e-6;

    it('stacks below the drawing where it lies, not where it was anchored', () => {
      const [page] = layOutPages('tbRl', [P, T, shifted, T, P]);
      const [first, second] = page!.tables as [Page['tables'][number], Page['tables'][number]];
      const drawing = page!.drawings[0]!;
      expect(drawing.top).toBeCloseTo(first.bottom, 6);
      expect(second.top).toBeCloseTo(drawing.top + 20, 6);
      for (const placed of page!.tables) expect(apart(drawing, placed)).toBe(true);
      // The strip after the second table meets both tables and the drawing
      // between them: it starts below all three.
      expect(page!.strips[2]![0]!.inlineStart).toBeCloseTo(second.bottom, 6);
    });

    it('stacks below the drawing moved below the rows a table continues with', () => {
      // The rows continued on page 2 cover the rectangle, which moves below
      // them; so does the text beside it.
      const pages = layOutPages('tbRl', [table(80), P, shifted, T, P, P]);
      expect(pages).toHaveLength(2);
      const [rest, later] = pages[1]!.tables as [Page['tables'][number], Page['tables'][number]];
      const drawing = pages[1]!.drawings[0]!;
      expect(drawing.top).toBeCloseTo(rest.bottom, 6);
      expect(later.top).toBeCloseTo(drawing.top + 20, 6);
      for (const placed of pages[1]!.tables) expect(apart(drawing, placed)).toBe(true);
      // The strips beside the rectangle start below the later table, which
      // lies past it.
      for (const lines of pages[1]!.strips.slice(2)) expect(lines[0]!.inlineStart).toBeCloseTo(later.bottom, 6);
    });

    it('stacks below the drawing moved below a run that continues a table', () => {
      const pages = layOutPages('tbRl', [table(80), T, shifted, T, P]);
      const tables = pages[1]!.tables;
      const drawing = pages[1]!.drawings[0]!;
      expect(drawing.top).toBeCloseTo(tables[1]!.bottom, 6);
      expect(tables[2]!.top).toBeCloseTo(drawing.top + 20, 6);
      for (const placed of tables) expect(apart(drawing, placed)).toBe(true);
    });
  });
  describe('tables and square-wrapped drawings', () => {
    it.each([
      ['tbLrV', anchoring(35, anchoredRectangle(45, 0, 38, 50, 'paragraph', 'character'))],
      ['tbRl', anchoring(5, anchoredRectangle(440, 60, 38, 50))],
    ] as const)('places a table of two columns below a drawing only where it ends within its column (%s)', (direction, before) => {
      // Two columns, each 340 pt along the column. The rectangle of the
      // paragraph before the table lies at the column top over the table's
      // place. A 250 pt table ends within the column below the rectangle
      // and stacks there; a 320 pt table would pass the column's end there,
      // so it stays at the column top, 7.14 pt past the rectangle along the
      // block axis.
      const below = layOutPages(direction, [before, table(1, 250), P], 2)[0]!;
      expect(below.tables[0]!.top).toBeCloseTo(below.drawings[0]!.top + below.drawings[0]!.height, 6);
      const past = layOutPages(direction, [before, table(1, 320), P], 2)[0]!;
      const [placed] = past.tables;
      const [drawing] = past.drawings;
      expect(placed!.top).toBeCloseTo(MARGIN.top, 6);
      expect(drawing!.top).toBeCloseTo(MARGIN.top, 6);
      if (direction === 'tbRl') expect(placed!.paintedRight).toBeCloseTo(drawing!.left - 7.14, 6);
      else expect(placed!.paintedLeft).toBeCloseTo(drawing!.left + drawing!.width + 7.14, 6);
      expect(collisions([below, past])).toEqual([]);
    });

    it('lays out a table before a long one whose rows a later drawing lies below', () => {
      // The second table's first fragment covers the rectangle anchored
      // after it and is pinned; the first table, placed before it, is not
      // obstructed by that pin when the page is laid out again.
      const pages = layOutPages('tbRl', [
        table(1), table(70), P.replace('</w:p>', `${anchoredRectangle(440, 480, 80, 10)}</w:p>`), P,
      ]);
      expect(pages[0]!.tables[0]!.top).toBeCloseTo(MARGIN.top, 6);
      expect(pages[0]!.tables[1]!.top).toBeCloseTo(pages[0]!.tables[0]!.bottom, 6);
      expect(collisions(pages)).toEqual([]);
    });

    it.each(['page', 'paragraph'] as const)(
      'stacks a table below a drawing positioned relative to the %s of the paragraph before it',
      (verticalFrom) => {
        // The drawing of the strip the table is planned with meets the plan
        // before the strip is committed.
        const y = verticalFrom === 'page' ? 64 : 4;
        const [page] = layOutPages('tbRl', [
          P.replace('</w:p>', `${anchoredRectangle(400, y, 80, 40, verticalFrom)}</w:p>`), T, P,
        ]);
        expect(page!.drawings[0]!.top).toBeCloseTo(64, 6);
        expect(page!.tables[0]!.top).toBeCloseTo(64 + 40, 6);
      },
    );

    it.each([0, 1])('moves a keepNext set with a table that a drawing of the set sends to the next page (%i before)', (before) => {
      // The last keepNext paragraph anchors a 90 x 400 pt rectangle relative
      // to itself; the 350 pt table after it stacks below the rectangle,
      // which leaves it no room on page 1. The look-ahead meets the
      // rectangle as the placement does, whether it measures the paragraph
      // that anchors it or a keepNext paragraph before that one, so the
      // keep set starts page 2 with the table.
      const keepNext = P.replace('<w:pPr>', '<w:pPr><w:keepNext/>');
      const anchoring = keepNext.replace('</w:p>', `${anchoredRectangle(160, 4, 90, 400, 'paragraph')}</w:p>`);
      const pages = layOutPages('tbRl', [
        ...Array.from({ length: 14 - before }, () => P), ...Array.from({ length: before }, () => keepNext),
        anchoring, table(1, 350), P,
      ]);
      expect(pages).toHaveLength(2);
      expect(pages[0]!.strips).toHaveLength(14 - before);
      expect(pages[0]!.tables).toHaveLength(0);
      expect(pages[1]!.drawings).toHaveLength(1);
      expect(pages[1]!.tables).toHaveLength(1);
      expect(pages[1]!.strips[0]![0]!.block).toBe(0);
      expect(collisions(pages)).toEqual([]);
    });

    it('moves a keepNext set with a table that a character-positioned drawing of the set sends to the next page', () => {
      // The second keepNext paragraph anchors a 38 x 400 pt rectangle 16 pt
      // from its character along the block axis and 4 pt from itself along
      // the inline axis. The look-ahead measures that paragraph where the
      // keep set places it, one strip after the first, so the rectangle and
      // the table below it are where the placement puts them: neither fits
      // page 1 and the whole keep set starts page 2.
      const keepNext = P.replace('<w:pPr>', '<w:pPr><w:keepNext/>');
      const anchoring = keepNext.replace('</w:p>', `${anchoredRectangle(16, 4, 38, 400, 'paragraph', 'character')}</w:p>`);
      const pages = layOutPages('tbRl', [...Array.from({ length: 13 }, () => P), keepNext, anchoring, table(1, 350), P]);
      expect(pages).toHaveLength(2);
      expect(pages[0]!.strips).toHaveLength(13);
      expect(pages[1]!.strips[0]![0]!.block).toBe(0);
      expect(pages[1]!.tables).toHaveLength(1);
      expect(collisions(pages)).toEqual([]);
    });

    it('lays out a drawing that follows its character where a table has moved it into its own text', () => {
      // Both rectangles lie over the 45-row table before their paragraphs
      // and move below it, into the text of the second paragraph. The second
      // is positioned along the block axis from its character, which its
      // own exclusion moves from strip to strip: it keeps the farthest
      // position it has reached and the paragraph lays out.
      let pages: Page[] = [];
      expect(() => {
        pages = layOutPages('tbLrV', [
          table(45),
          anchoring(35, anchoredRectangle(70, 170, 200, 50, 'paragraph')),
          anchoring(90, anchoredRectangle(-20, 170, 200, 150, 'paragraph', 'character')),
        ]);
      }).not.toThrow();
      expect(pages).toHaveLength(1);
      const [placed] = pages[0]!.tables;
      expect(pages[0]!.drawings).toHaveLength(2);
      for (const drawing of pages[0]!.drawings) expect(drawing.top).toBeCloseTo(placed!.bottom, 6);
      expect(collisions(pages)).toEqual([]);
    });

    it('keeps a drawing of the strip before a table clear of the table when wrapping the strip moves it', () => {
      // The rectangle follows its character, the last of the strip before
      // the table. Laid out without the table it is clear of the table's
      // place; wrapped around the table the strip's last characters, and
      // the rectangle with them, move to the strip past the table, where
      // it would lie over the table. The table keeps its place and the
      // rectangle lies directly below it.
      const pages = layOutPages('tbRl', [
        anchoring(35, anchoredRectangle(45, 0, 38, 50, 'paragraph', 'character')), table(1, 400),
      ], 2);
      expect(pages).toHaveLength(1);
      const [placed] = pages[0]!.tables;
      expect(placed).toEqual(layOutPages('tbRl', [anchoring(35), table(1, 400)], 2)[0]!.tables[0]);
      const [drawing] = pages[0]!.drawings;
      expect(shared(drawing!.left, drawing!.left + drawing!.width, placed!.paintedLeft, placed!.paintedRight)).toBe(true);
      expect(drawing!.top).toBeCloseTo(placed!.bottom, 6);
      expect(collisions(pages)).toEqual([]);
    });

    it.each([
      ['one drawing positioned from its line and character', 'tbRl', 1,
        [anchoring(140, anchoredRectangle(40, 8, 340, 1, 'line', 'character'))]],
      ['two drawings positioned from their character', 'tbLrV', 2,
        [anchoring(66, anchoredRectangle(19, 119, 340, 150, 'paragraph', 'character'),
          anchoredRectangle(-50, 207, 200, 50, 'line', 'character'))]],
    ] as const)('lays out a paragraph whose own drawings keep moving its text: %s', (_name, direction, columns, body) => {
      // The drawings follow characters that their own exclusions move, and
      // no position of them repeats from one reflow pass to the next. The
      // paragraph is laid out against the drawings where the last pass left
      // them, the same way every time.
      let pages: Page[] = [];
      expect(() => { pages = layOutPages(direction, body, columns); }).not.toThrow();
      expect(pages).toHaveLength(1);
      expect(pages[0]!.drawings).toHaveLength(body[0]!.split('<w:drawing>').length - 1);
      expect(layOutPages(direction, body, columns)).toEqual(pages);
    });

    it('moves a drawing carried from an earlier layout of its page below the table it would cover', () => {
      // The last-but-one paragraph's rectangle lies over strips of earlier
      // paragraphs, so the page is laid out again with the rectangle
      // registered from its start at the frame it had; the tables are
      // placed differently in that layout, and the rectangle lies below
      // the table it would then cover.
      const pages = layOutPages('tbLrV', [
        anchoring(5), anchoring(35, anchoredRectangle(-55, 325, 20, 10, 'page', 'character')),
        table(1, 800), table(6),
        anchoring(90, anchoredRectangle(21, 80, 60, 1, 'paragraph', 'character')),
        anchoring(20, anchoredRectangle(-36, 73, 60, 150, 'paragraph', 'character')),
        anchoring(90, anchoredRectangle(105, 142, 200, 1, 'line', 'page')), table(6),
      ], 2);
      expect(pages[0]!.tables).toHaveLength(3);
      expect(pages[0]!.drawings).toHaveLength(4);
      expect(collisions(pages)).toEqual([]);
    });

    it.each([
      ['tbRl', 2, 360], ['tbLrV', 2, 100], ['tbRl', 1, 360], ['tbLrV', 1, 100],
    ] as const)('moves a drawing with no room below its table 7.14 pt past the table\'s far edge (%s, %i column(s))', (direction, columns, x) => {
      // The table leaves the rectangle anchored after it no place below it
      // on the page: raised to the page's bottom edge in two columns (an
      // 80 x 60 pt rectangle), and ending 91 pt before it in one (80 x 100
      // pt). The rectangle keeps its position down the page and lies
      // 7.14 pt past the table's far edge.
      const pages = layOutPages(direction, [
        table(1, columns === 2 ? 800 : 690),
        anchoring(5, anchoredRectangle(x, 300, 80, columns === 2 ? 60 : 100)), P,
      ], columns);
      const [placed] = pages[0]!.tables;
      const [drawing] = pages[0]!.drawings;
      expect(placed!.bottom + drawing!.height).toBeGreaterThan(PAGE_H);
      expect(drawing!.top).toBeCloseTo(300, 6);
      if (direction === 'tbRl') expect(drawing!.left + drawing!.width).toBeCloseTo(placed!.paintedLeft - 7.14, 6);
      else expect(drawing!.left).toBeCloseTo(placed!.paintedRight + 7.14, 6);
      expect(collisions(pages)).toEqual([]);
    });

    it('keeps a drawing where it is when the page has no place for it clear of its table', () => {
      // Twenty strips in, the 800 pt table of a two-column section runs
      // past the page's side edge and reaches its bottom edge: neither
      // below it nor past its far edge is there a place on the page, and
      // the rectangle stays where it was anchored, on the page.
      const pages = layOutPages('tbRl', [
        ...Array.from({ length: 20 }, () => P), table(1, 800), anchoring(5, anchoredRectangle(60, 300, 80, 60)), P,
      ], 2);
      expect(pages[0]!.tables[0]!.paintedLeft).toBeLessThan(0);
      expect(pages[0]!.drawings).toEqual([{ left: 60, top: 300, width: 80, height: 60 }]);
    });

    it('keeps drawings that may not overlap apart when a table moves one of them', () => {
      // Both rectangles may not overlap (ECMA-376 §20.4.2.3) and are
      // anchored 200 pt apart down the page. The second lies over the table
      // and moves below it, where it would meet the first: it lies directly
      // below the first instead, clear of both.
      const pages = layOutPages('tbRl', [
        table(1, 200),
        anchoring(5, anchoredRectangle(360, 300, 80, 60, 'page', 'page', false)),
        anchoring(5, anchoredRectangle(360, 100, 80, 60, 'page', 'page', false)),
      ]);
      const [first, second] = pages[0]!.drawings;
      expect(first).toEqual({ left: 360, top: 300, width: 80, height: 60 });
      expect(second!.left).toBeCloseTo(360, 6);
      expect(second!.top).toBeCloseTo(first!.top + first!.height, 6);
      expect(collisions(pages)).toEqual([]);
    });

    it.each(['tbRl', 'tbLrV'] as const)(
      'places a table against a page-positioned drawing where its collision with an earlier one leaves it (%s)',
      (direction) => {
        // Both rectangles are positioned on the page, so their exclusions
        // are registered when the page starts. The second may not overlap
        // (ECMA-376 §20.4.2.3) and lies directly below the first. The table
        // after them, whose box meets both, stacks below the second where
        // it lies, not where it was anchored.
        const x = direction === 'tbRl' ? 380 : 100;
        const pages = layOutPages(direction, [
          anchoring(5, anchoredRectangle(x, 60, 80, 60)),
          anchoring(5, anchoredRectangle(x, 80, 80, 60, 'page', 'page', false)),
          P, T, P,
        ]);
        const [first, second] = pages[0]!.drawings;
        expect(first!.top).toBeCloseTo(60, 6);
        expect(second!.top).toBeCloseTo(120, 6);
        expect(pages[0]!.tables[0]!.top).toBeCloseTo(180, 6);
        expect(collisions(pages)).toEqual([]);
      },
    );

    it.each([
      ['tbRl', 40, 360], ['tbLrV', 155, 100],
    ] as const)('moves a drawing past its table when another drawing leaves it no room below (%s)', (direction, firstX, x) => {
      // Neither rectangle may overlap. Below the 650 pt table the second
      // would meet the first, which reaches the page's bottom edge, and
      // past the first there is no page left. From its own position it has
      // room past the table's far edge, and lies there.
      const pages = layOutPages(direction, [
        table(1, 650),
        anchoring(5, anchoredRectangle(firstX, 715, 400, 85, 'page', 'page', false)),
        anchoring(5, anchoredRectangle(x, 100, 80, 60, 'page', 'page', false)),
      ]);
      const [placed] = pages[0]!.tables;
      const [first, second] = pages[0]!.drawings;
      expect(first).toEqual({ left: firstX, top: 715, width: 400, height: 85 });
      expect(second!.top).toBeCloseTo(100, 6);
      if (direction === 'tbRl') expect(second!.left + second!.width).toBeCloseTo(placed!.paintedLeft - 7.14, 6);
      else expect(second!.left).toBeCloseTo(placed!.paintedRight + 7.14, 6);
      expect(collisions(pages)).toEqual([]);
    });

    it('keeps a drawing at its position when every place clear of a table and a drawing is off the page', () => {
      // Neither rectangle may overlap. The second lies over the table and
      // the first; below both and past both there is no page left, and
      // clear of the first alone it would leave the page too. It stays
      // where it was anchored.
      const pages = layOutPages('tbRl', [
        table(1, 650),
        anchoring(5, anchoredRectangle(0, 680, 315, 110, 'page', 'page', false)),
        anchoring(5, anchoredRectangle(300, 680, 80, 60, 'page', 'page', false)),
      ]);
      expect(pages[0]!.drawings).toEqual([
        { left: 0, top: 680, width: 315, height: 110 },
        { left: 300, top: 680, width: 80, height: 60 },
      ]);
    });

    it('keeps a page-positioned drawing where the further layout of its page registered it', () => {
      // The first rectangle follows its line; the second, positioned on
      // the page, may not overlap and is placed below it. Laid out again
      // with the second registered there, the first paragraph's text and
      // rectangle move. The second stays where it is registered, so the
      // table after them meets it where it is painted.
      const pages = layOutPages('tbRl', [
        anchoring(5, anchoredRectangle(435, -30, 60, 60, 'line')),
        anchoring(5, anchoredRectangle(435, 90, 60, 60, 'page', 'page', false)),
        P, T, P,
      ]);
      expect(pages).toHaveLength(1);
      expect(pages[0]!.drawings).toHaveLength(2);
      expect(collisions(pages)).toEqual([]);
      const [first, second] = pages[0]!.drawings;
      expect(shared(first!.top, first!.top + first!.height, second!.top, second!.top + second!.height)).toBe(false);
    });

    /** A run anchoring a square-wrapped text box at page position (100, 60),
     * authored 100 x 20 pt and fitted to twelve paragraphs (a:spAutoFit). */
    const fittedTextBox = '<w:r><w:drawing><wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
      + ' distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
      + '<wp:simplePos x="0" y="0"/>'
      + '<wp:positionH relativeFrom="page"><wp:posOffset>1270000</wp:posOffset></wp:positionH>'
      + '<wp:positionV relativeFrom="page"><wp:posOffset>762000</wp:posOffset></wp:positionV>'
      + '<wp:extent cx="1270000" cy="254000"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="9" name="Box"/>'
      + '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
      + '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
      + '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:cNvSpPr txBox="1"/><wps:spPr>'
      + '<a:xfrm><a:off x="0" y="0"/><a:ext cx="1270000" cy="254000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>'
      + `<wps:txbx><w:txbxContent>${paragraph('ab').repeat(12)}</w:txbxContent></wps:txbx>`
      + '<wps:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"><a:spAutoFit/></wps:bodyPr>'
      + '</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>';

    it('wraps the strips of an earlier column around a fitted text box anchored in a continued paragraph', () => {
      // The paragraph that anchors the box runs on from column 1 into
      // column 2, where its last strips and the anchor lie. The box is
      // registered when the page starts with its authored height and fitted
      // by the paragraph's second fragment: the page is laid out again with
      // the fitted exclusion, and every strip that shares the box's block
      // range starts below it.
      const pages = layOutPages('tbLrV', [T, P, anchoring(825, fittedTextBox), P], 2);
      const [fitted] = pages[0]!.drawings;
      expect(fitted!.height).toBeGreaterThan(20);
      const beside = pages[0]!.strips.flat().filter((strip) =>
        shared(MARGIN.left + strip.block, MARGIN.left + strip.block + PITCH, fitted!.left, fitted!.left + fitted!.width));
      expect(beside.length).toBeGreaterThan(0);
      for (const strip of beside) {
        expect(strip.inlineStart).toBeGreaterThanOrEqual(fitted!.top + fitted!.height - 1e-6);
      }
    }, 60_000);

    it('places a table below a text box that grows to fit its text', () => {
      // A 100 x 20 pt text box positioned on the page, fitted to twelve
      // paragraphs (a:spAutoFit): its exclusion is registered when the page
      // starts with the authored height, and its paragraph gives it the
      // fitted one. The page is laid out again with the fitted exclusion,
      // and the table that meets the box stacks below it.
      const pages = layOutPages('tbLrV', [anchoring(5, fittedTextBox), P, T, P]);
      const [fitted] = pages[0]!.drawings;
      expect(fitted!.top).toBeCloseTo(60, 6);
      expect(fitted!.height).toBeGreaterThan(20);
      expect(pages[0]!.tables[0]!.top).toBeCloseTo(fitted!.top + fitted!.height, 6);
      expect(collisions(pages)).toEqual([]);
    });

    it.each(['tbRl', 'tbLrV'] as const)(
      'lays the page out again when a page-positioned drawing is placed past a text-positioned one (%s)',
      (direction) => {
        // The first rectangle is positioned from its paragraph, so it is
        // not known when the page starts and the second, positioned on the
        // page, has its exclusion registered where it was anchored. Its
        // paragraph then places it directly below the first, which it may
        // not overlap. The page is laid out once more with the exclusion
        // there: the table after them stacks below the second where it is
        // painted.
        const x = direction === 'tbRl' ? 380 : 100;
        const pages = layOutPages(direction, [
          anchoring(5, anchoredRectangle(x, 0, 80, 60, 'paragraph')),
          anchoring(5, anchoredRectangle(x, 80, 80, 60, 'page', 'page', false)),
          P, T, P,
        ]);
        const [first, second] = pages[0]!.drawings;
        expect(first!.top).toBeCloseTo(60, 6);
        expect(second!.top).toBeCloseTo(120, 6);
        expect(pages[0]!.tables[0]!.top).toBeCloseTo(180, 6);
        expect(collisions(pages)).toEqual([]);
      },
    );
  });
});
