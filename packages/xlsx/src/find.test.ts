import { describe, it, expect } from 'vitest';
import { XlsxFindController, type FindCell } from './find.js';
import { formatCellValueWithColor } from './number-format.js';
import type { Styles } from './types.js';

/** Minimal Styles carrying a single custom numFmt at style index 0. */
function makeStyles(formatCode: string): Styles {
  return {
    fonts: [], fills: [], borders: [], dxfs: [],
    cellXfs: [{ fontId: 0, fillId: 0, borderId: 0, numFmtId: 164, alignH: null, alignV: null, wrapText: false }],
    numFmts: [{ numFmtId: 164, formatCode }],
  };
}

/**
 * IX2 xlsx find controller. Cell-based: each non-empty cell's rendered text is
 * searched; a match lands wholly in one cell and reports `{ sheet, sheetName,
 * ref, row, col }`. Every occurrence within a cell counts (browser find-count).
 */
function controllerFor(sheets: { name: string; cells: FindCell[] }[]): XlsxFindController {
  return new XlsxFindController(
    () => sheets.length,
    (sheet) => sheets[sheet]?.name ?? '',
    (sheet) => Promise.resolve(sheets[sheet]?.cells ?? []),
  );
}

const cell = (row: number, col: number, text: string): FindCell => ({ row, col, text });

describe('XlsxFindController.find', () => {
  it('searches several terms at once, and treats an all-empty list as a clear', async () => {
    const c = controllerFor([{ name: 'S', cells: [cell(1, 1, 'cat'), cell(1, 2, 'concatenate'), cell(2, 1, 'dog')] }]);
    const matches = await c.find(['dog', 'cat'], { wholeWord: true });
    expect(matches.map((m) => m.location.ref)).toEqual(['A1', 'A2']);
    expect(await c.find([''])).toEqual([]);
    expect(c.matches()).toEqual([]);
  });

  it("carries each term's colour into the matches and the highlights", async () => {
    const c = controllerFor([{ name: 'S', cells: [cell(1, 1, 'cat'), cell(2, 1, 'dog')] }]);
    const matches = await c.find([{ text: 'dog', color: 'red' }, 'cat']);
    expect(matches.map((m) => [m.location.ref, m.color])).toEqual([['A1', undefined], ['A2', 'red']]);
    expect(c.sheetHighlights(0).map((h) => h.color)).toEqual([undefined, 'red']);
    c.next();
    expect(c.next()?.color).toBe('red');
  });

  it('commits only the latest overlapping search', async () => {
    const pending: Array<(cells: FindCell[]) => void> = [];
    const c = new XlsxFindController(
      () => 1,
      () => 'S',
      () => new Promise<FindCell[]>((resolve) => pending.push(resolve)),
    );

    const stale = c.find('old');
    const current = c.find('new');
    pending[1]([cell(1, 1, 'new')]);
    await expect(current).resolves.toHaveLength(1);
    pending[0]([cell(1, 1, 'old')]);
    await expect(stale).resolves.toEqual([]);
    expect(c.matches().map((match) => match.text)).toEqual(['new']);
  });

  it('does not commit a search invalidated while collecting cells', async () => {
    let resolveCells: (cells: FindCell[]) => void = () => undefined;
    const c = new XlsxFindController(
      () => 1,
      () => 'S',
      () => new Promise<FindCell[]>((resolve) => { resolveCells = resolve; }),
    );

    const stale = c.find('old');
    c.invalidate();
    resolveCells([cell(1, 1, 'old')]);
    await expect(stale).resolves.toEqual([]);
    expect(c.matches()).toEqual([]);
  });

  it('suppresses a stale collection rejection after invalidation', async () => {
    let rejectCells: (reason: Error) => void = () => undefined;
    const c = new XlsxFindController(
      () => 1,
      () => 'S',
      () => new Promise<FindCell[]>((_resolve, reject) => { rejectCells = reject; }),
    );

    const stale = c.find('old');
    c.invalidate();
    rejectCells(new Error('old workbook closed'));

    await expect(stale).resolves.toEqual([]);
  });

  it('propagates a collection rejection for the current query', async () => {
    const c = new XlsxFindController(
      () => 1,
      () => 'S',
      () => Promise.reject(new Error('current workbook failed')),
    );

    await expect(c.find('current')).rejects.toThrow('current workbook failed');
  });

  it('finds a match in a cell and reports its A1 ref + row/col', async () => {
    const c = controllerFor([{ name: 'Sales', cells: [cell(7, 2, 'Revenue')] }]);
    const matches = await c.find('rev');
    expect(matches).toHaveLength(1);
    expect(matches[0].location).toMatchObject({
      sheet: 0,
      sheetName: 'Sales',
      ref: 'B7',
      row: 7,
      col: 2,
    });
    // Reported text is the matched substring in the cell's original case.
    expect(matches[0].text).toBe('Rev');
  });

  it('counts every occurrence within a cell', async () => {
    const c = controllerFor([{ name: 'S', cells: [cell(1, 1, 'the cat and the dog')] }]);
    const matches = await c.find('the');
    expect(matches).toHaveLength(2);
    expect(matches.every((m) => m.location.ref === 'A1')).toBe(true);
  });

  it('searches across sheets in document order', async () => {
    const c = controllerFor([
      { name: 'One', cells: [cell(1, 1, 'apple')] },
      { name: 'Two', cells: [cell(2, 3, 'grape')] },
    ]);
    const matches = await c.find('ap'); // "apple" and "grape" both contain "ap"
    expect(matches).toHaveLength(2);
    expect(matches[0].location.sheet).toBe(0);
    expect(matches[1].location.sheet).toBe(1);
    expect(matches[1].location.ref).toBe('C2');
  });

  it('is case-insensitive by default; caseSensitive honored', async () => {
    const ci = await controllerFor([{ name: 'S', cells: [cell(1, 1, 'FOO foo')] }]).find('foo');
    expect(ci).toHaveLength(2);
    const cs = await controllerFor([{ name: 'S', cells: [cell(1, 1, 'FOO foo')] }]).find('foo', {
      caseSensitive: true,
    });
    expect(cs).toHaveLength(1);
  });

  it('returns [] for an empty query', async () => {
    const c = controllerFor([{ name: 'S', cells: [cell(1, 1, 'x')] }]);
    expect(await c.find('')).toEqual([]);
  });

  it('matches a query against the number-format display string (XL3)', async () => {
    // findText searches the *rendered* text (the same string the grid draws via
    // formatCellValue), so a value stored as 1234.5 but formatted with a yen
    // currency + grouping code is found by its displayed grouping — "1,234" —
    // and by the currency glyph, but NOT by the raw "1234.5".
    const fc = formatCellValueWithColor(
      { row: 5, col: 3, value: { type: 'number', number: 1234.5 }, styleIndex: 0 },
      makeStyles('[$¥-411]#,##0.00'),
    );
    expect(fc.text).toBe('¥1,234.50'); // the display string
    const c = controllerFor([{ name: 'S', cells: [cell(5, 3, fc.text)] }]);
    expect(await c.find('1,234')).toHaveLength(1);
    expect(await c.find('¥')).toHaveLength(1);
    expect(await c.find('1234.5')).toHaveLength(0); // raw value is not searchable
  });
});

describe('XlsxFindController cursor + highlights', () => {
  it('cycles the active match across sheets with wrap-around', async () => {
    const c = controllerFor([
      { name: 'A', cells: [cell(1, 1, 'x')] },
      { name: 'B', cells: [cell(1, 1, 'x')] },
    ]);
    await c.find('x');
    expect(c.next()?.matchIndex).toBe(0);
    expect(c.activeLocation()?.sheet).toBe(0);
    expect(c.next()?.matchIndex).toBe(1);
    expect(c.activeLocation()?.sheet).toBe(1);
    expect(c.next()?.matchIndex).toBe(0); // wrap
    expect(c.prev()?.matchIndex).toBe(1); // wrap back
  });

  it('sheetHighlights scopes to sheet and marks active', async () => {
    const c = controllerFor([{ name: 'S', cells: [cell(1, 1, 'a'), cell(2, 1, 'a')] }]);
    await c.find('a');
    c.next(); // active = match 0 (A1)
    const hl = c.sheetHighlights(0);
    expect(hl).toHaveLength(2);
    expect(hl.find((h) => h.row === 1)?.active).toBe(true);
    expect(hl.find((h) => h.row === 2)?.active).toBe(false);
    expect(c.sheetHighlights(1)).toHaveLength(0);
  });

  it('invalidate clears matches + cursor', async () => {
    const c = controllerFor([{ name: 'S', cells: [cell(1, 1, 'a')] }]);
    await c.find('a');
    c.invalidate();
    expect(c.matches()).toHaveLength(0);
    expect(c.activeLocation()).toBeNull();
    expect(c.next()).toBeNull();
  });
});
