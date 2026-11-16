import { describe, it, expect, vi } from 'vitest';
import { XlsxFindController, type FindCell } from './find.js';
import { cellDisplayTexts, formatCellValueWithColor, formatGeneralNumber } from './number-format.js';
import type { Styles, Worksheet } from './types.js';
import { XlsxWorkbook, xlsxCellDisplayTexts } from './workbook.js';

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

  it('finds a General number by the shortened form paint shows (issue #2179)', async () => {
    const general = makeStyles('General');
    const pi = { row: 2, col: 3, styleIndex: 0, value: { type: 'number' as const, number: Math.PI } };
    const [text, ...alternates] = cellDisplayTexts(pi, general);
    const c = controllerFor([{ name: 'S', cells: [{ row: 2, col: 3, text, alternates }] }]);
    // The width-fitted text (e.g. an 8-character column) and the full text both match.
    const shown = formatGeneralNumber(Math.PI, (form) => form.length <= 8);
    expect(shown).toBe('3.141593');
    const fitted = await c.find('3.141593');
    expect(fitted).toHaveLength(1);
    expect(fitted[0].text).toBe('3.141593');
    expect(fitted[0].location).toMatchObject({ ref: 'C2', row: 2, col: 3 });
    // A query every form contains still counts the cell once.
    expect(await c.find('3.14')).toHaveLength(1);
    expect(await c.find('3.141592654')).toHaveLength(1);
    // The highlight is the whole matched cell, independent of which form matched.
    c.next();
    expect(c.sheetHighlights(0)).toEqual([{ row: 2, col: 3, active: true }]);
  });

  it('indexes the forms of a General format chosen by conditional formatting', async () => {
    // Base 0.00 shows 3.14; a matched rule whose dxf selects General paints a
    // width-fitted General form such as 3.14159, which find must locate.
    const styles: Styles = {
      ...makeStyles('0.00'),
      dxfs: [{ font: null, fill: null, border: null, numFmt: { numFmtId: 0, formatCode: 'General' } }],
    };
    const pi = { row: 1, col: 1, styleIndex: 0, value: { type: 'number' as const, number: Math.PI } };
    const ws = {
      name: 'S', rows: [{ index: 1, height: null, cells: [pi] }],
      colWidths: {}, rowHeights: {}, defaultColWidth: 8.43, defaultRowHeight: 15,
      mergeCells: [], freezeRows: 0, freezeCols: 0, images: [], charts: [],
      conditionalFormats: [{
        sqref: [{ top: 1, left: 1, bottom: 1, right: 1 }],
        rules: [{ type: 'expression', formula: 'TRUE', dxfId: 0, priority: 1, stopIfTrue: false }],
      }],
    } as unknown as Worksheet;
    const hook = XlsxWorkbook.prototype[xlsxCellDisplayTexts] as
      (this: unknown, ws: Worksheet) => (cell: typeof pi) => string[];
    const [text, ...alternates] = hook.call({ parsedWorkbook: { styles } }, ws)(pi);
    expect(text).toBe('3.141592654');
    const c = controllerFor([{ name: 'S', cells: [{ row: 1, col: 1, text, alternates }] }]);
    expect(await c.find('3.14159')).toHaveLength(1);
    // Without the rule the base format alone applies.
    const plain = hook.call({ parsedWorkbook: { styles } }, { ...ws, conditionalFormats: [] } as Worksheet)(pi);
    expect(plain).toEqual(['3.14']);
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

describe('XlsxFindController text highlights', () => {
  const sheets = [
    { name: 'Summary', cells: [cell(1, 1, 'Revenue and guaranteed returns'), cell(2, 2, 'revenue, revenue')] },
    { name: 'Notes', cells: [cell(3, 1, 'Guaranteed returns')] },
  ];
  const groups = [
    { id: 'user', terms: ['revenue'], color: 'rgba(0,0,255,0.3)' },
    { id: 'validated', terms: ['guaranteed returns', 'returns'] },
  ];

  it('returns every group match in sheet and cell order with group and term', async () => {
    const c = controllerFor(sheets);
    const matches = await c.highlight(groups);
    expect(matches.map((m) => [m.matchIndex, m.group, m.term, m.text, m.location.sheetName, m.location.ref])).toEqual([
      [0, 'user', 'revenue', 'Revenue', 'Summary', 'A1'],
      [1, 'validated', 'guaranteed returns', 'guaranteed returns', 'Summary', 'A1'],
      [2, 'user', 'revenue', 'revenue', 'Summary', 'B2'],
      [3, 'user', 'revenue', 'revenue', 'Summary', 'B2'],
      [4, 'validated', 'guaranteed returns', 'Guaranteed returns', 'Notes', 'A3'],
    ]);
  });

  it('draws one box per group per cell, nested by depth, before the find boxes', async () => {
    const c = controllerFor(sheets);
    await c.find('and');
    await c.highlight(groups);
    // Group by group in the order of the set, then the find.
    expect(c.sheetHighlights(0)).toEqual([
      { row: 1, col: 1, active: false, color: 'rgba(0,0,255,0.3)', depth: 0 },
      { row: 2, col: 2, active: false, color: 'rgba(0,0,255,0.3)', depth: 0 },
      { row: 1, col: 1, active: false, color: 'rgba(0, 184, 148, 0.35)', depth: 1 },
      { row: 1, col: 1, active: false },
    ]);
  });

  it('nests the boxes of a cell in the order of the set, wherever the matches start', async () => {
    const c = controllerFor([{ name: 'S', cells: [cell(1, 1, 'returns before revenue')] }]);
    // The second group matches first in the text; it is still the inner box.
    const matches = await c.highlight(groups);
    expect(matches.map((m) => m.group)).toEqual(['validated', 'user']);
    expect(c.sheetHighlights(0).map((box) => [box.color, box.depth])).toEqual([
      ['rgba(0,0,255,0.3)', 0], ['rgba(0, 184, 148, 0.35)', 1],
    ]);
  });

  it('navigates independently of the find and marks the cell of the active occurrence', async () => {
    const c = controllerFor(sheets);
    await c.find('revenue');
    await c.highlight(groups);
    expect(c.next()?.location.ref).toBe('A1');
    expect(c.highlightStep(true)).toMatchObject({ matchIndex: 4, group: 'validated', location: { sheet: 1, ref: 'A3' } });
    expect(c.highlightStep()).toMatchObject({ matchIndex: 0, group: 'user' });
    c.highlightStep();
    c.highlightStep();
    expect(c.highlightStep()).toMatchObject({ matchIndex: 3, location: { ref: 'B2' } });
    expect(c.sheetHighlights(0).map((box) => box.active)).toEqual([false, true, false, true, false, false]);
    // The find cursor did not move.
    expect(c.next()?.matchIndex).toBe(1);
  });

  it('keeps each kind when the other is cleared, and drops both on invalidate', async () => {
    const c = controllerFor(sheets);
    await c.find('revenue');
    await c.highlight(groups);
    c.clearFind();
    expect(c.matches()).toEqual([]);
    expect(c.sheetHighlights(0)).toHaveLength(3);
    await c.find('revenue');
    c.clearHighlights();
    expect(c.highlightStep()).toBeNull();
    expect(c.matches()).toHaveLength(3);
    await c.highlight(groups);
    c.invalidate();
    expect(c.sheetHighlights(0)).toEqual([]);
  });

  it('replaces the set, clears on an empty set, and rejects a repeated id', async () => {
    const c = controllerFor(sheets);
    await c.highlight(groups);
    expect((await c.highlight([{ id: 'only', terms: ['notes', 'summary', 'returns'] }])).map((m) => m.location.ref))
      .toEqual(['A1', 'A3']);
    await expect(c.highlight([{ id: 'x', terms: ['a'] }, { id: 'x', terms: ['b'] }])).rejects.toThrow('duplicate group id');
    // Also when the set is otherwise empty: nothing is cleared.
    await expect(c.highlight([{ id: 'x', terms: [] }, { id: 'x', terms: [] }])).rejects.toThrow('duplicate group id');
    expect(c.sheetHighlights(0)).toHaveLength(1);
    expect(await c.highlight([{ id: 'only', terms: [] }])).toEqual([]);
    expect(c.sheetHighlights(0)).toEqual([]);
  });

  it('matches whole words per group and searches a cell\'s alternate texts per group', async () => {
    const c = new XlsxFindController(
      () => 1,
      () => 'S',
      async () => [
        { row: 1, col: 1, text: 'category cat' },
        { row: 2, col: 1, text: '0.123456789', alternates: ['0.1235'] },
      ],
    );
    const matches = await c.highlight([
      { id: 'word', terms: ['cat'], wholeWord: true },
      { id: 'full', terms: ['456'] },
      { id: 'short', terms: ['235'] },
    ]);
    expect(matches.map((m) => [m.group, m.text, m.location.ref])).toEqual([
      ['word', 'cat', 'A1'], ['short', '235', 'A2'], ['full', '456', 'A2'],
    ]);
  });

  it('commits only the latest highlight request and never a cleared one', async () => {
    const pending: Array<(cells: FindCell[]) => void> = [];
    const c = new XlsxFindController(
      () => 1,
      () => 'S',
      () => new Promise<FindCell[]>((resolve) => pending.push(resolve)),
    );
    const stale = c.highlight([{ id: 'a', terms: ['old'] }]);
    const current = c.highlight([{ id: 'a', terms: ['new'] }]);
    // The term search loads on demand; the superseded request never collects.
    await vi.waitFor(() => expect(pending).toHaveLength(1));
    pending[0]([cell(1, 1, 'old new')]);
    expect((await current).map((m) => m.text)).toEqual(['new']);
    await expect(stale).resolves.toEqual([]);
    const cleared = c.highlight([{ id: 'a', terms: ['old'] }]);
    await vi.waitFor(() => expect(pending).toHaveLength(2));
    c.clearHighlights();
    pending[1]([cell(1, 1, 'old new')]);
    await expect(cleared).resolves.toEqual([]);
    expect(c.sheetHighlights(0)).toEqual([]);
  });
});

describe('XlsxFindController.find wholeWord', () => {
  it('matches whole words within each cell', async () => {
    const c = controllerFor([{ name: 'S', cells: [cell(1, 1, 'category cat'), cell(2, 1, 'Cat'), cell(3, 1, 'cats')] }]);
    expect((await c.find('cat', { wholeWord: true })).map((m) => [m.location.ref, m.text]))
      .toEqual([['A1', 'cat'], ['A2', 'Cat']]);
    expect(await c.find('cat')).toHaveLength(4);
  });
});

describe('XlsxFindController request results', () => {
  it('reports each request its own result, and no error of a superseded request', async () => {
    const pending: Array<{ resolve: (cells: FindCell[]) => void; reject: (error: Error) => void }> = [];
    const c = new XlsxFindController(
      () => 1,
      () => 'S',
      () => new Promise<FindCell[]>((resolve, reject) => pending.push({ resolve, reject })),
    );
    const first = c.find('alpha');
    pending[0].resolve([cell(1, 1, 'alpha beta')]);
    const second = c.find('beta');
    expect((await first).map((m) => m.text)).toEqual([]);
    pending[1].resolve([cell(1, 1, 'alpha beta')]);
    expect((await second).map((m) => m.text)).toEqual(['beta']);

    const failing = c.find('alpha');
    c.clearFind();
    pending[2].reject(new Error('sheet closed'));
    await expect(failing).resolves.toEqual([]);
    const current = c.find('alpha');
    pending[3].reject(new Error('sheet failed'));
    await expect(current).rejects.toThrow('sheet failed');
  });
});
