/// <reference types="node" />
import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import {
  expandSharedFormula,
  recalculateWorksheet,
  recalculateWorksheets,
  worksheetReferencesSheet,
} from './recalc.js';
import type { Cell, Row, SharedString, Worksheet } from './types.js';

function cell(row: number, col: number, value: Cell['value'], formula?: string): Cell {
  return formula === undefined
    ? { row, col, value }
    : { row, col, value, formula };
}

/** A `shared`-formula follower as the extended parser emits it: no `<f>`
 *  text, only `t`/`si` attributes plus its cached `<v>`. */
function follower(row: number, col: number, value: Cell['value'], si: number): Cell {
  return { row, col, value, formulaType: 'shared', formulaSi: si };
}

function sheet(
  name: string,
  cells: Cell[],
  opts?: { date1904?: boolean },
): Worksheet {
  const byRow = new Map<number, Cell[]>();
  for (const c of cells) {
    const list = byRow.get(c.row) ?? [];
    list.push(c);
    byRow.set(c.row, list);
  }
  const rows: Row[] = [...byRow.entries()]
    .sort((a, b) => a[0] - b[0])
    .map(([index, rowCells]) => ({ index, height: null, cells: rowCells }));
  return {
    name,
    rows,
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
    date1904: opts?.date1904,
  } as Worksheet;
}

const NO_STRINGS: SharedString[] = [];

/** Engine bytes for environments without a fetchable `?url` asset (vitest). */
const WASM_BYTES = new Uint8Array(
  readFileSync(new URL('./vendor/formualizer_wasm_bg.wasm', import.meta.url)),
);

function get(ws: Worksheet, row: number, col: number): Cell {
  const found = ws.rows.flatMap((r) => r.cells).find((c) => c.row === row && c.col === col);
  if (!found) throw new Error(`cell r${row}c${col} missing`);
  return found;
}

describe('recalc: SUM / chain across cells', () => {
  it('recomputes a cached value that disagrees with the formula', async () => {
    const ws = sheet('S1', [
      cell(1, 1, { type: 'number', number: 1 }),
      cell(2, 1, { type: 'number', number: 2 }),
      // Stale cache (would be 3): recalc overwrites with the computed value.
      cell(3, 1, { type: 'number', number: 999 }, 'SUM(A1:A2)'),
    ]);
    const stats = await recalculateWorksheet(ws, { sharedStrings: NO_STRINGS, date1904: false, formualizerWasm: WASM_BYTES });
    expect(get(ws, 3, 1).value).toEqual({ type: 'number', number: 3 });
    expect(get(ws, 3, 1).formula).toBe('SUM(A1:A2)');
    expect(stats).toMatchObject({ recalculated: 1, failed: 0, skipped: 0 });
  });

  it('follows multi-step chains within one run', async () => {
    const ws = sheet('S1', [
      cell(1, 1, { type: 'number', number: 4 }),
      cell(2, 1, { type: 'number', number: 0 }, 'A1*2'),
      cell(3, 1, { type: 'number', number: 0 }, 'A2+1'),
    ]);
    await recalculateWorksheet(ws, { sharedStrings: NO_STRINGS, date1904: false, formualizerWasm: WASM_BYTES });
    expect(get(ws, 2, 1).value).toEqual({ type: 'number', number: 8 });
    expect(get(ws, 3, 1).value).toEqual({ type: 'number', number: 9 });
  });

  it('leaves literal cells untouched', async () => {
    const ws = sheet('S1', [
      cell(1, 1, { type: 'number', number: 1.5 }),
      cell(1, 2, { type: 'text', text: 'hi' }),
      cell(1, 3, { type: 'number', number: 0 }, 'A1*2'),
    ]);
    await recalculateWorksheet(ws, { sharedStrings: NO_STRINGS, date1904: false, formualizerWasm: WASM_BYTES });
    expect(get(ws, 1, 1).value).toEqual({ type: 'number', number: 1.5 });
    expect(get(ws, 1, 2).value).toEqual({ type: 'text', text: 'hi' });
  });
});

describe('recalc: cross-sheet references', () => {
  it('recalculates across sheets sharing one engine (two-phase)', async () => {
    const s1 = sheet('Sheet1', [
      cell(1, 1, { type: 'number', number: 0 }, 'Sheet2!A1*2'),
    ]);
    const s2 = sheet('Sheet2', [cell(1, 1, { type: 'number', number: 10 })]);
    const stats = await recalculateWorksheets(
      [
        { name: 'Sheet1', ws: s1 },
        { name: 'Sheet2', ws: s2 },
      ],
      { sharedStrings: NO_STRINGS, date1904: false, formualizerWasm: WASM_BYTES },
    );
    expect(get(s1, 1, 1).value).toEqual({ type: 'number', number: 20 });
    expect(stats.get('sheet1')?.recalculated).toBe(1);
  });

  it('resolves cross-sheet inputs via resolveSheet for single-sheet recalc', async () => {
    const s1 = sheet('Sheet1', [
      cell(1, 1, { type: 'number', number: 0 }, 'Sheet2!A1+5'),
    ]);
    const s2 = sheet('Sheet2', [cell(1, 1, { type: 'number', number: 7 })]);
    const stats = await recalculateWorksheet(s1, {
      sharedStrings: NO_STRINGS,
      date1904: false,
      formualizerWasm: WASM_BYTES,
      resolveSheet: (name) => (name.toLowerCase() === 'sheet2' ? s2 : undefined),
    });
    expect(get(s1, 1, 1).value).toEqual({ type: 'number', number: 12 });
    // The referenced sheet is read-only input: never mutated.
    expect(get(s2, 1, 1).value).toEqual({ type: 'number', number: 7 });
    expect(stats.recalculated).toBe(1);
  });

  it('awaits async resolvers (lazy lease loads)', async () => {
    const s1 = sheet('Sheet1', [
      cell(1, 1, { type: 'number', number: 0 }, 'Sheet2!A1*3'),
    ]);
    const s2 = sheet('Sheet2', [cell(1, 1, { type: 'number', number: 5 })]);
    const stats = await recalculateWorksheet(s1, {
      sharedStrings: NO_STRINGS,
      date1904: false,
      formualizerWasm: WASM_BYTES,
      resolveSheet: async (name) => (name.toLowerCase() === 'sheet2' ? s2 : undefined),
    });
    expect(get(s1, 1, 1).value).toEqual({ type: 'number', number: 15 });
    expect(stats.recalculated).toBe(1);
  });

  it('keeps cached values when a referenced sheet does not resolve', async () => {
    const s1 = sheet('Sheet1', [
      cell(1, 1, { type: 'number', number: 42 }, 'Missing!A1+1'),
    ]);
    const stats = await recalculateWorksheet(s1, {
      sharedStrings: NO_STRINGS,
      date1904: false,
      formualizerWasm: WASM_BYTES,
      resolveSheet: () => undefined,
    });
    // Formualizer reports #REF! for the unknown sheet — an Excel code, so
    // the recomputed error is written (not silently cached).
    expect(get(s1, 1, 1).value).toEqual({ type: 'error', error: '#REF!' });
    expect(stats.recalculated).toBe(1);
  });

  it('captures overwritten cells for worker-boundary forwarding', async () => {
    const ws = sheet('S1', [
      cell(1, 1, { type: 'number', number: 2 }),
      cell(2, 1, { type: 'number', number: 0 }, 'A1*2'),
      cell(3, 1, { type: 'text', text: 'plain' }),
    ]);
    const captured: Array<{ row: number; col: number; value: Cell['value'] }> = [];
    await recalculateWorksheet(ws, { sharedStrings: NO_STRINGS, date1904: false, formualizerWasm: WASM_BYTES }, captured);
    expect(captured).toEqual([{ row: 2, col: 1, value: { type: 'number', number: 4 } }]);
  });
});

describe('recalc: volatile consistency (TODAY)', () => {
  it('recalculates TODAY() and its dependents in one graph run', async () => {
    const ws = sheet('S1', [
      cell(1, 1, { type: 'number', number: 0 }, 'TODAY()'),
      cell(2, 1, { type: 'number', number: 0 }, 'A1+1'),
    ]);
    await recalculateWorksheet(ws, { sharedStrings: NO_STRINGS, date1904: false, formualizerWasm: WASM_BYTES });
    const today = get(ws, 1, 1).value;
    const tomorrow = get(ws, 2, 1).value;
    expect(today.type).toBe('number');
    expect(tomorrow.type).toBe('number');
    if (today.type === 'number' && tomorrow.type === 'number') {
      // Coherent within the run: the dependent trails the volatile by exactly 1.
      expect(tomorrow.number - today.number).toBe(1);
      // Sanity: a plausible serial, not a blank/zero fallback.
      expect(today.number).toBeGreaterThan(40000);
    }
  });
});

describe('recalc: errors', () => {
  it('maps =1/0 to the #DIV/0! error branch (not text/number)', async () => {
    const ws = sheet('S1', [cell(1, 1, { type: 'number', number: 0 }, '1/0')]);
    const stats = await recalculateWorksheet(ws, { sharedStrings: NO_STRINGS, date1904: false, formualizerWasm: WASM_BYTES });
    expect(get(ws, 1, 1).value).toEqual({ type: 'error', error: '#DIV/0!' });
    expect(stats.recalculated).toBe(1);
  });

  it('propagates cached input errors through dependents', async () => {
    const ws = sheet('S1', [
      cell(1, 1, { type: 'error', error: '#N/A' }),
      cell(2, 1, { type: 'number', number: 0 }, 'A1+1'),
    ]);
    await recalculateWorksheet(ws, { sharedStrings: NO_STRINGS, date1904: false, formualizerWasm: WASM_BYTES });
    expect(get(ws, 2, 1).value).toEqual({ type: 'error', error: '#N/A' });
  });

  it('keeps the cached value on circular references (#CYCLE! is not Excel)', async () => {
    const ws = sheet('S1', [
      cell(1, 1, { type: 'number', number: 5 }, 'B1+1'),
      cell(1, 2, { type: 'number', number: 6 }, 'A1+1'),
    ]);
    const stats = await recalculateWorksheet(ws, { sharedStrings: NO_STRINGS, date1904: false, formualizerWasm: WASM_BYTES });
    expect(get(ws, 1, 1).value).toEqual({ type: 'number', number: 5 });
    expect(get(ws, 1, 2).value).toEqual({ type: 'number', number: 6 });
    expect(stats.failed).toBe(2);
    expect(stats.recalculated).toBe(0);
  });
});

describe('recalc: shared-string inputs', () => {
  it('resolves shared{si} before feeding so text refs compute', async () => {
    const strings: SharedString[] = [{ text: 'hello' }];
    const ws = sheet('S1', [
      cell(1, 1, { type: 'shared', si: 0 }),
      cell(2, 1, { type: 'text', text: '' }, 'LEN(A1)'),
    ]);
    await recalculateWorksheet(ws, { sharedStrings: strings, date1904: false, formualizerWasm: WASM_BYTES });
    expect(get(ws, 2, 1).value).toEqual({ type: 'number', number: 5 });
  });
});

describe('recalc: date systems', () => {
  it('matches the 1900 epoch by default (DATE serial 44927)', async () => {
    const ws = sheet('S1', [cell(1, 1, { type: 'number', number: 0 }, 'DATE(2023,1,1)')]);
    await recalculateWorksheet(ws, { sharedStrings: NO_STRINGS, date1904: false, formualizerWasm: WASM_BYTES });
    expect(get(ws, 1, 1).value).toEqual({ type: 'number', number: 44927 });
  });

  it('matches the 1904 epoch when date1904 is set (serial 43465, 1462 lower)', async () => {
    const ws = sheet('S1', [cell(1, 1, { type: 'number', number: 0 }, 'DATE(2023,1,1)')], {
      date1904: true,
    });
    await recalculateWorksheet(ws, { sharedStrings: NO_STRINGS, date1904: true, formualizerWasm: WASM_BYTES });
    expect(get(ws, 1, 1).value).toEqual({ type: 'number', number: 43465 });
  });
});

describe('recalc: uncached formula cells (no <v>)', () => {
  it('renders a computed value where the default path shows blank', async () => {
    const ws = sheet('S1', [
      cell(1, 1, { type: 'number', number: 3 }),
      cell(1, 2, { type: 'number', number: 4 }),
      cell(1, 3, { type: 'empty' }, 'SUM(A1:B1)'),
    ]);
    await recalculateWorksheet(ws, { sharedStrings: NO_STRINGS, date1904: false, formualizerWasm: WASM_BYTES });
    expect(get(ws, 1, 3).value).toEqual({ type: 'number', number: 7 });
  });
});

describe('recalc: resource cap', () => {
  it('skips the whole sheet over budget and keeps cached values', async () => {
    const ws = sheet('S1', [
      cell(1, 1, { type: 'number', number: 1 }),
      cell(2, 1, { type: 'number', number: 0 }, 'A1+1'),
    ]);
    const stats = await recalculateWorksheet(ws, {
      sharedStrings: NO_STRINGS,
      date1904: false,
      formualizerWasm: WASM_BYTES,
      maxCells: 1,
    });
    expect(get(ws, 2, 1).value).toEqual({ type: 'number', number: 0 });
    expect(stats).toMatchObject({ recalculated: 0, failed: 0, skipped: 1 });
  });

  it('maps an over-limit REPT to #VALUE! (engine caps >32,767 chars)', async () => {
    const ws = sheet('S1', [
      cell(1, 1, { type: 'text', text: 'cached' }, 'REPT("x",40000)'),
    ]);
    const stats = await recalculateWorksheet(ws, { sharedStrings: NO_STRINGS, date1904: false, formualizerWasm: WASM_BYTES });
    // Formualizer caps the result as Excel does: #VALUE! is an Excel error,
    // so the recomputed error is written (not silently cached).
    expect(get(ws, 1, 1).value).toEqual({ type: 'error', error: '#VALUE!' });
    expect(stats).toMatchObject({ recalculated: 1, failed: 0, skipped: 0 });
  });

  it('keeps the cached value when a text result exceeds the 32,767-char cell limit', async () => {
    const ws = sheet('S1', [
      cell(1, 1, { type: 'text', text: 'y'.repeat(40000) }),
      cell(1, 2, { type: 'text', text: 'cached' }, 'A1'),
    ]);
    const stats = await recalculateWorksheet(ws, { sharedStrings: NO_STRINGS, date1904: false, formualizerWasm: WASM_BYTES });
    expect(get(ws, 1, 2).value).toEqual({ type: 'text', text: 'cached' });
    expect(stats).toMatchObject({ recalculated: 0, failed: 1, skipped: 0 });
  });
});

describe('worksheetReferencesSheet', () => {
  it('matches qualified refs case-insensitively, ignoring string literals', () => {
    const ws = sheet('S1', [
      cell(1, 1, { type: 'number', number: 0 }, 'Sheet2!A1+1'),
      cell(2, 1, { type: 'text', text: 'x' }, '"Sheet2!A1"'),
    ]);
    expect(worksheetReferencesSheet(ws, 'sheet2')).toBe(true);
    expect(worksheetReferencesSheet(ws, 'SHEET2')).toBe(true);
    expect(worksheetReferencesSheet(ws, 'Sheet3')).toBe(false);
  });

  it('matches quoted sheet names', () => {
    const ws = sheet('S1', [
      cell(1, 1, { type: 'number', number: 0 }, "'My Sheet'!A1"),
    ]);
    expect(worksheetReferencesSheet(ws, 'My Sheet')).toBe(true);
    expect(worksheetReferencesSheet(ws, 'My')).toBe(false);
  });

  it('is false without formulas or qualifiers', () => {
    const ws = sheet('S1', [
      cell(1, 1, { type: 'number', number: 1 }),
      cell(2, 1, { type: 'number', number: 0 }, 'A1+1'),
    ]);
    expect(worksheetReferencesSheet(ws, 'S1')).toBe(false);
  });
});

describe('recalc: shared formulas (anchor + followers)', () => {
  it('expands the anchor to followers with shifted relative refs', async () => {
    const ws = sheet('S1', [
      cell(1, 1, { type: 'number', number: 3 }),
      cell(2, 1, { type: 'number', number: 4 }),
      // Anchor B1 as the extended parser emits it (text + t/ref/si).
      { row: 1, col: 2, value: { type: 'number', number: 0 }, formula: 'A1*2', formulaType: 'shared', formulaRef: 'B1:B2', formulaSi: 0 },
      // Follower B2: no text, only t/si, with a stale cache.
      follower(2, 2, { type: 'number', number: 0 }, 0),
    ]);
    const stats = await recalculateWorksheet(ws, { sharedStrings: NO_STRINGS, date1904: false, formualizerWasm: WASM_BYTES });
    expect(get(ws, 1, 2).value).toEqual({ type: 'number', number: 6 });
    // =A1*2 shifted by (2-1, 2-2) → =A2*2 → 8, not the stale 0.
    expect(get(ws, 2, 2).value).toEqual({ type: 'number', number: 8 });
    expect(stats).toMatchObject({ recalculated: 2, failed: 0, skipped: 0 });
  });

  it('respects $ pins when expanding to followers', async () => {
    const ws = sheet('S1', [
      cell(1, 1, { type: 'number', number: 3 }),
      cell(2, 1, { type: 'number', number: 4 }),
      cell(1, 3, { type: 'number', number: 10 }),
      { row: 1, col: 2, value: { type: 'number', number: 0 }, formula: '$A1+$C$1', formulaType: 'shared', formulaRef: 'B1:B2', formulaSi: 0 },
      follower(2, 2, { type: 'number', number: 0 }, 0),
    ]);
    await recalculateWorksheet(ws, { sharedStrings: NO_STRINGS, date1904: false, formualizerWasm: WASM_BYTES });
    expect(get(ws, 1, 2).value).toEqual({ type: 'number', number: 13 });
    // $A1 → $A2 (13→14), $C$1 pinned → 4+10.
    expect(get(ws, 2, 2).value).toEqual({ type: 'number', number: 14 });
  });

  it('keeps orphan followers (no anchor for si) on cached values', async () => {
    const ws = sheet('S1', [follower(1, 1, { type: 'number', number: 42 }, 7)]);
    const stats = await recalculateWorksheet(ws, { sharedStrings: NO_STRINGS, date1904: false, formualizerWasm: WASM_BYTES });
    expect(get(ws, 1, 1).value).toEqual({ type: 'number', number: 42 });
    expect(stats).toMatchObject({ recalculated: 0, failed: 0, skipped: 0 });
  });
});

describe('recalc: dataTable formulas are never evaluated', () => {
  it('keeps the cached value instead of reporting #NAME? for TABLE()', async () => {
    const ws = sheet('S1', [
      {
        row: 1, col: 1, value: { type: 'number', number: 5 },
        formula: 'TABLE(,C1)', formulaType: 'dataTable', formulaRef: 'A1:A2',
      },
    ]);
    const stats = await recalculateWorksheet(ws, { sharedStrings: NO_STRINGS, date1904: false, formualizerWasm: WASM_BYTES });
    expect(get(ws, 1, 1).value).toEqual({ type: 'number', number: 5 });
    expect(stats).toMatchObject({ recalculated: 0, failed: 1, skipped: 0 });
  });
});

describe('expandSharedFormula: relative shifting with $ pins', () => {
  it('shifts relative refs by the follower delta', () => {
    expect(expandSharedFormula('A1*2', 1, 2, 2, 2)).toBe('A2*2');
    expect(expandSharedFormula('SUM(A1:A3)', 1, 1, 4, 1)).toBe('SUM(A4:A6)');
  });

  it('pins $ coordinates', () => {
    expect(expandSharedFormula('$A1*2', 1, 2, 2, 3)).toBe('$A2*2');
    expect(expandSharedFormula('A$1*2', 1, 2, 2, 3)).toBe('B$1*2');
    expect(expandSharedFormula('$A$1*2', 1, 2, 9, 9)).toBe('$A$1*2');
  });

  it('shifts through sheet qualifiers but never rewrites string literals', () => {
    expect(expandSharedFormula('Sheet2!A1*2', 1, 1, 2, 1)).toBe('Sheet2!A2*2');
    expect(expandSharedFormula("'My Sheet'!A1", 1, 1, 3, 1)).toBe("'My Sheet'!A3");
    expect(expandSharedFormula('"A1"&A1', 1, 1, 2, 1)).toBe('"A1"&A2');
  });

  it('leaves names, functions and structured refs untouched', () => {
    expect(expandSharedFormula('MyName+A1', 1, 1, 2, 1)).toBe('MyName+A2');
    expect(expandSharedFormula('LOG10(A1)', 1, 1, 2, 1)).toBe('LOG10(A2)');
  });

  it('emits #REF! when shifting off-grid', () => {
    expect(expandSharedFormula('A1', 2, 1, 1, 1)).toBe('#REF!');
  });

  it('returns the anchor formula unchanged at zero delta', () => {
    expect(expandSharedFormula('A1*2', 1, 1, 1, 1)).toBe('A1*2');
  });
});
