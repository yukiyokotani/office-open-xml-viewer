import { describe, it, expect } from 'vitest';
import { evalFormulaToBool } from './formula.js';
import type { Cell } from './types.js';

function numCell(row: number, col: number, n: number): Cell {
  return { row, col, value: { type: 'number', number: n }, styleIndex: 0 };
}

function ctx(opts: { cells?: Cell[]; row?: number; col?: number } = {}) {
  const cellIndex = new Map<string, Cell>();
  for (const c of opts.cells ?? []) cellIndex.set(`${c.row}:${c.col}`, c);
  return {
    row: opts.row ?? 1,
    col: opts.col ?? 1,
    anchorRow: 1,
    anchorCol: 1,
    cellIndex,
    definedNames: new Map(),
    depth: 0,
  };
}

const ev = (f: string, c = ctx()) => evalFormulaToBool(f, c);

describe('evalFormulaToBool — comparisons', () => {
  it('numeric comparisons', () => {
    expect(ev('1>0')).toBe(true);
    expect(ev('1<0')).toBe(false);
    expect(ev('2>=2')).toBe(true);
    expect(ev('2<=1')).toBe(false);
    expect(ev('3=3')).toBe(true);
    expect(ev('3<>3')).toBe(false);
    expect(ev('3<>4')).toBe(true);
  });

  it('arithmetic before comparison', () => {
    expect(ev('2+2=4')).toBe(true);
    expect(ev('10-3*2=4')).toBe(true);
    expect(ev('(10-2)/4=2')).toBe(true);
  });

  it('string comparison', () => {
    expect(ev('"a"="a"')).toBe(true);
    expect(ev('"a"="b"')).toBe(false);
  });
});

describe('evalFormulaToBool — logical functions', () => {
  it('AND / OR / NOT', () => {
    expect(ev('AND(1>0,2>1)')).toBe(true);
    expect(ev('AND(1>0,2>3)')).toBe(false);
    expect(ev('OR(1>2,3>2)')).toBe(true);
    expect(ev('OR(1>2,3>4)')).toBe(false);
    expect(ev('NOT(1>2)')).toBe(true);
    expect(ev('NOT(1>0)')).toBe(false);
  });

  it('IF', () => {
    expect(ev('IF(1>0,1>0,1<0)')).toBe(true);
    expect(ev('IF(1<0,1>0,1<0)')).toBe(false);
  });

  it('nested logic', () => {
    expect(ev('AND(OR(1>2,2>1),NOT(3>4))')).toBe(true);
  });
});

describe('evalFormulaToBool — cell references', () => {
  it('reads a referenced cell (relative shift from anchor)', () => {
    // anchor A1, evaluated at A1 → A1 resolves to the cell at (1,1)
    const c = ctx({ cells: [numCell(1, 1, 90)], row: 1, col: 1 });
    expect(evalFormulaToBool('A1>=90', c)).toBe(true);
    expect(evalFormulaToBool('A1>90', c)).toBe(false);
  });
});

describe('evalFormulaToBool — numeric functions', () => {
  it('ABS / INT / MOD', () => {
    expect(ev('ABS(-5)=5')).toBe(true);
    expect(ev('INT(5.9)=5')).toBe(true);
    expect(ev('INT(-5.1)=-6')).toBe(true); // Excel INT rounds toward -infinity
    expect(ev('MOD(7,3)=1')).toBe(true);
  });
  it('CEILING / FLOOR', () => {
    expect(ev('CEILING(4.2,1)=5')).toBe(true);
    expect(ev('FLOOR(4.8,1)=4')).toBe(true);
  });
});

describe('evalFormulaToBool — IS / text functions', () => {
  it('ISBLANK', () => {
    const c = ctx({ cells: [numCell(1, 1, 0)], row: 1, col: 1 });
    expect(evalFormulaToBool('ISBLANK(B1)', c)).toBe(true);  // B1 absent → blank
    expect(evalFormulaToBool('ISBLANK(A1)', c)).toBe(false); // A1 present
  });
  it('EXACT', () => {
    expect(ev('EXACT("abc","abc")')).toBe(true);
    expect(ev('EXACT("abc","abC")')).toBe(false);
  });
});

describe('evalFormulaToBool — conditional aggregates', () => {
  it('COUNTIF over a range', () => {
    // A1:A3 = 90, 50, 95 ; count of >=90 is 2
    const cells = [numCell(1, 1, 90), numCell(2, 1, 50), numCell(3, 1, 95)];
    const c = ctx({ cells, row: 1, col: 1 });
    expect(evalFormulaToBool('COUNTIF(A1:A3,">=90")=2', c)).toBe(true);
  });
});

describe('evalFormulaToBool — DATE', () => {
  it('builds an Excel serial', () => {
    expect(ev('DATE(2024,1,1)=45292')).toBe(true);
    expect(ev('DATE(2024,1,15)=45306')).toBe(true);
  });

  it('maps the 1900 leap-bug boundary to serials 1/59/61 (never the phantom 60)', () => {
    // Now that DATE delegates to the shared excel-date module, the 1900 Lotus
    // leap-year-bug compensation is honoured: 1900-01-01 → 1, 1900-02-28 → 59,
    // and 1900-03-01 → 61. Serial 60 (the non-existent 1900-02-29) is skipped.
    expect(ev('DATE(1900,1,1)=1')).toBe(true);
    expect(ev('DATE(1900,2,28)=59')).toBe(true);
    expect(ev('DATE(1900,3,1)=61')).toBe(true);
  });
});

describe('evalFormulaToBool — YEAR / MONTH / DAY', () => {
  it('reads a modern serial (45292 → 2024-01-01)', () => {
    expect(ev('YEAR(45292)=2024')).toBe(true);
    expect(ev('MONTH(45292)=1')).toBe(true);
    expect(ev('DAY(45292)=1')).toBe(true);
  });

  it('reads serial 1 as 1900-01-01 (1900 leap-bug compensation)', () => {
    // Before routing through the shared module the formula engine had no
    // leap-bug correction and read serial 1 as 1899-12-31 (YEAR → 1899). It
    // must now return the Excel value 1900-01-01.
    expect(ev('YEAR(1)=1900')).toBe(true);
    expect(ev('MONTH(1)=1')).toBe(true);
    expect(ev('DAY(1)=1')).toBe(true);
  });

  it('reads the leap-bug boundary serials 59 and 61', () => {
    // Serial 59 = 1900-02-28 (last day before the phantom leap day); serial 61
    // = 1900-03-01 (first day after it).
    expect(ev('MONTH(59)=2')).toBe(true);
    expect(ev('DAY(59)=28')).toBe(true);
    expect(ev('MONTH(61)=3')).toBe(true);
    expect(ev('DAY(61)=1')).toBe(true);
  });
});

describe('evalFormulaToBool — WEEKDAY', () => {
  it('returns Sun=1..Sat=7 by default (2024-01-01 is a Monday → 2)', () => {
    expect(ev('WEEKDAY(45292)=2')).toBe(true);
    expect(ev('WEEKDAY(45292,1)=2')).toBe(true);
  });

  it('honours return-type 2 (Mon=1..Sun=7) and 3 (Mon=0..Sun=6)', () => {
    // Monday: type 2 → 1, type 3 → 0.
    expect(ev('WEEKDAY(45292,2)=1')).toBe(true);
    expect(ev('WEEKDAY(45292,3)=0')).toBe(true);
  });

  it('reads the correct weekday for serial 1 after leap-bug compensation', () => {
    // 1900-01-01 is a Monday → default return-type 1 gives 2.
    expect(ev('WEEKDAY(1)=2')).toBe(true);
  });
});

describe('evalFormulaToBool — error literals', () => {
  it('an error literal makes the rule not apply instead of reading as 0', () => {
    const c = ctx({ cells: [numCell(1, 2, 46180)] });
    expect(evalFormulaToBool('MONTH(#REF!)<>MONTH(B1)', c)).toBe(false);
    expect(evalFormulaToBool('#N/A=1', c)).toBe(false);
    expect(evalFormulaToBool('MONTH(B1)=6', c)).toBe(true);
  });
});
