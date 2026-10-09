import { describe, it, expect } from 'vitest';
import { evaluateFormula, evalFormulaToBool } from './formula.js';
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

describe('CF formula evaluation boundary', () => {
  it('distinguishes a valid false/zero from unsupported, invalid and error results', () => {
    expect(evaluateFormula('FALSE', ctx())).toEqual({ kind: 'value', value: false });
    expect(evaluateFormula('0', ctx())).toEqual({ kind: 'value', value: 0 });
    expect(evaluateFormula('NOT(UNKNOWN())', ctx())).toEqual({ kind: 'unsupported' });
    expect(evaluateFormula('1+', ctx())).toEqual({ kind: 'invalid' });
    expect(evaluateFormula('#REF!', ctx())).toEqual({ kind: 'error' });
  });

  it('retains decimal/exponent literals, escaped strings and nested defined names', () => {
    const c = ctx();
    c.definedNames.set('Limit', { name: 'Limit', formula: 'Base+1' });
    c.definedNames.set('Base', { name: 'Base', formula: '2' });
    expect(ev('Limit=3', c)).toBe(true);
    expect(ev('.5+1e-2=0.51')).toBe(true);
    expect(ev('1.E2=100')).toBe(true);
    expect(ev('"a""b"="a""b"')).toBe(true);
    expect(ev('=TRUE()')).toBe(true);
    expect(ev('=FALSE()')).toBe(false);
    c.definedNames.set('Base', { name: 'Base', formula: 'UNKNOWN()' });
    expect(ev('Limit=1', c)).toBe(false);
  });

  it('rejects excessive nesting without aborting rendering, including name expansion', () => {
    const nested = (body: string, count: number) => '('.repeat(count) + body + ')'.repeat(count);
    const c = ctx();
    expect(ev(nested('1', 30), c)).toBe(true);
    expect(evaluateFormula(nested('1', 1000), c)).toEqual({ kind: 'unsupported' });
    expect(ev('-'.repeat(1000) + '1', c)).toBe(false);
    c.definedNames.set('Inner', { name: 'Inner', formula: nested('1', 40) });
    expect(ev(nested('Inner', 40), c)).toBe(false);
    expect(ev('Inner', c)).toBe(true);
    // A long left-associative operator chain is not nesting: evaluating it
    // must not recurse once per operator.
    expect(ev('1' + '+1'.repeat(100_000) + '=100001', c)).toBe(true);
  });

  it('reads cell-like spellings outside the grid or before "(" as names, not cells', () => {
    const c = ctx({ cells: [numCell(1, 1, 5)] });
    c.definedNames.set('Limit1', { name: 'Limit1', formula: '10' });
    // LIMIT is past column XFD, so Limit1 is the name, not a blank cell.
    expect(ev('A1<Limit1', c)).toBe(true);
    expect(ev('A1>Limit1', c)).toBe(false);
    expect(ev('ISBLANK(XFD1048576)', c)).toBe(true);
    for (const f of ['ISBLANK(XFE1)', 'ISBLANK(A1048577)', 'ISBLANK(A0)']) {
      expect(evaluateFormula(f, c), f).toEqual({ kind: 'unsupported' });
    }
    // LOG is a column in the grid; followed by "(" it is a function name.
    expect(evaluateFormula('LOG10(100)', c)).toEqual({ kind: 'value', value: 2 });
  });

  it('resolves defined names ASCII case-insensitively, later case variants shadowing', () => {
    const c = ctx();
    c.definedNames.set('Limit', { name: 'Limit', formula: '3' });
    expect(ev('LIMIT=3', c)).toBe(true);
    expect(ev('limit=3', c)).toBe(true);
    // A case variant is the same name, so it shadows like an exact duplicate.
    const shadowed = ctx();
    shadowed.definedNames.set('Rate', { name: 'Rate', formula: '1' });
    shadowed.definedNames.set('RATE', { name: 'RATE', formula: '2' });
    expect(ev('Rate=2', shadowed)).toBe(true);
    // Only ASCII letters fold: KELVIN SIGN lowercases to "k" under Unicode
    // rules, but it is not the name "k" here.
    const kelvinSign = String.fromCodePoint(0x212a);
    const kelvin = ctx();
    kelvin.definedNames.set(kelvinSign, { name: kelvinSign, formula: '1' });
    expect(evaluateFormula('k', kelvin)).toEqual({ kind: 'unsupported' });
  });

  it('bounds retained expansion of branching defined names', () => {
    const c = ctx();
    c.definedNames.set('Layer_0', { name: 'Layer_0', formula: '1' });
    for (let level = 1; level <= 3; level++) {
      const name = `Layer_${level}`;
      c.definedNames.set(name, { name, formula: `SUM(${Array(20).fill(`Layer_${level - 1}`).join(',')})` });
    }
    expect(ev('Layer_2=400', c)).toBe(true);
    // A tiny workbook name graph can otherwise retain exponentially many
    // copies of its parsed bodies before evaluating a single CF cell.
    expect(evaluateFormula('Layer_3', c)).toEqual({ kind: 'unsupported' });
  });
});

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

  it('AVERAGEIF averages matched numeric cells, ignoring matched blanks', () => {
    // A1:A4 = 1, 1, 1, 2 ; B1 = 4, B2 blank, B3 = 0, B4 = 100 ; C1:C4 blank
    const cells = [
      numCell(1, 1, 1), numCell(2, 1, 1), numCell(3, 1, 1), numCell(4, 1, 2),
      numCell(1, 2, 4), numCell(3, 2, 0), numCell(4, 2, 100),
    ];
    const c = ctx({ cells });
    // The matched 0 counts; the matched blank and the unmatched 100 do not.
    expect(evaluateFormula('AVERAGEIF(A1:A4,1,B1:B4)', c)).toEqual({ kind: 'value', value: 2 });
    expect(evalFormulaToBool('AVERAGEIF(A1:A4,1,B1:B4)=2', c)).toBe(true);
    // Every matched cell is blank: no qualifying value is #DIV/0!, not 0.
    expect(evaluateFormula('AVERAGEIF(A1:A4,1,C1:C4)', c)).toEqual({ kind: 'error' });
    expect(evalFormulaToBool('AVERAGEIF(A1:A4,1,C1:C4)=0', c)).toBe(false);
    // Preserved library policy, not Excel evidence: indexes past a shorter
    // average_range are not actual blank cells, so they keep counting.
    expect(evaluateFormula('AVERAGEIF(A1:A4,1,B1:B1)', c)).toEqual({ kind: 'value', value: 4 / 3 });
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

function errCell(row: number, col: number, error: string): Cell {
  return { row, col, value: { type: 'error', error }, styleIndex: 0 };
}

describe('CF formula error values and lazy evaluation', () => {
  it('produces and propagates error values instead of blank or zero placeholders', () => {
    // A1=5, B1 blank, A2=#N/A (cached), A3=7.
    const c = ctx({ cells: [numCell(1, 1, 5), errCell(2, 1, '#N/A'), numCell(3, 1, 7)] });
    expect(evaluateFormula('A1/B1', c)).toEqual({ kind: 'error' });
    expect(ev('A1/B1<0.5', c)).toBe(false);
    expect(ev('MOD(A1,B1)=0', c)).toBe(false);
    expect(ev('A2=0', c)).toBe(false);
    expect(ev('A3>AVERAGE(A1:A3)', c)).toBe(false);
    expect(ev('A3>A1', c)).toBe(true);
  });

  it('evaluates only the selected IF branch and catches errors only in IFERROR/IS*', () => {
    const c = ctx({ cells: [numCell(1, 1, 5), errCell(2, 1, '#N/A'), errCell(3, 1, '#DIV/0!')] });
    expect(ev('IF(B1=0,TRUE,A1/B1>1)', c)).toBe(true);
    expect(ev('IF(A1>0,TRUE,#N/A)', c)).toBe(true);
    expect(ev('IF(A2,TRUE,TRUE)', c)).toBe(false);
    expect(evaluateFormula('IFS(A1>0,TRUE,TRUE,#N/A)', c)).toEqual({ kind: 'error' });
    expect(ev('ISNA(IFS(A1<0,TRUE))', c)).toBe(true);
    expect(ev('IFERROR(A1/B1,-1)=-1', c)).toBe(true);
    expect(ev('IFERROR(B1,"x")="x"', c)).toBe(false); // blank is not an error
    expect(ev('IFERROR(UNKNOWN(),TRUE)', c)).toBe(false); // unsupported is not caught
    expect(ev('ISERROR(B1)', c)).toBe(false);
    expect(ev('ISNA(A2)', c)).toBe(true);
    expect(ev('ISERR(A2)', c)).toBe(false);
    expect(ev('ISERR(A3)', c)).toBe(true);
    expect(ev('ISBLANK("")', c)).toBe(false);
    // Compound errors remain values that type/count functions can inspect.
    expect(ev('NOT(ISNUMBER(1/0+1))', c)).toBe(true);
    expect(ev('COUNT(1/0+1)=0', c)).toBe(true);
    expect(ev('COUNTA(1/0+1)=1', c)).toBe(true);
    expect(ev('ISBLANK(IFERROR(B1,TRUE))', c)).toBe(false);
    expect(ev('ISTEXT(IFERROR(B1,TRUE))', c)).toBe(true);
  });

  it('checks support for the whole formula before selecting an IF branch', () => {
    const c = ctx();
    c.definedNames.set('Hidden', { name: 'Hidden', formula: 'UNKNOWN()' });
    for (const formula of ['IF(TRUE,TRUE,UNKNOWN())', 'IF(TRUE,TRUE,Missing)', 'IF(TRUE,TRUE,Hidden)']) {
      expect(evaluateFormula(formula, c), formula).toEqual({ kind: 'unsupported' });
    }
    // A supported Excel error value can still be ignored by a lazy IF.
    expect(ev('IF(TRUE,TRUE,#NAME?)', c)).toBe(true);
  });

  it('enforces supported-function arity and reads ROW/COLUMN of a reference argument', () => {
    expect(evaluateFormula('NOT()', ctx())).toEqual({ kind: 'invalid' });
    expect(evaluateFormula('IF(TRUE)', ctx())).toEqual({ kind: 'invalid' });
    expect(evaluateFormula('ROUND(1.5)', ctx())).toEqual({ kind: 'invalid' });
    // CONCAT is a later Excel function with a 253-text-argument limit.
    const concat = (count: number) => `CONCAT(${Array(count).fill('""').join(',')})`;
    expect(evaluateFormula(concat(253), ctx())).toEqual({ kind: 'value', value: '' });
    expect(evaluateFormula(concat(254), ctx())).toEqual({ kind: 'invalid' });
    const c = ctx({ row: 5, col: 3 });
    expect(ev('ROW($A$1)=1', c)).toBe(true);
    expect(ev('COLUMN(B1)=4', c)).toBe(true); // relative B1 shifts by +2 columns
    expect(ev('MOD(ROW()-ROW($A$2),2)=1', c)).toBe(true);
  });

  it('rejects ranges larger than the evaluation cap instead of truncating them', () => {
    const c = ctx({ cells: [numCell(4500, 1, 1)] });
    expect(evaluateFormula('COUNTIF($A$1:$A$5000,1)', c)).toEqual({ kind: 'unsupported' });
  });
});

describe('scalar coercion and comparison types', () => {
  const textCell = (row: number, col: number, text: string): Cell =>
    ({ row, col, value: { type: 'text', text }, styleIndex: 0 });

  it('converts only a whole numeric text operand; other text is #VALUE!', () => {
    expect(evaluateFormula('" 1.5e1 "+1', ctx())).toEqual({ kind: 'value', value: 16 });
    expect(evaluateFormula('"-.5"*2', ctx())).toEqual({ kind: 'value', value: -1 });
    for (const f of ['"abc"+1', '"1x"+1', '""+1', '-"abc"', 'ROUND("1x",0)']) {
      expect(evaluateFormula(f, ctx()), f).toEqual({ kind: 'error' });
    }
    // A text cell must not read as 0 and make the rule match.
    const c = ctx({ cells: [textCell(1, 1, 'pending')] });
    expect(evalFormulaToBool('A1+0=0', c)).toBe(false);
  });

  it('keeps the arithmetic #VALUE! rule out of comparison operands', () => {
    // Comparison coercion is unchanged by the strict arithmetic conversion; a
    // digit-led text cell must still equal the identical text.
    const c = ctx({ cells: [textCell(1, 1, '2024 Q1')] });
    expect(evaluateFormula('A1="2024 Q1"', c)).toEqual({ kind: 'value', value: true });
  });
});


describe('existing ROUND numeric semantics', () => {
  it('ROUND rounds the decimal digits of x half away from zero (§18.17.7.278)', () => {
    const round = (f: string) => evaluateFormula(f, ctx());
    // The standard's own examples.
    expect(round('ROUND(2.15,1)')).toEqual({ kind: 'value', value: 2.2 });
    expect(round('ROUND(2.149,1)')).toEqual({ kind: 'value', value: 2.1 });
    expect(round('ROUND(-1.475,2)')).toEqual({ kind: 'value', value: -1.48 });
    expect(round('ROUND(21.5,-1)')).toEqual({ kind: 'value', value: 20 });
    // A negative tie rounds away from zero (Math.round gives -2).
    expect(round('ROUND(-2.5,0)')).toEqual({ kind: 'value', value: -3 });
    // A written 5 is a tie even when its binary product falls below it.
    expect(round('ROUND(1.005,2)')).toEqual({ kind: 'value', value: 1.01 });
    // Exponent-form magnitudes and out-of-range digit counts stay finite.
    expect(round('ROUND(1.5E-7,7)')).toEqual({ kind: 'value', value: 2e-7 });
    expect(round('ROUND(1.5,400)')).toEqual({ kind: 'value', value: 1.5 });
    expect(round('ROUND(123,-400)')).toEqual({ kind: 'value', value: 0 });
  });

  it('ROUND reports unrepresentable numbers and fractional digit counts', () => {
    const round = (f: string) => evaluateFormula(f, ctx());
    // §18.17.3 #NUM! range error: the rounded maximum double overflows.
    expect(round('ROUND(1.7976931348623157E308,-308)')).toEqual({ kind: 'error' });
    // An overflowed or NaN operand is not a number Excel can hold.
    expect(round('ROUND(1E308*10,0)')).toEqual({ kind: 'error' });
    expect(round('ROUND(1.5,1E308*10-1E308*10)')).toEqual({ kind: 'error' });
    // §18.17.7.278 does not say how a fractional number-digits is applied.
    expect(round('ROUND(1.25,1.5)')).toEqual({ kind: 'unsupported' });
  });
});

describe('typed numeric failures in existing error predicates', () => {
  it('supports error predicates and IFERROR around numeric coercion', () => {
    expect(evaluateFormula('ISERROR("1x"+1)', ctx())).toEqual({ kind: 'value', value: true });
    expect(evaluateFormula('ISNA("1x"+1)', ctx())).toEqual({ kind: 'value', value: false });
    expect(evaluateFormula('IFERROR(ROUND(1E308*10,0),7)', ctx())).toEqual({ kind: 'value', value: 7 });
  });
});

describe('decimal ROUND conversion stability', () => {
  it('preserves an operand when the requested precision discards no written digit', () => {
    for (const [n, places] of [[9.320345665328205, 18], [4.1432000626809895, 25]]) {
      expect(evaluateFormula(`ROUND(${n},${places})`, ctx())).toEqual({ kind: 'value', value: n });
    }
  });
  it('carries decimal digits without an intermediate floating-point shift', () => {
    expect(evaluateFormula('ROUND(9.995,2)', ctx())).toEqual({ kind: 'value', value: 10 });
    expect(evaluateFormula('ROUND(5.7118551805615425,15)', ctx())).toEqual({ kind: 'value', value: 5.711855180561543 });
    expect(evaluateFormula('ROUND(23.545329668559134,14)', ctx())).toEqual({ kind: 'value', value: 23.54532966855913 });
    expect(evaluateFormula('ROUND(-0.0005,3)', ctx())).toEqual({ kind: 'value', value: -0.001 });
    // The finite double's shortest spelling is 5e-324, a decimal tie.
    expect(evaluateFormula('ROUND(4.9406564584124654E-324,323)', ctx())).toEqual({ kind: 'value', value: 1e-323 });
  });
});
