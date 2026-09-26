import { describe, expect, it } from 'vitest';
import { evalCfExact } from './cf-exact-formula.js';
import type { Cell, CellValue } from './types.js';

function cell(row: number, col: number, value: CellValue): Cell {
  return { row, col, value, styleIndex: 0 };
}

// B1 = 3, C1 = "ab", D1 = TRUE, E1 = #N/A, F1 = 1e200; G1 is blank.
const cells = [
  cell(1, 2, { type: 'number', number: 3 }),
  cell(1, 3, { type: 'text', text: 'ab' }),
  cell(1, 4, { type: 'bool', bool: true }),
  cell(1, 5, { type: 'error', error: '#N/A' }),
  cell(1, 6, { type: 'number', number: 1e200 }),
];
const cellIndex = new Map(cells.map((c) => [`${c.row}:${c.col}`, c]));
const at = (row = 1, col = 1) => ({ row, col, anchorRow: 1, anchorCol: 1, cellIndex });
const ev = (f: string, ctx = at()) => evalCfExact(f, ctx);

describe('evalCfExact — the exactly evaluated grammar', () => {
  it('literals, typed references, arithmetic and concatenation', () => {
    expect(ev('0+10')).toBe(10);
    expect(ev('-2*3+1')).toBe(-5);
    expect(ev('-(1+2)/4')).toBe(-0.75);
    expect(ev('1.5E+2')).toBe(150);
    expect(ev('"a""b"')).toBe('a"b');
    expect(ev('"x"&C1')).toBe('xab');
    expect(ev('$B$1*2')).toBe(6);
    expect(ev('C1')).toBe('ab');
    expect(ev('D1')).toBe(true);
    expect(ev('G1')).toBeNull();
  });

  it('shifts relative references by the offset from the anchor', () => {
    // From A2 (one row down), `B1` is B2 (blank) and `B$1` stays B1.
    expect(ev('B1', at(2, 1))).toBeNull();
    expect(ev('B$1', at(2, 1))).toBe(3);
    // From B1, `A1` is B1.
    expect(ev('A1', at(1, 2))).toBe(3);
    expect(ev('$A1', at(1, 2))).toBeNull(); // column pinned to A: A1 is blank
  });
});

describe('evalCfExact — unevaluable instead of guessed', () => {
  it.each([
    // functions and names (round-3 review: OR/IF/overflow)
    'IF(OR(0,C1),0,10)', 'IF(1,G1,0)+1', 'AND(1,1)', 'NOT(0)', 'ABS(-1)', 'INT(1.5)', 'TRUE', 'TRUE()', 'SomeName',
    // Excel errors
    'F1*F1', '1E400', '1/0', 'E1', '(F1*F1)-F1*F1',
    // implicit conversions
    'G1+1', 'D1+1', 'C1+1', '"3"+1', '1&"x"', '"x"&G1', '-C1',
    // operators and references outside the grammar
    '1>0', '"a"="a"', '2^2', '50%', '+1', 'A1:B2', 'Sheet2!A1', 'A0', 'XFE1', 'A1048577',
    // malformed
    '', '"open', '1.2.3', '1E', '(1', '1)', '1 2', 'LOG10(1)',
  ])('%s', (f) => {
    expect(ev(f)).toBeUndefined();
  });

  it('a relative reference shifted off the sheet is #REF!', () => {
    expect(ev('A1', at(1, 1))).toBeNull();
    expect(evalCfExact('A1', { ...at(1, 1), anchorRow: 2 })).toBeUndefined();
  });
});
