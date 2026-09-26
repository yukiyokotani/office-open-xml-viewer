import { describe, expect, it } from 'vitest';
import { decodeCfOperand } from './cf-operand.js';
import type { Cell, CellValue } from './types.js';

function cell(row: number, col: number, value: CellValue): Cell {
  return { row, col, value, styleIndex: 0 };
}

// B1 = 3, C1 = "ab", D1 = TRUE, E1 = #N/A; F1 is blank.
const cellIndex = new Map([
  cell(1, 2, { type: 'number', number: 3 }),
  cell(1, 3, { type: 'text', text: 'ab' }),
  cell(1, 4, { type: 'bool', bool: true }),
  cell(1, 5, { type: 'error', error: '#N/A' }),
].map((c) => [`${c.row}:${c.col}`, c]));
const at = (row = 1, col = 1, anchorRow = 1) => ({ row, col, anchorRow, anchorCol: 1, cellIndex });
const decode = (f: string, ctx = at()) => decodeCfOperand(f, ctx);

describe('decodeCfOperand — literals and cached cell values only', () => {
  it('decodes plain literals', () => {
    expect(decode('10')).toBe(10);
    expect(decode('-5')).toBe(-5);
    expect(decode('+.5')).toBe(0.5);
    expect(decode('1.5E+2')).toBe(150);
    expect(decode('"a""b"')).toBe('a"b');
    expect(decode('""')).toBe('');
  });

  it('reads a single cell reference as the typed cached value', () => {
    expect(decode('$B$1')).toBe(3);
    expect(decode('C1')).toBe('ab');
    expect(decode('d1')).toBe(true);
    expect(decode('F1')).toBeNull();
    expect(decode('E1')).toBeUndefined();
  });

  it('anchors relative references at the rule range top-left', () => {
    // From A2 (one row down), `B1` is B2 (blank) and `B$1` stays B1.
    expect(decode('B1', at(2, 1))).toBeNull();
    expect(decode('B$1', at(2, 1))).toBe(3);
    // From B1, `A1` is B1; `$A1` stays in column A.
    expect(decode('A1', at(1, 2))).toBe(3);
    expect(decode('$A1', at(1, 2))).toBeNull();
    // Shifted off the sheet: not decoded.
    expect(decode('A1', at(1, 1, 2))).toBeUndefined();
  });

  it.each([
    '0+10', '-B1', '1/0', '1E400', 'IF(1,B1,0)', 'OR(0,C1)', 'TRUE', 'SomeName', '1>0',
    'B1:B2', 'Sheet2!B1', "'My Sheet'!B1", 'R1C1', 'A0', 'XFE1', 'A1048577', '"open', '1.2.3', '',
  ])('does not decode %s', (f) => {
    expect(decode(f)).toBeUndefined();
  });
});
