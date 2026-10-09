import { describe, expect, it } from 'vitest';
import { decodeCfOperand, decodeCfOperandResult } from './cf-operand.js';
import type { Cell } from './types.js';

const ctx = (cells: Cell[], row = 1, col = 1) => ({
  row, col, anchorRow: 1, anchorCol: 1,
  cellIndex: new Map(cells.map((c) => [`${c.row}:${c.col}`, c] as const)),
});

describe('tagged CF operand decode (#1547)', () => {
  const cells: Cell[] = [
    { row: 1, col: 2, value: { type: 'error', error: '#N/A' }, styleIndex: 0 },
    { row: 1, col: 3, value: { type: 'number', number: 7 }, styleIndex: 0 },
  ];

  it('keeps values, unsupported syntax, out-of-grid shifts and cached errors distinct', () => {
    expect(decodeCfOperandResult('4', ctx(cells))).toEqual({ kind: 'value', value: 4 });
    expect(decodeCfOperandResult('"a"', ctx(cells))).toEqual({ kind: 'value', value: 'a' });
    expect(decodeCfOperandResult('$C$1', ctx(cells))).toEqual({ kind: 'value', value: 7 });
    expect(decodeCfOperandResult('D1', ctx(cells))).toEqual({ kind: 'value', value: null });
    expect(decodeCfOperandResult('$B$1', ctx(cells))).toEqual({ kind: 'error' });
    for (const formula of ['A1+1', 'SUM(A1:A2)', 'B1:B2', 'Sheet2!A1', 'XFE1', 'A0', 'Name']) {
      expect(decodeCfOperandResult(formula, ctx(cells)), formula)
        .toEqual({ kind: 'unsupported', reason: 'syntax' });
    }
    expect(decodeCfOperandResult('A1048576', ctx(cells, 2, 1)))
      .toEqual({ kind: 'unsupported', reason: 'outOfGrid' });
  });

  it('keeps the compatibility wrapper contract', () => {
    expect(decodeCfOperand('$C$1', ctx(cells))).toBe(7);
    expect(decodeCfOperand('D1', ctx(cells))).toBeNull();
    expect(decodeCfOperand('$B$1', ctx(cells))).toBeUndefined();
    expect(decodeCfOperand('A1+1', ctx(cells))).toBeUndefined();
  });
});
