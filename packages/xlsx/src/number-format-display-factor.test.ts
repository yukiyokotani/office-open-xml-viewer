import { describe, expect, it } from 'vitest';
import { formatCellValue, formatCellValueWithColor } from './number-format.js';
import type { Cell, Styles } from './types.js';

// An OpenDocument `number:display-factor` (ODF 1.3 Part 3 §19.346) divides
// the displayed value by the factor. The model carries it per numeric
// section of the code; the section is still chosen on the undivided value.

function styles(formatCode: string, sectionDivisors: number[]): Styles {
  return {
    fonts: [],
    fills: [],
    borders: [],
    cellXfs: [{ fontId: 0, fillId: 0, borderId: 0, numFmtId: 164, alignH: null, alignV: null, wrapText: false }],
    numFmts: [{ numFmtId: 164, formatCode, sectionDivisors }],
    dxfs: [],
  };
}

const number = (value: number): Cell => ({ col: 1, row: 1, value: { type: 'number', number: value }, styleIndex: 0 });

describe('section divisors', () => {
  it('divides by any factor, not only powers of 1000', () => {
    expect(formatCellValue(number(100), styles('0', [10]))).toBe('10');
    expect(formatCellValue(number(5000), styles('0.0', [2000]))).toBe('2.5');
    expect(formatCellValue(number(1234567), styles('#,##0', [1000]))).toBe('1,235');
  });

  it('shows nothing for a zero whose zero section is empty', () => {
    // A formula result shown as nothing: the numeric zero with an empty
    // zero section.
    expect(formatCellValue(number(0), styles('0;-0;;@', []))).toBe('');
  });

  it("formats General between a section's literals", () => {
    expect(formatCellValue(number(1234.5), styles('"$"General', []))).toBe('$1234.5');
    expect(formatCellValue(number(-1234.5), styles('"$"General', []))).toBe('-$1234.5');
    expect(formatCellValue(number(-2), styles('General" kg";(General)', []))).toBe('(2)');
    expect(formatCellValue(number(0.25), styles('General%', []))).toBe('25%');
  });

  it('carries a fraction that rounds to a whole', () => {
    // 2.999 with one-digit denominators is 3, not 2 8/9 or 26/9.
    expect(formatCellValue(number(2.999), styles('# ?/?', [])).trim()).toBe('3');
    expect(formatCellValue(number(2.999), styles('?/?', []))).toBe('3/1');
    expect(formatCellValue(number(0.0625), styles('# ?/##?', [])).trim()).toBe('1/16');
  });

  it('divides a General code', () => {
    expect(formatCellValue(number(250), styles('General', [100]))).toBe('2.5');
  });

  it('formats a General section as General, divided by its own factor', () => {
    // A mapped data style composes `General;General` with per-section factors.
    const code = styles('General;General', [100, 10]);
    expect(formatCellValue(number(250), code)).toBe('2.5');
    // The negative section formats the magnitude.
    expect(formatCellValue(number(-250), code)).toBe('25');
    expect(formatCellValue(number(-250), styles('[Red]General;[Blue]General', []))).toBe('250');
  });

  it('colours a General section, bright green by its indexed colour', () => {
    // An OpenDocument data style's `fo:color` beside a number without
    // decimal places; LibreOffice's green is 00FF00, `[Color4]`.
    expect(formatCellValueWithColor(number(250), styles('[Red]General', []))).toMatchObject({ text: '250', color: '#FF0000' });
    expect(formatCellValueWithColor(number(250), styles('[Color4]0', []))).toMatchObject({ text: '250', color: '#00FF00' });
  });

  it('selects the section on the undivided value, then divides by its own factor', () => {
    // [>=1000] picks the first section for 1500 even though 1500/1000 < 1000.
    const code = styles('[>=1000]0.0" k";0', [1000, 1]);
    expect(formatCellValue(number(1500), code)).toBe('1.5 k');
    expect(formatCellValue(number(999), code)).toBe('999');
    // Negative section: magnitude of the divided value.
    expect(formatCellValue(number(-50), styles('0;(0)', [1, 10]))).toBe('(5)');
  });

  it('passes a conditional format divisor through', () => {
    const cell = number(300);
    expect(formatCellValue(cell, styles('0', []), { numFmtId: 165, formatCode: '0', sectionDivisors: [3] })).toBe('100');
  });

  it('ignores non-positive or non-finite divisors', () => {
    expect(formatCellValue(number(42), styles('0', [0]))).toBe('42');
    expect(formatCellValue(number(42), styles('0', [Number.POSITIVE_INFINITY]))).toBe('42');
  });
});
