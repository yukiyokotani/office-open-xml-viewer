import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { formatCellValue } from './number-format.js';
import type { Cell, Styles } from './types.js';

const FMT_ID = 164; // first free custom id

function styles(formatCode: string): Styles {
  return {
    fonts: [],
    fills: [],
    borders: [],
    cellXfs: [{ fontId: 0, fillId: 0, borderId: 0, numFmtId: FMT_ID, alignH: null, alignV: null, wrapText: false }],
    numFmts: [{ numFmtId: FMT_ID, formatCode }],
    dxfs: [],
  };
}

function builtinStyles(numFmtId: number): Styles {
  return {
    fonts: [],
    fills: [],
    borders: [],
    cellXfs: [{
      fontId: 0,
      fillId: 0,
      borderId: 0,
      numFmtId,
      alignH: null,
      alignV: null,
      wrapText: false,
    }],
    numFmts: [],
    dxfs: [],
  };
}

function numCell(n: number): Cell {
  return { row: 1, col: 1, value: { type: 'number', number: n }, styleIndex: 0 };
}

/** Format a number with a custom format code, as Excel would render it. */
const fmt = (n: number, code: string) => formatCellValue(numCell(n), styles(code));

describe('number formats — integers & decimals', () => {
  it('plain integer', () => {
    expect(fmt(5, '0')).toBe('5');
    expect(fmt(5.6, '0')).toBe('6'); // rounds
  });
  it('fixed decimals', () => {
    expect(fmt(5, '0.00')).toBe('5.00');
    expect(fmt(5.125, '0.00')).toBe('5.13');
  });
  it('thousands separator', () => {
    expect(fmt(1234567, '#,##0')).toBe('1,234,567');
    expect(fmt(1234.5, '#,##0.0')).toBe('1,234.5');
  });
});

describe('number formats — percent', () => {
  it('scales by 100', () => {
    expect(fmt(0.5, '0%')).toBe('50%');
    expect(fmt(0.1234, '0.0%')).toBe('12.3%');
  });
});

describe('number formats — sign sections (§18.8.30)', () => {
  it('positive / negative / zero selection', () => {
    // positive;negative;zero
    expect(fmt(5, '0;(0);"-"')).toBe('5');
    expect(fmt(-5, '0;(0);"-"')).toBe('(5)');
    expect(fmt(0, '0;(0);"-"')).toBe('-');
  });
  it('negative falls back to positive section when absent', () => {
    expect(fmt(-5, '0.0')).toBe('-5.0');
  });
});

describe('number formats — literals', () => {
  it('keeps quoted literal text around the number', () => {
    expect(fmt(3, '0" units"')).toBe('3 units');
  });
});

describe('General format code (§18.8.30 / LibreOffice custom numFmt)', () => {
  // LibreOffice Calc writes a custom numFmt (id ≥ 164) with formatCode="General"
  // for every saved workbook. "General" is the reserved General-format keyword,
  // so a cell must render its value — not the literal text "General" (issue #358).
  it('renders the number for a custom numFmt whose code is "General"', () => {
    expect(fmt(10, 'General')).toBe('10');
    expect(fmt(42, 'General')).toBe('42');
    expect(fmt(3.14, 'General')).toBe('3.14');
  });
  it('is case-insensitive and tolerates surrounding whitespace', () => {
    expect(fmt(10, 'general')).toBe('10');
    expect(fmt(10, 'GENERAL')).toBe('10');
    expect(fmt(10, ' General ')).toBe('10');
  });
});

describe('General format — 11 significant digit rounding (XL2)', () => {
  // Excel's General format is not raw float round-trip: the display engine
  // rounds to 11 significant digits (15-digit internal precision minus the
  // ~4 digits Excel reserves for display robustness), so binary floating
  // point noise from arithmetic (e.g. 0.1 + 0.2) never surfaces to the user.
  // This table pins the rounding + trailing-zero-trim + exponential-switch
  // rules against `formatGeneralNumber` (see number-format.ts for the exact
  // exponent thresholds and their rationale).
  it('rounds binary floating point noise away', () => {
    expect(fmt(0.1 + 0.2, 'General')).toBe('0.3');
    expect(fmt(-(0.1 + 0.2), 'General')).toBe('-0.3');
  });
  it('rounds a repeating decimal to 11 significant digits', () => {
    expect(fmt(1 / 3, 'General')).toBe('0.33333333333');
  });
  it('leaves an 11-digit integer untouched', () => {
    expect(fmt(12345678901, 'General')).toBe('12345678901');
  });
  it('switches a 12-digit integer to Excel exponential notation', () => {
    // Mantissa capped at 6 significant digits (5 decimal places) once the
    // General format has already committed to scientific notation.
    expect(fmt(123456789012, 'General')).toBe('1.23457E+11');
  });
  it('rounds a many-decimal value to 11 significant digits', () => {
    expect(fmt(1234.5678901234, 'General')).toBe('1234.5678901');
  });
  it('applies the same rounding to negative numbers, sign excluded from digit count', () => {
    expect(fmt(-0.30000000000000004, 'General')).toBe('-0.3');
  });
  it('renders negative zero as "0"', () => {
    expect(fmt(-0, 'General')).toBe('0');
  });
  it('switches a very small number to exponential once fixed-point would bury it past 11 significant digits', () => {
    expect(fmt(0.000000001234567890123, 'General')).toBe('1.23457E-09');
  });
  it('keeps a small-but-not-tiny decimal in fixed-point form', () => {
    expect(fmt(0.00001, 'General')).toBe('0.00001');
  });
  it('switches at the documented exponent boundary (1e-6 range)', () => {
    expect(fmt(0.000001, 'General')).toBe('1E-06');
  });
  it('trims trailing zeros from an exact decimal', () => {
    expect(fmt(100, 'General')).toBe('100');
    expect(fmt(0.5, 'General')).toBe('0.5');
  });
  it('handles the rounding-carry boundary: an 11-digit value that rounds up to a 12-digit exponent', () => {
    // 99999999999.6 rounds to 11 significant digits as 100000000000, whose
    // decimal exponent (11) crosses the fixed→exponential threshold. The
    // exponent must be derived from the *rounded* form (this is exactly the
    // non-obvious case the code comment cites), so it renders "1E+11" rather
    // than a spurious "99999999999" or "100000000000".
    expect(fmt(99999999999.6, 'General')).toBe('1E+11');
  });
});

describe('non-numeric cells', () => {
  it('passes text through when no 4th section', () => {
    const cell: Cell = { row: 1, col: 1, value: { type: 'text', text: 'hello' }, styleIndex: 0 };
    expect(formatCellValue(cell, styles('0.00'))).toBe('hello');
  });
});

describe('date formats (Excel serial; 45292 = 2024-01-01)', () => {
  it('ISO and slash dates', () => {
    expect(fmt(45306, 'yyyy-mm-dd')).toBe('2024-01-15');
    expect(fmt(45306, 'm/d/yy')).toBe('1/15/24');
    expect(fmt(45306, 'mm/dd/yyyy')).toBe('01/15/2024');
  });
  it('day and month parts', () => {
    expect(fmt(45292, 'yyyy')).toBe('2024');
    expect(fmt(45292, 'd')).toBe('1');
    expect(fmt(45292, 'dd')).toBe('01');
  });

  it('localizes built-in short-date format 14 to the application UI language', () => {
    vi.stubGlobal('navigator', { language: 'ja-JP' });
    try {
      expect(formatCellValue(numCell(45306), builtinStyles(14))).toBe('2024/1/15');
    } finally {
      vi.unstubAllGlobals();
    }
  });

  it('keeps built-in short-date format 14 in month/day/year order for en-US', () => {
    vi.stubGlobal('navigator', { language: 'en-US' });
    try {
      expect(formatCellValue(numCell(45306), builtinStyles(14))).toBe('1/15/2024');
    } finally {
      vi.unstubAllGlobals();
    }
  });

  it('reuses the localized short-date formatter for cells in the same UI locale', () => {
    vi.stubGlobal('navigator', { language: 'ko-KR' });
    const NativeDateTimeFormat = Intl.DateTimeFormat;
    const formatter = vi.spyOn(Intl, 'DateTimeFormat').mockImplementation(
      function dateTimeFormat(locales, options) {
        return new NativeDateTimeFormat(locales, options);
      },
    );
    try {
      formatCellValue(numCell(45306), builtinStyles(14));
      formatCellValue(numCell(45307), builtinStyles(14));
      expect(formatter).toHaveBeenCalledTimes(1);
    } finally {
      formatter.mockRestore();
      vi.unstubAllGlobals();
    }
  });
});

describe('date formats — 1900 Lotus leap-year-bug compat (§18.17.4.1)', () => {
  // The cell formatter now delegates serial → date to the shared core
  // `excelSerialToUtcDate`, which shifts serials < 60 by +1 day to reproduce
  // Excel's phantom 1900-02-29. This changes output ONLY for serials ≤ 59.
  it('serial 1 renders 1900-01-01', () => {
    expect(fmt(1, 'yyyy-mm-dd')).toBe('1900-01-01');
  });
  it('serial 59 renders 1900-02-28 (was off-by-one before the compat fix)', () => {
    expect(fmt(59, 'yyyy-mm-dd')).toBe('1900-02-28');
  });
  it('serial 61 renders 1900-03-01 (day after the phantom leap day)', () => {
    expect(fmt(61, 'yyyy-mm-dd')).toBe('1900-03-01');
  });
  it('modern serials (≥ 60) are unchanged: 45292 → 2024-01-01', () => {
    expect(fmt(45292, 'yyyy-mm-dd')).toBe('2024-01-01');
  });
});

describe('date formats — 1904 date system (§18.2.28 / §18.17.4.1)', () => {
  // A 1904 (Mac-authored) workbook stores serials 1462 days lower than a 1900
  // workbook for the same calendar date. `formatCellValue`'s 4th arg carries
  // `<workbookPr date1904>` and shifts the epoch accordingly.
  const fmt1904 = (n: number, code: string) =>
    formatCellValue(numCell(n), styles(code), null, true);

  it('renders the same calendar date from the 1904-system serial (43830 → 2024-01-01)', () => {
    // 1900-system serial 45292 and 1904-system serial 43830 are both 2024-01-01.
    expect(fmt1904(43830, 'yyyy-mm-dd')).toBe('2024-01-01');
    // Without the date1904 flag the same serial reads 1462 days early.
    expect(fmt(43830, 'yyyy-mm-dd')).toBe('2019-12-31');
  });

  it('serial 0 is the 1904 base date 1904-01-01', () => {
    expect(fmt1904(0, 'yyyy-mm-dd')).toBe('1904-01-01');
  });

  it('serial 1 is 1904-01-02 (no 1900 leap-year bug in the 1904 system)', () => {
    expect(fmt1904(1, 'yyyy-mm-dd')).toBe('1904-01-02');
  });
});

describe('formula cells render their cached value, never a recalculation', () => {
  // Library policy (number-format.ts): cell formulas are never calculated. A
  // TODAY()/NOW() cell shows the cached `<v>` saved with the file, so it stays
  // consistent with every cell derived from it. The clock is frozen far from
  // the cached dates so a recalculation could not pass by coincidence.
  beforeEach(() => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date('2031-07-20T12:00:00.000Z'));
  });

  afterEach(() => {
    vi.useRealTimers();
  });

  function formulaCell(formula: string, value: Cell['value']): Cell {
    return { row: 1, col: 1, value, styleIndex: 0, formula };
  }

  it('renders the cached TODAY() serial in the workbook date system', () => {
    // 1900-system serial 45306 = 2024-01-15; 1904-system serial 43830 = 2024-01-01.
    const cell1900 = formulaCell('TODAY()', { type: 'number', number: 45306 });
    expect(formatCellValue(cell1900, styles('yyyy-mm-dd'))).toBe('2024-01-15');
    const cell1904 = formulaCell('TODAY()', { type: 'number', number: 43830 });
    expect(formatCellValue(cell1904, styles('yyyy-mm-dd'), null, true)).toBe('2024-01-01');
  });

  it('renders the cached NOW() serial, including its time fraction', () => {
    const cell = formulaCell('NOW()', { type: 'number', number: 45306.75 });
    expect(formatCellValue(cell, styles('yyyy-mm-dd hh:mm'))).toBe('2024-01-15 18:00');
  });

  it('renders a TODAY() cell without a cached value like any other uncached formula', () => {
    const empty: Cell['value'] = { type: 'empty' };
    const uncachedSum = formatCellValue(formulaCell('SUM(A1:A2)', empty), styles('yyyy-mm-dd'));
    expect(uncachedSum).toBe('');
    expect(formatCellValue(formulaCell('TODAY()', empty), styles('yyyy-mm-dd'))).toBe(uncachedSum);
  });
});
