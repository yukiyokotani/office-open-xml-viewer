import { describe, it, expect } from 'vitest';
import {
  formatChartVal,
  formatChartValWithCode,
  formatCategoryLabel,
  formatLocalizedExcelShortDate,
} from './chart-number-format.js';

describe('formatChartVal (General)', () => {
  it('shows integers raw with no abbreviation', () => {
    expect(formatChartVal(7507)).toBe('7507');
    expect(formatChartVal(0)).toBe('0');
  });
  it('trims trailing decimal zeros', () => {
    expect(formatChartVal(0.5)).toBe('0.5');
    expect(formatChartVal(1.25)).toBe('1.25');
  });
});

describe('formatChartValWithCode — thousands separator', () => {
  it('adds commas with #,##0', () => {
    expect(formatChartValWithCode(1234567, '#,##0')).toBe('1,234,567');
    expect(formatChartValWithCode(7507, '#,##0')).toBe('7,507');
  });
  it('keeps decimals alongside the separator', () => {
    expect(formatChartValWithCode(1234.5, '#,##0.0')).toBe('1,234.5');
  });
});

describe('formatChartValWithCode — sample-2 slide-16 data labels (§18.8.30)', () => {
  // The bar-chart series number format is `#,##0_);[Red]\(#,##0\)`. Before the
  // parser wired `valFormatCode` through, these labels rendered as "7507" with
  // no comma; PowerPoint shows "7,507". The positive section's `_)` reserves a
  // glyph-width of ')' so the value aligns with parenthesised negatives — we
  // render it as a trailing space, matching Excel.
  const CODE = '#,##0_);[Red]\\(#,##0\\)';
  it('positive value gets a comma (the reported bug)', () => {
    expect(formatChartValWithCode(7507, CODE)).toBe('7,507 ');
    expect(formatChartValWithCode(2208, CODE)).toBe('2,208 ');
    expect(formatChartValWithCode(6117, CODE)).toBe('6,117 ');
  });
  it('negative value uses the parenthesised section', () => {
    expect(formatChartValWithCode(-7507, CODE)).toBe('(7,507)');
  });
});

describe('formatChartValWithCode — percent & null', () => {
  it('scales by 100 with %', () => {
    expect(formatChartValWithCode(0.5, '0%')).toBe('50%');
  });
  it('falls back to General when code is null/General', () => {
    expect(formatChartValWithCode(7507, null)).toBe('7507');
    // "General" is the reserved General-format keyword (ECMA-376 §18.8.30), not a
    // literal pattern. LibreOffice charts emit <c:numFmt formatCode="General">,
    // which must render the value — not the literal text "General" (issue #358).
    expect(formatChartValWithCode(7507, 'General')).toBe('7507');
    expect(formatChartValWithCode(7507, 'general')).toBe('7507');
    expect(formatChartValWithCode(0.5, ' General ')).toBe('0.5');
  });
});

describe('chart date1904 (c:date1904 §21.2.2.38 / §18.17.4.1)', () => {
  it('shares fractional-second carry with worksheet cells in either date system', () => {
    expect(formatChartValWithCode(59.96 / 86400, 'mm:ss.0')).toBe('01:00.0');
    expect(formatCategoryLabel(String(59.96 / 86400), 'mm:ss.0', true)).toBe('01:00.0');
  });

  it('localizes worksheet built-in short-date format 14', () => {
    expect(formatLocalizedExcelShortDate(45_658, false, 'ja-JP')).toBe('2025/1/1');
    expect(formatLocalizedExcelShortDate(45_658, false, 'en-US')).toBe('1/1/2025');
  });

  it('formatChartValWithCode shifts a date serial to the 1904 epoch when date1904=true', () => {
    // 1900-system serial 44927 and 1904-system serial 43465 are both
    // 2023-01-01 (offset 1462 days).
    expect(formatChartValWithCode(44927, 'yyyy-mm-dd')).toBe('2023-01-01');
    expect(formatChartValWithCode(43465, 'yyyy-mm-dd', true)).toBe('2023-01-01');
    // Same serial without the flag reads 1462 days early.
    expect(formatChartValWithCode(43465, 'yyyy-mm-dd')).toBe('2018-12-31');
  });
  it('formatCategoryLabel threads date1904 to the date formatter', () => {
    expect(formatCategoryLabel('43465', 'yyyy-mm-dd', true)).toBe('2023-01-01');
    expect(formatCategoryLabel('43465', 'yyyy-mm-dd')).toBe('2018-12-31');
  });
  it('date1904 does not affect non-date (numeric) codes', () => {
    expect(formatChartValWithCode(1234567, '#,##0', true)).toBe('1,234,567');
  });
});

describe('chart time fields (§18.17.4.2 time fraction)', () => {
  it('rounds the stored fraction to the millisecond instead of truncating', () => {
    // 8:00 is stored a hair under 28 800 000 ms.
    expect(formatChartValWithCode(45292.33333333333, 'hh:mm')).toBe('08:00');
    expect(formatChartValWithCode(0.33333333333333331, 'h:mm:ss')).toBe('8:00:00');
  });
  it('keeps minutes after an escaped time separator', () => {
    expect(formatChartValWithCode(45292.34375, 'h\\:mm')).toBe('8:15');
  });
  it('reads a negative elapsed duration from one absolute duration', () => {
    expect(formatChartValWithCode(-1.5104166666666667, '[h]:mm', true)).toBe('-36:15');
  });
  it('treats an era-only code as a date', () => {
    expect(formatCategoryLabel('45292', '[$-411]ggge')).toBe('令和6');
  });
  it('carries a rounded midnight into the next day', () => {
    // 1 ms short of 2024-01-02 00:00 rounds up to it.
    const serial = 45293 - 0.4 / 86_400_000;
    expect(formatChartValWithCode(serial, 'yyyy-mm-dd hh:mm')).toBe('2024-01-02 00:00');
  });
});

describe('formatCategoryLabel — category-axis numFmt (§21.2.2.71)', () => {
  it('formats a numeric serial category through a date code', () => {
    // 44927 = 2023-01-01 in the 1900 date system.
    expect(formatCategoryLabel('44927', 'm/d/yyyy')).toBe('1/1/2023');
  });
  it('formats a numeric category through a plain number code', () => {
    expect(formatCategoryLabel('1234567', '#,##0')).toBe('1,234,567');
  });
  it('leaves a string category verbatim even when a code is present', () => {
    expect(formatCategoryLabel('Q1', 'm/d/yyyy')).toBe('Q1');
    expect(formatCategoryLabel('North', '#,##0')).toBe('North');
  });
  it('leaves a numeric category verbatim when no code is present', () => {
    expect(formatCategoryLabel('44927', null)).toBe('44927');
    expect(formatCategoryLabel('44927', undefined)).toBe('44927');
    expect(formatCategoryLabel('44927', '')).toBe('44927');
  });
  it('does not format blank or whitespace-only categories (Number("") is a false 0)', () => {
    expect(formatCategoryLabel('', 'm/d/yyyy')).toBe('');
    expect(formatCategoryLabel('   ', '#,##0')).toBe('   ');
  });
  it('a General code leaves the raw numeric text unchanged', () => {
    // "General" routes through formatChartValWithCode → formatChartVal, which
    // shows the raw number (44927), matching the pre-format behavior.
    expect(formatCategoryLabel('44927', 'General')).toBe('44927');
  });

  it('renders abbreviated and full month names for date-axis format tokens', () => {
    // 45658 = 2025-01-01 in the 1900 date system. The `mmm` token displays
    // "Jan-25" and must not be folded to the numeric form used by `mm`.
    expect(formatCategoryLabel('45658', 'mmm\\-yy')).toBe('Jan-25');
    expect(formatCategoryLabel('45658', 'mmmm yyyy')).toBe('January 2025');
    expect(formatCategoryLabel('45658', 'mm/dd/yy')).toBe('01/01/25');
  });
});


describe('shared date formatting range (#1710)', () => {
  it('marks invalid calendar and clock serials without throwing in charts or format 14', () => {
    expect(formatLocalizedExcelShortDate(12345678901, false, 'en-US')).toBe('#');
    expect(formatChartValWithCode(12345678901, 'd-mmm-yy')).toBe('#');
    expect(formatChartValWithCode(-1, 'h:mm:ss')).toBe('#');
    expect(formatChartValWithCode(2957004, 'yyyy-mm-dd', true)).toBe('#');
  });
  it('preserves elapsed durations but rejects invalid calendar fields mixed with them', () => {
    expect(formatChartValWithCode(-1.5104166666666667, '[h]:mm')).toBe('-36:15');
    expect(formatChartValWithCode(12345678901, '[h]:mm')).toBe('296296293624:00');
    expect(formatChartValWithCode(12345678901, '[h] yyyy')).toBe('#');
    expect(formatChartValWithCode(12345678901, '[h] "yyyy"')).toBe('296296293624 yyyy');
    expect(formatChartValWithCode(Infinity, '[h]:mm')).toBe('#');
  });
  it('keeps the last millisecond on the final day when binary rounding reaches midnight', () => {
    expect(formatChartValWithCode(2958466 - 0.4 / 86400000, 'yyyy-mm-dd h:mm:ss')).toBe('9999-12-31 23:59:59');
  });
});

describe('section divisors (OpenDocument display factor)', () => {
  it('divides any factor once the section is chosen on the undivided value', () => {
    expect(formatChartValWithCode(100, '0', false, [10])).toBe('10');
    expect(formatChartValWithCode(5000, '0.0', false, [2000])).toBe('2.5');
    expect(formatChartValWithCode(-50, '0;(0)', false, [1, 10])).toBe('(5)');
    expect(formatCategoryLabel('100', '0', false, [10])).toBe('10');
  });
  it('divides General codes and General sections', () => {
    expect(formatChartValWithCode(2500, 'General', false, [1000])).toBe('2.5');
    expect(formatChartValWithCode(250, 'General;General', false, [100, 10])).toBe('2.5');
    expect(formatChartValWithCode(-250, 'General;General', false, [100, 10])).toBe('25');
    expect(formatChartValWithCode(-250, 'General;General')).toBe('250');
  });
  it('chooses explicitly conditioned sections as worksheet cells do', () => {
    const code = '[>=1000]0.0" k";0';
    expect(formatChartValWithCode(999, code, false, [1000, 1])).toBe('999');
    expect(formatChartValWithCode(2500, code, false, [1000, 1])).toBe('2.5 k');
    expect(formatChartValWithCode(-50, code, false, [1000, 1])).toBe('-50');
    expect(formatChartValWithCode(-5, '[<0]\\-0.0;0', false)).toBe('-5.0');
    expect(formatChartValWithCode(5, '[<0]0;[>10]0', false)).toBe('#');
  });
  it('reads an unclosed bracket run in linear time', () => {
    const started = performance.now();
    expect(typeof formatChartValWithCode(1, '['.repeat(200_000))).toBe('string');
    expect(performance.now() - started).toBeLessThan(5_000);
  });
  it('formats decorated General sections as General', () => {
    expect(formatChartValWithCode(250, '[Red]General', false, [100])).toBe('2.5');
    expect(formatChartValWithCode(-250, '0;[Red]General', false, [1, 10])).toBe('25');
    expect(formatChartValWithCode(7, '[Blue]General')).toBe('7');
  });
  it('ignores absent, non-positive or non-finite divisors', () => {
    expect(formatChartValWithCode(42, '0', false, null)).toBe('42');
    expect(formatChartValWithCode(42, '0', false, [0])).toBe('42');
    expect(formatChartValWithCode(42, '0', false, [Number.NaN])).toBe('42');
  });
});
