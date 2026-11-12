import { describe, expect, it } from 'vitest';
import { utcDateToExcelSerial } from './excel-date.js';
import {
  excelDateTimeSectionNameTexts, formatExcelDateTime, isDateFormatSection,
} from './excel-number-format.js';

// The quarter and week-of-year fields of an OpenDocument date style (ODF 1.3
// Part 3 §16.29.18 / §16.29.17) reach the formatter as the bracket tokens
// `[Q]`, `[QQ]` and `[WW]` / `[WW:d,n]`. The expected texts are what
// LibreOffice 26.8 displays for the same dates and locales.

const format = (date: string, section: string): string =>
  formatExcelDateTime(utcDateToExcelSerial(new Date(`${date}T00:00:00Z`)), section);

describe('calendar-field brackets', () => {
  it('formats the quarter and its name', () => {
    expect(format('2024-03-31', '[Q]')).toBe('Q1');
    expect(format('2024-04-01', '[Q]')).toBe('Q2');
    expect(format('2024-08-15', '[QQ]')).toBe('3rd quarter');
    expect(format('2024-12-30', 'yyyy "/" [QQ]')).toBe('2024 / 4th quarter');
    // A ja-JP section names the quarter in Japanese, as its eras are.
    expect(format('2024-04-01', '[$-411][QQ]')).toBe('第２四半期');
  });

  it('numbers weeks from Sunday with 1 January in week 1 by default (en-US)', () => {
    const weeks = ['2024-01-01', '2024-01-06', '2024-01-07', '2024-12-29', '2021-01-01', '2021-01-03']
      .map(date => format(date, '[WW]'));
    expect(weeks).toEqual(['1', '1', '2', '1', '1', '2']);
  });

  it('numbers weeks by the rule the token carries', () => {
    // de-DE: Monday, at least four days (ISO 8601).
    const iso = ['2024-01-07', '2024-01-08', '2024-12-29', '2024-12-30', '2021-01-01', '2021-01-04', '2023-01-01']
      .map(date => format(date, '[WW:1,4]'));
    expect(iso).toEqual(['1', '2', '52', '1', '53', '1', '52']);
    // ar-SA: Saturday, one day.
    const saturday = ['2024-01-05', '2024-01-06', '2021-01-01', '2021-01-02', '2022-01-01']
      .map(date => format(date, '[WW:6,1]'));
    expect(saturday).toEqual(['1', '2', '1', '2', '1']);
    // ru-RU: Monday, one day.
    expect(['2021-01-01', '2021-01-03', '2021-01-04'].map(date => format(date, '[WW:1,1]')))
      .toEqual(['1', '1', '2']);
  });

  it('makes a section a date section and lists the quarter texts', () => {
    expect(isDateFormatSection('"Q"[Q]')).toBe(true);
    expect(isDateFormatSection('[WW:1,4]')).toBe(true);
    // Any other bracket stays metadata.
    expect(isDateFormatSection('[Red]0.00')).toBe(false);
    expect(isDateFormatSection('[W]0')).toBe(false);
    expect(excelDateTimeSectionNameTexts('[QQ]').join('')).toContain('quarter');
    expect(excelDateTimeSectionNameTexts('[WW]')).toEqual([]);
  });

  it('shows the invalid-date marker for a serial without a calendar date', () => {
    expect(formatExcelDateTime(-1, '[Q]')).toBe('#');
    expect(formatExcelDateTime(-1, '[WW]')).toBe('#');
  });
});
