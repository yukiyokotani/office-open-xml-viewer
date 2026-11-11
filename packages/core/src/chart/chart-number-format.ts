// Chart value number-formatting. Pure string/number logic (no canvas),
// extracted from the chart renderer so it can be unit-tested and reused.
// Implements the subset of ECMA-376 §18.8.30 number-format codes that chart
// axis ticks and data labels need: section syntax (positive;negative;zero),
// literal escapes, thousands separators, decimals, percent, and Excel serial
// dates.

import { excelSerialToUtcDate, isExcelDisplayDateSerial } from '../excel-date';
import {
  formatExcelDateTime,
  isDateFormatSection,
  splitFormatSections,
  textSectionIndex,
} from '../excel-number-format';
import { roundDecimalHalfUp } from '../text/round-decimal';

const localizedShortDateFormatters = new Map<string, Intl.DateTimeFormat>();

/** Excel built-in number format 14 is locale-sensitive rather than the
 * invariant `m/d/yy` pattern stored in a chart cache. */
export function formatLocalizedExcelShortDate(
  serial: number,
  date1904 = false,
  locale = typeof navigator === 'undefined' ? undefined : navigator.language,
): string {
  if (!isExcelDisplayDateSerial(serial, date1904)) return '#';
  const cacheKey = locale ?? '';
  let formatter = localizedShortDateFormatters.get(cacheKey);
  if (!formatter) {
    formatter = new Intl.DateTimeFormat(locale, {
      year: 'numeric', month: 'numeric', day: 'numeric', timeZone: 'UTC',
    });
    localizedShortDateFormatters.set(cacheKey, formatter);
  }
  return formatter.format(excelSerialToUtcDate(serial, date1904));
}

/**
 * The texts {@link formatLocalizedExcelShortDate} paints for `locale`, for
 * font-source demand (library policy). `complete` is true when they are the
 * whole output set: a Gregorian calendar with Latin (ASCII) digits paints, for
 * Excel's date range (years 1900–9999), only ASCII digits and the pattern's
 * literal separators, which do not depend on the date. Any other calendar
 * (`roc`, `japanese`, `buddhist`, …) or numbering system can paint era names,
 * signs and digits no finite sample proves, so `complete` is false; the texts
 * are then a sample over every year and month of the range's eras, enough to
 * prove which scripts the output uses (CJK era names, native digits), while
 * scalar demand must be `'all'`.
 */
export function localizedExcelShortDateTexts(
  locale = typeof navigator === 'undefined' ? undefined : navigator.language,
): { readonly texts: readonly string[]; readonly complete: boolean } {
  const formatter = new Intl.DateTimeFormat(locale, {
    year: 'numeric', month: 'numeric', day: 'numeric', timeZone: 'UTC',
  });
  const resolved = formatter.resolvedOptions();
  const sample = (years: Iterable<number>, months: readonly number[]) => {
    const texts: string[] = [];
    for (const year of years) for (const month of months) texts.push(formatter.format(Date.UTC(year, month, 28)));
    return texts;
  };
  const decades = (from: number, to: number, step: number) =>
    Array.from({ length: Math.floor((to - from) / step) + 1 }, (_, index) => from + index * step);
  if (resolved.calendar !== 'gregory' || resolved.numberingSystem !== 'latn') {
    // Every year of 1899–2101 (any era boundary of a modern calendar), then
    // every decade to 9999, each in January and December, and serial 1.
    return Object.freeze({
      texts: Object.freeze([
        formatter.format(Date.UTC(1900, 0, 1)),
        ...sample(decades(1899, 2101, 1), [0, 11]),
        ...sample(decades(2110, 9990, 10), [0]),
      ]),
      complete: false,
    });
  }
  const texts = ['0123456789'];
  for (const date of [Date.UTC(1900, 0, 1), Date.UTC(2000, 11, 31), Date.UTC(9999, 11, 31)]) {
    for (const part of formatter.formatToParts(date)) {
      if (part.type === 'literal') texts.push(part.value);
      else if ((part.type !== 'year' && part.type !== 'month' && part.type !== 'day') || !/^[0-9]+$/.test(part.value)) {
        return Object.freeze({ texts: Object.freeze([...texts, part.value]), complete: false });
      }
    }
  }
  return Object.freeze({ texts: Object.freeze(texts), complete: true });
}

/** Excel's default `formatCode="General"` for charts: raw numbers with no
 *  "k"/"M" abbreviation, trailing decimal zeros trimmed. */
export function formatChartVal(v: number): string {
  // Matches Excel's default `<c:valAx><c:numFmt formatCode="General">` which
  // shows raw numbers — no "k"/"M" abbreviation.
  if (Number.isInteger(v)) return String(v);
  // Cap at 6 decimals and trim trailing zeros (0.50 → "0.5"). Round half-UP on
  // the decimal value like Excel (matters only at a `.xxxxxx5` boundary; equal
  // to `toFixed(6)` otherwise), so this path never disagrees with the explicit
  // number-format sections below.
  return roundDecimalHalfUp(v, 6).replace(/\.?0+$/, '');
}

/**
 * Format a chart value with an Excel number-format code. Honors ECMA-376
 * §18.8.30 section syntax (positive;negative;zero;text), common literal
 * escapes (`"..."`, `\x`, `_x` → space), and numeric patterns built from
 * `#`, `0`, `.`, `,`. Unknown tokens are emitted verbatim so currency
 * symbols like `¥` or `$` keep working even when the workbook stored them
 * unquoted. Returns the default `formatChartVal` output when `code` is null
 * or an empty section tells the caller to hide the value.
 */
export function formatChartValWithCode(
  v: number,
  code: string | null | undefined,
  /** Chart date system (`<c:date1904>`, §21.2.2.38). `true` resolves serial
   *  date codes against the 1904 epoch. Defaults to false (1900 system). */
  date1904 = false,
  /** Per numeric section of `code`, the divisor applied once the section is
   *  chosen on the undivided value (an OpenDocument display factor, ODF 1.3
   *  Part 3 §19.346), exactly as worksheet cells apply it. */
  divisors?: readonly number[] | null,
): string {
  // Absent `code`, or the reserved "General" keyword (ECMA-376 §18.8.30), both
  // mean the General number format. LibreOffice charts emit
  // `<c:numFmt formatCode="General">`; tokenizing it as a literal pattern would
  // render the word "General" instead of the value (issue #358).
  if (!code || isGeneralSection(code)) return formatChartVal(v / sectionDivisor(divisors, 0));
  // Section selection per §18.8.30: positive;negative;zero, with the text
  // section left out (it never formats a number). When the negative section
  // is omitted, a negative number is formatted with the positive section and
  // a leading minus.
  const sections = splitFormatSections(code);
  const textIndex = textSectionIndex(sections);
  const numeric = sections.filter((_, i) => i !== textIndex);
  if (numeric.length === 0) return formatChartVal(v);
  const index = v > 0 ? 0 : v < 0 ? (numeric.length > 1 ? 1 : 0) : (numeric.length > 2 ? 2 : 0);
  const section = numeric[index];
  if (section === '') return '';
  const explicitNegative = v < 0 && numeric.length >= 2;
  v /= sectionDivisor(divisors, index);
  // A section that is the General keyword (`General;General`) formats its
  // value as General: the magnitude in an explicit negative section.
  if (isGeneralSection(section)) return formatChartVal(explicitNegative ? -v : v);
  // Date/time is a property of the selected section: the value is a serial
  // date (scatter / time-series axes), rendered by the same formatter as
  // worksheet cells.
  if (isDateFormatSection(section)) {
    return formatExcelDateTime(explicitNegative ? -v : v, section, date1904);
  }
  const needsLeadingMinus = v < 0 && !explicitNegative;
  return (needsLeadingMinus ? '-' : '') + applyChartNumberSection(Math.abs(v), section);
}

/**
 * Format a category-axis tick label. Categories arrive as the raw cached
 * strings the workbook stored (`<c:cat><c:strCache>` text, or the numeric
 * text of a `<c:numCache>` on a date/number axis). ECMA-376 §21.2.2.71
 * (`c:numFmt`) lets the category axis carry a number-format code: Excel
 * applies it to the tick labels just as it does on the value axis. When the
 * axis has such a code AND the raw category parses to a finite number, format
 * it (which routes serial dates through `formatExcelDateTime` automatically); a
 * missing code, `"General"`, or a non-numeric category string falls through
 * to the raw text unchanged — no new interpretation is invented.
 */
export function formatCategoryLabel(
  raw: string,
  code: string | null | undefined,
  /** Chart date system (`<c:date1904>`, §21.2.2.38). Threaded to the date
   *  formatter for serial-date category labels. Defaults to false. */
  date1904 = false,
  /** Section divisors of `code` ({@link formatChartValWithCode}). */
  divisors?: readonly number[] | null,
): string {
  if (!code) return raw;
  // Only numeric-looking categories are formatted. `Number('')` and
  // `Number('  ')` are 0 (falsely finite), so reject blank / whitespace text
  // explicitly; keep genuine string categories (e.g. "Q1", "North") verbatim.
  if (raw.trim() === '') return raw;
  const num = Number(raw);
  if (!Number.isFinite(num)) return raw;
  return formatChartValWithCode(num, code, date1904, divisors);
}

/** The divisor of numeric section `index`: a finite positive factor, else 1. */
function sectionDivisor(divisors: readonly number[] | null | undefined, index: number): number {
  const divisor = divisors?.[index];
  return divisor != null && Number.isFinite(divisor) && divisor > 0 ? divisor : 1;
}

/** Whether a code or section is the reserved General keyword (§18.8.30). */
function isGeneralSection(section: string): boolean {
  return section.trim().toLowerCase() === 'general';
}

function applyChartNumberSection(abs: number, section: string): string {
  // Tokenize the section, separating numeric-pattern runs (`#`, `0`, `.`,
  // `,`, `?`) from literal runs so percent / decimal handling runs once.
  type Tok = { kind: 'lit' | 'num'; text: string };
  const toks: Tok[] = [];
  let i = 0;
  let pushedNum = false;
  let percent = false;
  while (i < section.length) {
    const c = section[i];
    if (c === '"') {
      i++;
      let s = '';
      while (i < section.length && section[i] !== '"') { s += section[i]; i++; }
      if (i < section.length) i++;
      toks.push({ kind: 'lit', text: s });
      continue;
    }
    if (c === '\\' && i + 1 < section.length) {
      toks.push({ kind: 'lit', text: section[i + 1] });
      i += 2;
      continue;
    }
    if (c === '_' && i + 1 < section.length) {
      // `_x` pads a width of x — render as a single space, matching Excel
      // alignment padding without caring about exact glyph metrics.
      toks.push({ kind: 'lit', text: ' ' });
      i += 2;
      continue;
    }
    if (c === '*' && i + 1 < section.length) {
      // `*x` fills the remaining column width with x; we can't know the
      // column width at this layer so we drop it.
      i += 2;
      continue;
    }
    if (c === '[') {
      i++;
      while (i < section.length && section[i] !== ']') i++;
      if (i < section.length) i++;
      continue;
    }
    if (c === '%') { percent = true; toks.push({ kind: 'lit', text: '%' }); i++; continue; }
    if (c === '#' || c === '0' || c === '.' || c === ',' || c === '?') {
      let run = '';
      while (
        i < section.length &&
        (section[i] === '#' || section[i] === '0' || section[i] === '.' ||
         section[i] === ',' || section[i] === '?')
      ) { run += section[i]; i++; }
      toks.push({ kind: 'num', text: run });
      pushedNum = true;
      continue;
    }
    // Everything else (currency symbols like ¥, $, parens, spaces) is literal.
    toks.push({ kind: 'lit', text: c });
    i++;
  }
  if (!pushedNum) {
    // No numeric pattern at all — section is purely literal (e.g. `"N/A"`).
    return toks.map(t => t.text).join('');
  }
  const value = percent ? abs * 100 : abs;
  // Merge numeric tokens into one pattern — Excel treats `#,##0.00` as a
  // single pattern even when flanked by literals. We keep the literal tokens
  // where they are and replace the first num token with the formatted number,
  // dropping subsequent num tokens (they're all part of the same pattern).
  let pattern = '';
  for (const t of toks) if (t.kind === 'num') pattern += t.text;
  const formatted = formatNumericPattern(value, pattern);
  let seenNum = false;
  return toks.map(t => {
    if (t.kind === 'lit') return t.text;
    if (seenNum) return '';
    seenNum = true;
    return formatted;
  }).join('');
}

function formatNumericPattern(value: number, pattern: string): string {
  // Detect thousands separator (a `,` between digit placeholders) and the
  // number of decimal places (digit chars after `.`).
  const dotIdx = pattern.indexOf('.');
  const intPart = dotIdx >= 0 ? pattern.slice(0, dotIdx) : pattern;
  const fracPart = dotIdx >= 0 ? pattern.slice(dotIdx + 1) : '';
  const thousands = /,/.test(intPart);
  const fracDigits = (fracPart.match(/[#0?]/g) ?? []).length;
  // Minimum integer digits = count of `0` in integer part.
  const minIntDigits = (intPart.replace(/,/g, '').match(/0/g) ?? []).length;
  // Office rounds a chart label's decimals half-UP on the decimal value
  // (2.675 → "2.68"); `toFixed` would round the binary double down.
  const rounded = roundDecimalHalfUp(value, fracDigits);
  const [ints, fracs = ''] = rounded.split('.');
  const paddedInts = ints.padStart(minIntDigits, '0');
  const withSeparators = thousands ? paddedInts.replace(/\B(?=(\d{3})+(?!\d))/g, ',') : paddedInts;
  if (fracDigits === 0) return withSeparators;
  return `${withSeparators}.${fracs.padEnd(fracDigits, '0')}`;
}
