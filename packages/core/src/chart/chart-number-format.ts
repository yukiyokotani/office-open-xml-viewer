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
  generalSectionAffixes,
  selectFormatSection,
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
  // Section selection per §18.8.30, shared with worksheet cells: explicit
  // conditions (`[>=1000]`) when any section has one, else
  // positive;negative;zero, with the text section left out (it never formats
  // a number). When the negative section is omitted, a negative number is
  // formatted with the positive section and a leading minus.
  const selected = selectFormatSection(v, code);
  if (selected === null) return formatChartVal(v);
  // No condition met and no unconditional section: Excel shows `#`.
  if (selected === 'none') return '#';
  const sections = splitFormatSections(code);
  const textIndex = textSectionIndex(sections);
  const section = sections.filter((_, i) => i !== textIndex)[selected.index];
  if (section === '') return '';
  // The section formats the magnitude: its literals carry the sign.
  const explicitNegative = selected.useMagnitude;
  v /= sectionDivisor(divisors, selected.index);
  // A section whose body is the General keyword (`General;General`,
  // `[Red]General`) formats its value as General: the magnitude in a section
  // that carries the sign itself.
  const general = generalSectionAffixes(selected.section.body);
  if (general) {
    const scaled = general.percent ? v * 100 : v;
    const shown = explicitNegative ? Math.abs(scaled) : scaled;
    return (shown < 0 ? '-' : '') + general.prefix + formatChartVal(Math.abs(shown)) + general.suffix;
  }
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
  // A scientific section (§18.8.30 `E+` / `E-` after the digit pattern) has
  // one `exp` token where its exponent goes; a fraction section (`# ?/?`,
  // `?/8`) one `den` token, the digit run before its `/` being the numerator.
  type Tok = { kind: 'lit' | 'num' | 'exp' | 'den' | 'slash'; text: string };
  const toks: Tok[] = [];
  let i = 0;
  let pushedNum = false;
  let percent = false;
  let exponent: { plus: boolean; digits: number } | null = null;
  let fraction = false;
  const isPlaceholder = (ch: string | undefined): boolean => ch === '0' || ch === '#' || ch === '?';
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
    if ((c === 'E' || c === 'e') && pushedNum && !exponent && !fraction
        && (section[i + 1] === '+' || section[i + 1] === '-') && isPlaceholder(section[i + 2])) {
      let j = i + 2;
      let digits = 0;
      while (isPlaceholder(section[j])) { if (section[j] === '0') digits++; j++; }
      exponent = { plus: section[i + 1] === '+', digits };
      toks.push({ kind: 'exp', text: '' });
      i = j;
      continue;
    }
    if (c === '/' && !fraction && !exponent && toks[toks.length - 1]?.kind === 'num'
        && (isPlaceholder(section[i + 1]) || /[1-9]/.test(section[i + 1] ?? ''))) {
      let j = i + 1;
      while (isPlaceholder(section[j]) || /[0-9]/.test(section[j] ?? '')) j++;
      fraction = true;
      toks.push({ kind: 'slash', text: '/' }, { kind: 'den', text: section.slice(i + 1, j) });
      i = j;
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
  if (fraction) return formatFractionSection(value, toks);
  if (exponent) {
    let pattern = '';
    for (const t of toks) if (t.kind === 'num') pattern += t.text;
    // The exponent is a multiple of the integer placeholders (`##0.0E+0`
    // is engineering notation); a mantissa that rounds up to the next
    // power takes the next exponent.
    const places = Math.max(1, (pattern.split('.')[0].match(/[#0?]/g) ?? []).length);
    const [digits, power] = value === 0 ? ['0', '0'] : value.toExponential().split('e');
    let exp = Math.floor(Number(power) / places) * places;
    let mantissa = formatNumericPattern(Number(`${digits}e${Number(power) - exp}`), pattern);
    if (Number(mantissa.replace(/,/g, '')) >= 10 ** places) {
      exp += places;
      mantissa = formatNumericPattern(Number(`${digits}e${Number(power) - exp}`), pattern);
    }
    const sign = exp < 0 ? '-' : exponent.plus ? '+' : '';
    const exponentText = `E${sign}${String(Math.abs(exp)).padStart(exponent.digits, '0')}`;
    let seen = false;
    return toks.map(t => {
      if (t.kind === 'lit') return t.text;
      if (t.kind === 'exp') return exponentText;
      if (seen) return '';
      seen = true;
      return mantissa;
    }).join('');
  }
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

/** The fraction `numerator / denominator` nearest to `value` (0 ≤ value < 1
 *  or an improper value) with a denominator of at most `maxDenominator`,
 *  by the Stern–Brocot mediants between its floor and ceiling. */
function nearestFraction(value: number, maxDenominator: number): [number, number] {
  const whole = Math.floor(value);
  const rest = value - whole;
  let [ln, ld, hn, hd] = [0, 1, 1, 1];
  let best: [number, number] = rest < 0.5 ? [0, 1] : [1, 1];
  for (;;) {
    const n = ln + hn;
    const d = ld + hd;
    if (d > maxDenominator) break;
    if (Math.abs(n / d - rest) < Math.abs(best[0] / best[1] - rest)) best = [n, d];
    if (n / d === rest) break;
    if (n / d < rest) { ln = n; ld = d; } else { hn = n; hd = d; }
  }
  return [whole * best[1] + best[0], best[1]];
}

/** A fraction section: the digit run before `/` is the numerator, the runs
 *  before that the whole part (none: an improper fraction), and the
 *  denominator is fixed (`?/8`) or has at most as many digits as its
 *  placeholders. A value without a fractional part shows its whole part
 *  alone, and one below 1 its fraction alone. */
function formatFractionSection(
  value: number,
  toks: { kind: 'lit' | 'num' | 'exp' | 'den' | 'slash'; text: string }[],
): string {
  const slash = toks.findIndex(t => t.kind === 'slash');
  const numerator = slash - 1;
  const denominatorSpec = toks[slash + 1].text;
  const wholeToks = toks.slice(0, numerator).filter(t => t.kind === 'num');
  const mixed = wholeToks.length > 0;
  const fixed = /^[0-9]+$/.test(denominatorSpec) && Number(denominatorSpec) > 0 ? Number(denominatorSpec) : null;
  const whole = mixed ? Math.floor(value) : 0;
  const rest = value - whole;
  let [n, d] = fixed !== null
    ? [Math.round(rest * fixed), fixed]
    : nearestFraction(rest, 10 ** Math.min(denominatorSpec.length, 7) - 1);
  let shownWhole = whole;
  if (mixed && n >= d) { shownWhole += 1; n -= d; }
  const wholeText = mixed ? formatNumericPattern(shownWhole, wholeToks.map(t => t.text).join('')) : '';
  const first = toks.findIndex(t => t.kind === 'num');
  const showWhole = mixed && (shownWhole !== 0 || n === 0);
  const showFraction = !mixed || n !== 0;
  return toks.map((t, index) => {
    // Literals between the whole part and the fraction (the space of
    // `# ?/?`) belong to showing both.
    if (mixed && index > first && index < numerator) return showWhole && showFraction && t.kind === 'lit' ? t.text : '';
    if (index === first && mixed) return showWhole ? (wholeText || '0') : '';
    if (index === numerator) return showFraction ? String(n) : '';
    if (t.kind === 'slash') return showFraction ? '/' : '';
    if (t.kind === 'den') return showFraction ? String(d) : '';
    return t.kind === 'lit' ? t.text : '';
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
