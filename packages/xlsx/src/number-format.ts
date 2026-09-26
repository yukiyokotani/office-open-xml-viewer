import {
  excelSerialToUtcDate,
  formatLocalizedExcelShortDate,
  roundDecimalHalfUp,
} from '@silurus/ooxml-core';
import type { Cell, CellValue, Styles } from './types.js';

function cellValueText(value: CellValue): string {
  switch (value.type) {
    case 'empty': return '';
    case 'text': return value.text;
    case 'number': return String(value.number);
    case 'bool': return value.bool ? 'TRUE' : 'FALSE';
    case 'error': return value.error;
    // `shared` cells are resolved to `text` (see shared-strings.ts) before any
    // consumer runs, so this is unreachable at runtime — present only to keep
    // the switch exhaustive over CellValue.
    case 'shared': return '';
  }
}

/**
 * A formatted cell value together with any text colour the number-format code
 * asked for (§18.8.30 "Specify colors"). `color` is a `#RRGGBB` hex when the
 * matched section began with a `[Red]`/…/`[ColorN]` token, otherwise absent.
 */
export interface FormattedCell {
  text: string;
  color?: string;
}

/**
 * Backward-compatible string entry point. Returns exactly the display string
 * (colour discarded). Kept as the primary export so existing call sites
 * (find, validation-list, workbook.cellText) are unaffected — the renderer,
 * which needs the colour, calls {@link formatCellValueWithColor} instead.
 */
export function formatCellValue(
  cell: Cell,
  styles: Styles,
  cfNumFmt?: { numFmtId: number; formatCode: string | null } | null,
  date1904 = false,
): string {
  return formatCellValueWithColor(cell, styles, cfNumFmt, date1904).text;
}

export function formatCellValueWithColor(
  cell: Cell,
  styles: Styles,
  cfNumFmt?: { numFmtId: number; formatCode: string | null } | null,
  /** Workbook date system (`<workbookPr date1904>`, §18.2.28). `true` resolves
   *  serial dates against the 1904 epoch (§18.17.4.1). Defaults to false (1900
   *  date system) so callers that don't thread the flag are unaffected. */
  date1904 = false,
): FormattedCell {
  // Resolve the effective format once so both the numeric and text paths
  // honour the same precedence: CF dxf numFmt > style numFmt (§18.8.17).
  const xf = styles.cellXfs[cell.styleIndex ?? 0];
  const styleNumFmtId = xf?.numFmtId ?? 0;
  const styleFmt = styles.numFmts?.find(f => f.numFmtId === styleNumFmtId)?.formatCode ?? null;
  const effectiveFmtId = cfNumFmt?.numFmtId ?? styleNumFmtId;
  const effectiveFmt = cfNumFmt?.formatCode ?? styleFmt;

  // Non-numeric cells still need to honour the 4th format section (text).
  // §18.8.30: format sections are positive;negative;zero;text. An empty text
  // section hides the value (Excel's `;;;` trick used for chart-placeholder
  // cells like D3 in the holiday-budget sample), and `@` substitutes the
  // original text. Cells without a 4-section format pass through unchanged.
  if (cell.value.type !== 'number') {
    const text = cellValueText(cell.value);
    return { text: effectiveFmt ? applyTextSection(text, effectiveFmt) : text };
  }

  // Library policy: cell values are rendered from the cached `<v>` only
  // (ECMA-376 §18.3.1.96). Cell formulas (`<f>`, §18.3.1.40) are never
  // calculated, including volatile functions such as TODAY()/NOW(). Replacing
  // one volatile cell with the current time would leave every cell derived from
  // it at its cached value, so the render would be internally inconsistent; the
  // cached values the producing application saved together are authoritative.
  // A formula cell without a cached value is rendered like any other empty
  // cell, regardless of which function it calls.
  return applyFormat(cell.value.number, effectiveFmtId, effectiveFmt, date1904);
}

/**
 * Apply the text section of an Excel number format to a text value.
 * ECMA-376 §18.8.30 "Include a section for text entry":
 *   - The text section, if present, is the *last* section.
 *   - With four sections, section[3] is unconditionally the text section (so
 *     the `;;;` "hide everything" idiom hides text via an empty 4th section).
 *   - With fewer sections, the format has a text section only if its last
 *     section contains `@`; otherwise "text entered in a cell is not affected
 *     by the format code" and passes through unchanged.
 *   - `@` substitutes the original text; quoted / escaped literals are emitted;
 *     `[...]` metadata and `_`/`*` pad pairs follow the numeric conventions.
 */
function applyTextSection(text: string, formatCode: string): string {
  const sections = splitSections(formatCode);
  let section: string;
  if (sections.length >= 4) {
    section = sections[3];
  } else {
    const last = sections[sections.length - 1];
    // A text section is one that contains an `@` placeholder. Without one, the
    // format has no text section and text is unaffected.
    if (!last.includes('@')) return text;
    section = last;
  }
  if (section === '') return '';
  let out = '';
  let i = 0;
  while (i < section.length) {
    const ch = section[i];
    if (ch === '"') {
      i++;
      while (i < section.length && section[i] !== '"') out += section[i++];
      if (i < section.length) i++;
    } else if (ch === '\\') {
      if (i + 1 < section.length) out += section[i + 1];
      i += 2;
    } else if (ch === '[') {
      while (i < section.length && section[i] !== ']') i++;
      if (i < section.length) i++;
    } else if (ch === '@') {
      out += text;
      i++;
    } else if (ch === '_' || ch === '*') {
      i += 2;
    } else {
      out += ch;
      i++;
    }
  }
  return out;
}

// ────────────────────────────────────────────────────────────────
// Date / time formatting  (ECMA-376 §18.8.30)
// ────────────────────────────────────────────────────────────────

// Built-in numFmtId → format code. ID 14 is handled separately because
// ECMA-376 §18.8.30 permits built-in formats to be interpreted differently
// according to the implementing application's UI language. IDs 15-22 use the
// generic table's patterns here; IDs 27-31 and 50-58 are East-Asian (Japanese)
// locale built-ins that Office ships pre-assigned when authored in ja-JP.
const BUILTIN_DATE_FMT: Record<number, string> = {
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'm/d/yyyy h:mm',
  // Japanese locale built-ins (East-Asian Office). Values mirror what
  // Excel ja-JP writes for these IDs.
  27: '[$-411]ge.m.d',
  28: '[$-411]ggge"年"m"月"d"日"',
  29: '[$-411]ggge"年"m"月"d"日"',
  30: 'm/d/yy',
  31: 'yyyy"年"m"月"d"日"',
  50: '[$-411]ge.m.d',
  51: '[$-411]ggge"年"m"月"d"日"',
  52: 'yyyy"年"m"月"',
  53: 'm"月"d"日"',
  54: '[$-411]ggge"年"m"月"d"日"',
  55: 'yyyy"年"m"月"',
  56: 'm"月"d"日"',
  57: '[$-411]ge.m.d',
  58: '[$-411]ggge"年"m"月"d"日"',
};

const MONTH_NAMES = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
];
const WEEKDAY_NAMES = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
/** Japanese short weekday names (aaa format code, e.g. "水"). */
const JP_WEEKDAY_SHORT = ['日', '月', '火', '水', '木', '金', '土'];
/** Japanese long weekday names (aaaa format code, e.g. "水曜日"). */
const JP_WEEKDAY_LONG = ['日曜日', '月曜日', '火曜日', '水曜日', '木曜日', '金曜日', '土曜日'];

/** Japanese imperial eras, newest-first. First entry whose `start` is
 *  ≤ the target date wins (ECMA-376 §18.8.30 — g/gg/ggg and e/ee codes). */
const JP_ERAS: Array<{ start: Date; abbr: string; short: string; long: string }> = [
  { start: new Date(Date.UTC(2019, 4,  1)), abbr: 'R', short: '令', long: '令和' },
  { start: new Date(Date.UTC(1989, 0,  8)), abbr: 'H', short: '平', long: '平成' },
  { start: new Date(Date.UTC(1926, 11, 25)), abbr: 'S', short: '昭', long: '昭和' },
  { start: new Date(Date.UTC(1912, 6,  30)), abbr: 'T', short: '大', long: '大正' },
  { start: new Date(Date.UTC(1868, 0,  25)), abbr: 'M', short: '明', long: '明治' },
];

function resolveJpEra(date: Date): { abbr: string; short: string; long: string; year: number } {
  for (const era of JP_ERAS) {
    if (date.getTime() >= era.start.getTime()) {
      return {
        abbr: era.abbr,
        short: era.short,
        long: era.long,
        year: date.getUTCFullYear() - era.start.getUTCFullYear() + 1,
      };
    }
  }
  // Pre-Meiji: fall back to Gregorian year, keep Meiji names as a best effort.
  const last = JP_ERAS[JP_ERAS.length - 1];
  return { abbr: last.abbr, short: last.short, long: last.long, year: date.getUTCFullYear() };
}

/**
 * Format an Excel date serial using an ECMA-376 format code.
 * Supports: y/yy/yyy/yyyy, m/mm/mmm/mmmm/mmmmm, d/dd/ddd/dddd,
 *           h/hh, m/mm (minutes when after h), s/ss, AM/PM, A/P,
 *           quoted literals, bracket escapes, _ padding, * fill.
 *
 * `date1904` selects the date system (`<workbookPr date1904>`, §18.2.28). The
 * serial → calendar-date conversion is delegated to the shared core
 * `excelSerialToUtcDate` (§18.17.4.1), which carries the 1900 Lotus
 * leap-year-bug compat and the 1904 epoch. It defaults to false so 1900-system
 * workbooks are unchanged (apart from the serial ≤ 59 leap-bug compat, which is
 * now correct in both systems).
 */
function formatExcelDateCode(serial: number, fmtCode: string, date1904 = false): string {
  const date = excelSerialToUtcDate(serial, date1904);
  const yr = date.getUTCFullYear();
  const mo = date.getUTCMonth() + 1;   // 1-12
  const dy = date.getUTCDate();
  const wd = date.getUTCDay();          // 0=Sun
  const hr = date.getUTCHours();
  const mi = date.getUTCMinutes();
  const sc = date.getUTCSeconds();

  // Take the first section (positive / no-sign section)
  const section = fmtCode.split(';')[0];
  const hasAmPm = /am\/pm|a\/p/i.test(section);
  let era: ReturnType<typeof resolveJpEra> | null = null;
  const getEra = (): ReturnType<typeof resolveJpEra> => era ?? (era = resolveJpEra(date));

  let result = '';
  let i = 0;
  let prevWasHour = false;

  while (i < section.length) {
    const ch = section[i];

    if (ch === '"') {
      // Quoted string literal
      i++;
      while (i < section.length && section[i] !== '"') result += section[i++];
      if (i < section.length) i++;
      prevWasHour = false;

    } else if (ch === '[') {
      // ECMA-376 §18.8.30: `[h]` / `[m]` / `[s]` are elapsed-time tokens that
      // suppress the h < 24 / m < 60 / s < 60 wrap-around and instead render
      // the full duration. Any other bracket content (locale IDs, colours,
      // conditions) is metadata and skipped.
      const end = section.indexOf(']', i);
      const inner = end > i ? section.slice(i + 1, end) : '';
      const elapsed = inner.match(/^([hms])\1*$/i);
      if (elapsed) {
        const kind = elapsed[1].toLowerCase();
        const sign = serial < 0 ? '-' : '';
        const absSec = Math.floor(Math.abs(serial) * 86400);
        let v: number;
        if      (kind === 'h') v = Math.floor(absSec / 3600);
        else if (kind === 'm') v = Math.floor(absSec / 60);
        else                   v = absSec;
        const padded = inner.length >= 2 ? String(v).padStart(inner.length, '0') : String(v);
        result += sign + padded;
        i = end + 1;
        prevWasHour = kind === 'h';
      } else {
        while (i < section.length && section[i] !== ']') i++;
        if (i < section.length) i++;
      }

    } else if (ch === '_') {
      i += 2; // _ followed by a padding character — skip both

    } else if (ch === '*') {
      i += 2; // * followed by fill character — skip both

    } else if (ch === '\\') {
      if (i + 1 < section.length) result += section[i + 1];
      i += 2;
      prevWasHour = false;

    } else if (ch === 'y' || ch === 'Y') {
      let n = 0;
      while (i < section.length && section[i].toLowerCase() === 'y') { n++; i++; }
      result += n <= 2 ? String(yr).slice(-2) : String(yr).padStart(4, '0');
      prevWasHour = false;

    } else if (ch === 'm' || ch === 'M') {
      let n = 0;
      while (i < section.length && section[i].toLowerCase() === 'm') { n++; i++; }
      // Determine month vs minutes:
      //   minutes when immediately after h/hh, OR immediately before :s/:ss
      const rest = section.slice(i).replace(/\[[^\]]*\]/g, '');
      const isMinutes = prevWasHour || /^:s/i.test(rest);
      if (isMinutes) {
        result += n >= 2 ? String(mi).padStart(2, '0') : String(mi);
      } else {
        if      (n === 1) result += String(mo);
        else if (n === 2) result += String(mo).padStart(2, '0');
        else if (n === 3) result += MONTH_NAMES[mo - 1].slice(0, 3);
        else if (n === 4) result += MONTH_NAMES[mo - 1];
        else              result += MONTH_NAMES[mo - 1][0]; // mmmmm = first letter
      }
      prevWasHour = false;

    } else if (ch === 'd' || ch === 'D') {
      let n = 0;
      while (i < section.length && section[i].toLowerCase() === 'd') { n++; i++; }
      if      (n === 1) result += String(dy);
      else if (n === 2) result += String(dy).padStart(2, '0');
      else if (n === 3) result += WEEKDAY_NAMES[wd].slice(0, 3);
      else              result += WEEKDAY_NAMES[wd];
      prevWasHour = false;

    } else if (ch === 'h' || ch === 'H') {
      let n = 0;
      while (i < section.length && section[i].toLowerCase() === 'h') { n++; i++; }
      const h = hasAmPm ? (hr % 12 || 12) : hr;
      result += n >= 2 ? String(h).padStart(2, '0') : String(h);
      prevWasHour = true;

    } else if (ch === 's' || ch === 'S') {
      let n = 0;
      while (i < section.length && section[i].toLowerCase() === 's') { n++; i++; }
      result += n >= 2 ? String(sc).padStart(2, '0') : String(sc);
      prevWasHour = false;

    } else if (ch === 'g' || ch === 'G') {
      // Japanese era name (ECMA-376 §18.8.30 ja locale):
      //   g   → 'R' / 'H' / 'S' / 'T' / 'M'
      //   gg  → '令' / '平' / '昭' / '大' / '明'
      //   ggg → '令和' / '平成' / '昭和' / '大正' / '明治'
      let n = 0;
      while (i < section.length && section[i].toLowerCase() === 'g') { n++; i++; }
      const e = getEra();
      if      (n === 1) result += e.abbr;
      else if (n === 2) result += e.short;
      else              result += e.long;
      prevWasHour = false;

    } else if (ch === 'e' || ch === 'E') {
      // Japanese era year: `e` → unpadded, `ee` → 2-digit zero-padded.
      let n = 0;
      while (i < section.length && section[i].toLowerCase() === 'e') { n++; i++; }
      const y = getEra().year;
      result += n >= 2 ? String(y).padStart(2, '0') : String(y);
      prevWasHour = false;

    } else if (ch === 'r' || ch === 'R') {
      // Some Japanese Excel variants expose `r` / `rr` as era-year aliases.
      let n = 0;
      while (i < section.length && section[i].toLowerCase() === 'r') { n++; i++; }
      const y = getEra().year;
      result += n >= 2 ? String(y).padStart(2, '0') : String(y);
      prevWasHour = false;

    } else if (ch === 'A' || ch === 'a') {
      const upper = section.slice(i).toUpperCase();
      // Japanese weekday format codes (Excel ja locale). `aaaa` = "水曜日",
      // `aaa` = "水". Checked before AM/PM because those are shorter matches
      // and would otherwise swallow the leading 'a'.
      if (upper.startsWith('AAAA')) {
        result += JP_WEEKDAY_LONG[wd]; i += 4;
      } else if (upper.startsWith('AAA')) {
        result += JP_WEEKDAY_SHORT[wd]; i += 3;
      } else if (upper.startsWith('AM/PM')) {
        result += hr < 12 ? 'AM' : 'PM'; i += 5;
      } else if (upper.startsWith('A/P')) {
        result += hr < 12 ? 'A' : 'P'; i += 3;
      } else {
        result += ch; i++;
      }
      prevWasHour = false;

    } else {
      result += ch;
      i++;
      // Separators (:/-. space) don't reset the hour context for m/mm lookahead
      if (ch !== ':' && ch !== '/' && ch !== '-' && ch !== '.' && ch !== ' ') {
        prevWasHour = false;
      }
    }
  }

  return result;
}

/** Returns true if a custom formatCode is a date/time format. */
function isDateFormatCode(code: string): boolean {
  // Elapsed-time brackets `[h]`, `[m]`, `[s]` (ECMA-376 §18.8.30) are themselves
  // time formats, so detect those *before* stripping bracket content below.
  if (/\[[hms]+\]/i.test(code)) return true;
  // Strip quoted literals and bracket content, then look for unambiguous date specifiers.
  // 'y' = year, 'd' = day — both are unambiguous. 'm' alone is ambiguous (month or minutes).
  const stripped = code.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '');
  // y / d are unambiguous date specifiers. `aaa+` is the Japanese-locale
  // weekday code and implies a date format even without y/d (e.g. the
  // bare `aaa` custom format).
  return /[yd]/i.test(stripped) || /a{3,}/i.test(stripped);
}

// Excel's General format does not round-trip the raw IEEE-754 double: the
// display engine rounds to 11 significant digits (of the 15-17 significant
// digits a double can carry), which is what keeps binary floating point
// noise from arithmetic (e.g. 0.1 + 0.2 === 0.30000000000000004) from ever
// reaching the screen. See "Floating-point arithmetic may give inaccurate
// results in Excel" (Microsoft KB78113) for the 15-digit internal precision
// this display rounding sits on top of.
const GENERAL_SIGNIFICANT_DIGITS = 11;
// Once General has committed to scientific notation (see thresholds below),
// Excel caps the mantissa at 6 significant digits (1 integer + 5 decimal),
// e.g. 123456789012 -> "1.23457E+11", independent of the 11-digit budget
// used for fixed-point display.
const GENERAL_EXPONENTIAL_MANTISSA_DIGITS = 6;

/** Strips trailing fractional zeros (and a dangling ".") from a fixed-point
 *  digit string. No-op for strings without a decimal point. */
function trimTrailingZeros(digits: string): string {
  if (!digits.includes('.')) return digits;
  return digits.replace(/0+$/, '').replace(/\.$/, '');
}

/** Formats `exponent` the way Excel's General exponential notation does:
 *  always signed, at least two digits (E+05, E+11, E-09, ...). */
function formatExcelExponent(exponent: number): string {
  const sign = exponent >= 0 ? '+' : '-';
  const digits = Math.abs(exponent).toString().padStart(2, '0');
  return `${sign}${digits}`;
}

/** Renders a finite, non-zero, non-negative number in Excel's General
 *  exponential style: mantissa trimmed to `GENERAL_EXPONENTIAL_MANTISSA_DIGITS`
 *  significant digits with trailing zeros dropped, uppercase `E`, signed
 *  2+-digit exponent (e.g. 123456789012 -> "1.23457E+11"). */
function formatGeneralExponential(abs: number): string {
  const [mantissa, exponent] = abs.toExponential(GENERAL_EXPONENTIAL_MANTISSA_DIGITS - 1).split('e');
  return `${trimTrailingZeros(mantissa)}E${formatExcelExponent(Number(exponent))}`;
}

/**
 * Formats a number the way Excel's "General" cell format does: round to 11
 * significant digits (hiding binary floating-point round-trip noise like
 * 0.1 + 0.2), trim trailing fractional zeros, and switch to Excel-style
 * exponential notation ("1.23457E+11") once the value's decimal exponent
 * falls outside the fixed-point display budget.
 *
 * Thresholds:
 * - Integer part >= 12 digits (decimal exponent >= 11): Excel General is
 *   documented to switch 12+ digit numbers to scientific notation.
 * - Decimal exponent < -5 (value would need 6+ leading fractional zeros
 *   before the first significant digit): mirrors the same "11 significant
 *   digits must fit in the fixed-point budget" rule on the small-number
 *   side — not a numerically-documented Microsoft threshold, but the
 *   consistent extrapolation of the documented large-number rule.
 *
 * Column-width narrowing is not modeled: Excel shrinks a General value's
 * displayed precision further to fit a narrow column, but this function always
 * emits the full 11-significant-digit form regardless of the destination
 * column width, matching how the rest of this renderer treats layout as
 * independent of formatting.
 */
function formatGeneralNumber(num: number): string {
  if (!Number.isFinite(num)) return String(num);
  if (num === 0) return '0'; // canonicalizes -0 to "0"

  const negative = num < 0;
  const abs = Math.abs(num);

  // Decimal exponent of the value once rounded to the target significant
  // digits, derived from the (already-rounded) exponential form so a
  // rounding carry that bumps the digit count (e.g. 99999999999.6 -> 1E+11)
  // is reflected before the fixed-vs-exponential branch is chosen.
  const exponent = Number(abs.toExponential(GENERAL_SIGNIFICANT_DIGITS - 1).split('e')[1]);
  const useExponential = exponent >= GENERAL_SIGNIFICANT_DIGITS || exponent < -5;

  const body = useExponential
    ? formatGeneralExponential(abs)
    : trimTrailingZeros(abs.toPrecision(GENERAL_SIGNIFICANT_DIGITS));

  return negative ? `-${body}` : body;
}

function applyFormat(num: number, numFmtId: number, formatCode: string | null, date1904 = false): FormattedCell {
  // Built-in 14 is Excel's locale-sensitive short-date format rather than a
  // file-authored pattern. Use the host application's UI locale, while keeping
  // UTC fields so the Excel serial cannot cross a calendar-day boundary in a
  // non-UTC timezone.
  if (numFmtId === 14 && !formatCode) {
    return {
      text: formatLocalizedExcelShortDate(num, date1904),
    };
  }
  // Built-in date/time numFmtIds (ECMA-376 §18.8.30 table)
  const builtinFmt = BUILTIN_DATE_FMT[numFmtId];
  if (builtinFmt) return { text: formatExcelDateCode(num, builtinFmt, date1904) };
  // ECMA-376 §18.8.30: "General" is the reserved General number format regardless
  // of numFmtId. LibreOffice writes a custom numFmt (id ≥ 164) with
  // formatCode="General"; tokenizing it as a literal pattern would render the
  // word "General" instead of the value (issue #358).
  if (formatCode && formatCode.trim().toLowerCase() === 'general') return { text: formatGeneralNumber(num) };
  if (formatCode) {
    if (isDateFormatCode(formatCode)) return { text: formatExcelDateCode(num, formatCode, date1904) };
    return applyFormatCode(num, formatCode);
  }
  switch (numFmtId) {
    // Built-in numeric numFmtIds without an explicit formatCode. Route the ones
    // that have a well-defined pattern (§18.8.30 p.1776 "All Languages" table)
    // through the same grammar engine as custom codes so their placeholder /
    // sign-section semantics match exactly.
    case 0: return { text: formatGeneralNumber(num) };
    case 1: return applyFormatCode(num, '0');
    case 2: return applyFormatCode(num, '0.00');
    case 3: return applyFormatCode(num, '#,##0');
    case 4: return applyFormatCode(num, '#,##0.00');
    case 9: return applyFormatCode(num, '0%');
    case 10: return applyFormatCode(num, '0.00%');
    case 11: return applyFormatCode(num, '0.00E+00');
    case 37: return applyFormatCode(num, '#,##0 ;(#,##0)');
    case 38: return applyFormatCode(num, '#,##0 ;[Red](#,##0)');
    case 39: return applyFormatCode(num, '#,##0.00;(#,##0.00)');
    case 40: return applyFormatCode(num, '#,##0.00;[Red](#,##0.00)');
    case 48: return applyFormatCode(num, '##0.0E+0');
    case 49: return { text: String(num) };
    default: return { text: formatGeneralNumber(num) };
  }
}

// (formatExcelDate removed; all date formatting now goes through formatExcelDateCode)

// ════════════════════════════════════════════════════════════════════════════
// Number-format grammar engine (ECMA-376 §18.8.30 / §18.8.31)
//
// A number format is up to four `;`-separated sections (positive;negative;zero;
// text). Each numeric section is a token stream mixing literal text with a
// single numeric placeholder run built from `0` / `#` / `?` (and `.` , `E±`).
// Modifiers attached to a section: a leading colour `[Red]`/`[ColorN]`, a
// condition `[>=100]`, a `%` multiplier, and trailing `,` scaling. The engine
// is a tokenizer → AST → renderer, replacing the old regex-driven `.toFixed`
// approximation so `#`/`?` placeholder semantics, fractions, comma-scaling and
// conditional/coloured sections all follow the spec exactly.
// ════════════════════════════════════════════════════════════════════════════

/** §18.8.30 "Specify colors": the eight named section colours (p.1787). */
const NAMED_COLORS: Record<string, string> = {
  black: '#000000', blue: '#0000FF', cyan: '#00FFFF', green: '#008000',
  magenta: '#FF00FF', red: '#FF0000', white: '#FFFFFF', yellow: '#FFFF00',
};

// Legacy indexed colour palette (§18.8.27, indices 0-63). `[ColorN]` maps N→
// indexed=(N+7): the spec note says "[Color1] refers to indexed=8 ... [Color3]
// for Red". Kept in sync with the Rust parser's INDEXED_COLORS.
const INDEXED_COLORS: readonly string[] = [
  '#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF', // 0-7
  '#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF', // 8-15
  '#800000', '#008000', '#000080', '#808000', '#800080', '#008080', '#C0C0C0', '#808080', // 16-23
  '#9999FF', '#993366', '#FFFFCC', '#CCFFFF', '#660066', '#FF8080', '#0066CC', '#CCCCFF', // 24-31
  '#000080', '#FF00FF', '#FFFF00', '#00FFFF', '#800080', '#800000', '#008080', '#0000FF', // 32-39
  '#00CCFF', '#CCFFFF', '#CCFFCC', '#FFFF99', '#99CCFF', '#FF99CC', '#CC99FF', '#FFCC99', // 40-47
  '#3366FF', '#33CCCC', '#99CC00', '#FFCC00', '#FF9900', '#FF6600', '#666699', '#969696', // 48-55
  '#003366', '#339966', '#003300', '#333300', '#993300', '#993366', '#333399', '#333333', // 56-63
];

interface SectionCondition {
  op: '<' | '<=' | '>' | '>=' | '=' | '<>';
  value: number;
}

interface ParsedSection {
  /** Raw section body with `[...]` modifiers stripped out. */
  body: string;
  color?: string;
  condition?: SectionCondition;
}

/**
 * Split a whole format code into its `;`-separated sections. `;` never appears
 * inside quotes, escapes or `[...]` in a valid code, but we scan structurally
 * so a stray one inside those never splits the section (defensive; matches how
 * Excel lexes).
 */
function splitSections(code: string): string[] {
  const out: string[] = [];
  let cur = '';
  let i = 0;
  while (i < code.length) {
    const ch = code[i];
    if (ch === '"') {
      cur += ch; i++;
      while (i < code.length && code[i] !== '"') cur += code[i++];
      if (i < code.length) cur += code[i++];
    } else if (ch === '\\') {
      cur += ch;
      if (i + 1 < code.length) cur += code[i + 1];
      i += 2;
    } else if (ch === '[') {
      cur += ch; i++;
      while (i < code.length && code[i] !== ']') cur += code[i++];
      if (i < code.length) cur += code[i++];
    } else if (ch === ';') {
      out.push(cur); cur = ''; i++;
    } else {
      cur += ch; i++;
    }
  }
  out.push(cur);
  return out;
}

/** Parse a section's leading `[...]` modifiers (colour, condition, currency).
 *  Currency `[$sym-LCID]` is left *in* the body (its `$sym` is emitted as a
 *  literal by the tokenizer); only colour and condition brackets are consumed
 *  here. */
function parseSection(section: string): ParsedSection {
  let body = '';
  let color: string | undefined;
  let condition: SectionCondition | undefined;
  let i = 0;
  while (i < section.length) {
    const ch = section[i];
    if (ch === '"') { // pass a quoted literal straight through
      body += ch; i++;
      while (i < section.length && section[i] !== '"') body += section[i++];
      if (i < section.length) body += section[i++];
    } else if (ch === '\\') {
      body += ch;
      if (i + 1 < section.length) body += section[i + 1];
      i += 2;
    } else if (ch === '[') {
      const end = section.indexOf(']', i);
      if (end < 0) { body += ch; i++; continue; }
      const inner = section.slice(i + 1, end);
      const lower = inner.toLowerCase();
      const idxMatch = lower.match(/^color(\d{1,2})$/);
      const condMatch = inner.match(/^(<=|>=|<>|<|>|=)\s*(-?[0-9.]+(?:[eE][-+]?\d+)?)$/);
      if (lower in NAMED_COLORS) {
        color = NAMED_COLORS[lower];
      } else if (idxMatch) {
        const n = parseInt(idxMatch[1], 10);
        // §18.8.30: [ColorN] → indexed=(N+7); valid N is 1..56.
        if (n >= 1 && n <= 56) color = INDEXED_COLORS[n + 7] ?? color;
      } else if (condMatch) {
        condition = { op: condMatch[1] as SectionCondition['op'], value: Number(condMatch[2]) };
      } else {
        // Currency / locale / elapsed brackets stay in the body for the
        // tokenizer to interpret (`[$sym-LCID]`), so re-emit them verbatim.
        body += section.slice(i, end + 1);
      }
      i = end + 1;
    } else {
      body += ch; i++;
    }
  }
  return { body, color, condition };
}

function testCondition(cond: SectionCondition, num: number): boolean {
  switch (cond.op) {
    case '<': return num < cond.value;
    case '<=': return num <= cond.value;
    case '>': return num > cond.value;
    case '>=': return num >= cond.value;
    case '=': return num === cond.value;
    case '<>': return num !== cond.value;
  }
}

/**
 * The lexed pieces of one numeric section. Literal text and placeholder runs
 * are kept in order (`parts`) so digits can be substituted into placeholder
 * positions *in place* — this is what lets embedded literals such as the `-` in
 * a phone mask (`000\-00`) or the parentheses in `(000)` survive at their
 * original spot. `intSpec` / `fracSpec` are the placeholder-only strings (for
 * digit/decimal counting); `exp`, `hasPercent`, `commaScale`, `grouping`,
 * `fraction` are section-wide modifiers.
 */
interface LexedSection {
  parts: SecPart[];
  intSpec: string;    // integer placeholder chars only (0/#/?)
  fracSpec: string;   // fraction placeholder chars only (0/#/?)
  hasPercent: boolean;
  commaScale: number; // trailing commas after the last integer placeholder
  grouping: boolean;  // a grouping comma inside the integer placeholders
  exp?: { plus: boolean; width: number };
  fraction?: {
    /** Placeholder run for the whole-number part (`#` in `# ?/?`), or '' when
     *  the format is a pure fraction (`?/?` with no leading whole group). */
    wholeSpec: string;
    /** Placeholder run for the numerator (the group just before `/`). */
    numSpec: string;
    denSpec: string;
    fixedDen: number | null;
  };
}

type SecPart =
  | { kind: 'lit'; text: string }         // verbatim literal (quoted / escaped / symbol / space)
  | { kind: 'intph'; ph: string }          // one integer placeholder char (0/#/?) — positional fill
  | { kind: 'dot' }                        // the decimal point
  | { kind: 'fracph'; ph: string }         // one fraction placeholder char
  | { kind: 'percent' }
  | { kind: 'exp' }                        // marker: emit the exponent block here
  | { kind: 'fraction' };                  // marker: emit the whole `n/d` block here

/**
 * Lex one numeric section body (already stripped of `[...]` colour/condition
 * modifiers) into a `LexedSection`. Currency `[$sym-LCID]` brackets that
 * survived `parseSection` are expanded here to their literal symbol.
 */
function lexSection(body: string): LexedSection {
  const parts: SecPart[] = [];
  let intSpec = '';
  let fracSpec = '';
  let hasPercent = false;
  let inFrac = false;   // past the decimal point
  let exp: LexedSection['exp'];
  let sawSlash = false;
  // For fractions: the length of `intSpec` at the most recent gap (a literal
  // between integer placeholders). The numerator run is everything after it, so
  // `# ?/?` splits into whole `#` and numerator `?`.
  let intGapLen = 0;
  // Commas trailing the last placeholder (in int or frac position). Each scales
  // the value by 1000 (§18.8.30 comma-scaling rule). Reset by any placeholder.
  let trailingCommas = 0;

  const pushLit = (s: string) => {
    if (!s) return;
    if (!inFrac && !sawSlash) intGapLen = intSpec.replace(/,/g, '').length;
    const last = parts[parts.length - 1];
    if (last && last.kind === 'lit') last.text += s;
    else parts.push({ kind: 'lit', text: s });
  };

  let i = 0;
  while (i < body.length) {
    const ch = body[i];
    if (ch === '"') {
      i++;
      let s = '';
      while (i < body.length && body[i] !== '"') s += body[i++];
      if (i < body.length) i++;
      pushLit(s);
    } else if (ch === '\\') {
      if (i + 1 < body.length) pushLit(body[i + 1]);
      i += 2;
    } else if (ch === '[') {
      const end = body.indexOf(']', i);
      const inner = end > i ? body.slice(i + 1, end) : '';
      // Currency: `[$sym-LCID]` → emit `sym` (between `$` and `-`), drop LCID.
      if (inner.startsWith('$')) {
        const rest = inner.slice(1);
        const dash = rest.indexOf('-');
        pushLit(dash >= 0 ? rest.slice(0, dash) : rest);
      }
      i = end < 0 ? body.length : end + 1;
    } else if (ch === '_') {
      // `_x` — a space the width of x (§18.8.30 p.1786). Render as one space.
      pushLit(' ');
      i += 2;
    } else if (ch === '*') {
      // `*x` — repeat x to fill the column width (§18.8.30 p.1784). No column
      // width is available in this pure formatter, so emit a single x (Excel's
      // minimum). Layout-driven fill is out of scope for the display string.
      pushLit(body[i + 1] ?? '');
      i += 2;
    } else if (ch === '#' || ch === '0' || ch === '?') {
      if (inFrac) { fracSpec += ch; parts.push({ kind: 'fracph', ph: ch }); }
      else { intSpec += ch; parts.push({ kind: 'intph', ph: ch }); }
      trailingCommas = 0; // a placeholder after commas cancels trailing scaling
      i++;
    } else if (ch === '.') {
      inFrac = true;
      parts.push({ kind: 'dot' });
      i++;
    } else if (ch === ',') {
      // A comma between integer placeholders is a grouping (thousands) comma;
      // a comma trailing the last placeholder scales the value by 1000 each.
      if (!inFrac) intSpec += ',';
      trailingCommas++;
      i++;
    } else if (ch === '/' && (intSpec.replace(/,/g, '').length > 0)) {
      sawSlash = true;
      parts.push({ kind: 'fraction' });
      // Consume the denominator placeholders / literal.
      i++;
      let den = '';
      while (i < body.length && /[0-9#?]/.test(body[i])) den += body[i++];
      // Store on a temporary marker via closure vars (handled below).
      (parts[parts.length - 1] as { den?: string }).den = den;
    } else if (ch === '%') {
      hasPercent = true;
      parts.push({ kind: 'percent' });
      i++;
    } else if ((ch === 'E' || ch === 'e') && (body[i + 1] === '+' || body[i + 1] === '-')) {
      const plus = body[i + 1] === '+';
      i += 2;
      let width = 0;
      while (i < body.length && (body[i] === '0' || body[i] === '#' || body[i] === '?')) { width++; i++; }
      exp = { plus, width: Math.max(width, 1) };
      parts.push({ kind: 'exp' });
    } else {
      pushLit(ch);
      i++;
    }
  }

  // Trailing scaling commas (recorded during the scan, covering both the
  // integer-tail `#,##0,` and the fraction-tail `0.0,,` forms).
  const commaScale = trailingCommas;
  const grouping = /,(?=[#0?])/.test(intSpec);
  const intPlaceholders = intSpec.replace(/,/g, '');

  let fraction: LexedSection['fraction'];
  if (sawSlash) {
    const fracPart = parts.find(p => p.kind === 'fraction') as (SecPart & { den?: string }) | undefined;
    const denRaw = fracPart?.den ?? '?';
    const denLit = denRaw.match(/[0-9]+/);
    // Split the integer placeholders into whole (before the last gap) and
    // numerator (after it): `# ?/?` → whole `#`, numerator `?`.
    const wholeSpec = intPlaceholders.slice(0, intGapLen);
    const numSpec = intPlaceholders.slice(intGapLen) || '?';
    fraction = {
      wholeSpec,
      numSpec,
      denSpec: denRaw.replace(/[^0#?]/g, ''),
      fixedDen: denLit ? parseInt(denLit[0], 10) : null,
    };
  }

  return { parts, intSpec: intPlaceholders, fracSpec, hasPercent, commaScale, grouping, exp, fraction };
}

// ── Numeric rendering primitives ────────────────────────────────────────────

/** Group an integer digit string in threes with commas (thousands separator). */
function groupThousands(intDigits: string): string {
  return intDigits.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

/**
 * Fill an integer digit string into an integer placeholder template, right to
 * left, per §18.8.30 `0`/`#`/`?` semantics:
 *  - every actual digit is shown (extra digits beyond the placeholders spill
 *    out on the left);
 *  - `0` positions with no digit show `0`; `?` positions show a space; `#`
 *    positions show nothing.
 * Grouping commas are applied to the emitted digit run when `grouping` is set.
 */
function fillIntegerTemplate(intDigits: string, phString: string, grouping: boolean): string {
  const placeholders = phString.split('');
  const digits = intDigits.split('');
  const out: string[] = [];
  let di = digits.length - 1;
  let emitted: string[] = []; // collected actual digits (for grouping)

  // Walk placeholders right→left.
  for (let p = placeholders.length - 1; p >= 0; p--) {
    if (di >= 0) { out.unshift(digits[di]); emitted.unshift(digits[di]); di--; }
    else if (placeholders[p] === '0') { out.unshift('0'); emitted.unshift('0'); }
    else if (placeholders[p] === '?') { out.unshift(' '); }
    // '#' with no digit → nothing.
  }
  // Any remaining (higher-order) digits spill out to the left.
  while (di >= 0) { out.unshift(digits[di]); emitted.unshift(digits[di]); di--; }

  if (grouping) {
    const grouped = groupThousands(emitted.join(''));
    // Re-attach any leading `?` spaces that preceded the digit run.
    const leadSpaces = out.length - emitted.length > 0 ? out.slice(0, out.length - emitted.length).join('') : '';
    return leadSpaces + grouped;
  }
  return out.join('');
}

/** Render the fraction digits into a `fracph` template, keeping `0`, dropping
 *  trailing `#`, and padding trailing `?` with spaces (§18.8.30). */
function fillFractionText(fracDigits: string, fracSpec: string): string {
  const decCount = fracSpec.length;
  if (decCount === 0) return '';
  const chars = fracDigits.padEnd(decCount, '0').slice(0, decCount).split('');
  for (let k = decCount - 1; k >= 0; k--) {
    const ph = fracSpec[k] ?? '#';
    if (chars[k] === '0' && ph === '#') chars[k] = '';
    else if (chars[k] === '0' && ph === '?') chars[k] = ' ';
    else break;
  }
  return chars.join('');
}

/** Best rational approximation of `frac` (0<=frac<1) with a denominator of at
 *  most `maxDenDigits` digits, or an exact fixed denominator. Uses a
 *  Stern-Brocot mediant search (Excel's fraction display algorithm). */
function approximateFraction(frac: number, maxDenDigits: number, fixedDen: number | null): [number, number] {
  if (fixedDen !== null) return [Math.round(frac * fixedDen), fixedDen];
  const maxDen = Math.pow(10, Math.max(maxDenDigits, 1)) - 1;
  let bestN = 0, bestD = 1, bestErr = Math.abs(frac);
  let lo: [number, number] = [0, 1];
  let hi: [number, number] = [1, 1];
  for (let iter = 0; iter < 100; iter++) {
    const mN = lo[0] + hi[0];
    const mD = lo[1] + hi[1];
    if (mD > maxDen) break;
    const val = mN / mD;
    const err = Math.abs(val - frac);
    if (err < bestErr) { bestErr = err; bestN = mN; bestD = mD; }
    if (val < frac) lo = [mN, mD];
    else if (val > frac) hi = [mN, mD];
    else break;
  }
  return [bestN, bestD];
}

/**
 * Format `num` against one already-parsed numeric section. `useMagnitude`
 * strips the sign (the negative section supplies its own minus / parentheses).
 */
function renderNumericSection(num: number, body: string, useMagnitude: boolean): string {
  const lex = lexSection(body);

  let value = useMagnitude ? Math.abs(num) : num;
  if (lex.hasPercent) value = value * 100;
  if (lex.commaScale > 0) value = value / Math.pow(1000, lex.commaScale);

  const negative = value < 0;
  const sign = negative ? '-' : '';
  const abs = Math.abs(value);

  // ── Fraction section (`# ?/?`, `?/8`, …) ──────────────────────────────────
  if (lex.fraction) {
    const whole = Math.floor(abs);
    const frac = abs - whole;
    const { wholeSpec, numSpec, denSpec, fixedDen } = lex.fraction;
    const hasWholePart = wholeSpec.length > 0; // e.g. `# ?/?` has a `#` whole part
    const [numr, den] = approximateFraction(frac, denSpec.length, fixedDen);

    // §18.8.30: `?` pads insignificant positions with spaces so fractions align
    // on the slash — the numerator right-aligns (pad left), the denominator
    // left-aligns (pad right). `0` would zero-pad instead.
    const padNum = (n: number, ph: string): string => {
      let s = String(n);
      const pad = ph.includes('0') ? '0' : ' ';
      while (s.length < ph.length) s = pad + s;
      return s;
    };
    const padDen = (n: number, ph: string): string => {
      let s = String(n);
      const pad = ph.includes('0') ? '0' : ' ';
      while (s.length < ph.length) s = s + pad;
      return s;
    };

    let out = sign;
    if (hasWholePart) {
      const wholeText = whole > 0 ? String(whole) : (wholeSpec.includes('0') ? '0' : '');
      if (numr === 0) {
        // Integral value: Excel blanks the whole " n/d" group (including the
        // slash) with spaces so the whole number lines up with fractional
        // neighbours in the column. Width = space + numerator + slash + denom.
        const denWidth = fixedDen !== null ? String(fixedDen).length : (denSpec.length || 1);
        out += wholeText + ' '.repeat(1 + numSpec.length + 1 + denWidth);
      } else {
        const denText = fixedDen !== null ? String(fixedDen) : padDen(den, denSpec);
        out += wholeText + ' ' + padNum(numr, numSpec) + '/' + denText;
      }
    } else {
      // Pure fraction: fold the whole part back into the numerator.
      const totalNumr = numr + whole * den;
      const denText = fixedDen !== null ? String(fixedDen) : padDen(den, denSpec);
      out += padNum(totalNumr, numSpec) + '/' + denText;
    }
    return out;
  }

  // ── Scientific section ────────────────────────────────────────────────────
  if (lex.exp) {
    const intPlaceCount = Math.max(lex.intSpec.length, 1);
    const decCount = lex.fracSpec.length;
    let mantissa = 0, e = 0;
    if (abs !== 0) {
      e = Math.floor(Math.log10(abs));
      // Engineering grouping: exponent shifts to a multiple of the integer
      // placeholder count so `#0.0E+0` on 1.22e7 → `12.2E+6`.
      e = Math.floor(e / intPlaceCount) * intPlaceCount;
      mantissa = abs / Math.pow(10, e);
      if (parseFloat(roundDecimalHalfUp(mantissa, decCount)) >= Math.pow(10, intPlaceCount)) {
        e += intPlaceCount;
        mantissa = abs / Math.pow(10, e);
      }
    }
    // Round the mantissa's decimals half-UP like the plain section (Excel is
    // consistent across sections); `toFixed` would round the binary double down.
    const mantStr = roundDecimalHalfUp(mantissa, decCount);
    const [mInt, mFrac = ''] = mantStr.split('.');
    const intText = fillIntegerTemplate(mInt, lex.intSpec, false);
    const fracText = fillFractionText(mFrac, lex.fracSpec);
    const expSign = e < 0 ? '-' : (lex.exp.plus ? '+' : '');
    const expText = 'E' + expSign + String(Math.abs(e)).padStart(lex.exp.width, '0');
    return sign + assembleFixed(lex, intText, fracText, expText);
  }

  // ── Plain fixed-point section ─────────────────────────────────────────────
  const decCount = lex.fracSpec.length;
  // Excel rounds the displayed decimals half-UP on the decimal value
  // (2.675 under `0.00` → "2.68"); `toFixed` rounds the binary double down.
  const rounded = roundDecimalHalfUp(abs, decCount);
  const [intDigitsRaw, fracDigits = ''] = rounded.split('.');
  let intDigits = intDigitsRaw.replace(/^0+/, '');
  // Keep a leading zero when a `0`/`?` placeholder forces the units digit, or
  // when there is no fraction to lead with a bare dot.
  const forcesLeadingZero = /[0]/.test(lex.intSpec) || (lex.intSpec === '' && false);
  if (intDigits === '' && forcesLeadingZero) intDigits = '0';
  const intText = fillIntegerTemplate(intDigits, lex.intSpec, lex.grouping);
  const fracText = fillFractionText(fracDigits, lex.fracSpec);

  return sign + assembleFixed(lex, intText, fracText, '');
}

/**
 * Reassemble a section's literal parts around the rendered integer and fraction
 * blocks, filling `intph`/`fracph` placeholders positionally so embedded
 * literals (the `-` in a phone mask, the `(` `)` around accounting negatives)
 * land at their original positions.
 */
function assembleFixed(lex: LexedSection, intText: string, fracText: string, expText: string): string {
  // Split the pre-rendered integer text back across the intph placeholder
  // positions. We fill from the right: the last intph gets the last char of
  // intText, earlier ones the preceding chars, and the very first intph
  // absorbs all remaining (overflow) characters.
  const intChars = intText.split('');
  const intPhIndices: number[] = [];
  lex.parts.forEach((p, idx) => { if (p.kind === 'intph') intPhIndices.push(idx); });
  const fracChars = fracText.split('');
  const fracPhIndices: number[] = [];
  lex.parts.forEach((p, idx) => { if (p.kind === 'fracph') fracPhIndices.push(idx); });

  const intAssign = new Map<number, string>();
  let ci = intChars.length - 1;
  for (let k = intPhIndices.length - 1; k >= 0; k--) {
    if (k === 0) {
      // First placeholder absorbs everything remaining (overflow digits).
      let s = '';
      while (ci >= 0) s = intChars[ci--] + s;
      intAssign.set(intPhIndices[k], s);
    } else if (ci >= 0) {
      intAssign.set(intPhIndices[k], intChars[ci--]);
    } else {
      intAssign.set(intPhIndices[k], '');
    }
  }
  const fracAssign = new Map<number, string>();
  for (let k = 0; k < fracPhIndices.length; k++) {
    fracAssign.set(fracPhIndices[k], fracChars[k] ?? '');
  }

  // A dot is only shown when there is fraction content or a forced `0`/`?`.
  const showDot = lex.fracSpec.length > 0 && (fracText.length > 0 || /[0?]/.test(lex.fracSpec));

  let out = '';
  for (let idx = 0; idx < lex.parts.length; idx++) {
    const p = lex.parts[idx];
    if (p.kind === 'lit') out += p.text;
    else if (p.kind === 'intph') out += intAssign.get(idx) ?? '';
    else if (p.kind === 'fracph') out += fracAssign.get(idx) ?? '';
    else if (p.kind === 'dot') out += showDot ? '.' : '';
    else if (p.kind === 'percent') out += '%';
    else if (p.kind === 'exp') out += expText;
  }
  return out;
}


/**
 * Apply a full custom number-format code (§18.8.30) to a numeric value.
 * Handles section selection (positive/negative/zero + conditional overrides),
 * per-section colour, and the numeric grammar. Returns the display string and
 * any section colour.
 */
function applyFormatCode(num: number, formatCode: string): FormattedCell {
  const rawSections = splitSections(formatCode);
  const parsed = rawSections.map(parseSection);

  // Conditional sections (§18.8.30 "Specify conditions"): if any section
  // carries a `[cond]`, section selection is condition-driven — the first
  // section whose condition matches wins; a trailing section without a
  // condition is the "else". This overrides the positional pos/neg/zero rule.
  const hasConditions = parsed.some(s => s.condition);
  let chosen: ParsedSection | undefined;
  let useMagnitude = false;

  if (hasConditions) {
    let matchedByCondition = false;
    for (const sec of parsed) {
      if (sec.condition) {
        if (testCondition(sec.condition, num)) { chosen = sec; matchedByCondition = true; break; }
      } else {
        // Unconditional section acts as the default/else clause.
        chosen = chosen ?? sec;
        if (chosen === sec) break;
      }
    }
    if (!chosen) {
      // No criterion met and no else → Excel shows "#" across the cell.
      return { text: '#' };
    }
    // Sign semantics (§18.8.30 / §18.8.31): a section selected by its own
    // matching condition formats the value's *magnitude* — the section's
    // literals carry the sign presentation, exactly like the positional
    // negative section (the spec's `$0.00" Surplus";$-0.00" Shortage"` example
    // on p.1785 shows -125.74 as "$-125.74 Shortage": magnitude plus the
    // section's own literal `-`; built-ins 37-40 use parentheses the same way).
    // Prepending the sign on top would double it: `[<0]\-0.0` @ -5 would print
    // "--5.0". The unconditional "else" section, by contrast, mirrors the
    // positional fallback rule (a negative formatted by the only/positive
    // section keeps its sign, e.g. `0.0` @ -5 → "-5.0"), so a value that no
    // condition claimed keeps its sign there.
    useMagnitude = matchedByCondition && num < 0;
  } else {
    // Positional selection: positive;negative;zero (§18.8.31 p.1783).
    if (num > 0) chosen = parsed[0];
    else if (num < 0) {
      if (parsed.length > 1) { chosen = parsed[1]; useMagnitude = true; }
      else chosen = parsed[0];
    } else {
      chosen = parsed.length > 2 ? parsed[2] : parsed[0];
    }
  }

  const text = renderNumericSection(num, chosen.body, useMagnitude);
  return chosen.color ? { text, color: chosen.color } : { text };
}
