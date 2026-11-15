import {
  excelDateTimeSectionNameTexts,
  formatExcelDateTime,
  generalSectionAffixes,
  formatLocalizedExcelShortDate,
  localizedExcelShortDateTexts,
  isDateFormatSection,
  isExcelDisplayDateSerial,
  roundDecimalHalfUp,
  selectFormatSection,
  splitFormatSections,
  textSectionIndex,
} from '@silurus/ooxml-core';
import type { Cell, CellValue, NumFmt, Styles } from './types.js';

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
  /** Paint '#' across the available cell width, rather than a literal hash. */
  fill?: boolean;
  /** The text is a number in the General format, which shortens to the
   *  available width (§18.8.30). */
  general?: boolean;
}

/** A conditional-format number format: a dxf's id and code, with the
 *  section divisors of an OpenDocument display factor. */
type CfNumFmt = { numFmtId: number; formatCode: string | null; sectionDivisors?: number[] };

/** A `numFmtId` → `<numFmt>` lookup of one `styles.numFmts`, built once by a
 *  caller that formats many cells in one synchronous pass (the renderer)
 *  instead of a linear `numFmts` search per cell. */
export interface StyleNumFmtIndex {
  get(numFmtId: number): NumFmt | undefined;
}

/** Build a {@link StyleNumFmtIndex} that answers as
 *  `numFmts?.find(f => f.numFmtId === id)` does: a repeated id keeps its
 *  FIRST entry, and `NaN` is never indexed (`===` never matches it, whereas
 *  `Map` keys compare by SameValueZero). A collection that is not an array of
 *  objects keeps the linear search, so it succeeds or fails as before.
 *
 *  Contract: `numFmts` is library-owned plain data — parser output decoded
 *  by `JSON.parse` or a structured clone, or the delimited-text loader's own
 *  literal (data properties only, no getters), which no public API hands to
 *  callers for mutation (Node exposes its copy
 *  recursively frozen). The index reflects `numFmts` as it is now; build it
 *  per pass, never cache it across passes. Callers formatting caller-supplied
 *  styles (find, validation lists, the OpenDocument facade) keep the linear
 *  search by not passing an index. */
export function indexStyleNumFmts(numFmts: readonly NumFmt[] | undefined): StyleNumFmtIndex {
  if (Array.isArray(numFmts)) {
    let wellFormed = true;
    for (let i = 0; i < numFmts.length; i++) {
      const format: unknown = numFmts[i];
      if (typeof format !== 'object' || format === null) { wellFormed = false; break; }
    }
    if (wellFormed) {
      const index = new Map<number, NumFmt>();
      for (const format of numFmts) {
        if (Number.isNaN(format.numFmtId) || index.has(format.numFmtId)) continue;
        index.set(format.numFmtId, format);
      }
      return index;
    }
  }
  return { get: (numFmtId) => numFmts?.find(f => f.numFmtId === numFmtId) };
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
  cfNumFmt?: CfNumFmt | null,
  date1904 = false,
): string {
  return formatCellValueWithColor(cell, styles, cfNumFmt, date1904).text;
}

export function formatCellValueWithColor(
  cell: Cell,
  styles: Styles,
  cfNumFmt?: CfNumFmt | null,
  /** Workbook date system (`<workbookPr date1904>`, §18.2.28). `true` resolves
   *  serial dates against the 1904 epoch (§18.17.4.1). Defaults to false (1900
   *  date system) so callers that don't thread the flag are unaffected. */
  date1904 = false,
  /** {@link indexStyleNumFmts} of `styles.numFmts`, built in the caller's
   *  current pass; absent → a linear search. */
  numFmtIndex?: StyleNumFmtIndex,
): FormattedCell {
  // Resolve the effective format once so both the numeric and text paths
  // honour the same precedence: CF dxf numFmt > style numFmt (§18.8.15).
  const xf = styles.cellXfs[cell.styleIndex ?? 0];
  const styleNumFmtId = xf?.numFmtId ?? 0;
  const styleNumFmt = numFmtIndex
    ? numFmtIndex.get(styleNumFmtId)
    : styles.numFmts?.find(f => f.numFmtId === styleNumFmtId);
  const styleFmt = styleNumFmt?.formatCode ?? null;
  const effectiveFmtId = cfNumFmt?.numFmtId ?? styleNumFmtId;
  // A matched differential format replaces the id/code pair. An absent CF
  // code implies its built-in format, rather than inheriting the base code.
  const effectiveFmt = cfNumFmt ? cfNumFmt.formatCode : styleFmt;
  const divisors = cfNumFmt ? cfNumFmt.sectionDivisors : styleNumFmt?.sectionDivisors;

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
  return applyFormat(cell.value.number, effectiveFmtId, effectiveFmt, date1904, divisors);
}

/** The divisor of numeric section `index`: a finite positive factor, else 1. */
function sectionDivisor(divisors: readonly number[] | undefined, index: number): number {
  const divisor = divisors?.[index];
  return divisor !== undefined && Number.isFinite(divisor) && divisor > 0 ? divisor : 1;
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
  const sections = splitFormatSections(formatCode);
  const textIndex = textSectionIndex(sections);
  // Without a text section, text is unaffected by the format.
  if (textIndex < 0) return text;
  const section = sections[textIndex];
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
  45: 'mm:ss',
  46: '[h]:mm:ss',
  // MS-OE376 §2.1.739(a): Office resolves the ECMA table's `mmss.0`
  // to `mm:ss.0`. Use that documented Office pattern, as for built-in 14.
  47: 'mm:ss.0',
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

// Built-in numeric numFmtIds with a well-defined pattern (§18.8.30 p.1776
// "All Languages" table), formatted by the same grammar engine as custom codes.
const BUILTIN_NUMERIC_FMT: Record<number, string> = {
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  37: '#,##0 ;(#,##0)',
  38: '#,##0 ;[Red](#,##0)',
  39: '#,##0.00;(#,##0.00)',
  40: '#,##0.00;[Red](#,##0.00)',
  48: '##0.0E+0',
};

// ECMA-376 Part 1 §18.8.30 ("General" format, Numbers) specifies the General
// display of a number:
// - "display the full number up to 11 digits (inc. decimal point)"; its
//   "Floating point rule" makes that a cap of 11 characters, "not including
//   negative sign, but includes leading zeros and decimal separator".
// - A number too large switches to exponential format; too many significant
//   digits truncate (round) the display. "The optimal method of display is based
//   on the available cell width", and when no form fits that width the cell
//   shows "#" across its width.
// - Exponential format needs room for at least five characters (`xE-xx`). It
//   is used when the integer part does not fit, when not even one non-zero
//   digit fits in fixed-point form (negative exponent), and when a lossy
//   fixed-point form of a negative-exponent value would show fewer digits than
//   the exponential form (§18.8.30 conditions 2-4).
// The display works on Excel's 15-significant-digit value (Microsoft KB78113),
// which also keeps binary noise (0.1 + 0.2 = 0.30000000000000004) off screen.
// An exponential mantissa carries at most 6 significant digits
// (`1.23457E+11`), as Excel and the SheetJS SSF / Apache POI General
// emulations show.
const GENERAL_MAX_CHARS = 11;
const GENERAL_EXPONENTIAL_MAX_DECIMALS = 5;
const GENERAL_SOURCE_DIGITS = 15;

/** A positive finite value as 15 significant decimal digits `d0.d1d2… × 10^exp`. */
interface DecimalDigits { digits: number[]; exp: number }

function decimalDigits(abs: number): DecimalDigits {
  const [mantissa, exponent] = abs.toExponential(GENERAL_SOURCE_DIGITS - 1).split('e');
  return {
    digits: mantissa.replace('.', '').split('').map(Number),
    exp: Number(exponent),
  };
}

/** Rounds `value` half-up to `count` significant digits. `count === 0` rounds
 *  to the decade above the leading digit: one unit there, or zero (no digits);
 *  a negative count is zero. */
function roundSignificant(value: DecimalDigits, count: number): DecimalDigits {
  if (count < 0) return { digits: [], exp: value.exp };
  const kept = value.digits.slice(0, count);
  if ((value.digits[count] ?? 0) < 5) return { digits: kept, exp: value.exp };
  let i = kept.length - 1;
  for (; i >= 0; i--) {
    if (kept[i] === 9) kept[i] = 0;
    else { kept[i] += 1; break; }
  }
  // Carry out of the leading digit (or count === 0 rounding up): 10^(exp+1).
  if (i < 0) return { digits: [1], exp: value.exp + 1 };
  return { digits: kept, exp: value.exp };
}

/** Fixed-point text of `value` rounded to `decimals` places, trailing
 *  fractional zeros dropped; `null` when the rounded value is zero. */
function fixedText(value: DecimalDigits, decimals: number): string | null {
  const { digits, exp } = roundSignificant(value, value.exp + 1 + decimals);
  if (!digits.some(d => d !== 0)) return null;
  const digitAt = (power: number): number => {
    const index = exp - power;
    return index >= 0 && index < digits.length ? digits[index] : 0;
  };
  let intPart = '';
  for (let power = Math.max(exp, 0); power >= 0; power--) intPart += digitAt(power);
  let frac = '';
  for (let power = -1; power >= -decimals; power--) frac += digitAt(power);
  frac = frac.replace(/0+$/, '');
  return frac ? `${intPart}.${frac}` : intPart;
}

/** Excel's General exponent: uppercase E, always signed, at least two digits. */
function formatExcelExponent(exponent: number): string {
  const sign = exponent >= 0 ? '+' : '-';
  return `E${sign}${Math.abs(exponent).toString().padStart(2, '0')}`;
}

/** The fixed-point General form within `budget` characters, or `null` when the
 *  integer part does not fit or no non-zero digit would remain. */
function generalFixed(value: DecimalDigits, budget: number): string | null {
  if (value.exp >= 0) {
    const intDigits = value.exp + 1;
    if (intDigits > budget) return null;
    // A rounding carry (9999.6 -> 10000) can add an integer digit; drop one
    // more decimal and retry until the text fits the budget.
    for (let decimals = Math.max(0, budget - intDigits - 1); decimals >= 0; decimals--) {
      const text = fixedText(value, decimals);
      if (text !== null && text.length <= budget) return text;
    }
    return null;
  }
  // "0." plus the decimals left in the budget; a one-character budget can only
  // show a value that rounds to 1.
  const text = fixedText(value, Math.max(0, budget - 2));
  return text !== null && text.length <= budget ? text : null;
}

/** The exponential General form within `budget` characters (mantissa up to 6
 *  significant digits, trailing zeros dropped), or `null` when even `dE+xx`
 *  does not fit. */
function generalExponential(value: DecimalDigits, budget: number): string | null {
  for (let decimals = GENERAL_EXPONENTIAL_MAX_DECIMALS; decimals >= 0; decimals--) {
    const { digits, exp } = roundSignificant(value, decimals + 1);
    const frac = digits.slice(1).join('').replace(/0+$/, '');
    const text = `${digits[0]}${frac ? `.${frac}` : ''}${formatExcelExponent(exp)}`;
    if (text.length <= budget) return text;
  }
  return null;
}

/** Number of significant digits a General form shows. */
function shownDigits(text: string): number {
  return text.split('E')[0].replace('.', '').replace(/^0+/, '').length;
}

/** The General form of a positive value within a `budget`-character cap
 *  (§18.8.30 conditions for switching to exponential format). */
function generalForBudget(value: DecimalDigits, budget: number): string | null {
  const fixed = generalFixed(value, budget);
  const exponential = generalExponential(value, budget);
  if (fixed === null || exponential === null) return fixed ?? exponential;
  if (value.exp >= 0) return fixed;
  // Negative exponent: keep fixed point while it shows every significant
  // digit, or while exponential form would not show more of them.
  const significant = value.digits.join('').replace(/0+$/, '').length;
  const fixedDigits = shownDigits(fixed);
  if (fixedDigits >= significant) return fixed;
  return shownDigits(exponential) > fixedDigits ? exponential : fixed;
}

/**
 * Formats a number the way Excel's General format does (ECMA-376 §18.8.30).
 * Each character budget from the 11-character cap (sign excluded) down to one
 * has a single General form; the first that `fits` accepts is the display.
 * Narrower budgets first drop rounded decimal places, then switch to
 * exponential form as the §18.8.30 conditions require. Without `fits`, only
 * the cap applies (the text of a wide enough column). Returns `null` when no
 * form fits (§18.8.30 then shows "#"; {@link fitGeneralNumber} keeps the shortest form).
 */
export function formatGeneralNumber(num: number, fits?: (text: string) => boolean): string | null {
  if (!Number.isFinite(num)) return String(num);
  if (num === 0) return !fits || fits('0') ? '0' : null; // canonicalizes -0 to "0"
  const sign = num < 0 ? '-' : '';
  const value = decimalDigits(Math.abs(num));
  let previous: string | null = null;
  for (let budget = GENERAL_MAX_CHARS; budget >= 1; budget--) {
    const body = generalForBudget(value, budget);
    if (body === null || body === previous) continue;
    previous = body;
    const text = sign + body;
    if (!fits || fits(text)) return text;
  }
  return null;
}

/** The General text without a width: the 11-character cap alone. */
function formatGeneralText(num: number): string {
  return formatGeneralNumber(num) ?? '';
}

/**
 * Shortens a General number to its available width, as
 * {@link formatGeneralNumber} specifies (§18.8.30). When even the shortest
 * form does not fit, that shortest form is kept (and clipped by paint) rather
 * than replaced by Excel's "#" fill, so a value is never hidden. Every other
 * number format, non-numeric cells and texts that fit are returned unchanged.
 */
export function fitGeneralNumber(
  cell: Cell,
  formatted: FormattedCell,
  fits: (text: string) => boolean,
): FormattedCell {
  if (cell.value.type !== 'number' || !formatted.general || fits(formatted.text)) return formatted;
  let shortest: string | null = null;
  const text = formatGeneralNumber(cell.value.number, (candidate) => {
    shortest = candidate;
    return fits(candidate);
  });
  const shown = text ?? shortest;
  if (shown === null) return formatted;
  return formatted.color === undefined
    ? { text: shown, general: true }
    : { text: shown, color: formatted.color, general: true };
}

/**
 * Every text a cell can display under its effective number format (the
 * conditional-formatting override included): one entry, or, for a General
 * number, each distinct form {@link fitGeneralNumber} can shorten it to,
 * widest first with the unconstrained text leading. Find indexes these,
 * because the shown form depends on the column width and font, which only
 * paint knows.
 */
export function cellDisplayTexts(
  cell: Cell,
  styles: Styles,
  cfNumFmt?: { numFmtId: number; formatCode: string | null } | null,
  date1904 = false,
): string[] {
  const formatted = formatCellValueWithColor(cell, styles, cfNumFmt, date1904);
  if (cell.value.type !== 'number' || !formatted.general) return [formatted.text];
  const forms: string[] = [];
  formatGeneralNumber(cell.value.number, (candidate) => {
    forms.push(candidate);
    return false;
  });
  return forms.length > 0 ? forms : [formatted.text];
}

function applyFormat(
  num: number,
  numFmtId: number,
  formatCode: string | null,
  date1904 = false,
  divisors?: readonly number[],
): FormattedCell {
  // Built-in 14 is Excel's locale-sensitive short-date format rather than a
  // file-authored pattern. Use the host application's UI locale, while keeping
  // UTC fields so the Excel serial cannot cross a calendar-day boundary in a
  // non-UTC timezone.
  if (numFmtId === 14 && !formatCode) {
    return dateFormattedCell(formatLocalizedExcelShortDate(num, date1904), num, date1904);
  }
  // ECMA-376 §18.8.30: "General" is the reserved General number format regardless
  // of numFmtId. LibreOffice writes a custom numFmt (id ≥ 164) with
  // formatCode="General"; tokenizing it as a literal pattern would render the
  // word "General" instead of the value (issue #358).
  if (formatCode && formatCode.trim().toLowerCase() === 'general') {
    const divisor = sectionDivisor(divisors, 0);
    // Width fitting re-derives the General forms from the cell's own value,
    // so only an undivided General stays fittable.
    return divisor === 1
      ? { text: formatGeneralText(num), general: true }
      : { text: formatGeneralText(num / divisor) };
  }
  // §18.8.30 implies a built-in code when no numFmt is authored. An explicit
  // code (including a differential format layered by §18.8.15) takes precedence
  // over that fallback, even when its numFmtId also names a built-in format.
  if (formatCode) return applyFormatCode(num, formatCode, date1904, divisors);
  const builtinFmt = BUILTIN_DATE_FMT[numFmtId];
  if (builtinFmt) return dateFormattedCell(formatExcelDateTime(num, builtinFmt, date1904), num, date1904);
  // Built-in numeric numFmtIds without an explicit formatCode. Route the ones
  // that have a well-defined pattern (§18.8.30 p.1776 "All Languages" table)
  // through the same grammar engine as custom codes so their placeholder /
  // sign-section semantics match exactly.
  const builtinNumeric = BUILTIN_NUMERIC_FMT[numFmtId];
  if (builtinNumeric) return applyFormatCode(num, builtinNumeric);
  if (numFmtId === 49) return { text: String(num) };
  // 0 (General) and every unknown id.
  return { text: formatGeneralText(num), general: true };
}

/** Number-format paint texts of a workbook's styles (see
 * {@link numberFormatPaintTexts}). */
export interface NumberFormatPaintTexts {
  /** Per `cellXfs` index, the texts that style's number format can paint. */
  readonly cellXfs: readonly (readonly string[])[];
  /** Texts the conditional-format differential numFmts can paint; any cell
   * may take them, whatever its font. */
  readonly dxfs: readonly string[];
  /** False when some format's output set is not enumerated (a
   * locale/calendar-dependent built-in): scalar demand must then be `'all'`,
   * while the texts still prove the output's scripts. */
  readonly complete: boolean;
}

/**
 * Every non-value text the cell formatter can paint for number cells under
 * `styles`, per style: the input of font-source demand (library policy; see
 * `font-source-load.ts`), which charges it to each style's font as painted
 * cell text. Number cells paint, besides ASCII (digits, signs, separators,
 * `E`, `%`, space, `#` fill, General output), only
 *
 * - characters of the format code that formats them: literals, quoted and
 *   escaped text, currency symbols, `*` fill characters (the whole code is
 *   returned, so bracket metadata adds harmless extra scalars);
 * - the formatter's own name tables for a date / time section (month and
 *   weekday names, AM/PM, Japanese weekdays and eras:
 *   `excelDateTimeSectionNameTexts`);
 * - for built-in 14 without a code, the host locale's numeric short date
 *   (`localizedExcelShortDateTexts`): its complete output set for a
 *   Gregorian, ASCII-digit locale; otherwise a sample over the date range,
 *   with `complete: false` (another calendar or numbering system can paint
 *   era names and digits no finite sample proves).
 *
 * Codes come from the style numFmts and the differential formats a
 * conditional format may apply (§18.8.15); built-in ids are resolved as
 * {@link formatCellValueWithColor} resolves them.
 */
export function numberFormatPaintTexts(styles: Pick<Styles, 'numFmts' | 'cellXfs' | 'dxfs'>): NumberFormatPaintTexts {
  const codeOf = new Map<number, string>();
  for (const format of styles.numFmts ?? []) {
    if (format.formatCode !== null && format.formatCode !== undefined && !codeOf.has(format.numFmtId)) {
      codeOf.set(format.numFmtId, format.formatCode);
    }
  }
  let complete = true;
  const ofCode = new Map<string, readonly string[]>();
  const codeTexts = (code: string): readonly string[] => {
    let texts = ofCode.get(code);
    if (!texts) {
      const list = [code];
      for (const section of splitFormatSections(code)) {
        if (isDateFormatSection(section)) list.push(...excelDateTimeSectionNameTexts(section));
      }
      texts = Object.freeze(list);
      ofCode.set(code, texts);
    }
    return texts;
  };
  const ofBuiltin = new Map<number, readonly string[]>();
  const builtinTexts = (id: number): readonly string[] => {
    let texts = ofBuiltin.get(id);
    if (!texts) {
      if (id === 14) {
        const shortDate = localizedExcelShortDateTexts();
        if (!shortDate.complete) complete = false;
        texts = shortDate.texts;
      } else {
        const code = BUILTIN_DATE_FMT[id] ?? BUILTIN_NUMERIC_FMT[id];
        texts = code ? codeTexts(code) : Object.freeze([]);
      }
      ofBuiltin.set(id, texts);
    }
    return texts;
  };
  const cellXfs = (styles.cellXfs ?? []).map((xf) => {
    const id = xf.numFmtId ?? 0;
    const code = codeOf.get(id);
    return code ? codeTexts(code) : builtinTexts(id);
  });
  const dxfs: string[] = [];
  for (const dxf of styles.dxfs ?? []) {
    if (!dxf.numFmt) continue;
    dxfs.push(...(dxf.numFmt.formatCode ? codeTexts(dxf.numFmt.formatCode) : builtinTexts(dxf.numFmt.numFmtId)));
  }
  return Object.freeze({ cellXfs: Object.freeze(cellXfs), dxfs: Object.freeze(dxfs), complete });
}

/** Invalid dates use '#'. Keep width-dependent painting out of cellText/find;
 * a valid section may also return a literal '#', e.g. g"#" without ja-JP. */
function dateFormattedCell(text: string, serial: number, date1904: boolean): FormattedCell {
  return text === '#' && !isExcelDisplayDateSerial(serial, date1904) ? { text, fill: true } : { text };
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

// Section modifiers (colour, condition) and section selection are shared
// with chart labels: `parseFormatSection` / `selectFormatSection` in core.

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
 * survived `parseFormatSection` are expanded here to their literal symbol.
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
  // The ends of the search, 0/1 and 1/1, are candidates too: 0.999 with
  // one-digit denominators is 1/1 (the caller carries it), not 8/9.
  let bestN = 0, bestD = 1, bestErr = Math.abs(frac);
  if (Math.abs(1 - frac) < bestErr) { bestN = 1; bestD = 1; bestErr = Math.abs(1 - frac); }
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
    let whole = Math.floor(abs);
    const frac = abs - whole;
    const { wholeSpec, numSpec, denSpec, fixedDen } = lex.fraction;
    const hasWholePart = wholeSpec.length > 0; // e.g. `# ?/?` has a `#` whole part
    let [numr, den] = approximateFraction(frac, denSpec.length, fixedDen);
    // A fraction that rounds to a whole carries into the whole part.
    if (hasWholePart && numr === den) { whole += 1; numr = 0; }

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
function applyFormatCode(
  value: number,
  formatCode: string,
  date1904 = false,
  divisors?: readonly number[],
): FormattedCell {
  let num = value;
  // Positional (positive;negative;zero) or condition-driven selection, the
  // text section left out (§18.8.30); shared with chart labels.
  const selected = selectFormatSection(num, formatCode);
  if (selected === null) return { text: formatGeneralText(num), general: true };
  // No criterion met and no else → Excel shows "#" across the cell.
  if (selected === 'none') return { text: '#' };
  const { section: chosen, useMagnitude } = selected;

  // An OpenDocument display factor (ODF 1.3 Part 3 §19.346) divides the
  // value once its section is chosen on the undivided value.
  num /= sectionDivisor(divisors, selected.index);
  // Date/time is a property of the selected section (§18.8.30): `0.00;h:mm`
  // formats a positive value as a number and only a negative one as a time.
  const serial = useMagnitude ? Math.abs(num) : num;
  // A section whose number token is the General keyword (`[Red]General`,
  // `General;General`, `"$"General`) formats its (magnitude) value as
  // General between the section's literals (§18.8.30).
  const general = generalSectionAffixes(chosen.body);
  if (general) {
    const scaled = general.percent ? serial * 100 : serial;
    // A lone section carries no sign of its own: the minus leads it.
    const text = (scaled < 0 ? '-' : '') + general.prefix + formatGeneralText(Math.abs(scaled)) + general.suffix;
    return chosen.color ? { text, color: chosen.color } : { text };
  }
  const formatted = isDateFormatSection(chosen.body)
    ? dateFormattedCell(formatExcelDateTime(serial, chosen.body, date1904), serial, date1904)
    : { text: renderNumericSection(num, chosen.body, useMagnitude) };
  return chosen.color ? { ...formatted, color: chosen.color } : formatted;
}
