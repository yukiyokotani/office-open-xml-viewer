// Shared SpreadsheetML number-format grammar pieces (ECMA-376 §18.8.30) used
// by worksheet cells and chart labels alike: section splitting, the text
// section, date/time section detection and date/time rendering.

import { excelSerialToUtcDate, isExcelDisplayDateSerial } from './excel-date';
import { hasOwnKey } from './internal/own-key';

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

/** One numeric format section with its leading modifiers read. */
export interface ParsedFormatSection {
  /** Raw section body with colour and condition `[...]` modifiers removed. */
  body: string;
  color?: string;
  condition?: SectionCondition;
}

/** Parse a section's `[...]` modifiers (colour, condition, currency).
 *  Currency `[$sym-LCID]` is left *in* the body (its `$sym` is emitted as a
 *  literal by the tokenizer); only colour and condition brackets are consumed
 *  here. */
export function parseFormatSection(section: string): ParsedFormatSection {
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
    } else if (ch === '\\' || ch === '_' || ch === '*') {
      // An escape or a pad / fill pair: its operand is never a delimiter.
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
      if (hasOwnKey(NAMED_COLORS, lower)) {
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

/** The numeric section a value formats with (§18.8.30). */
export interface SelectedFormatSection {
  section: ParsedFormatSection;
  /** Index among the numeric sections (the text section left out): the
   *  index of the section's display divisor. */
  index: number;
  /** The section formats the value's magnitude (its literals carry the
   *  sign): a positional negative section, or a negative value claimed by
   *  a section's own condition. */
  useMagnitude: boolean;
}

/**
 * Choose the numeric section of `formatCode` for `value`, as worksheet
 * cells and chart labels both do. `null` when the code has no numeric
 * section; `'none'` when sections carry conditions and none matches with no
 * unconditional section to fall back on (Excel shows `#`).
 */
export function selectFormatSection(
  value: number,
  formatCode: string,
): SelectedFormatSection | 'none' | null {
  const raw = splitFormatSections(formatCode);
  // The text section never formats a number (§18.8.30).
  const textIndex = textSectionIndex(raw);
  const parsed = raw.filter((_, i) => i !== textIndex).map(parseFormatSection);
  if (parsed.length === 0) return null;
  // Conditional sections (§18.8.30 "Specify conditions"): if any section
  // carries a `[cond]`, section selection is condition-driven: the first
  // section whose condition matches wins; a section without a condition is
  // the "else". This overrides the positional pos/neg/zero rule.
  if (parsed.some(s => s.condition)) {
    for (let index = 0; index < parsed.length; index++) {
      const section = parsed[index];
      if (!section.condition) return { section, index, useMagnitude: false };
      if (testCondition(section.condition, value)) {
        // A section selected by its own matching condition formats the
        // value's magnitude: its literals carry the sign presentation, like
        // the positional negative section (§18.8.30 p.1785 example
        // `$0.00" Surplus";$-0.00" Shortage"`). The unconditional "else"
        // keeps the value's sign, like the positional fallback.
        return { section, index, useMagnitude: value < 0 };
      }
    }
    return 'none';
  }
  // Positional selection: positive;negative;zero (§18.8.31 p.1783).
  if (value > 0) return { section: parsed[0], index: 0, useMagnitude: false };
  if (value < 0) {
    return parsed.length > 1
      ? { section: parsed[1], index: 1, useMagnitude: true }
      : { section: parsed[0], index: 0, useMagnitude: false };
  }
  const index = parsed.length > 2 ? 2 : 0;
  return { section: parsed[index], index, useMagnitude: false };
}

/** Whether a section body is the reserved General keyword (§18.8.30). */
export function isGeneralSectionBody(body: string): boolean {
  return body.trim().toLowerCase() === 'general';
}

/** Index of the text section (§18.8.30), or -1: the fourth section when
 *  there are four, otherwise the last section when it holds an `@`
 *  placeholder outside quotes, escapes and pad / fill pairs. */
export function textSectionIndex(sections: string[]): number {
  if (sections.length >= 4) return 3;
  const last = sections[sections.length - 1];
  let i = 0;
  while (i < last.length) {
    const ch = last[i];
    if (ch === '\\' || ch === '_' || ch === '*') {
      i += 2;
    } else if (ch === '"') {
      const end = last.indexOf('"', i + 1);
      i = end < 0 ? last.length : end + 1;
    } else if (ch === '[') {
      const end = last.indexOf(']', i);
      i = end < 0 ? last.length : end + 1;
    } else if (ch === '@') {
      return sections.length - 1;
    } else {
      i++;
    }
  }
  return -1;
}

/**
 * Split a whole format code into its `;`-separated sections. `;` never appears
 * inside quotes, escapes or `[...]` in a valid code, but we scan structurally
 * so a stray one inside those never splits the section (defensive; matches how
 * Excel lexes).
 */
export function splitFormatSections(code: string): string[] {
  const out: string[] = [];
  let cur = '';
  let i = 0;
  while (i < code.length) {
    const ch = code[i];
    if (ch === '"') {
      cur += ch; i++;
      while (i < code.length && code[i] !== '"') cur += code[i++];
      if (i < code.length) cur += code[i++];
    } else if (ch === '\\' || ch === '_' || ch === '*') {
      // An escape or a pad / fill pair: its operand is never a delimiter.
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
 * Format an Excel date serial with one date/time format section (callers
 * select the section; see splitFormatSections / isDateFormatSection).
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
export function formatExcelDateTime(serial: number, section: string, date1904 = false): string {
  const elapsedSection = hasElapsedBracket(section);
  const subSecondDigits = fractionalSecondDigits(section);
  // Elapsed-only sections describe durations (§18.8.30), not calendar dates.
  // Validate after section selection/magnitude handling, before Date or Intl
  // can throw or produce invented/NaN text. '#' is independent of cell width;
  // the worksheet painter expands it across the available width.
  const validCalendar = isExcelDisplayDateSerial(serial, date1904);
  if (!Number.isFinite(serial) || !Number.isFinite(serial * 86_400_000)
      || (!elapsedSection && !validCalendar)) return '#';
  // Duration sections may mix clock and calendar tokens. Check actual
  // calendar fields below in their token context, so quoted calendar letters
  // and minute tokens cannot accidentally restrict a duration.
  // MS-OE376 §2.1.739(b), NFPartSubSecond: one to three zero placeholders
  // after a decimal separator display tenths/hundredths/milliseconds. Round
  // the entire clock once, so a fraction carry also updates seconds/minutes.
  // Formats without subsecond tokens retain the existing millisecond clock.
  const quantumMs = 10 ** (3 - (subSecondDigits || 3));
  const absMs = Math.round(Math.abs(serial) * 86_400_000 / quantumMs) * quantumMs;
  const roundedSerial = (serial < 0 ? -absMs : absMs) / 86_400_000;
  // Preserve the final-day saturation provided by excelSerialToUtcDate: a
  // fractional-second carry must not invent a calendar date in year 10000.
  const dateSerial = subSecondDigits && (!validCalendar || isExcelDisplayDateSerial(roundedSerial, date1904))
    ? roundedSerial : serial;
  const date = excelSerialToUtcDate(dateSerial, date1904);
  const yr = date.getUTCFullYear();
  const mo = date.getUTCMonth() + 1;   // 1-12
  const dy = date.getUTCDate();
  const wd = date.getUTCDay();          // 0=Sun
  // An elapsed-time section (`[h]:mm`, `[mm]:ss`) reads its clock fields as
  // remainders of the same absolute, millisecond-rounded duration as its
  // elapsed total, so a negative duration keeps consistent minutes.
  const hr = elapsedSection ? Math.floor(absMs / 3_600_000) % 24 : date.getUTCHours();
  const mi = elapsedSection ? Math.floor(absMs / 60_000) % 60 : date.getUTCMinutes();
  const sc = elapsedSection ? Math.floor(absMs / 1_000) % 60 : date.getUTCSeconds();

  // Take the first section (positive / no-sign section)
  const hasAmPm = hasAmPmToken(section);
  const japaneseEra = sectionHasJapaneseLocale(section);
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
        // Whole seconds of the millisecond-rounded duration, the same
        // rounding the clock fields above get from excelSerialToUtcDate.
        const absSec = Math.floor(absMs / 1000);
        let v: number;
        if      (kind === 'h') v = Math.floor(absSec / 3600);
        else if (kind === 'm') v = Math.floor(absSec / 60);
        else                   v = absSec;
        const padded = inner.length >= 2 ? String(v).padStart(inner.length, '0') : String(v);
        result += sign + padded;
        i = end + 1;
        const fraction = kind === 's' ? /^\.(0{1,3})(?!0)/.exec(section.slice(i)) : null;
        if (fraction) {
          const digits = fraction[1].length;
          result += '.' + String(Math.floor((absMs % 1000) / 10 ** (3 - digits))).padStart(digits, '0');
          i += fraction[0].length;
        }
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
      const escaped = section[i + 1];
      if (escaped !== undefined) result += escaped;
      i += 2;
      // An escaped time separator (`h\:mm`) keeps the hour context, like the
      // bare separators below.
      if (escaped !== ':' && escaped !== '/' && escaped !== '-' && escaped !== '.' && escaped !== ' ') {
        prevWasHour = false;
      }

    } else if (ch === 'y' || ch === 'Y') {
      if (!validCalendar) return '#';
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
        if (!validCalendar) return '#';
        if      (n === 1) result += String(mo);
        else if (n === 2) result += String(mo).padStart(2, '0');
        else if (n === 3) result += MONTH_NAMES[mo - 1].slice(0, 3);
        else if (n === 4) result += MONTH_NAMES[mo - 1];
        else              result += MONTH_NAMES[mo - 1][0]; // mmmmm = first letter
      }
      prevWasHour = false;

    } else if (ch === 'd' || ch === 'D') {
      if (!validCalendar) return '#';
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
      const fraction = /^\.(0{1,3})(?!0)/.exec(section.slice(i));
      if (fraction) {
        const digits = fraction[1].length;
        const milliseconds = elapsedSection ? absMs % 1000 : date.getUTCMilliseconds();
        result += '.' + String(Math.floor(milliseconds / 10 ** (3 - digits))).padStart(digits, '0');
        i += fraction[0].length;
      }
      prevWasHour = false;

    } else if (ch === 'g' || ch === 'G') {
      // Era name. Under a `[$-411]` (ja-JP) section, as Excel renders it:
      //   g → 'R' / 'H' / 'S' / 'T' / 'M', gg → '令' …, ggg → '令和' ….
      // Excel draws no era name in any other section (other LCIDs or none).
      let n = 0;
      while (i < section.length && section[i].toLowerCase() === 'g') { n++; i++; }
      if (japaneseEra) {
        if (!validCalendar) return '#';
        const e = getEra();
        if      (n === 1) result += e.abbr;
        else if (n === 2) result += e.short;
        else              result += e.long;
      }
      prevWasHour = false;

    } else if (ch === 'e' || ch === 'E') {
      if (!validCalendar) return '#';
      // Era year: under `[$-411]` the Japanese era year (`ee` zero-padded);
      // Excel renders `e` and `ee` as the four-digit year in any other section.
      let n = 0;
      while (i < section.length && section[i].toLowerCase() === 'e') { n++; i++; }
      if (japaneseEra) {
        const y = getEra().year;
        result += n >= 2 ? String(y).padStart(2, '0') : String(y);
      } else {
        result += String(yr).padStart(4, '0');
      }
      prevWasHour = false;

    } else if (ch === 'r' || ch === 'R') {
      if (!validCalendar) return '#';
      // Under `[$-411]`, `r` is `ee` and `rr` is `gggee` (§18.8.30); Excel
      // renders both as the four-digit year in any other section.
      let n = 0;
      while (i < section.length && section[i].toLowerCase() === 'r') { n++; i++; }
      if (japaneseEra) {
        const e = getEra();
        result += (n >= 2 ? e.long : '') + String(e.year).padStart(2, '0');
      } else {
        result += String(yr).padStart(4, '0');
      }
      prevWasHour = false;

    } else if (ch === 'A' || ch === 'a') {
      const upper = section.slice(i).toUpperCase();
      // Japanese weekday format codes (Excel ja locale). `aaaa` = "水曜日",
      // `aaa` = "水". Checked before AM/PM because those are shorter matches
      // and would otherwise swallow the leading 'a'.
      if (upper.startsWith('AAAA')) {
        if (!validCalendar) return '#';
        result += JP_WEEKDAY_LONG[wd]; i += 4;
      } else if (upper.startsWith('AAA')) {
        if (!validCalendar) return '#';
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

/**
 * Every name-table text {@link formatExcelDateTime} can paint for one date /
 * time section, beyond the section's own characters and ASCII digits: the
 * English month and weekday names (`mmm…`, `ddd…`), AM/PM and A/P, the
 * Japanese weekday names (`aaa`, `aaaa`) and, under a ja-JP locale, the
 * Japanese era names (`g…`). Font-source scalar demand (library policy) uses
 * it so a date cell never paints a scalar the demand left out. The token test
 * deliberately over-approximates (a quoted `"mmm"` also adds the month names):
 * a superset only loads more subsets, a missing scalar would paint a fallback.
 */
export function excelDateTimeSectionNameTexts(section: string): string[] {
  const texts: string[] = [];
  if (/m{3}/i.test(section)) texts.push(...MONTH_NAMES);
  if (/d{3}/i.test(section)) texts.push(...WEEKDAY_NAMES);
  if (/a{3}/i.test(section)) texts.push(...JP_WEEKDAY_SHORT, ...JP_WEEKDAY_LONG);
  if (/[ap]/i.test(section)) texts.push('AM', 'PM', 'A', 'P');
  if (/g/i.test(section) && sectionHasJapaneseLocale(section)) {
    for (const era of JP_ERAS) texts.push(era.abbr, era.short, era.long);
  }
  return texts;
}

/** Whether era codes in a section read as the Japanese era: the section
 *  carries a `[$…-LCID]` bracket (outside quotes) whose language ID is
 *  0x0411 ja-JP. Measured in Excel (ja-JP macOS): `[$-411]` renders the
 *  Japanese era, while `[$-404]`, `[$-409]` and codes without an LCID all
 *  render `g` as nothing and `e` / `ee` / `r` / `rr` as the four-digit year. */
function sectionHasJapaneseLocale(section: string): boolean {
  let i = 0;
  while (i < section.length) {
    const ch = section[i];
    if (ch === '\\' || ch === '_' || ch === '*') {
      i += 2;
    } else if (ch === '"') {
      const end = section.indexOf('"', i + 1);
      i = end < 0 ? section.length : end + 1;
    } else if (ch === '[') {
      const end = section.indexOf(']', i);
      if (end < 0) break;
      const lcid = /^\$[^-]*-([0-9a-f]+)$/i.exec(section.slice(i + 1, end));
      if (lcid && (parseInt(lcid[1], 16) & 0xffff) === 0x0411) return true;
      i = end + 1;
    } else {
      i++;
    }
  }
  return false;
}

/** Fractional seconds following an unescaped seconds token. Quoted text and
 * metadata never change clock precision. MS-OE376 limits precision to 3. */
function fractionalSecondDigits(section: string): number {
  let digits = 0;
  for (let i = 0; i < section.length;) {
    const ch = section[i];
    if (ch === '\\' || ch === '_' || ch === '*') {
      i += 2;
    } else if (ch === '[') {
      // ECMA-376 §18.8.30 includes `[ss].00` (elapsed seconds/hundredths).
      const token = /^\[s+\]\.(0{1,3})(?!0)/i.exec(section.slice(i));
      if (token) digits = Math.max(digits, token[1].length);
      const end = section.indexOf(']', i + 1);
      i = token ? i + token[0].length : end < 0 ? section.length : end + 1;
    } else if (ch === '"') {
      const end = section.indexOf('"', i + 1);
      i = end < 0 ? section.length : end + 1;
    } else if (ch === 's' || ch === 'S') {
      const token = /^s+\.(0{1,3})(?!0)/i.exec(section.slice(i));
      if (token) {
        digits = Math.max(digits, token[1].length);
        i += token[0].length;
      } else {
        while (i < section.length && section[i].toLowerCase() === 's') i++;
      }
    } else {
      i++;
    }
  }
  return digits;
}

/** Whether a section holds an AM/PM or A/P token outside quotes, escapes and
 *  pad / fill pairs (a quoted "AM/PM" is literal text). */
function hasAmPmToken(section: string): boolean {
  let i = 0;
  while (i < section.length) {
    const ch = section[i];
    if (ch === '\\' || ch === '_' || ch === '*') {
      i += 2;
    } else if (ch === '"') {
      const end = section.indexOf('"', i + 1);
      i = end < 0 ? section.length : end + 1;
    } else if (ch === '[') {
      const end = section.indexOf(']', i);
      i = end < 0 ? section.length : end + 1;
    } else if (/^(am\/pm|a\/p)/i.test(section.slice(i))) {
      return true;
    } else {
      i++;
    }
  }
  return false;
}

/** Whether a section holds an elapsed-time bracket `[h]` / `[mm]` / `[ss]`
 *  outside quotes, escapes and pad / fill pairs. */
function hasElapsedBracket(section: string): boolean {
  let i = 0;
  while (i < section.length) {
    const ch = section[i];
    if (ch === '\\' || ch === '_' || ch === '*') {
      i += 2;
    } else if (ch === '"') {
      const end = section.indexOf('"', i + 1);
      i = end < 0 ? section.length : end + 1;
    } else if (ch === '[') {
      const end = section.indexOf(']', i);
      if (end < 0) return false;
      if (/^([hms])\1*$/i.test(section.slice(i + 1, end))) return true;
      i = end + 1;
    } else {
      i++;
    }
  }
  return false;
}

/** Whether one format section (§18.8.30) is a date/time format. The body is
 *  scanned left to right so each token is read in context: `\x` escapes,
 *  `_x` padding and `*x` fill pairs (whose operand may itself be `"`), quoted
 *  literals and bracket content are skipped; what remains is date/time when
 *  it holds y / m / d / h / s, AM/PM or A/P, the Japanese weekday code
 *  `aaa+`, the Japanese era `g`, era year `e` and locale codes `r` / `rr`,
 *  or an elapsed-time bracket `[h]` / `[mm]` / `[ss]`. An `e` after a numeric
 *  placeholder (`0` `#` `?`) or before `+` / `-` / a placeholder is the
 *  scientific exponent, and `General` is the General keyword. */
export function isDateFormatSection(body: string): boolean {
  let i = 0;
  let afterPlaceholder = false;
  while (i < body.length) {
    const ch = body[i];
    if (ch === '\\' || ch === '_' || ch === '*') {
      i += 2;
    } else if (ch === '"') {
      const end = body.indexOf('"', i + 1);
      i = end < 0 ? body.length : end + 1;
    } else if (ch === '[') {
      const end = body.indexOf(']', i);
      if (end < 0) return false;
      if (/^([hms])\1*$/i.test(body.slice(i + 1, end))) return true;
      i = end + 1;
    } else if (/^general/i.test(body.slice(i))) {
      i += 'general'.length;
    } else if (/[ymdhsgr]/i.test(ch)) {
      return true;
    } else if (ch === 'e' || ch === 'E') {
      if (!afterPlaceholder && !/[-+0#?]/.test(body[i + 1] ?? '')) return true;
      i++;
    } else if (ch === '0' || ch === '#' || ch === '?') {
      afterPlaceholder = true;
      i++;
    } else if (/^(am\/pm|a\/p|a{3,})/i.test(body.slice(i))) {
      return true;
    } else {
      i++;
    }
  }
  return false;
}
