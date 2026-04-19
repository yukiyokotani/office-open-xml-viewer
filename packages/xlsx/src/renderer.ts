import type {
  Worksheet, Styles, Cell, CellValue, Font, Fill, Border, BorderEdge, CellXf,
  ViewportRange, RenderViewportOptions,
  CfRule, CellRange, CfStop, CfValue, Dxf, Hyperlink,
  Run, ChartData,
} from './types.js';
import { renderChart, type ChartModel } from '@silurus/ooxml-core';

const DEFAULT_FONT_FAMILY = 'Calibri, Arial, sans-serif';
const DEFAULT_FONT_SIZE = 11;
const MDW = 7;
const ROW_HEIGHT_TO_PX = 4 / 3;

export const HEADER_W = 50;
export const HEADER_H = 22;

// Thin line drawn between frozen and scrollable areas
const FREEZE_LINE_COLOR = '#7a7a7a';

export function colWidthToPx(w: number): number {
  return Math.trunc(((256 * w + 128 / MDW) / 256) * MDW);
}

export function rowHeightToPx(h: number): number {
  return Math.round(h * ROW_HEIGHT_TO_PX);
}

function hexToRgba(hex: string, alpha = 1): string {
  const h = hex.replace('#', '');
  const r = parseInt(h.slice(0, 2), 16);
  const g = parseInt(h.slice(2, 4), 16);
  const b = parseInt(h.slice(4, 6), 16);
  return alpha === 1 ? `rgb(${r},${g},${b})` : `rgba(${r},${g},${b},${alpha})`;
}

function buildFont(font: Font, cs = 1): string {
  const style = font.italic ? 'italic ' : '';
  const weight = font.bold ? 'bold ' : '';
  const sizePx = Math.max(1, Math.round(font.size * ROW_HEIGHT_TO_PX * cs));
  const family = font.name ? `"${font.name}", ${DEFAULT_FONT_FAMILY}` : DEFAULT_FONT_FAMILY;
  return `${style}${weight}${sizePx}px ${family}`;
}

/**
 * Resolve a Run's font against a base Font. Per ECMA-376, a run's <rPr>
 * completely specifies bold/italic/underline/strike for that run, while
 * size/color/name fall back to the base when omitted. A run with no
 * <rPr> (run.font undefined) inherits the base entirely.
 */
function applyRunFont(base: Font, run: Run): Font {
  const rf = run.font;
  if (!rf) return base;
  return {
    bold: rf.bold,
    italic: rf.italic,
    underline: rf.underline,
    strike: rf.strike,
    size: rf.size ?? base.size,
    color: rf.color ?? base.color,
    name: rf.name ?? base.name,
  };
}

function resolveXf(styles: Styles, styleIndex: number): { font: Font; fill: Fill; border: Border; xf: CellXf } {
  const xf: CellXf = styles.cellXfs[styleIndex] ?? styles.cellXfs[0] ?? {
    fontId: 0, fillId: 0, borderId: 0, numFmtId: 0, alignH: null, alignV: null, wrapText: false,
  };
  const font: Font = styles.fonts[xf.fontId] ?? { bold: false, italic: false, underline: false, strike: false, size: DEFAULT_FONT_SIZE, color: null, name: null };
  const fill: Fill = styles.fills[xf.fillId] ?? { patternType: 'none', fgColor: null, bgColor: null };
  const border: Border = styles.borders[xf.borderId] ?? { left: null, right: null, top: null, bottom: null };
  return { font, fill, border, xf };
}

function cellValueText(value: CellValue): string {
  switch (value.type) {
    case 'empty': return '';
    case 'text': return value.text;
    case 'number': return String(value.number);
    case 'bool': return value.bool ? 'TRUE' : 'FALSE';
    case 'error': return value.error;
  }
}

function formatCellValue(cell: Cell, styles: Styles): string {
  if (cell.value.type !== 'number') return cellValueText(cell.value);
  const xf = styles.cellXfs[cell.styleIndex ?? 0];
  const numFmtId = xf?.numFmtId ?? 0;
  const num = cell.value.number;
  const customFmt = styles.numFmts?.find(f => f.numFmtId === numFmtId);
  return applyFormat(num, numFmtId, customFmt?.formatCode ?? null);
}

// ────────────────────────────────────────────────────────────────
// Date / time formatting  (ECMA-376 §18.8.30)
// ────────────────────────────────────────────────────────────────

// Built-in numFmtId → format code (US English locale, as per the spec)
const BUILTIN_DATE_FMT: Record<number, string> = {
  14: 'm/d/yyyy',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'm/d/yyyy h:mm',
};

const MONTH_NAMES = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
];
const WEEKDAY_NAMES = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

/** Convert an Excel date serial to a UTC Date (avoids local-timezone off-by-one errors). */
function excelSerialToUTCDate(serial: number): Date {
  return new Date((serial - 25569) * 86400 * 1000);
}

/**
 * Format an Excel date serial using an ECMA-376 format code.
 * Supports: y/yy/yyy/yyyy, m/mm/mmm/mmmm/mmmmm, d/dd/ddd/dddd,
 *           h/hh, m/mm (minutes when after h), s/ss, AM/PM, A/P,
 *           quoted literals, bracket escapes, _ padding, * fill.
 */
function formatExcelDateCode(serial: number, fmtCode: string): string {
  const date = excelSerialToUTCDate(serial);
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
      // Locale / colour / condition bracket — skip entirely
      while (i < section.length && section[i] !== ']') i++;
      if (i < section.length) i++;

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

    } else if (ch === 'A' || ch === 'a') {
      const upper = section.slice(i).toUpperCase();
      if (upper.startsWith('AM/PM')) {
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
  // Strip quoted literals and bracket content, then look for unambiguous date specifiers.
  // 'y' = year, 'd' = day — both are unambiguous. 'm' alone is ambiguous (month or minutes).
  const stripped = code.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '');
  return /[yd]/i.test(stripped);
}

function applyFormat(num: number, numFmtId: number, formatCode: string | null): string {
  // Built-in date/time numFmtIds (ECMA-376 §18.8.30 table)
  const builtinFmt = BUILTIN_DATE_FMT[numFmtId];
  if (builtinFmt) return formatExcelDateCode(num, builtinFmt);
  if (formatCode) {
    if (isDateFormatCode(formatCode)) return formatExcelDateCode(num, formatCode);
    return applyFormatCode(num, formatCode);
  }
  switch (numFmtId) {
    case 0: return String(num);
    case 1: return Math.round(num).toString();
    case 2: return num.toFixed(2);
    case 3: return formatThousands(num, 0);
    case 4: return formatThousands(num, 2);
    case 9: return Math.round(num * 100) + '%';
    case 10: return (num * 100).toFixed(2) + '%';
    case 11: return num.toExponential(2);
    case 37: case 38: return formatThousands(num, 0);
    case 39: case 40: return formatThousands(num, 2);
    case 49: return String(num);
    default: return String(num);
  }
}

function formatThousands(num: number, decimals: number): string {
  return num.toLocaleString('en-US', { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
}

// (formatExcelDate removed; all date formatting now goes through formatExcelDateCode)

function countDecimalPlaces(fmt: string): number {
  const m = fmt.match(/\.([0#]+)/);
  return m ? m[1].length : 0;
}

function applyFormatCode(num: number, formatCode: string): string {
  const sections = formatCode.split(';');
  const section = num < 0 && sections.length > 1 ? sections[1] : sections[0];
  const cleaned = section.replace(/\[.*?\]/g, '').replace(/_./g, '').replace(/\*/g, '');
  if (cleaned.includes('%')) return (num * 100).toFixed(countDecimalPlaces(cleaned)) + '%';
  const hasThousands = cleaned.includes(',') && (cleaned.includes('#') || cleaned.includes('0'));
  const dec = countDecimalPlaces(cleaned);
  if (hasThousands) return formatThousands(num, dec);
  if (cleaned.includes('.')) return num.toFixed(dec);
  if (cleaned.match(/[#0]/)) return Math.round(num).toString();
  return String(num);
}

function wrapTextLines(ctx: CanvasRenderingContext2D, text: string, maxWidth: number): string[] {
  const words = text.split(' ');
  const lines: string[] = [];
  let current = '';
  for (const word of words) {
    const test = current ? `${current} ${word}` : word;
    if (ctx.measureText(test).width <= maxWidth || !current) {
      current = test;
    } else {
      lines.push(current);
      current = word;
    }
  }
  if (current) lines.push(current);
  return lines;
}

interface RichSeg {
  text: string;
  font: Font;
  width: number; // px
}

interface RichLine {
  segments: RichSeg[];
  maxFontSize: number; // pt (line-height source)
}

/**
 * Layout rich text runs into wrapped lines. Each run is split into words (and
 * CJK characters for granular wrapping). Per-run font is preserved so measurement
 * and drawing use the correct font.
 *
 * Follows ECMA-376 §18.3.1.53 (w:r) semantics: runs are inline and share the
 * paragraph width. wrapText breaks at word boundaries (ASCII spaces) and at any
 * CJK code point boundary.
 */
function layoutRichTextLines(
  ctx: CanvasRenderingContext2D,
  runs: Run[],
  baseFont: Font,
  cs: number,
  maxWidth: number,
): RichLine[] {
  const lines: RichLine[] = [];
  let cur: RichSeg[] = [];
  let curW = 0;
  let curMaxSize = 0;

  const flush = () => {
    if (cur.length === 0) return;
    lines.push({ segments: cur, maxFontSize: curMaxSize });
    cur = []; curW = 0; curMaxSize = 0;
  };

  const push = (text: string, font: Font) => {
    if (!text) return;
    ctx.font = buildFont(font, cs);
    const w = ctx.measureText(text).width;
    if (cur.length > 0 && curW + w > maxWidth) flush();
    cur.push({ text, font, width: w });
    curW += w;
    if (font.size > curMaxSize) curMaxSize = font.size;
  };

  const isCJK = (cp: number) =>
    (cp >= 0x3000 && cp <= 0x9FFF) ||
    (cp >= 0xF900 && cp <= 0xFAFF) ||
    (cp >= 0xAC00 && cp <= 0xD7AF) ||
    (cp >= 0xFF00 && cp <= 0xFFEF);

  for (const run of runs) {
    const font = applyRunFont(baseFont, run);
    // Tokenize: runs of non-space latin, spaces, or individual CJK chars
    const tokens: string[] = [];
    let i = 0;
    while (i < run.text.length) {
      const ch = run.text[i];
      const cp = ch.codePointAt(0) ?? 0;
      if (cp === 0x000A) {
        // Explicit newline: force break
        tokens.push('\n'); i += 1;
      } else if (isCJK(cp)) {
        tokens.push(ch);
        i += cp > 0xFFFF ? 2 : 1;
      } else if (ch === ' ') {
        let j = i;
        while (j < run.text.length && run.text[j] === ' ') j++;
        tokens.push(run.text.slice(i, j));
        i = j;
      } else {
        let j = i;
        while (j < run.text.length) {
          const c = run.text[j];
          const p = c.codePointAt(0) ?? 0;
          if (c === ' ' || c === '\n' || isCJK(p)) break;
          j += p > 0xFFFF ? 2 : 1;
        }
        tokens.push(run.text.slice(i, j));
        i = j;
      }
    }
    for (const tok of tokens) {
      if (tok === '\n') flush();
      else push(tok, font);
    }
  }
  flush();
  return lines;
}

function colToLetter(col: number): string {
  let result = '';
  while (col > 0) {
    const rem = (col - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    col = Math.floor((col - 1) / 26);
  }
  return result;
}

// ────────────────────────────────────────────────────────────────
// Conditional formatting
// ────────────────────────────────────────────────────────────────
interface CompiledCfRule {
  rule: CfRule;
  sqref: CellRange[];
  scaleMin?: number;
  scaleMax?: number;
  scaleStops?: number[];
  barMin?: number;
  barMax?: number;
  top10Threshold?: number;
  top10IsTop?: boolean;
  avgValue?: number;
  avgIsAbove?: boolean;
  iconThresholds?: number[];
}

interface CfContext {
  compiled: CompiledCfRule[];
}

interface CfResult {
  fill?: Fill;
  fontColor?: string;
  fontBold?: boolean;
  fontItalic?: boolean;
  fontUnderline?: boolean;
  fontStrike?: boolean;
  dataBar?: { color: string; ratio: number };
  iconSet?: { name: string; index: number };
}

function rangeContains(ranges: CellRange[], row: number, col: number): boolean {
  for (const r of ranges) {
    if (row >= r.top && row <= r.bottom && col >= r.left && col <= r.right) return true;
  }
  return false;
}

function cellNumericValue(cell: Cell | undefined): number | null {
  if (!cell) return null;
  if (cell.value.type === 'number') return cell.value.number;
  return null;
}

function collectNumericValuesInRanges(worksheet: Worksheet, ranges: CellRange[]): number[] {
  const out: number[] = [];
  for (const row of worksheet.rows) {
    for (const c of row.cells) {
      if (c.value.type !== 'number') continue;
      if (rangeContains(ranges, c.row, c.col)) out.push(c.value.number);
    }
  }
  return out;
}

function resolveCfvoValue(cfv: CfValue | CfStop, samples: number[]): number {
  const minv = samples.length ? Math.min(...samples) : 0;
  const maxv = samples.length ? Math.max(...samples) : 0;
  const n = cfv.value != null ? parseFloat(cfv.value) : NaN;
  switch (cfv.kind) {
    case 'min': return minv;
    case 'max': return maxv;
    case 'num': return isNaN(n) ? 0 : n;
    case 'percent': {
      const p = isNaN(n) ? 50 : n;
      return minv + (maxv - minv) * (p / 100);
    }
    case 'percentile': {
      if (!samples.length) return 0;
      const sorted = [...samples].sort((a, b) => a - b);
      const p = (isNaN(n) ? 50 : n) / 100;
      const idx = Math.max(0, Math.min(sorted.length - 1, Math.round(p * (sorted.length - 1))));
      return sorted[idx];
    }
    default: return isNaN(n) ? 0 : n;
  }
}

function compileCf(worksheet: Worksheet): CfContext {
  const compiled: CompiledCfRule[] = [];
  for (const cf of worksheet.conditionalFormats ?? []) {
    const samples = collectNumericValuesInRanges(worksheet, cf.sqref);
    for (const rule of cf.rules) {
      const entry: CompiledCfRule = { rule, sqref: cf.sqref };
      if (rule.type === 'colorScale') {
        entry.scaleStops = rule.stops.map(s => resolveCfvoValue(s, samples));
      } else if (rule.type === 'dataBar') {
        entry.barMin = resolveCfvoValue(rule.min, samples);
        entry.barMax = resolveCfvoValue(rule.max, samples);
      } else if (rule.type === 'top10') {
        const sorted = [...samples].sort((a, b) => a - b);
        const n = sorted.length;
        if (n > 0) {
          const rank = Math.min(rule.rank, n);
          if (rule.percent) {
            const p = rule.top ? (1 - rank / 100) : (rank / 100);
            const idx = Math.max(0, Math.min(n - 1, Math.round(p * (n - 1))));
            entry.top10Threshold = sorted[idx];
          } else {
            entry.top10Threshold = rule.top ? sorted[Math.max(0, n - rank)] : sorted[Math.min(n - 1, rank - 1)];
          }
          entry.top10IsTop = rule.top;
        }
      } else if (rule.type === 'aboveAverage') {
        if (samples.length > 0) {
          entry.avgValue = samples.reduce((a, b) => a + b, 0) / samples.length;
          entry.avgIsAbove = rule.aboveAverage;
        }
      } else if (rule.type === 'iconSet') {
        entry.iconThresholds = rule.cfvos.map(cfv => resolveCfvoValue(cfv, samples));
      }
      compiled.push(entry);
    }
  }
  // Higher priority (lower number) wins
  compiled.sort((a, b) => {
    const pa = (a.rule as { priority: number }).priority ?? 0;
    const pb = (b.rule as { priority: number }).priority ?? 0;
    return pa - pb;
  });
  return { compiled };
}

function cellIsMatch(num: number, operator: string, args: number[]): boolean {
  switch (operator) {
    case 'greaterThan': return num > (args[0] ?? 0);
    case 'greaterThanOrEqual': return num >= (args[0] ?? 0);
    case 'lessThan': return num < (args[0] ?? 0);
    case 'lessThanOrEqual': return num <= (args[0] ?? 0);
    case 'equal': return num === (args[0] ?? 0);
    case 'notEqual': return num !== (args[0] ?? 0);
    case 'between': return num >= (args[0] ?? 0) && num <= (args[1] ?? 0);
    case 'notBetween': return num < (args[0] ?? 0) || num > (args[1] ?? 0);
    default: return false;
  }
}

function interpolateHex(a: string, b: string, t: number): string {
  const pa = a.replace('#', '');
  const pb = b.replace('#', '');
  const ar = parseInt(pa.slice(0, 2), 16), ag = parseInt(pa.slice(2, 4), 16), ab = parseInt(pa.slice(4, 6), 16);
  const br = parseInt(pb.slice(0, 2), 16), bg = parseInt(pb.slice(2, 4), 16), bb = parseInt(pb.slice(4, 6), 16);
  const r = Math.round(ar + (br - ar) * t);
  const g = Math.round(ag + (bg - ag) * t);
  const bl = Math.round(ab + (bb - ab) * t);
  return `#${r.toString(16).padStart(2, '0').toUpperCase()}${g.toString(16).padStart(2, '0').toUpperCase()}${bl.toString(16).padStart(2, '0').toUpperCase()}`;
}

function colorScaleAt(num: number, stops: CfStop[], stopValues: number[]): string {
  if (!stops.length) return '#FFFFFF';
  if (num <= stopValues[0]) return stops[0].color;
  if (num >= stopValues[stopValues.length - 1]) return stops[stops.length - 1].color;
  for (let i = 1; i < stopValues.length; i++) {
    if (num <= stopValues[i]) {
      const lo = stopValues[i - 1];
      const hi = stopValues[i];
      const t = hi === lo ? 0 : (num - lo) / (hi - lo);
      return interpolateHex(stops[i - 1].color, stops[i].color, t);
    }
  }
  return stops[stops.length - 1].color;
}

function applyDxfToResult(result: CfResult, dxf: Dxf | null | undefined): void {
  if (!dxf) return;
  if (dxf.fill?.fgColor) result.fill = dxf.fill;
  if (dxf.font?.color) result.fontColor = dxf.font.color;
  if (dxf.font?.bold) result.fontBold = true;
  if (dxf.font?.italic) result.fontItalic = true;
  if (dxf.font?.underline) result.fontUnderline = true;
  if (dxf.font?.strike) result.fontStrike = true;
}

function evaluateCf(cell: Cell | undefined, row: number, col: number, cfCtx: CfContext, dxfs: Dxf[]): CfResult {
  const result: CfResult = {};
  if (!cfCtx.compiled.length) return result;
  for (const entry of cfCtx.compiled) {
    if (!rangeContains(entry.sqref, row, col)) continue;
    const rule = entry.rule;
    const numVal = cellNumericValue(cell);

    if (rule.type === 'cellIs') {
      if (numVal == null) continue;
      const args = rule.formulas.map(f => parseFloat(f)).filter(n => !isNaN(n));
      if (cellIsMatch(numVal, rule.operator, args)) {
        applyDxfToResult(result, rule.dxfId != null ? dxfs[rule.dxfId] : null);
      }
    } else if (rule.type === 'top10') {
      if (numVal == null || entry.top10Threshold == null) continue;
      const matches = entry.top10IsTop ? numVal >= entry.top10Threshold : numVal <= entry.top10Threshold;
      if (matches) applyDxfToResult(result, rule.dxfId != null ? dxfs[rule.dxfId] : null);
    } else if (rule.type === 'aboveAverage') {
      if (numVal == null || entry.avgValue == null) continue;
      const matches = entry.avgIsAbove ? numVal > entry.avgValue : numVal < entry.avgValue;
      if (matches) applyDxfToResult(result, rule.dxfId != null ? dxfs[rule.dxfId] : null);
    } else if (rule.type === 'iconSet') {
      if (numVal == null || !entry.iconThresholds?.length) continue;
      const thresholds = entry.iconThresholds;
      const n = thresholds.length;
      let iconIdx = 0;
      for (let i = 1; i < n; i++) {
        if (numVal >= thresholds[i]) iconIdx = i;
      }
      if (rule.reverse) iconIdx = n - 1 - iconIdx;
      result.iconSet = { name: rule.iconSet, index: iconIdx };
    } else if (rule.type === 'colorScale') {
      if (numVal == null || !entry.scaleStops) continue;
      const color = colorScaleAt(numVal, rule.stops, entry.scaleStops);
      result.fill = { patternType: 'solid', fgColor: color, bgColor: color };
    } else if (rule.type === 'dataBar') {
      if (numVal == null || entry.barMin == null || entry.barMax == null) continue;
      const range = entry.barMax - entry.barMin;
      const ratio = range === 0 ? 0 : Math.max(0, Math.min(1, (numVal - entry.barMin) / range));
      result.dataBar = { color: rule.color, ratio };
    }
  }
  return result;
}

// ────────────────────────────────────────────────────────────────
// Shared state for a single renderViewport call
// ────────────────────────────────────────────────────────────────
interface RenderContext {
  worksheet: Worksheet;
  styles: Styles;
  cellMap: Map<string, Cell>;
  mergeAnchorMap: Map<string, { totalW: number; totalH: number }>;
  mergeSkipSet: Set<string>;
  cfContext: CfContext;
  colWidths: number[];
  rowHeights: number[];
  frozenColWidths: number[];
  frozenRowHeights: number[];
  frozenW: number;
  frozenH: number;
  startRow: number;
  startCol: number;
  cs: number;
  dpr: number;
  autoFilterCells: Set<string>;
  hyperlinkMap: Map<string, string>;
}

// ────────────────────────────────────────────────────────────────
// Icon Set drawing
// ────────────────────────────────────────────────────────────────
const ICON_COLORS_3 = ['#FF0000', '#FFFF00', '#00B050'];
const ICON_COLORS_4 = ['#FF0000', '#FF6600', '#FFFF00', '#00B050'];
const ICON_COLORS_5 = ['#FF0000', '#FF6600', '#FFFF00', '#92D050', '#00B050'];

function drawCfIcon(ctx: CanvasRenderingContext2D, name: string, index: number, x: number, y: number, sz: number): void {
  const safeName = name || '3TrafficLights1';
  const nIcons = parseInt(safeName[0]) || 3;
  const palette = nIcons === 5 ? ICON_COLORS_5 : nIcons === 4 ? ICON_COLORS_4 : ICON_COLORS_3;
  const color = palette[Math.max(0, Math.min(index, palette.length - 1))];
  ctx.save();
  ctx.fillStyle = color;
  if (safeName.includes('Arrow')) {
    const half = sz / 2;
    ctx.beginPath();
    if (index === nIcons - 1) {
      ctx.moveTo(x + half, y); ctx.lineTo(x + sz, y + sz); ctx.lineTo(x, y + sz);
    } else if (index === 0) {
      ctx.moveTo(x, y); ctx.lineTo(x + sz, y); ctx.lineTo(x + half, y + sz);
    } else {
      ctx.moveTo(x, y + sz * 0.3); ctx.lineTo(x + sz, y + half); ctx.lineTo(x, y + sz * 0.7);
    }
    ctx.closePath();
    ctx.fill();
  } else if (safeName.includes('Flag')) {
    ctx.beginPath();
    ctx.moveTo(x, y); ctx.lineTo(x + sz, y); ctx.lineTo(x, y + sz);
    ctx.closePath();
    ctx.fill();
  } else {
    ctx.beginPath();
    ctx.arc(x + sz / 2, y + sz / 2, sz / 2, 0, Math.PI * 2);
    ctx.fill();
  }
  ctx.restore();
}

function drawAutoFilterArrow(ctx: CanvasRenderingContext2D, cx: number, cy: number, cw: number, ch: number): void {
  const sz = Math.max(6, Math.round(Math.min(cw, ch) * 0.45));
  const x = cx + cw - sz - 1;
  const y = cy + ch - sz - 1;
  ctx.save();
  ctx.fillStyle = '#D0D0D0';
  ctx.fillRect(x, y, sz, sz);
  ctx.fillStyle = '#444444';
  const tri = sz * 0.55;
  const tx = x + (sz - tri) / 2;
  const ty = y + (sz - tri * 0.5) / 2;
  ctx.beginPath();
  ctx.moveTo(tx, ty);
  ctx.lineTo(tx + tri, ty);
  ctx.lineTo(tx + tri / 2, ty + tri * 0.5);
  ctx.closePath();
  ctx.fill();
  ctx.restore();
}

// ────────────────────────────────────────────────────────────────
// Render one rectangular region of cells
// ────────────────────────────────────────────────────────────────
function renderQuadrant(
  ctx: CanvasRenderingContext2D,
  rc: RenderContext,
  startRow: number, startCol: number,
  colWidths: number[], rowHeights: number[],
  pixOffsetX: number, pixOffsetY: number,
  originX: number, originY: number,
  clipX: number, clipY: number, clipW: number, clipH: number,
): void {
  if (clipW <= 0 || clipH <= 0) return;

  const { styles, cellMap, mergeAnchorMap, mergeSkipSet, cfContext, cs, dpr } = rc;
  const numCols = colWidths.length;
  const numRows = rowHeights.length;

  // Canvas x for each column
  const colXs: number[] = [];
  let x = -pixOffsetX;
  for (let ci = 0; ci < numCols; ci++) {
    colXs.push(x);
    x += colWidths[ci];
  }

  // Canvas y for each row
  const rowYs: number[] = [];
  let y = -pixOffsetY;
  for (let ri = 0; ri < numRows; ri++) {
    rowYs.push(y);
    y += rowHeights[ri];
  }

  ctx.save();
  ctx.beginPath();
  ctx.rect(clipX, clipY, clipW, clipH);
  ctx.clip();

  // Pre-pass: merge cells whose anchor lies outside this viewport quadrant but whose
  // span overlaps it (e.g. scrolled past the anchor row/col, or the anchor is in a
  // frozen quadrant while we are rendering the scrollable quadrant).
  for (const mc of rc.worksheet.mergeCells ?? []) {
    const aRow = mc.top, aCol = mc.left;
    // If anchor is within the main loop range, skip — handled normally below.
    if (aRow >= startRow && aRow < startRow + numRows &&
        aCol >= startCol && aCol < startCol + numCols) continue;
    // Skip if merge span has no overlap with this viewport.
    if (mc.bottom < startRow || mc.top >= startRow + numRows) continue;
    if (mc.right  < startCol || mc.left >= startCol + numCols) continue;

    const info = rc.mergeAnchorMap.get(`${aRow}:${aCol}`);
    if (!info) continue;

    // Canvas X of anchor col (may be negative = off-screen to the left).
    let aCx: number;
    if (aCol >= startCol) {
      aCx = originX + colXs[aCol - startCol];
    } else {
      let dx = 0;
      for (let c = aCol; c < startCol; c++) {
        dx += Math.round(colWidthToPx(rc.worksheet.colWidths[c] ?? rc.worksheet.defaultColWidth) * cs);
      }
      aCx = originX - pixOffsetX - dx;
    }
    // Canvas Y of anchor row (may be negative = off-screen above).
    let aCy: number;
    if (aRow >= startRow) {
      aCy = originY + rowYs[aRow - startRow];
    } else {
      let dy = 0;
      for (let r = aRow; r < startRow; r++) {
        dy += Math.round(rowHeightToPx(rc.worksheet.rowHeights[r] ?? rc.worksheet.defaultRowHeight) * cs);
      }
      aCy = originY - pixOffsetY - dy;
    }

    const cW = info.totalW, cH = info.totalH;
    const key = `${aRow}:${aCol}`;
    const cell = rc.cellMap.get(key);
    const { font, fill, border, xf } = resolveXf(styles, cell?.styleIndex ?? 0);
    const cf = evaluateCf(cell, aRow, aCol, cfContext, styles.dxfs ?? []);
    const effectiveFill = cf.fill ?? fill;

    if (effectiveFill.patternType !== 'none' && effectiveFill.patternType !== '' && effectiveFill.fgColor) {
      ctx.fillStyle = hexToRgba(effectiveFill.fgColor);
      ctx.fillRect(aCx, aCy, cW, cH);
    }
    if (cf.dataBar && cf.dataBar.ratio > 0) {
      const bInset = 2;
      const bW = Math.max(0, (cW - bInset * 2) * cf.dataBar.ratio);
      if (bW > 0) {
        ctx.fillStyle = hexToRgba(cf.dataBar.color, 0.6);
        ctx.fillRect(aCx + bInset, aCy + bInset, bW, cH - bInset * 2);
      }
    }
    renderBorder(ctx, border, aCx, aCy, cW, cH);

    if (!cell) continue;
    const text = formatCellValue(cell, styles);
    if (!text || (text === '0' && rc.worksheet.showZeros === false)) continue;

    const effectiveBold = font.bold || !!cf.fontBold;
    const effectiveItalic = font.italic || !!cf.fontItalic;
    const effectiveUnderline = font.underline || !!cf.fontUnderline;
    const effectiveStrike = font.strike || !!cf.fontStrike;
    const fontForDraw: Font = (
      effectiveBold !== font.bold || effectiveItalic !== font.italic ||
      effectiveUnderline !== font.underline || effectiveStrike !== font.strike
    ) ? { ...font, bold: effectiveBold, italic: effectiveItalic, underline: effectiveUnderline, strike: effectiveStrike }
      : font;
    ctx.font = buildFont(fontForDraw, cs);
    const hyperlinkUrl = rc.hyperlinkMap.get(key);
    const textColor = hyperlinkUrl ? '#0563C1' : (cf.fontColor ?? font.color);
    ctx.fillStyle = textColor ? hexToRgba(textColor) : '#000000';

    const paddingX = 3, paddingY = 2;
    const isNumeric = cell.value.type === 'number';
    const alignH = xf.alignH ?? (isNumeric ? 'right' : 'left');
    const alignV = xf.alignV ?? 'bottom';
    const indentPx = xf.indent ? Math.round(xf.indent * font.size * ROW_HEIGHT_TO_PX * 0.5) : 0;
    const leftPad = paddingX + (alignH === 'left' || !xf.alignH ? indentPx : 0);

    ctx.save();
    ctx.beginPath();
    ctx.rect(aCx, aCy, cW, cH);
    ctx.clip();

    let textX: number;
    if (alignH === 'right') { textX = aCx + cW - paddingX; ctx.textAlign = 'right'; }
    else if (alignH === 'center') { textX = aCx + cW / 2; ctx.textAlign = 'center'; }
    else { textX = aCx + leftPad; ctx.textAlign = 'left'; }

    let textY: number;
    if (alignV === 'top') { ctx.textBaseline = 'top'; textY = aCy + paddingY; }
    else if (alignV === 'center') { ctx.textBaseline = 'middle'; textY = aCy + cH / 2; }
    else { ctx.textBaseline = 'bottom'; textY = aCy + cH - paddingY; }

    ctx.fillText(text, textX, textY);
    ctx.restore();
  }

  for (let ri = 0; ri < numRows; ri++) {
    const rowIndex = startRow + ri;
    const cy = originY + rowYs[ri];
    const ch = rowHeights[ri];
    if (cy + ch <= clipY || cy >= clipY + clipH) continue;

    for (let ci = 0; ci < numCols; ci++) {
      const colIndex = startCol + ci;
      const cx = originX + colXs[ci];
      const cw = colWidths[ci];
      if (cx + cw <= clipX || cx >= clipX + clipW) continue;

      const key = `${rowIndex}:${colIndex}`;
      if (mergeSkipSet.has(key)) continue;

      const mergeInfo = mergeAnchorMap.get(key);
      const cellW = mergeInfo ? mergeInfo.totalW : cw;
      const cellH = mergeInfo ? mergeInfo.totalH : ch;

      const cell = cellMap.get(key);
      const styleIndex = cell?.styleIndex ?? 0;
      const { font, fill, border, xf } = resolveXf(styles, styleIndex);
      const cf = evaluateCf(cell, rowIndex, colIndex, cfContext, styles.dxfs ?? []);
      const effectiveFill = cf.fill ?? fill;

      // Background fill (base or CF override)
      if (effectiveFill.patternType !== 'none' && effectiveFill.patternType !== '' && effectiveFill.fgColor) {
        ctx.fillStyle = hexToRgba(effectiveFill.fgColor);
        ctx.fillRect(cx, cy, cellW, cellH);
      }

      // DataBar (drawn inside the cell, left-anchored)
      if (cf.dataBar && cf.dataBar.ratio > 0) {
        const barInset = 2;
        const barW = Math.max(0, (cellW - barInset * 2) * cf.dataBar.ratio);
        if (barW > 0) {
          ctx.fillStyle = hexToRgba(cf.dataBar.color, 0.6);
          ctx.fillRect(cx + barInset, cy + barInset, barW, cellH - barInset * 2);
        }
      }

      // Grid lines – draw only right + bottom edges once per cell (avoids double-drawing at
      // shared cell boundaries). Half-device-pixel offset (0.5/dpr) aligns each line to the
      // device pixel grid so we get a crisp 1-device-pixel result.
      {
        const hp = 0.5 / dpr;
        ctx.strokeStyle = '#d0d0d0';
        ctx.lineWidth = 0.5;
        ctx.beginPath();
        ctx.moveTo(cx + cellW + hp, cy);          // right edge
        ctx.lineTo(cx + cellW + hp, cy + cellH);
        ctx.moveTo(cx, cy + cellH + hp);           // bottom edge
        ctx.lineTo(cx + cellW, cy + cellH + hp);
        if (ri === 0) {                            // top edge for first row
          ctx.moveTo(cx, cy + hp);
          ctx.lineTo(cx + cellW, cy + hp);
        }
        if (ci === 0) {                            // left edge for first column
          ctx.moveTo(cx + hp, cy);
          ctx.lineTo(cx + hp, cy + cellH);
        }
        ctx.stroke();
      }

      // Cell borders
      renderBorder(ctx, border, cx, cy, cellW, cellH);

      // AutoFilter dropdown indicator
      if (rc.autoFilterCells.has(key)) {
        drawAutoFilterArrow(ctx, cx, cy, cw, cellH);
      }

      if (!cell) continue;
      const text = formatCellValue(cell, styles);
      if (!text || (text === '0' && rc.worksheet.showZeros === false)) continue;

      const effectiveBold = font.bold || !!cf.fontBold;
      const effectiveItalic = font.italic || !!cf.fontItalic;
      const effectiveUnderline = font.underline || !!cf.fontUnderline;
      const effectiveStrike = font.strike || !!cf.fontStrike;
      const fontForDraw: Font = (
        effectiveBold !== font.bold || effectiveItalic !== font.italic ||
        effectiveUnderline !== font.underline || effectiveStrike !== font.strike
      )
        ? { ...font, bold: effectiveBold, italic: effectiveItalic, underline: effectiveUnderline, strike: effectiveStrike }
        : font;
      ctx.font = buildFont(fontForDraw, cs);
      const hyperlinkUrl = rc.hyperlinkMap.get(key);
      const textColor = hyperlinkUrl ? '#0563C1' : (cf.fontColor ?? font.color);
      ctx.fillStyle = textColor ? hexToRgba(textColor) : '#000000';

      const paddingX = 3;
      const paddingY = 2;
      const isNumeric = cell.value.type === 'number';
      const alignH = xf.alignH ?? (isNumeric ? 'right' : 'left');
      const alignV = xf.alignV ?? 'bottom';
      // Indent: each level ≈ one character width (ECMA-376 §18.8.44)
      const indentPx = xf.indent ? Math.round(xf.indent * font.size * ROW_HEIGHT_TO_PX * 0.5) : 0;
      // IconSet: reserve space on the left for the icon
      const iconSz = cf.iconSet ? Math.max(8, Math.round(Math.min(cellW, cellH) * 0.55)) : 0;
      const iconPad = iconSz > 0 ? iconSz + 4 : 0;
      const leftPad = paddingX + (alignH === 'left' || !xf.alignH ? indentPx : 0) + iconPad;

      // Text overflow for left-aligned non-merged non-wrap cells
      let drawW = cellW;
      if (!mergeInfo && !xf.wrapText && alignH !== 'right' && alignH !== 'center') {
        const textPx = ctx.measureText(text).width + leftPad + paddingX;
        if (textPx > cellW) {
          for (let oci = ci + 1; oci < numCols; oci++) {
            const adjKey = `${rowIndex}:${startCol + oci}`;
            if (mergeSkipSet.has(adjKey) || mergeAnchorMap.has(adjKey)) break;
            const adjCell = cellMap.get(adjKey);
            if (adjCell && adjCell.value.type !== 'empty') break;
            drawW += colWidths[oci];
          }
        }
      }

      let textX: number;
      let textAlign: CanvasTextAlign;
      if (alignH === 'right') {
        textX = cx + cellW - paddingX;
        textAlign = 'right';
      } else if (alignH === 'center') {
        textX = cx + cellW / 2;
        textAlign = 'center';
      } else {
        textX = cx + leftPad;
        textAlign = 'left';
      }

      const rotation = xf.textRotation ?? 0;
      const isStacked = rotation === 255;
      const isRotated = rotation > 0 && rotation !== 255;

      // Draw icon set icon (before text clip block)
      if (cf.iconSet && iconSz > 0) {
        ctx.save();
        ctx.beginPath();
        ctx.rect(cx, cy, cellW, cellH);
        ctx.clip();
        drawCfIcon(ctx, cf.iconSet.name, cf.iconSet.index, cx + 2, cy + (cellH - iconSz) / 2, iconSz);
        ctx.restore();
      }

      ctx.save();
      ctx.beginPath();
      ctx.rect(cx, cy, drawW, cellH);
      ctx.clip();

      // Stacked text (textRotation=255): draw each character on its own line
      if (isStacked) {
        const charH = Math.round(font.size * ROW_HEIGHT_TO_PX * 1.1);
        const totalH = text.length * charH;
        let charY = alignV === 'top' ? cy + paddingY
          : alignV === 'center' ? cy + (cellH - totalH) / 2
          : cy + cellH - totalH - paddingY;
        ctx.textAlign = 'center';
        ctx.textBaseline = 'top';
        for (const ch of text) {
          ctx.fillText(ch, cx + cellW / 2, charY);
          charY += charH;
        }
        ctx.restore();
        continue;
      }

      // Rotated text: translate to cell center, rotate, draw, restore
      if (isRotated) {
        const angleRad = rotation <= 90
          ? -(rotation * Math.PI / 180)
          : ((rotation - 90) * Math.PI / 180);
        ctx.translate(cx + cellW / 2, cy + cellH / 2);
        ctx.rotate(angleRad);
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText(text, 0, 0);
        ctx.restore();
        continue;
      }

      // shrinkToFit: scale context horizontally if text is wider than cell
      if (xf.shrinkToFit) {
        const textW = ctx.measureText(text).width;
        const availW = cellW - leftPad - paddingX;
        if (textW > availW && textW > 0) {
          const scale = availW / textW;
          const pivotX = alignH === 'right' ? cx + cellW - paddingX
            : alignH === 'center' ? cx + cellW / 2
            : cx + leftPad;
          ctx.transform(scale, 0, 0, 1, pivotX * (1 - scale), 0);
        }
      }

      ctx.textAlign = textAlign;

      // Rich text: draw each run with its own font. Only supported for the
      // non-wrap path (wrap with mixed fonts is significantly more complex).
      const runs = cell.value.type === 'text' ? cell.value.runs : undefined;
      const hasRichText = runs && runs.length > 0;

      if (xf.wrapText && hasRichText) {
        // Rich text with wrapping: per-run fonts, break on spaces and CJK boundaries
        const wrapW = cellW - leftPad - paddingX;
        const rLines = layoutRichTextLines(ctx, runs, fontForDraw, cs, wrapW);
        const totalH = rLines.reduce((s, l) => s + Math.round(l.maxFontSize * ROW_HEIGHT_TO_PX * 1.2), 0);
        let yy: number;
        if (alignV === 'top') yy = cy + paddingY;
        else if (alignV === 'center') yy = cy + (cellH - totalH) / 2;
        else yy = cy + cellH - totalH - paddingY;
        ctx.textAlign = 'left';
        ctx.textBaseline = 'top';
        for (const line of rLines) {
          const lineH = Math.round(line.maxFontSize * ROW_HEIGHT_TO_PX * 1.2);
          const totalW = line.segments.reduce((s, seg) => s + seg.width, 0);
          let xx: number;
          if (alignH === 'right') xx = cx + cellW - paddingX - totalW;
          else if (alignH === 'center') xx = cx + cellW / 2 - totalW / 2;
          else xx = cx + leftPad;
          for (const seg of line.segments) {
            ctx.font = buildFont(seg.font, cs);
            const segColor = cf.fontColor ?? seg.font.color;
            ctx.fillStyle = segColor ? hexToRgba(segColor) : '#000000';
            ctx.fillText(seg.text, xx, yy);
            const rSizePx = Math.round(seg.font.size * ROW_HEIGHT_TO_PX);
            if (seg.font.underline) {
              ctx.save();
              ctx.strokeStyle = segColor ? hexToRgba(segColor) : '#000000';
              ctx.lineWidth = 0.5;
              ctx.beginPath(); ctx.moveTo(xx, yy + rSizePx + 1); ctx.lineTo(xx + seg.width, yy + rSizePx + 1); ctx.stroke();
              ctx.restore();
            }
            if (seg.font.strike) {
              ctx.save();
              ctx.strokeStyle = segColor ? hexToRgba(segColor) : '#000000';
              ctx.lineWidth = 0.5;
              const sy2 = yy + Math.round(rSizePx * 0.5);
              ctx.beginPath(); ctx.moveTo(xx, sy2); ctx.lineTo(xx + seg.width, sy2); ctx.stroke();
              ctx.restore();
            }
            xx += seg.width;
          }
          yy += lineH;
        }
      } else if (xf.wrapText) {
        const lines = wrapTextLines(ctx, text, cellW - leftPad - paddingX);
        const lineH = Math.round(font.size * ROW_HEIGHT_TO_PX * 1.2);
        const totalTextH = lines.length * lineH;
        let startY: number;
        if (alignV === 'top') { startY = cy + paddingY; ctx.textBaseline = 'top'; }
        else if (alignV === 'center') { startY = cy + (cellH - totalTextH) / 2; ctx.textBaseline = 'top'; }
        else { startY = cy + cellH - totalTextH - paddingY; ctx.textBaseline = 'top'; }
        for (let li = 0; li < lines.length; li++) {
          ctx.fillText(lines[li], textX, startY + li * lineH);
        }
      } else if (hasRichText) {
        // Per-run drawing: compute font for each run, measure widths, draw LTR
        const runFonts = runs.map(r => applyRunFont(fontForDraw, r));
        const runWidths: number[] = runs.map((r, i) => {
          ctx.font = buildFont(runFonts[i], cs);
          return ctx.measureText(r.text).width;
        });
        const totalWidth = runWidths.reduce((a, b) => a + b, 0);
        let startX: number;
        if (alignH === 'right') startX = cx + cellW - paddingX - totalWidth;
        else if (alignH === 'center') startX = cx + cellW / 2 - totalWidth / 2;
        else startX = cx + leftPad;
        // Use left alignment since we position each run ourselves
        ctx.textAlign = 'left';
        let textY: number;
        if (alignV === 'top') { ctx.textBaseline = 'top'; textY = cy + paddingY; }
        else if (alignV === 'center') { ctx.textBaseline = 'middle'; textY = cy + cellH / 2; }
        else { ctx.textBaseline = 'bottom'; textY = cy + cellH - paddingY; }
        let runX = startX;
        for (let i = 0; i < runs.length; i++) {
          const rf = runFonts[i];
          ctx.font = buildFont(rf, cs);
          const runColor = cf.fontColor ?? rf.color;
          ctx.fillStyle = runColor ? hexToRgba(runColor) : '#000000';
          ctx.fillText(runs[i].text, runX, textY);
          const rSizePx = Math.round(rf.size * ROW_HEIGHT_TO_PX);
          if (rf.underline) {
            const uy = alignV === 'top'
              ? cy + paddingY + rSizePx + 1
              : alignV === 'center'
                ? cy + cellH / 2 + Math.round(rSizePx * 0.55)
                : cy + cellH - paddingY + 1;
            ctx.save();
            ctx.strokeStyle = runColor ? hexToRgba(runColor) : '#000000';
            ctx.lineWidth = 0.5;
            ctx.beginPath(); ctx.moveTo(runX, uy); ctx.lineTo(runX + runWidths[i], uy); ctx.stroke();
            ctx.restore();
          }
          if (rf.strike) {
            const sy = alignV === 'top'
              ? cy + paddingY + Math.round(rSizePx * 0.5)
              : alignV === 'center'
                ? cy + cellH / 2
                : cy + cellH - paddingY - Math.round(rSizePx * 0.35);
            ctx.save();
            ctx.strokeStyle = runColor ? hexToRgba(runColor) : '#000000';
            ctx.lineWidth = 0.5;
            ctx.beginPath(); ctx.moveTo(runX, sy); ctx.lineTo(runX + runWidths[i], sy); ctx.stroke();
            ctx.restore();
          }
          runX += runWidths[i];
        }
      } else {
        // Measure once for both underline and strike
        let overlayMetrics: TextMetrics | null = null;
        const measureOverlay = () => overlayMetrics ??= ctx.measureText(text);
        const overlayX = () => {
          const tW = Math.min(measureOverlay().width, drawW - leftPad - paddingX);
          return {
            x: alignH === 'right' ? cx + cellW - paddingX - tW
              : alignH === 'center' ? cx + cellW / 2 - tW / 2
              : cx + leftPad,
            width: tW,
          };
        };
        const sizePx = Math.round(font.size * ROW_HEIGHT_TO_PX);

        if (fontForDraw.underline || hyperlinkUrl) {
          const { x: ux, width: tW } = overlayX();
          const uy = alignV === 'top'
            ? cy + paddingY + sizePx + 1
            : alignV === 'center'
              ? cy + cellH / 2 + Math.round(sizePx * 0.55)
              : cy + cellH - paddingY + 1;
          ctx.save();
          ctx.strokeStyle = hyperlinkUrl ? '#0563C1' : (textColor ? hexToRgba(textColor) : '#000000');
          ctx.lineWidth = 0.5;
          ctx.beginPath(); ctx.moveTo(ux, uy); ctx.lineTo(ux + tW, uy); ctx.stroke();
          ctx.restore();
        }
        if (fontForDraw.strike) {
          const { x: sx, width: tW } = overlayX();
          // Strike line sits roughly at the x-height mid-line (~45% up from baseline)
          const sy = alignV === 'top'
            ? cy + paddingY + Math.round(sizePx * 0.5)
            : alignV === 'center'
              ? cy + cellH / 2
              : cy + cellH - paddingY - Math.round(sizePx * 0.35);
          ctx.save();
          ctx.strokeStyle = textColor ? hexToRgba(textColor) : '#000000';
          ctx.lineWidth = 0.5;
          ctx.beginPath(); ctx.moveTo(sx, sy); ctx.lineTo(sx + tW, sy); ctx.stroke();
          ctx.restore();
        }
        let textY: number;
        if (alignV === 'top') { ctx.textBaseline = 'top'; textY = cy + paddingY; }
        else if (alignV === 'center') { ctx.textBaseline = 'middle'; textY = cy + cellH / 2; }
        else { ctx.textBaseline = 'bottom'; textY = cy + cellH - paddingY; }
        ctx.fillText(text, textX, textY);
      }

      ctx.restore();
    }
  }

  ctx.restore();
}

// ────────────────────────────────────────────────────────────────
// Main render function
// ────────────────────────────────────────────────────────────────
export function renderViewport(
  ctx: CanvasRenderingContext2D,
  worksheet: Worksheet,
  styles: Styles,
  viewport: ViewportRange,
  opts: RenderViewportOptions = {},
): void {
  const dpr = opts.dpr ?? 1;
  const cs = opts.cellScale ?? 1;
  const canvasW = ctx.canvas.width / dpr;
  const canvasH = ctx.canvas.height / dpr;

  ctx.clearRect(0, 0, canvasW, canvasH);
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, canvasW, canvasH);

  // Scaled pixel helper: apply cellScale to all cell/header dimensions
  const sp = (px: number) => Math.round(px * cs);
  const hw = sp(HEADER_W);  // scaled header column width
  const hh = sp(HEADER_H);  // scaled header row height

  const { row: startRow, col: startCol, rows: numRows, cols: numCols } = viewport;
  const scrollOffsetX = (opts.scrollOffsetX ?? 0) * cs;
  const scrollOffsetY = (opts.scrollOffsetY ?? 0) * cs;
  const freezeRows = opts.freezeRows ?? 0;
  const freezeCols = opts.freezeCols ?? 0;

  // ── Compute frozen area pixel sizes (scaled) ─────────────────
  const frozenColWidths: number[] = [];
  for (let c = 1; c <= freezeCols; c++) {
    frozenColWidths.push(sp(colWidthToPx(worksheet.colWidths[c] ?? worksheet.defaultColWidth)));
  }
  const frozenRowHeights: number[] = [];
  for (let r = 1; r <= freezeRows; r++) {
    frozenRowHeights.push(sp(rowHeightToPx(worksheet.rowHeights[r] ?? worksheet.defaultRowHeight)));
  }
  const frozenW = frozenColWidths.reduce((s, w) => s + w, 0);
  const frozenH = frozenRowHeights.reduce((s, h) => s + h, 0);

  // ── Scrollable col/row pixel widths (scaled) ─────────────────
  const scrollColWidths: number[] = [];
  for (let c = startCol; c < startCol + numCols; c++) {
    scrollColWidths.push(sp(colWidthToPx(worksheet.colWidths[c] ?? worksheet.defaultColWidth)));
  }
  const scrollRowHeights: number[] = [];
  for (let r = startRow; r < startRow + numRows; r++) {
    scrollRowHeights.push(sp(rowHeightToPx(worksheet.rowHeights[r] ?? worksheet.defaultRowHeight)));
  }

  // ── Build cell & merge lookup ────────────────────────────────
  const cellMap = new Map<string, Cell>();
  for (const row of worksheet.rows) {
    for (const cell of row.cells) {
      cellMap.set(`${cell.row}:${cell.col}`, cell);
    }
  }

  const mergeAnchorMap = new Map<string, { totalW: number; totalH: number }>();
  const mergeSkipSet = new Set<string>();
  for (const mc of worksheet.mergeCells ?? []) {
    let totalW = 0;
    for (let c = mc.left; c <= mc.right; c++) {
      totalW += sp(colWidthToPx(worksheet.colWidths[c] ?? worksheet.defaultColWidth));
    }
    let totalH = 0;
    for (let r = mc.top; r <= mc.bottom; r++) {
      totalH += sp(rowHeightToPx(worksheet.rowHeights[r] ?? worksheet.defaultRowHeight));
    }
    mergeAnchorMap.set(`${mc.top}:${mc.left}`, { totalW, totalH });
    for (let r = mc.top; r <= mc.bottom; r++) {
      for (let c = mc.left; c <= mc.right; c++) {
        if (r === mc.top && c === mc.left) continue;
        mergeSkipSet.add(`${r}:${c}`);
      }
    }
  }

  const cfContext = compileCf(worksheet);

  // Build autoFilter indicator cell set
  const autoFilterCells = new Set<string>();
  if (worksheet.autoFilter) {
    const af = worksheet.autoFilter;
    for (let c = af.left; c <= af.right; c++) {
      autoFilterCells.add(`${af.top}:${c}`);
    }
  }

  // Build hyperlink lookup map
  const hyperlinkMap = new Map<string, string>();
  for (const hl of worksheet.hyperlinks ?? []) {
    if (hl.url) hyperlinkMap.set(`${hl.row}:${hl.col}`, hl.url);
  }

  const rc: RenderContext = {
    worksheet, styles, cellMap, mergeAnchorMap, mergeSkipSet, cfContext,
    colWidths: scrollColWidths,
    rowHeights: scrollRowHeights,
    frozenColWidths, frozenRowHeights,
    frozenW, frozenH,
    startRow, startCol,
    cs,
    dpr,
    autoFilterCells,
    hyperlinkMap,
  };

  // Canvas areas for each quadrant
  const cellAreaX = hw;
  const cellAreaY = hh;
  const scrollAreaX = cellAreaX + frozenW;
  const scrollAreaY = cellAreaY + frozenH;
  const scrollAreaW = Math.max(0, canvasW - scrollAreaX);
  const scrollAreaH = Math.max(0, canvasH - scrollAreaY);

  // ── Q1: frozen rows × frozen cols ───────────────────────────
  if (freezeRows > 0 && freezeCols > 0) {
    renderQuadrant(ctx, rc,
      1, 1, frozenColWidths, frozenRowHeights,
      0, 0,
      cellAreaX, cellAreaY,
      cellAreaX, cellAreaY, frozenW, frozenH,
    );
  }

  // ── Q2: frozen rows × scrollable cols ───────────────────────
  if (freezeRows > 0) {
    renderQuadrant(ctx, rc,
      1, startCol, scrollColWidths, frozenRowHeights,
      scrollOffsetX, 0,
      scrollAreaX, cellAreaY,
      scrollAreaX, cellAreaY, scrollAreaW, frozenH,
    );
  }

  // ── Q3: scrollable rows × frozen cols ───────────────────────
  if (freezeCols > 0) {
    renderQuadrant(ctx, rc,
      startRow, 1, frozenColWidths, scrollRowHeights,
      0, scrollOffsetY,
      cellAreaX, scrollAreaY,
      cellAreaX, scrollAreaY, frozenW, scrollAreaH,
    );
  }

  // ── Q4: scrollable rows × scrollable cols (main area) ───────
  renderQuadrant(ctx, rc,
    startRow, startCol, scrollColWidths, scrollRowHeights,
    scrollOffsetX, scrollOffsetY,
    scrollAreaX, scrollAreaY,
    scrollAreaX, scrollAreaY, scrollAreaW, scrollAreaH,
  );

  // ── Anchored images (clipped to scrollable area) ─────────────
  if (worksheet.images && worksheet.images.length > 0 && opts.loadedImages) {
    renderImages(
      ctx, worksheet, opts.loadedImages, cs,
      startRow, startCol,
      scrollOffsetX, scrollOffsetY,
      scrollAreaX, scrollAreaY,
      scrollAreaW, scrollAreaH,
    );
  }

  // ── Anchored charts (clipped to scrollable area) ──────────────
  if (worksheet.charts && worksheet.charts.length > 0) {
    renderCharts(
      ctx, worksheet, cs,
      startRow, startCol,
      scrollOffsetX, scrollOffsetY,
      scrollAreaX, scrollAreaY,
      scrollAreaW, scrollAreaH,
    );
  }

  // ── Row/col headers (drawn last, always on top) ──────────────
  renderHeaders(ctx, canvasW, canvasH,
    startRow, startCol, numRows, numCols,
    scrollColWidths, scrollRowHeights,
    scrollOffsetX, scrollOffsetY,
    frozenColWidths, frozenRowHeights,
    frozenW, frozenH,
    hw, hh, cs, dpr,
  );

  // ── Freeze pane separator lines ──────────────────────────────
  if (freezeRows > 0) {
    ctx.save();
    ctx.strokeStyle = FREEZE_LINE_COLOR;
    ctx.lineWidth = 0.5;
    ctx.beginPath();
    ctx.moveTo(hw, scrollAreaY + 0.5);
    ctx.lineTo(canvasW, scrollAreaY + 0.5);
    ctx.stroke();
    ctx.restore();
  }
  if (freezeCols > 0) {
    ctx.save();
    ctx.strokeStyle = FREEZE_LINE_COLOR;
    ctx.lineWidth = 0.5;
    ctx.beginPath();
    ctx.moveTo(scrollAreaX + 0.5, hh);
    ctx.lineTo(scrollAreaX + 0.5, canvasH);
    ctx.stroke();
    ctx.restore();
  }
}

// ────────────────────────────────────────────────────────────────
// Headers
// ────────────────────────────────────────────────────────────────
function renderHeaders(
  ctx: CanvasRenderingContext2D,
  canvasW: number, canvasH: number,
  startRow: number, startCol: number,
  numRows: number, numCols: number,
  scrollColWidths: number[], scrollRowHeights: number[],
  scrollOffsetX: number, scrollOffsetY: number,
  frozenColWidths: number[], frozenRowHeights: number[],
  frozenW: number, frozenH: number,
  hw: number, hh: number, cs: number, dpr: number,
): void {
  const HEADER_BG = '#f8f9fa';
  const HEADER_BORDER = '#c8ccd0';
  const HEADER_TEXT = '#444';
  const headerFontSize = Math.max(1, Math.round(11 * cs));
  const HEADER_FONT = `${headerFontSize}px ${DEFAULT_FONT_FAMILY}`;
  const scrollAreaX = hw + frozenW;
  const scrollAreaY = hh + frozenH;
  const hp = 0.5 / dpr;  // half device-pixel offset for 1dp crisp lines

  // Corner – draw all 4 edges (standalone box)
  ctx.fillStyle = HEADER_BG;
  ctx.fillRect(0, 0, hw, hh);
  ctx.strokeStyle = HEADER_BORDER;
  ctx.lineWidth = 0.5;
  ctx.beginPath();
  ctx.moveTo(hp, 0); ctx.lineTo(hp, hh);          // left  (outward — canvas edge)
  ctx.moveTo(0, hp); ctx.lineTo(hw, hp);            // top   (outward — canvas edge)
  ctx.moveTo(hw - hp, 0); ctx.lineTo(hw - hp, hh);  // right  (inset — aligns with row-header right)
  ctx.moveTo(0, hh - hp); ctx.lineTo(hw, hh - hp);  // bottom (inset — aligns with col-header bottom)
  ctx.stroke();

  ctx.font = HEADER_FONT;
  ctx.fillStyle = HEADER_TEXT;

  // Helper: draw one column header cell.
  // Borders are drawn INSET (-hp) so that the next cell's fillRect (which starts at cx+cw)
  // never overwrites the current cell's right/bottom border line.
  const drawColHeader = (col: number, cx: number, cw: number) => {
    ctx.fillStyle = HEADER_BG;
    ctx.fillRect(cx, 0, cw, hh);
    ctx.strokeStyle = HEADER_BORDER;
    ctx.lineWidth = 0.5;
    ctx.beginPath();
    ctx.moveTo(cx + cw - hp, 0);     ctx.lineTo(cx + cw - hp, hh);  // right (inset)
    ctx.moveTo(cx, hh - hp);          ctx.lineTo(cx + cw, hh - hp);  // bottom (inset)
    ctx.moveTo(cx, hp);               ctx.lineTo(cx + cw, hp);        // top
    ctx.stroke();
    ctx.fillStyle = HEADER_TEXT;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText(colToLetter(col), cx + cw / 2, hh / 2);
  };

  // Helper: draw one row header cell.
  // Borders drawn inset so adjacent cell's fill never overwrites them.
  const drawRowHeader = (row: number, cy: number, ch: number) => {
    ctx.fillStyle = HEADER_BG;
    ctx.fillRect(0, cy, hw, ch);
    ctx.strokeStyle = HEADER_BORDER;
    ctx.lineWidth = 0.5;
    ctx.beginPath();
    ctx.moveTo(hw - hp, cy);  ctx.lineTo(hw - hp, cy + ch);   // right (inset)
    ctx.moveTo(0, cy + ch - hp); ctx.lineTo(hw, cy + ch - hp); // bottom (inset)
    ctx.moveTo(hp, cy);       ctx.lineTo(hp, cy + ch);          // left
    ctx.stroke();
    ctx.fillStyle = HEADER_TEXT;
    ctx.textAlign = 'right';
    ctx.textBaseline = 'middle';
    ctx.fillText(String(row), hw - Math.max(2, Math.round(4 * cs)), cy + ch / 2);
  };

  // Frozen col headers (no h-scroll, fixed positions)
  if (frozenColWidths.length > 0) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(hw, 0, frozenW, hh);
    ctx.clip();
    let cx = hw;
    for (let ci = 0; ci < frozenColWidths.length; ci++) {
      drawColHeader(ci + 1, cx, frozenColWidths[ci]);
      cx += frozenColWidths[ci];
    }
    ctx.restore();
  }

  // Scrollable col headers
  ctx.save();
  ctx.beginPath();
  ctx.rect(scrollAreaX, 0, canvasW - scrollAreaX, hh);
  ctx.clip();
  let cx = scrollAreaX - scrollOffsetX;
  for (let ci = 0; ci < scrollColWidths.length; ci++) {
    const cw = scrollColWidths[ci];
    if (cx + cw > scrollAreaX && cx < canvasW) {
      drawColHeader(startCol + ci, cx, cw);
    }
    cx += cw;
  }
  ctx.restore();

  // Frozen row headers (no v-scroll)
  if (frozenRowHeights.length > 0) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(0, hh, hw, frozenH);
    ctx.clip();
    let cy = hh;
    for (let ri = 0; ri < frozenRowHeights.length; ri++) {
      drawRowHeader(ri + 1, cy, frozenRowHeights[ri]);
      cy += frozenRowHeights[ri];
    }
    ctx.restore();
  }

  // Scrollable row headers
  ctx.save();
  ctx.beginPath();
  ctx.rect(0, scrollAreaY, hw, canvasH - scrollAreaY);
  ctx.clip();
  let cy = scrollAreaY - scrollOffsetY;
  for (let ri = 0; ri < scrollRowHeights.length; ri++) {
    const ch = scrollRowHeights[ri];
    if (cy + ch > scrollAreaY && cy < canvasH) {
      drawRowHeader(startRow + ri, cy, ch);
    }
    cy += ch;
  }
  ctx.restore();

  // Cover frozen area corner (where row/col headers meet)
  if (frozenW > 0 || frozenH > 0) {
    ctx.fillStyle = HEADER_BG;
    if (frozenW > 0) {
      ctx.fillRect(0, hh, hw, frozenH);
    }
    if (frozenH > 0) {
      ctx.fillRect(hw, 0, frozenW, hh);
    }
  }
}

// ────────────────────────────────────────────────────────────────
// Image anchors  (ECMA-376 §20.5, <xdr:twoCellAnchor>)
// ────────────────────────────────────────────────────────────────
const EMU_PER_PX = 9525;

/** Sum scaled column widths for cols 1..n-1 (sheet-space X of col n in scaled px). */
function sheetXForCol(
  ws: Worksheet,
  col1: number, // 1-indexed column number
  cs: number,
): number {
  let x = 0;
  for (let c = 1; c < col1; c++) {
    x += Math.round(colWidthToPx(ws.colWidths[c] ?? ws.defaultColWidth) * cs);
  }
  return x;
}

/** Sum scaled row heights for rows 1..n-1 (sheet-space Y of row n in scaled px). */
function sheetYForRow(
  ws: Worksheet,
  row1: number, // 1-indexed row number
  cs: number,
): number {
  let y = 0;
  for (let r = 1; r < row1; r++) {
    y += Math.round(rowHeightToPx(ws.rowHeights[r] ?? ws.defaultRowHeight) * cs);
  }
  return y;
}

function renderImages(
  ctx: CanvasRenderingContext2D,
  ws: Worksheet,
  loadedImages: Map<string, HTMLImageElement>,
  cs: number,
  startRow: number,
  startCol: number,
  scrollOffsetX: number,
  scrollOffsetY: number,
  scrollAreaX: number,
  scrollAreaY: number,
  scrollAreaW: number,
  scrollAreaH: number,
): void {
  if (scrollAreaW <= 0 || scrollAreaH <= 0) return;

  // Sheet-space origin of the current scroll viewport's first visible cell
  const scrollOriginSheetX = sheetXForCol(ws, startCol, cs);
  const scrollOriginSheetY = sheetYForRow(ws, startRow, cs);

  ctx.save();
  ctx.beginPath();
  ctx.rect(scrollAreaX, scrollAreaY, scrollAreaW, scrollAreaH);
  ctx.clip();

  for (const anchor of ws.images) {
    const img = loadedImages.get(anchor.dataUrl);
    if (!img) continue;

    // xdr col/row are 0-indexed; our widths map is 1-indexed.
    const fromCol1 = anchor.fromCol + 1;
    const fromRow1 = anchor.fromRow + 1;
    const toCol1   = anchor.toCol   + 1;
    const toRow1   = anchor.toRow   + 1;

    // Image sheet-space rect (in scaled px)
    const imgSheetX1 = sheetXForCol(ws, fromCol1, cs) + (anchor.fromColOff * cs) / EMU_PER_PX;
    const imgSheetY1 = sheetYForRow(ws, fromRow1, cs) + (anchor.fromRowOff * cs) / EMU_PER_PX;
    const imgSheetX2 = sheetXForCol(ws, toCol1,   cs) + (anchor.toColOff   * cs) / EMU_PER_PX;
    const imgSheetY2 = sheetYForRow(ws, toRow1,   cs) + (anchor.toRowOff   * cs) / EMU_PER_PX;

    const imgW = imgSheetX2 - imgSheetX1;
    const imgH = imgSheetY2 - imgSheetY1;
    if (imgW <= 0 || imgH <= 0) continue;

    // Translate to canvas coordinates of the scrollable viewport
    const canvasX = scrollAreaX + (imgSheetX1 - scrollOriginSheetX) - scrollOffsetX;
    const canvasY = scrollAreaY + (imgSheetY1 - scrollOriginSheetY) - scrollOffsetY;

    // Early out when entirely off-screen
    if (canvasX + imgW < scrollAreaX || canvasX > scrollAreaX + scrollAreaW) continue;
    if (canvasY + imgH < scrollAreaY || canvasY > scrollAreaY + scrollAreaH) continue;

    ctx.drawImage(img, canvasX, canvasY, imgW, imgH);
  }

  ctx.restore();
}

// ────────────────────────────────────────────────────────────────
// Border drawing
// ────────────────────────────────────────────────────────────────
function renderBorder(ctx: CanvasRenderingContext2D, border: Border, x: number, y: number, w: number, h: number): void {
  const edges: Array<{ edge: BorderEdge | null | undefined; x1: number; y1: number; x2: number; y2: number }> = [
    { edge: border.top,         x1: x,     y1: y,     x2: x + w, y2: y },
    { edge: border.bottom,      x1: x,     y1: y + h, x2: x + w, y2: y + h },
    { edge: border.left,        x1: x,     y1: y,     x2: x,     y2: y + h },
    { edge: border.right,       x1: x + w, y1: y,     x2: x + w, y2: y + h },
    { edge: border.diagonalUp,  x1: x,     y1: y + h, x2: x + w, y2: y },
    { edge: border.diagonalDown,x1: x,     y1: y,     x2: x + w, y2: y + h },
  ];
  for (const { edge, x1, y1, x2, y2 } of edges) {
    if (!edge || !edge.style || edge.style === 'none') continue;
    ctx.beginPath();
    ctx.strokeStyle = edge.color ? hexToRgba(edge.color) : '#000000';
    ctx.lineWidth = borderStyleWidth(edge.style);
    const dash = borderStyleDash(edge.style);
    ctx.setLineDash(dash);
    ctx.moveTo(x1, y1);
    ctx.lineTo(x2, y2);
    ctx.stroke();
    ctx.setLineDash([]);
  }
}

function borderStyleWidth(style: string): number {
  switch (style) {
    case 'thick': return 2;
    case 'medium': case 'mediumDashed': case 'mediumDashDot': case 'mediumDashDotDot': case 'slantDashDot': return 1.5;
    case 'hair': return 0.5;
    default: return 1;
  }
}

function borderStyleDash(style: string): number[] {
  switch (style) {
    case 'dashed': case 'mediumDashed': return [4, 3];
    case 'dotted': return [2, 2];
    case 'dashDot': case 'mediumDashDot': return [4, 2, 1, 2];
    case 'dashDotDot': case 'mediumDashDotDot': return [4, 2, 1, 2, 1, 2];
    case 'slantDashDot': return [5, 3, 1, 3];
    default: return [];
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// Chart rendering — delegated to @silurus/ooxml-core's unified renderer.
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Normalize the XLSX parser's raw `ChartData` (chartType="bar" + barDir + grouping)
 * into the canonical `ChartModel.chartType` vocabulary expected by core.
 */
function canonicalChartType(chart: ChartData): string {
  const t = chart.chartType;
  const g = chart.grouping;
  if (t === 'bar') {
    const isH = chart.barDir === 'bar';
    if (g === 'stacked')        return isH ? 'stackedBarH'    : 'stackedBar';
    if (g === 'percentStacked') return isH ? 'stackedBarHPct' : 'stackedBarPct';
    return isH ? 'clusteredBarH' : 'clusteredBar';
  }
  if (t === 'line') {
    if (g === 'stacked')        return 'stackedLine';
    if (g === 'percentStacked') return 'stackedLinePct';
    return 'line';
  }
  if (t === 'area') {
    if (g === 'stacked')        return 'stackedArea';
    if (g === 'percentStacked') return 'stackedAreaPct';
    return 'area';
  }
  return t;
}

function adaptChartData(chart: ChartData): ChartModel {
  return {
    chartType: canonicalChartType(chart),
    title: chart.title,
    categories: chart.categories,
    series: chart.series.map(s => ({
      name: s.name,
      color: s.color ?? null,
      values: s.values,
      seriesType: s.seriesType ?? null,
      categories: s.categories.length > 0 ? s.categories : null,
    })),
    showDataLabels: chart.showDataLabels ?? false,
    valMin: null,
    valMax: null,
    catAxisTitle: chart.catAxisTitle ?? null,
    valAxisTitle: chart.valAxisTitle ?? null,
    catAxisHidden: false,
    valAxisHidden: false,
    plotAreaBg: null,
    subtotalIndices: [],
  };
}

// ── renderCharts ────────────────────────────────────────────────────────────

function renderCharts(
  ctx: CanvasRenderingContext2D,
  ws: Worksheet,
  cs: number,
  startRow: number,
  startCol: number,
  scrollOffsetX: number,
  scrollOffsetY: number,
  scrollAreaX: number,
  scrollAreaY: number,
  scrollAreaW: number,
  scrollAreaH: number,
): void {
  if (scrollAreaW <= 0 || scrollAreaH <= 0) return;

  const scrollOriginSheetX = sheetXForCol(ws, startCol, cs);
  const scrollOriginSheetY = sheetYForRow(ws, startRow, cs);

  for (const anchor of ws.charts) {
    const fromCol1 = anchor.fromCol + 1;
    const fromRow1 = anchor.fromRow + 1;
    const toCol1   = anchor.toCol   + 1;
    const toRow1   = anchor.toRow   + 1;

    const shX1 = sheetXForCol(ws, fromCol1, cs) + (anchor.fromColOff * cs) / EMU_PER_PX;
    const shY1 = sheetYForRow(ws, fromRow1, cs) + (anchor.fromRowOff * cs) / EMU_PER_PX;
    const shX2 = sheetXForCol(ws, toCol1,   cs) + (anchor.toColOff   * cs) / EMU_PER_PX;
    const shY2 = sheetYForRow(ws, toRow1,   cs) + (anchor.toRowOff   * cs) / EMU_PER_PX;

    const cw = shX2 - shX1;
    const ch = shY2 - shY1;
    if (cw <= 0 || ch <= 0) continue;

    const cx = scrollAreaX + (shX1 - scrollOriginSheetX) - scrollOffsetX;
    const cy = scrollAreaY + (shY1 - scrollOriginSheetY) - scrollOffsetY;

    if (cx + cw < scrollAreaX || cx > scrollAreaX + scrollAreaW) continue;
    if (cy + ch < scrollAreaY || cy > scrollAreaY + scrollAreaH) continue;

    ctx.save();
    ctx.beginPath();
    ctx.rect(scrollAreaX, scrollAreaY, scrollAreaW, scrollAreaH);
    ctx.clip();

    renderChart(ctx, adaptChartData(anchor.chart), { x: cx, y: cy, w: cw, h: ch });
    ctx.restore();
  }
}
