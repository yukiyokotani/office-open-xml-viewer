import type {
  Worksheet, Styles, Cell, CellValue, Font, Fill, Border, BorderEdge, CellXf,
  ViewportRange, RenderViewportOptions,
  CfRule, CellRange, CfStop, CfValue, Dxf,
  Run, ChartData,
} from './types.js';

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
  scaleStops?: number[];  // numeric values at each color stop
  barMin?: number;
  barMax?: number;
}

interface CfContext {
  compiled: CompiledCfRule[];
}

interface CfResult {
  fill?: Fill;
  fontColor?: string;
  fontBold?: boolean;
  fontItalic?: boolean;
  dataBar?: { color: string; ratio: number };
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
        const dxf = rule.dxfId != null ? dxfs[rule.dxfId] : null;
        if (dxf?.fill?.fgColor) result.fill = dxf.fill;
        if (dxf?.font?.color) result.fontColor = dxf.font.color;
        if (dxf?.font?.bold) result.fontBold = true;
        if (dxf?.font?.italic) result.fontItalic = true;
      }
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
  colWidths: number[];  // widths of scrollable cols (viewport.col..)
  rowHeights: number[]; // heights of scrollable rows (viewport.row..)
  frozenColWidths: number[];  // widths of frozen cols (cols 1..freezeCols)
  frozenRowHeights: number[]; // heights of frozen rows (rows 1..freezeRows)
  frozenW: number;
  frozenH: number;
  startRow: number;  // first scrollable row index
  startCol: number;  // first scrollable col index
  cs: number;        // cell scale factor (default 1)
  dpr: number;       // device pixel ratio
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

      if (!cell) continue;
      const text = formatCellValue(cell, styles);
      if (!text) continue;

      const effectiveBold = font.bold || !!cf.fontBold;
      const effectiveItalic = font.italic || !!cf.fontItalic;
      const fontForDraw: Font = effectiveBold !== font.bold || effectiveItalic !== font.italic
        ? { ...font, bold: effectiveBold, italic: effectiveItalic }
        : font;
      ctx.font = buildFont(fontForDraw, cs);
      const textColor = cf.fontColor ?? font.color;
      ctx.fillStyle = textColor ? hexToRgba(textColor) : '#000000';

      const paddingX = 3;
      const paddingY = 2;
      const isNumeric = cell.value.type === 'number';
      const alignH = xf.alignH ?? (isNumeric ? 'right' : 'left');
      const alignV = xf.alignV ?? 'bottom';
      // Indent: each level ≈ one character width (ECMA-376 §18.8.44)
      const indentPx = xf.indent ? Math.round(xf.indent * font.size * ROW_HEIGHT_TO_PX * 0.5) : 0;
      const leftPad = paddingX + (alignH === 'left' || !xf.alignH ? indentPx : 0);

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
      const hasRichText = runs && runs.length > 0 && !xf.wrapText;

      if (xf.wrapText) {
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

        if (font.underline) {
          const { x: ux, width: tW } = overlayX();
          const uy = alignV === 'top'
            ? cy + paddingY + sizePx + 1
            : alignV === 'center'
              ? cy + cellH / 2 + Math.round(sizePx * 0.55)
              : cy + cellH - paddingY + 1;
          ctx.save();
          ctx.strokeStyle = font.color ? hexToRgba(font.color) : '#000000';
          ctx.lineWidth = 0.5;
          ctx.beginPath(); ctx.moveTo(ux, uy); ctx.lineTo(ux + tW, uy); ctx.stroke();
          ctx.restore();
        }
        if (font.strike) {
          const { x: sx, width: tW } = overlayX();
          // Strike line sits roughly at the x-height mid-line (~45% up from baseline)
          const sy = alignV === 'top'
            ? cy + paddingY + Math.round(sizePx * 0.5)
            : alignV === 'center'
              ? cy + cellH / 2
              : cy + cellH - paddingY - Math.round(sizePx * 0.35);
          ctx.save();
          ctx.strokeStyle = font.color ? hexToRgba(font.color) : '#000000';
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

  const rc: RenderContext = {
    worksheet, styles, cellMap, mergeAnchorMap, mergeSkipSet, cfContext,
    colWidths: scrollColWidths,
    rowHeights: scrollRowHeights,
    frozenColWidths, frozenRowHeights,
    frozenW, frozenH,
    startRow, startCol,
    cs,
    dpr,
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
// Chart rendering
// ═══════════════════════════════════════════════════════════════════════════

// Standard Excel chart colour palette (matches Office default theme)
const CHART_PALETTE = [
  '4472C4','ED7D31','A9D18E','FF0000','70AD47','4BACC6',
  'FFC000','9E480E','843C0C','636363','255E91','967300',
];

function chartColor(idx: number, explicit?: string | null): string {
  return explicit ? `#${explicit}` : `#${CHART_PALETTE[idx % CHART_PALETTE.length]}`;
}

function niceStep(range: number, targetSteps = 5): number {
  if (range === 0) return 1;
  const raw = range / targetSteps;
  const mag = Math.pow(10, Math.floor(Math.log10(raw)));
  const normed = raw / mag;
  const nice = normed < 1.5 ? 1 : normed < 3.5 ? 2 : normed < 7.5 ? 5 : 10;
  return nice * mag;
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

    renderChartData(ctx, anchor.chart, cx, cy, cw, ch);
    ctx.restore();
  }
}

// ── Dispatcher ──────────────────────────────────────────────────────────────

function renderChartData(
  ctx: CanvasRenderingContext2D,
  chart: ChartData,
  x: number, y: number, w: number, h: number,
): void {
  // White background + light border
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(x, y, w, h);
  ctx.strokeStyle = '#d0d0d0';
  ctx.lineWidth = 1;
  ctx.strokeRect(x + 0.5, y + 0.5, w - 1, h - 1);

  if (chart.series.length === 0) {
    ctx.fillStyle = '#888';
    ctx.font = '12px sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText('(no data)', x + w / 2, y + h / 2);
    return;
  }

  switch (chart.chartType) {
    case 'bar':     renderBarChart(ctx, chart, x, y, w, h);                break;
    case 'line':    renderLineChartXlsx(ctx, chart, x, y, w, h);           break;
    case 'area':    renderAreaChartXlsx(ctx, chart, x, y, w, h);           break;
    case 'pie':     renderPieChartXlsx(ctx, chart, x, y, w, h, false);     break;
    case 'doughnut':renderPieChartXlsx(ctx, chart, x, y, w, h, true);      break;
    case 'radar':   renderRadarChart(ctx, chart, x, y, w, h);              break;
    case 'scatter': renderScatterChartXlsx(ctx, chart, x, y, w, h);        break;
    default:
      ctx.fillStyle = '#888';
      ctx.font = '11px sans-serif';
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      ctx.fillText(`Chart: ${chart.chartType}`, x + w / 2, y + h / 2);
  }
}

// ── Legend helper ────────────────────────────────────────────────────────────

function drawLegend(
  ctx: CanvasRenderingContext2D,
  series: ChartData['series'],
  lx: number, ly: number, lw: number, lh: number,
): void {
  const fontSize = Math.max(9, Math.min(12, lh / (series.length + 1)));
  ctx.font = `${fontSize}px sans-serif`;
  ctx.textBaseline = 'middle';
  const sw = 10; const gap = 4;
  const rowH = fontSize + 4;
  let ry = ly + (lh - rowH * series.length) / 2;
  for (let i = 0; i < series.length; i++) {
    const color = chartColor(i, null);
    ctx.fillStyle = color;
    ctx.fillRect(lx, ry, sw, fontSize);
    ctx.fillStyle = '#333';
    ctx.textAlign = 'left';
    const label = series[i].name || `Series ${i + 1}`;
    ctx.fillText(label.slice(0, 20), lx + sw + gap, ry + fontSize / 2);
    ry += rowH;
  }
}

// ── Title helper ─────────────────────────────────────────────────────────────

function drawChartTitle(
  ctx: CanvasRenderingContext2D,
  title: string | null,
  x: number, y: number, w: number, fontSize: number,
): void {
  if (!title) return;
  ctx.font = `bold ${fontSize}px sans-serif`;
  ctx.fillStyle = '#333';
  ctx.textAlign = 'center';
  ctx.textBaseline = 'top';
  ctx.fillText(title, x + w / 2, y);
}

// ═══════════════════════════════════════════════════════════════════════════
// Bar chart (vertical columns + horizontal bars, clustered + stacked)
// Also handles mixed bar+line series (seriesType per series)
// ═══════════════════════════════════════════════════════════════════════════

function renderBarChart(
  ctx: CanvasRenderingContext2D,
  chart: ChartData,
  x: number, y: number, w: number, h: number,
): void {
  const isH    = chart.barDir === 'bar';
  const stacked = chart.grouping === 'stacked' || chart.grouping === 'percentStacked';
  const pct    = chart.grouping === 'percentStacked';

  // Separate bar and line series (for mixed charts)
  const barSeries  = chart.series.filter(s => s.seriesType !== 'line');
  const lineSeries = chart.series.filter(s => s.seriesType === 'line');
  const allSeries  = barSeries; // for axis calculation include bar only

  const cats = chart.categories.length > 0
    ? chart.categories
    : (chart.series[0]?.categories ?? []);
  const n = cats.length;
  if (n === 0) return;

  // Layout
  const titleH  = chart.title ? Math.max(14, h * 0.06) : 0;
  const legendW = chart.series.length >= 1 ? Math.max(80, w * 0.22) : 0;
  const pad = { t: titleH + h * 0.04, r: legendW + w * 0.03, b: h * 0.14, l: w * 0.12 };
  if (isH) { pad.l = w * 0.22; pad.b = h * 0.08; }

  drawChartTitle(ctx, chart.title, x, y + 2, w, Math.max(11, titleH * 0.7));

  const px0 = x + pad.l; const py0 = y + pad.t;
  const pw  = w - pad.l - pad.r; const ph = h - pad.t - pad.b;
  if (pw <= 0 || ph <= 0) return;

  // Compute data max
  let dataMax = 0;
  for (let ci = 0; ci < n; ci++) {
    let stackSum = 0;
    for (const s of allSeries) {
      const v = s.values[ci] ?? 0;
      if (stacked) stackSum += Math.abs(v);
      else dataMax = Math.max(dataMax, Math.abs(v));
    }
    if (stacked) dataMax = Math.max(dataMax, stackSum);
  }
  if (pct) dataMax = 100;
  if (dataMax === 0) dataMax = 1;

  const step  = niceStep(dataMax);
  const axMax = Math.ceil(dataMax / step) * step;

  // Draw grid + value axis
  const gridColor = '#e0e0e0';
  const steps = Math.round(axMax / step);
  ctx.textBaseline = 'middle';
  ctx.font = `${Math.max(8, Math.min(11, ph / 20))}px sans-serif`;
  ctx.fillStyle = '#555';

  for (let si = 0; si <= steps; si++) {
    const val = si * step;
    const label = pct ? `${Math.round(val)}%` : val >= 1000 ? `${(val / 1000).toFixed(1)}k` : String(val);
    if (!isH) {
      const gy = py0 + ph - (val / axMax) * ph;
      ctx.strokeStyle = si === 0 ? '#aaa' : gridColor;
      ctx.lineWidth = si === 0 ? 1 : 0.5;
      ctx.beginPath(); ctx.moveTo(px0, gy); ctx.lineTo(px0 + pw, gy); ctx.stroke();
      ctx.textAlign = 'right';
      ctx.fillText(label, px0 - 4, gy);
    } else {
      const gx = px0 + (val / axMax) * pw;
      ctx.strokeStyle = si === 0 ? '#aaa' : gridColor;
      ctx.lineWidth = si === 0 ? 1 : 0.5;
      ctx.beginPath(); ctx.moveTo(gx, py0); ctx.lineTo(gx, py0 + ph); ctx.stroke();
      ctx.textAlign = 'center';
      ctx.fillText(label, gx, py0 + ph + 10);
    }
  }

  // Draw category axis line
  ctx.strokeStyle = '#aaa'; ctx.lineWidth = 1;
  if (!isH) {
    ctx.beginPath(); ctx.moveTo(px0, py0 + ph); ctx.lineTo(px0 + pw, py0 + ph); ctx.stroke();
  } else {
    ctx.beginPath(); ctx.moveTo(px0, py0); ctx.lineTo(px0, py0 + ph); ctx.stroke();
  }

  // Draw bars
  const catGap = !isH ? pw / n : ph / n;
  const barW   = catGap * (stacked ? 0.6 : 0.6 / Math.max(1, barSeries.length));
  const clusterGap = stacked ? 0 : catGap * 0.6 / Math.max(1, barSeries.length);
  const catStart   = stacked ? catGap * 0.2 : catGap * 0.2;

  for (let ci = 0; ci < n; ci++) {
    let stackOffset = 0;
    let stackSum = 0;
    if (pct) {
      for (const s of barSeries) stackSum += Math.abs(s.values[ci] ?? 0);
      if (stackSum === 0) stackSum = 1;
    }

    for (let si = 0; si < barSeries.length; si++) {
      const s = barSeries[si];
      const raw = s.values[ci] ?? 0;
      const val = pct ? (Math.abs(raw) / stackSum) * 100 : Math.abs(raw);
      const color = chartColor(si, null);

      if (!isH) {
        const bx = stacked
          ? px0 + ci * catGap + catStart
          : px0 + ci * catGap + catStart + si * clusterGap;
        const barH = (val / axMax) * ph;
        const by   = py0 + ph - (stacked ? (stackOffset + val) : val) / axMax * ph;
        ctx.fillStyle = color;
        ctx.fillRect(bx, by, barW, barH);
      } else {
        const by = stacked
          ? py0 + (n - 1 - ci) * catGap + catStart
          : py0 + (n - 1 - ci) * catGap + catStart + si * clusterGap;
        const barL = (val / axMax) * pw;
        const bx   = stacked ? px0 + (stackOffset / axMax) * pw : px0;
        ctx.fillStyle = color;
        ctx.fillRect(bx, by, barL, barW);
      }
      if (stacked) stackOffset += val;
    }
  }

  // Draw category labels
  ctx.fillStyle = '#555';
  ctx.font = `${Math.max(8, Math.min(11, catGap * 0.5))}px sans-serif`;
  for (let ci = 0; ci < n; ci++) {
    const label = (cats[ci] ?? '').toString().slice(0, 12);
    if (!isH) {
      const lx = px0 + ci * catGap + catGap / 2;
      ctx.textAlign = 'center'; ctx.textBaseline = 'top';
      ctx.fillText(label, lx, py0 + ph + 3);
    } else {
      const ly = py0 + (n - 1 - ci) * catGap + catGap / 2;
      ctx.textAlign = 'right'; ctx.textBaseline = 'middle';
      ctx.fillText(label, px0 - 4, ly);
    }
  }

  // Overlay line series (mixed bar+line)
  if (lineSeries.length > 0) {
    for (let si = 0; si < lineSeries.length; si++) {
      const s = lineSeries[si];
      const color = chartColor(barSeries.length + si, null);
      ctx.strokeStyle = color; ctx.lineWidth = 2;
      ctx.setLineDash([]);
      ctx.beginPath();
      let started = false;
      for (let ci = 0; ci < n; ci++) {
        const v = s.values[ci];
        if (v == null) { started = false; continue; }
        const lx = px0 + ci * catGap + catGap / 2;
        const ly = py0 + ph - (v / axMax) * ph;
        if (!started) { ctx.moveTo(lx, ly); started = true; } else ctx.lineTo(lx, ly);
      }
      ctx.stroke();
      // Markers
      for (let ci = 0; ci < n; ci++) {
        const v = s.values[ci];
        if (v == null) continue;
        const lx = px0 + ci * catGap + catGap / 2;
        const ly = py0 + ph - (v / axMax) * ph;
        ctx.fillStyle = color;
        ctx.beginPath(); ctx.arc(lx, ly, 3, 0, Math.PI * 2); ctx.fill();
      }
    }
  }

  // Legend
  if (legendW > 0) {
    drawLegend(ctx, chart.series, x + w - legendW + 4, py0, legendW - 8, ph);
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// Line chart
// ═══════════════════════════════════════════════════════════════════════════

function renderLineChartXlsx(
  ctx: CanvasRenderingContext2D,
  chart: ChartData,
  x: number, y: number, w: number, h: number,
): void {
  const cats = chart.categories.length > 0 ? chart.categories : (chart.series[0]?.categories ?? []);
  const n = cats.length; if (n === 0) return;

  const titleH  = chart.title ? Math.max(14, h * 0.06) : 0;
  const legendW = chart.series.length >= 1 ? Math.max(80, w * 0.22) : 0;
  const pad = { t: titleH + h * 0.04, r: legendW + w * 0.05, b: h * 0.14, l: w * 0.12 };

  drawChartTitle(ctx, chart.title, x, y + 2, w, Math.max(11, titleH * 0.7));

  const px0 = x + pad.l; const py0 = y + pad.t;
  const pw = w - pad.l - pad.r; const ph = h - pad.t - pad.b;
  if (pw <= 0 || ph <= 0) return;

  // Y range
  let dataMin = Infinity; let dataMax = -Infinity;
  for (const s of chart.series) for (const v of s.values) if (v != null) { dataMin = Math.min(dataMin, v); dataMax = Math.max(dataMax, v); }
  if (!isFinite(dataMin)) { dataMin = 0; dataMax = 1; }
  if (dataMin > 0) dataMin = 0;
  if (dataMax < 0) dataMax = 0;
  if (dataMax === dataMin) dataMax = dataMin + 1;

  const step  = niceStep(dataMax - dataMin);
  const axMin = Math.floor(dataMin / step) * step;
  const axMax = Math.ceil(dataMax / step) * step;
  const range = axMax - axMin; if (range === 0) return;

  const toY = (v: number) => py0 + ph - ((v - axMin) / range) * ph;
  const toX = (i: number) => px0 + (n === 1 ? pw / 2 : (i / (n - 1)) * pw);

  // Grid
  const steps = Math.round((axMax - axMin) / step);
  ctx.font = `${Math.max(8, Math.min(11, ph / 20))}px sans-serif`;
  ctx.textBaseline = 'middle';
  for (let si = 0; si <= steps; si++) {
    const v = axMin + si * step;
    const gy = toY(v);
    ctx.strokeStyle = v === 0 ? '#aaa' : '#e0e0e0';
    ctx.lineWidth = v === 0 ? 1 : 0.5;
    ctx.beginPath(); ctx.moveTo(px0, gy); ctx.lineTo(px0 + pw, gy); ctx.stroke();
    ctx.fillStyle = '#555'; ctx.textAlign = 'right';
    ctx.fillText(String(v), px0 - 4, gy);
  }

  // Baseline
  ctx.strokeStyle = '#aaa'; ctx.lineWidth = 1;
  ctx.beginPath(); ctx.moveTo(px0, py0 + ph); ctx.lineTo(px0 + pw, py0 + ph); ctx.stroke();

  // Series
  const markerR = Math.max(3, ph * 0.015);
  for (let si = 0; si < chart.series.length; si++) {
    const s = chart.series[si];
    const color = chartColor(si, null);
    ctx.strokeStyle = color; ctx.lineWidth = 2; ctx.setLineDash([]);
    ctx.beginPath();
    let started = false;
    for (let ci = 0; ci < n; ci++) {
      const v = s.values[ci]; if (v == null) { started = false; continue; }
      const px = toX(ci); const py = toY(v);
      if (!started) { ctx.moveTo(px, py); started = true; } else ctx.lineTo(px, py);
    }
    ctx.stroke();
    ctx.fillStyle = color;
    for (let ci = 0; ci < n; ci++) {
      const v = s.values[ci]; if (v == null) continue;
      ctx.beginPath(); ctx.arc(toX(ci), toY(v), markerR, 0, Math.PI * 2); ctx.fill();
    }
  }

  // Category labels
  const labelInterval = Math.max(1, Math.ceil(n / 8));
  ctx.fillStyle = '#555'; ctx.textAlign = 'center'; ctx.textBaseline = 'top';
  ctx.font = `${Math.max(8, Math.min(11, pw / n * 0.8))}px sans-serif`;
  for (let ci = 0; ci < n; ci += labelInterval) {
    ctx.fillText((cats[ci] ?? '').toString().slice(0, 10), toX(ci), py0 + ph + 3);
  }

  if (legendW > 0) drawLegend(ctx, chart.series, x + w - legendW + 4, py0, legendW - 8, ph);
}

// ═══════════════════════════════════════════════════════════════════════════
// Area chart
// ═══════════════════════════════════════════════════════════════════════════

function renderAreaChartXlsx(
  ctx: CanvasRenderingContext2D,
  chart: ChartData,
  x: number, y: number, w: number, h: number,
): void {
  const cats = chart.categories.length > 0 ? chart.categories : (chart.series[0]?.categories ?? []);
  const n = cats.length; if (n === 0) return;
  const stacked = chart.grouping === 'stacked' || chart.grouping === 'percentStacked';

  const titleH  = chart.title ? Math.max(14, h * 0.06) : 0;
  const legendW = chart.series.length >= 1 ? Math.max(80, w * 0.22) : 0;
  const pad = { t: titleH + h * 0.04, r: legendW + w * 0.05, b: h * 0.14, l: w * 0.12 };

  drawChartTitle(ctx, chart.title, x, y + 2, w, Math.max(11, titleH * 0.7));

  const px0 = x + pad.l; const py0 = y + pad.t;
  const pw = w - pad.l - pad.r; const ph = h - pad.t - pad.b;
  if (pw <= 0 || ph <= 0) return;

  // Compute max (stacked: sum per category)
  let dataMax = 0;
  for (let ci = 0; ci < n; ci++) {
    if (stacked) {
      let sum = 0;
      for (const s of chart.series) sum += s.values[ci] ?? 0;
      dataMax = Math.max(dataMax, sum);
    } else {
      for (const s of chart.series) dataMax = Math.max(dataMax, s.values[ci] ?? 0);
    }
  }
  if (dataMax === 0) dataMax = 1;
  const step  = niceStep(dataMax);
  const axMax = Math.ceil(dataMax / step) * step;

  const toX = (i: number) => px0 + (n === 1 ? pw / 2 : (i / (n - 1)) * pw);
  const toY = (v: number) => py0 + ph - (v / axMax) * ph;

  // Grid
  ctx.font = `${Math.max(8, Math.min(11, ph / 20))}px sans-serif`;
  ctx.textBaseline = 'middle';
  const steps = Math.round(axMax / step);
  for (let si = 0; si <= steps; si++) {
    const v = si * step; const gy = toY(v);
    ctx.strokeStyle = si === 0 ? '#aaa' : '#e0e0e0';
    ctx.lineWidth = si === 0 ? 1 : 0.5;
    ctx.beginPath(); ctx.moveTo(px0, gy); ctx.lineTo(px0 + pw, gy); ctx.stroke();
    ctx.fillStyle = '#555'; ctx.textAlign = 'right';
    ctx.fillText(String(v), px0 - 4, gy);
  }
  ctx.strokeStyle = '#aaa'; ctx.lineWidth = 1;
  ctx.beginPath(); ctx.moveTo(px0, py0 + ph); ctx.lineTo(px0 + pw, py0 + ph); ctx.stroke();

  // Areas (render back-to-front)
  const stackBase = stacked ? new Array(n).fill(0) as number[] : null;
  for (let si = chart.series.length - 1; si >= 0; si--) {
    const s = chart.series[si];
    const color = chartColor(si, null);
    const baseY = py0 + ph;

    ctx.beginPath();
    if (stacked && stackBase) {
      // top edge: from first to last
      for (let ci = 0; ci < n; ci++) {
        const v = (s.values[ci] ?? 0) + stackBase[ci];
        const px = toX(ci); const py = toY(v);
        if (ci === 0) ctx.moveTo(px, py); else ctx.lineTo(px, py);
      }
      // bottom edge: base stack values in reverse
      for (let ci = n - 1; ci >= 0; ci--) {
        ctx.lineTo(toX(ci), toY(stackBase[ci]));
      }
      for (let ci = 0; ci < n; ci++) stackBase[ci] += s.values[ci] ?? 0;
    } else {
      ctx.moveTo(toX(0), baseY);
      for (let ci = 0; ci < n; ci++) ctx.lineTo(toX(ci), toY(s.values[ci] ?? 0));
      ctx.lineTo(toX(n - 1), baseY);
    }
    ctx.closePath();
    ctx.fillStyle = hexToRgba(color, 0.6);
    ctx.fill();
    ctx.strokeStyle = color; ctx.lineWidth = 1.5; ctx.setLineDash([]);
    ctx.stroke();
  }

  // Category labels
  const labelInterval = Math.max(1, Math.ceil(n / 8));
  ctx.fillStyle = '#555'; ctx.textAlign = 'center'; ctx.textBaseline = 'top';
  ctx.font = `${Math.max(8, Math.min(11, pw / n * 0.8))}px sans-serif`;
  for (let ci = 0; ci < n; ci += labelInterval) {
    ctx.fillText((cats[ci] ?? '').toString().slice(0, 10), toX(ci), py0 + ph + 3);
  }

  if (legendW > 0) drawLegend(ctx, chart.series, x + w - legendW + 4, py0, legendW - 8, ph);
}

// ═══════════════════════════════════════════════════════════════════════════
// Pie / Doughnut chart
// ═══════════════════════════════════════════════════════════════════════════

function renderPieChartXlsx(
  ctx: CanvasRenderingContext2D,
  chart: ChartData,
  x: number, y: number, w: number, h: number,
  isDoughnut: boolean,
): void {
  const s = chart.series[0]; if (!s) return;
  const cats = s.categories.length > 0 ? s.categories : chart.categories;
  const vals = s.values.map(v => Math.abs(v ?? 0));
  const total = vals.reduce((a, b) => a + b, 0);
  if (total === 0) return;

  const titleH = chart.title ? Math.max(14, h * 0.06) : 0;
  drawChartTitle(ctx, chart.title, x, y + 2, w, Math.max(11, titleH * 0.7));

  const legendW = Math.max(80, w * 0.28);
  const pw = w - legendW; const ph = h - titleH - h * 0.04;
  const cx2 = x + pw / 2; const cy2 = y + titleH + h * 0.04 + ph / 2;
  const outerR = Math.min(pw, ph) * 0.42;
  const innerR = isDoughnut ? outerR * 0.5 : 0;

  let angle = -Math.PI / 2;
  for (let i = 0; i < vals.length; i++) {
    const slice = (vals[i] / total) * Math.PI * 2;
    const color = chartColor(i, null);
    ctx.beginPath();
    ctx.moveTo(cx2, cy2);
    ctx.arc(cx2, cy2, outerR, angle, angle + slice);
    ctx.closePath();
    ctx.fillStyle = color; ctx.fill();
    ctx.strokeStyle = '#fff'; ctx.lineWidth = 1; ctx.stroke();
    angle += slice;
  }

  if (isDoughnut) {
    ctx.beginPath(); ctx.arc(cx2, cy2, innerR, 0, Math.PI * 2);
    ctx.fillStyle = '#fff'; ctx.fill();
  }

  // Legend
  const lx = x + pw + 4;
  const fontSize = Math.max(9, Math.min(12, h / (vals.length + 2)));
  ctx.font = `${fontSize}px sans-serif`;
  ctx.textBaseline = 'middle';
  const rowH = fontSize + 4;
  let ry = y + (h - rowH * vals.length) / 2;
  for (let i = 0; i < vals.length; i++) {
    ctx.fillStyle = chartColor(i, null);
    ctx.fillRect(lx, ry, 10, fontSize);
    ctx.fillStyle = '#333'; ctx.textAlign = 'left';
    ctx.fillText((cats[i] ?? `Item ${i + 1}`).toString().slice(0, 18), lx + 14, ry + fontSize / 2);
    ry += rowH;
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// Radar / Spider chart
// ═══════════════════════════════════════════════════════════════════════════

function renderRadarChart(
  ctx: CanvasRenderingContext2D,
  chart: ChartData,
  x: number, y: number, w: number, h: number,
): void {
  const cats = chart.categories.length > 0 ? chart.categories : (chart.series[0]?.categories ?? []);
  const n = cats.length; if (n < 3) return;

  const titleH  = chart.title ? Math.max(14, h * 0.06) : 0;
  const legendW = chart.series.length > 1 ? Math.max(70, w * 0.2) : 0;
  drawChartTitle(ctx, chart.title, x, y + 2, w, Math.max(11, titleH * 0.7));

  const pw = w - legendW; const ph = h - titleH - h * 0.04;
  const cx2 = x + pw / 2; const cy2 = y + titleH + h * 0.04 + ph / 2;
  const r   = Math.min(pw, ph) * 0.38;

  // Find max across all series
  let dataMax = 0;
  for (const s of chart.series) for (const v of s.values) dataMax = Math.max(dataMax, v ?? 0);
  if (dataMax === 0) dataMax = 1;
  const step  = niceStep(dataMax);
  const axMax = Math.ceil(dataMax / step) * step;

  const angle0 = -Math.PI / 2;
  const spoke  = (i: number) => angle0 + (i / n) * Math.PI * 2;

  // Concentric rings (grid)
  const rings = Math.round(axMax / step);
  ctx.strokeStyle = '#ddd'; ctx.lineWidth = 0.5;
  for (let ri = 1; ri <= rings; ri++) {
    const rr = (ri / rings) * r;
    ctx.beginPath();
    for (let i = 0; i < n; i++) {
      const a = spoke(i);
      const px = cx2 + Math.cos(a) * rr; const py = cy2 + Math.sin(a) * rr;
      if (i === 0) ctx.moveTo(px, py); else ctx.lineTo(px, py);
    }
    ctx.closePath(); ctx.stroke();
  }

  // Spokes
  ctx.strokeStyle = '#bbb'; ctx.lineWidth = 0.5;
  for (let i = 0; i < n; i++) {
    const a = spoke(i);
    ctx.beginPath(); ctx.moveTo(cx2, cy2);
    ctx.lineTo(cx2 + Math.cos(a) * r, cy2 + Math.sin(a) * r); ctx.stroke();
  }

  // Category labels
  ctx.font = `${Math.max(8, Math.min(11, r * 0.2))}px sans-serif`;
  ctx.fillStyle = '#444'; ctx.textBaseline = 'middle';
  for (let i = 0; i < n; i++) {
    const a = spoke(i);
    const lx = cx2 + Math.cos(a) * (r + 12);
    const ly = cy2 + Math.sin(a) * (r + 12);
    ctx.textAlign = Math.cos(a) < -0.1 ? 'right' : Math.cos(a) > 0.1 ? 'left' : 'center';
    ctx.fillText((cats[i] ?? '').toString().slice(0, 12), lx, ly);
  }

  // Series polygons
  for (let si = 0; si < chart.series.length; si++) {
    const s = chart.series[si];
    const color = chartColor(si, null);
    ctx.beginPath();
    for (let i = 0; i < n; i++) {
      const v = s.values[i] ?? 0;
      const frac = v / axMax;
      const a = spoke(i);
      const px = cx2 + Math.cos(a) * r * frac;
      const py = cy2 + Math.sin(a) * r * frac;
      if (i === 0) ctx.moveTo(px, py); else ctx.lineTo(px, py);
    }
    ctx.closePath();
    ctx.fillStyle = hexToRgba(color, 0.25); ctx.fill();
    ctx.strokeStyle = color; ctx.lineWidth = 2; ctx.stroke();
  }

  if (legendW > 0) {
    drawLegend(ctx, chart.series, x + w - legendW + 4, y + titleH + h * 0.04, legendW - 8, ph);
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// Scatter chart
// ═══════════════════════════════════════════════════════════════════════════

function renderScatterChartXlsx(
  ctx: CanvasRenderingContext2D,
  chart: ChartData,
  x: number, y: number, w: number, h: number,
): void {
  const titleH  = chart.title ? Math.max(14, h * 0.06) : 0;
  const legendW = chart.series.length >= 1 ? Math.max(80, w * 0.22) : 0;
  const pad = { t: titleH + h * 0.06, r: legendW + w * 0.05, b: h * 0.12, l: w * 0.12 };

  drawChartTitle(ctx, chart.title, x, y + 2, w, Math.max(11, titleH * 0.7));

  const px0 = x + pad.l; const py0 = y + pad.t;
  const pw = w - pad.l - pad.r; const ph = h - pad.t - pad.b;
  if (pw <= 0 || ph <= 0) return;

  // For scatter: categories hold X values (as strings), values hold Y
  let allX: number[] = []; let allY: number[] = [];
  for (const s of chart.series) {
    for (const c of s.categories) { const v = parseFloat(c); if (!isNaN(v)) allX.push(v); }
    for (const v of s.values) if (v != null) allY.push(v);
  }
  // If no numeric X, use index
  let useIndexX = allX.length === 0;
  if (useIndexX) {
    const maxLen = Math.max(...chart.series.map(s => s.values.length));
    for (let i = 0; i < maxLen; i++) allX.push(i);
  }

  let xMin = Math.min(...allX); let xMax = Math.max(...allX);
  let yMin = Math.min(...allY); let yMax = Math.max(...allY);
  if (xMin === xMax) { xMin -= 1; xMax += 1; }
  if (yMin === yMax) { yMin -= 1; yMax += 1; }
  if (yMin > 0) yMin = 0;

  const toX = (v: number) => px0 + ((v - xMin) / (xMax - xMin)) * pw;
  const toY = (v: number) => py0 + ph - ((v - yMin) / (yMax - yMin)) * ph;

  // Grid lines
  ctx.font = `${Math.max(8, Math.min(11, ph / 20))}px sans-serif`;
  const yStep = niceStep(yMax - yMin);
  const ySteps = Math.round((yMax - yMin) / yStep) + 1;
  for (let si = 0; si < ySteps; si++) {
    const v = yMin + si * yStep; if (v > yMax + yStep * 0.01) break;
    const gy = toY(v);
    ctx.strokeStyle = '#e0e0e0'; ctx.lineWidth = 0.5;
    ctx.beginPath(); ctx.moveTo(px0, gy); ctx.lineTo(px0 + pw, gy); ctx.stroke();
    ctx.fillStyle = '#555'; ctx.textAlign = 'right'; ctx.textBaseline = 'middle';
    ctx.fillText(String(v), px0 - 4, gy);
  }
  // Axes
  ctx.strokeStyle = '#aaa'; ctx.lineWidth = 1;
  ctx.beginPath(); ctx.moveTo(px0, py0 + ph); ctx.lineTo(px0 + pw, py0 + ph); ctx.stroke();
  ctx.beginPath(); ctx.moveTo(px0, py0); ctx.lineTo(px0, py0 + ph); ctx.stroke();

  // Points
  const markerR = Math.max(3, ph * 0.015);
  for (let si = 0; si < chart.series.length; si++) {
    const s = chart.series[si];
    const color = chartColor(si, null);
    ctx.fillStyle = color;
    for (let ci = 0; ci < s.values.length; ci++) {
      const yv = s.values[ci]; if (yv == null) continue;
      const xv = useIndexX ? ci : parseFloat(s.categories[ci] ?? '0');
      if (isNaN(xv)) continue;
      ctx.beginPath(); ctx.arc(toX(xv), toY(yv), markerR, 0, Math.PI * 2); ctx.fill();
    }
  }

  if (legendW > 0) drawLegend(ctx, chart.series, x + w - legendW + 4, py0, legendW - 8, ph);
}
