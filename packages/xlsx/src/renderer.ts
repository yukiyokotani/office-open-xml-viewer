import type {
  Worksheet, Styles, Cell, CellValue, Font, Fill, Border, BorderEdge, CellXf,
  ViewportRange, RenderViewportOptions, TextRunInfo,
  CfRule, CellRange, CfStop, CfValue, Dxf, Hyperlink, DefinedName,
  Run, ChartData, GradientFillSpec, ShapeInfo, SlicerItem,
} from './types.js';
import { renderChart, renderSparkline, type ChartModel, type SparklineModel } from '@silurus/ooxml-core';

// Default font stack. Calibri is the workbook default font in Excel; on
// systems without Office (macOS / Linux) the browser would otherwise fall
// back to Arial / Helvetica, which is meaningfully wider than Calibri at
// every weight/size combination. Carlito is the Google-released, metric-
// compatible Calibri clone (same advance widths and ascender / descender
// metrics) and is loaded opt-in by `XlsxWorkbook.load({ useGoogleFonts:
// true })`. Listing it in the cascade means: Calibri (Windows / Office)
// → Carlito (loaded webfont) → Arial → sans-serif. Caladea is the same
// for Cambria.
const DEFAULT_FONT_FAMILY = '"Calibri", "Carlito", "Cambria", "Caladea", Arial, sans-serif';
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

/**
 * Fill a data-bar rectangle. Excel 2010+ dataBars default to a horizontal
 * gradient (`x14:dataBar@gradient="1"`): solid color on the left, fading
 * to an ~85%-tinted-to-white version on the right. We render with alpha
 * stops rather than literally mixing toward white so underlying cell
 * background (including zebra-striping or fills) shows through. With
 * `gradient="0"` the bar is drawn as a flat solid color.
 */
function fillDataBar(
  ctx: CanvasRenderingContext2D,
  color: string,
  x: number, y: number, w: number, h: number,
  gradient: boolean,
): void {
  if (w <= 0 || h <= 0) return;
  if (gradient) {
    const grad = ctx.createLinearGradient(x, y, x + w, y);
    grad.addColorStop(0, hexToRgba(color, 0.85));
    grad.addColorStop(1, hexToRgba(color, 0.15));
    ctx.fillStyle = grad;
  } else {
    ctx.fillStyle = hexToRgba(color);
  }
  ctx.fillRect(x, y, w, h);
}

/**
 * Fractional fg coverage for an ECMA-376 ST_PatternType (§18.8.22). Values are
 * derived from the spec's verbal descriptions — gray125 = 12.5% fg on bg,
 * gray0625 = 6.25%, mediumGray ≈ 50%, darkGray ≈ 75%, lightGray ≈ 25%. Hatch
 * variants (darkHorizontal etc.) approximate their average ink density.
 * Unknown values default to 1 so they render as solid fg.
 */
function patternCoverage(pt: string): number {
  switch (pt) {
    case 'solid':         return 1;
    case 'darkGray':      return 0.75;
    case 'mediumGray':    return 0.50;
    case 'lightGray':     return 0.25;
    case 'gray125':       return 0.125;
    case 'gray0625':      return 0.0625;
    // Directional hatches: visual ink ratio is roughly 50/50 at default cell
    // sizes; without a true hatch tile we blend at 50% so the cell reads as a
    // middle-tone fill rather than a solid fg block.
    case 'darkHorizontal':
    case 'darkVertical':
    case 'darkDown':
    case 'darkUp':
    case 'darkGrid':
    case 'darkTrellis':   return 0.5;
    case 'lightHorizontal':
    case 'lightVertical':
    case 'lightDown':
    case 'lightUp':
    case 'lightGrid':
    case 'lightTrellis':  return 0.25;
    default:              return 1;
  }
}

/**
 * Cache of (pattern, fg, bg) → CanvasPattern. Keyed by a compound string so
 * that the same pattern across thousands of cells only builds the tile once.
 */
const PATTERN_CACHE = new Map<string, CanvasPattern | null>();

/**
 * Build a small repeating tile for ECMA-376 directional hatch patterns. Uses
 * an offscreen canvas painted with bgColor + fgColor lines and returns a
 * CanvasPattern for `ctx.fillStyle`. Non-hatch patterns return null so the
 * caller falls back to the fg/bg blend.
 */
function hatchPattern(
  ctx: CanvasRenderingContext2D,
  pt: string,
  fgHex: string,
  bgHex: string,
): CanvasPattern | null {
  const key = `${pt}|${fgHex}|${bgHex}`;
  if (PATTERN_CACHE.has(key)) return PATTERN_CACHE.get(key)!;

  const size = 8;
  const isDark = pt.startsWith('dark');
  const isLight = pt.startsWith('light');
  if (!isDark && !isLight) {
    PATTERN_CACHE.set(key, null);
    return null;
  }
  const off = document.createElement('canvas');
  off.width = size;
  off.height = size;
  const octx = off.getContext('2d');
  if (!octx) { PATTERN_CACHE.set(key, null); return null; }

  octx.fillStyle = hexToRgba(bgHex);
  octx.fillRect(0, 0, size, size);
  octx.strokeStyle = hexToRgba(fgHex);
  octx.lineWidth = isDark ? 2 : 1;
  octx.beginPath();
  switch (pt) {
    case 'darkHorizontal':
    case 'lightHorizontal':
      octx.moveTo(0, size / 2);
      octx.lineTo(size, size / 2);
      break;
    case 'darkVertical':
    case 'lightVertical':
      octx.moveTo(size / 2, 0);
      octx.lineTo(size / 2, size);
      break;
    case 'darkDown':
    case 'lightDown':
      // Diagonal from top-left to bottom-right; draw a second line shifted to
      // avoid seams where the tile wraps.
      octx.moveTo(0, 0); octx.lineTo(size, size);
      octx.moveTo(-size, 0); octx.lineTo(0, size);
      octx.moveTo(0, -size); octx.lineTo(size, 0);
      break;
    case 'darkUp':
    case 'lightUp':
      octx.moveTo(0, size); octx.lineTo(size, 0);
      octx.moveTo(-size, size); octx.lineTo(0, 0);
      octx.moveTo(0, 2 * size); octx.lineTo(size, size);
      break;
    case 'darkGrid':
    case 'lightGrid':
      octx.moveTo(0, size / 2); octx.lineTo(size, size / 2);
      octx.moveTo(size / 2, 0); octx.lineTo(size / 2, size);
      break;
    case 'darkTrellis':
    case 'lightTrellis':
      octx.moveTo(0, 0); octx.lineTo(size, size);
      octx.moveTo(0, size); octx.lineTo(size, 0);
      break;
    default:
      PATTERN_CACHE.set(key, null);
      return null;
  }
  octx.stroke();

  const pat = ctx.createPattern(off, 'repeat');
  PATTERN_CACHE.set(key, pat);
  return pat;
}

/**
 * Build a Canvas gradient object for an xlsx `<gradientFill>`. Linear uses
 * the degree attribute (0° = left→right, 90° = top→bottom). Path gradients
 * radiate from a rectangular inner bounds defined by left/right/top/bottom
 * as fractions of the cell.
 */
function buildGradientFill(
  ctx: CanvasRenderingContext2D,
  g: GradientFillSpec,
  x: number, y: number, w: number, h: number,
): CanvasGradient {
  let grad: CanvasGradient;
  if (g.gradientType === 'path') {
    // Use the inner rectangle's center as the radial origin; radius spans to
    // the farthest cell corner so stop=1 always reaches a cell edge.
    const cxg = x + w * (g.left + (1 - g.right - g.left) / 2);
    const cyg = y + h * (g.top + (1 - g.bottom - g.top) / 2);
    const r = Math.hypot(Math.max(cxg - x, x + w - cxg), Math.max(cyg - y, y + h - cyg));
    grad = ctx.createRadialGradient(cxg, cyg, 0, cxg, cyg, r);
  } else {
    // Linear: rotate around the cell's center and extend to the bounds.
    const rad = (g.degree * Math.PI) / 180;
    const cxg = x + w / 2;
    const cyg = y + h / 2;
    const ext = (Math.abs(Math.cos(rad)) * w + Math.abs(Math.sin(rad)) * h) / 2;
    grad = ctx.createLinearGradient(
      cxg - Math.cos(rad) * ext, cyg - Math.sin(rad) * ext,
      cxg + Math.cos(rad) * ext, cyg + Math.sin(rad) * ext,
    );
  }
  for (const stop of g.stops) {
    const pos = Math.min(1, Math.max(0, stop.position));
    grad.addColorStop(pos, hexToRgba(stop.color));
  }
  return grad;
}

/**
 * Parse an A1-style cell reference ("A1", "B12", "AA3") to 1-based row/col.
 * Returns null when the input doesn't match the expected shape (parser-side
 * data is trusted, but we still guard against malformed refs).
 */
function parseA1Ref(ref: string): { row: number; col: number } | null {
  const m = /^([A-Z]+)(\d+)$/.exec(ref);
  if (!m) return null;
  const colLetters = m[1];
  const row = parseInt(m[2], 10);
  let col = 0;
  for (let i = 0; i < colLetters.length; i++) {
    col = col * 26 + (colLetters.charCodeAt(i) - 64);
  }
  return { row, col };
}

/**
 * Draw Excel's comment marker — a small filled triangle in the top-right
 * corner of the cell — coloured like Excel's default red indicator. Scales
 * with cell size but is clamped so it stays legible at small zoom.
 */
function drawCommentMarker(
  ctx: CanvasRenderingContext2D,
  x: number, y: number, w: number, h: number,
): void {
  const size = Math.max(4, Math.min(8, Math.min(w, h) * 0.18));
  ctx.save();
  ctx.fillStyle = '#D40000';
  ctx.beginPath();
  ctx.moveTo(x + w - size, y);
  ctx.lineTo(x + w, y);
  ctx.lineTo(x + w, y + size);
  ctx.closePath();
  ctx.fill();
  ctx.restore();
}

/** Linear-interpolate two #RRGGBB values in RGB space at coverage fg weight. */
function blendHex(fgHex: string, bgHex: string, fgCoverage: number): string {
  const fh = fgHex.replace('#', '');
  const bh = bgHex.replace('#', '');
  const fr = parseInt(fh.slice(0, 2), 16);
  const fg = parseInt(fh.slice(2, 4), 16);
  const fb = parseInt(fh.slice(4, 6), 16);
  const br = parseInt(bh.slice(0, 2), 16);
  const bg = parseInt(bh.slice(2, 4), 16);
  const bb = parseInt(bh.slice(4, 6), 16);
  const c = Math.min(1, Math.max(0, fgCoverage));
  const r = Math.round(fr * c + br * (1 - c));
  const g = Math.round(fg * c + bg * (1 - c));
  const b = Math.round(fb * c + bb * (1 - c));
  return `rgb(${r},${g},${b})`;
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

function formatCellValue(
  cell: Cell,
  styles: Styles,
  cfNumFmt?: { numFmtId: number; formatCode: string | null } | null,
): string {
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
    return effectiveFmt ? applyTextSection(text, effectiveFmt) : text;
  }

  // Volatile builtins: TODAY()/NOW() cells have a cached `<v>` from the last
  // save, which the viewer would otherwise show as a stale date. Recompute
  // them against the current system clock at render time.
  const num = recomputeVolatile(cell.formula) ?? cell.value.number;
  return applyFormat(num, effectiveFmtId, effectiveFmt);
}

/**
 * Apply the 4th section (text section) of an Excel number format to a text
 * value. ECMA-376 §18.8.30:
 *   - Fewer than 4 sections → text passes through unchanged (Excel default).
 *   - Empty text section   → the value is hidden.
 *   - `@` in the section   → substituted by the original text.
 *   - Quoted / escaped literals are emitted; `[...]` metadata and
 *     `_`/`*` pad pairs are dropped (same conventions as the numeric path).
 */
function applyTextSection(text: string, formatCode: string): string {
  const sections = formatCode.split(';');
  if (sections.length < 4) return text;
  const section = sections[3];
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

/** If `formula` is a volatile builtin (TODAY/NOW), return the current Excel
 *  serial. Tolerates surrounding whitespace and an optional leading `=`. */
function recomputeVolatile(formula: string | undefined): number | null {
  if (!formula) return null;
  const f = formula.trim().replace(/^=/, '').toUpperCase().replace(/\s+/g, '');
  if (f === 'TODAY()') return todaySerial();
  if (f === 'NOW()') return nowSerial();
  return null;
}

// ────────────────────────────────────────────────────────────────
// Date / time formatting  (ECMA-376 §18.8.30)
// ────────────────────────────────────────────────────────────────

// Built-in numFmtId → format code. IDs 14-22 are the ECMA-376 US-English
// built-ins; IDs 27-31 and 50-58 are East-Asian (Japanese) locale built-ins
// that Office ships pre-assigned when the file was authored in ja-JP. The
// spec lists the codes under §18.8.30 Table "Built-in formats" (the
// locale-dependent block is given without format strings but the de-facto
// codes match the ones that Office writes back when opening and re-saving).
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

/**
 * Split a format section into an ordered list of tokens, preserving the exact
 * literal surroundings (quoted strings, backslash escapes, non-placeholder
 * characters like `$`, `€`, `¥` when unquoted) so they can be reassembled
 * around the formatted number. Drops bracket metadata (`[Red]`, `[>0]`),
 * underscore-pad pairs and `*`-fill pairs, per ECMA-376 §18.8.30.
 */
type FmtToken =
  | { kind: 'lit'; text: string }
  | { kind: 'num' }
  | { kind: 'percent' }
  | { kind: 'sci'; expSign: boolean };

function tokenizeNumberFormat(section: string): { tokens: FmtToken[]; numSpec: string } {
  const tokens: FmtToken[] = [];
  let numSpec = '';
  let numPushed = false;
  let sciPushed = false;
  const pushLit = (s: string) => {
    if (!s) return;
    const last = tokens[tokens.length - 1];
    if (last && last.kind === 'lit') last.text += s;
    else tokens.push({ kind: 'lit', text: s });
  };
  const ensureNum = () => {
    if (!numPushed) { tokens.push({ kind: 'num' }); numPushed = true; }
  };

  let i = 0;
  while (i < section.length) {
    const ch = section[i];
    if (ch === '"') {
      i++;
      let s = '';
      while (i < section.length && section[i] !== '"') s += section[i++];
      if (i < section.length) i++;
      pushLit(s);
    } else if (ch === '\\') {
      if (i + 1 < section.length) pushLit(section[i + 1]);
      i += 2;
    } else if (ch === '[') {
      while (i < section.length && section[i] !== ']') i++;
      if (i < section.length) i++;
    } else if (ch === '_') {
      i += 2;
    } else if (ch === '*') {
      i += 2;
    } else if (ch === '#' || ch === '0' || ch === '?' || ch === '.' || ch === ',') {
      ensureNum();
      numSpec += ch;
      i++;
    } else if (ch === '%') {
      tokens.push({ kind: 'percent' });
      i++;
    } else if ((ch === 'E' || ch === 'e') && (section[i + 1] === '+' || section[i + 1] === '-')) {
      if (!sciPushed) {
        tokens.push({ kind: 'sci', expSign: section[i + 1] === '+' });
        sciPushed = true;
      }
      i += 2;
      while (i < section.length && section[i] === '0') i++;
    } else {
      pushLit(ch);
      i++;
    }
  }
  return { tokens, numSpec };
}

function formatNumberSpec(value: number, numSpec: string): string {
  const hasThousands = numSpec.includes(',') && /[#0]/.test(numSpec);
  const dec = countDecimalPlaces(numSpec);
  if (hasThousands) return formatThousands(value, dec);
  if (numSpec.includes('.')) return value.toFixed(dec);
  if (/[#0?]/.test(numSpec)) return Math.round(value).toString();
  return String(value);
}

function applyFormatCode(num: number, formatCode: string): string {
  const sections = formatCode.split(';');
  // Excel number formats have up to 4 sections: positive;negative;zero;text
  // (§18.8.30). Pick the section matching `num`, falling back to the positive
  // section when the target one is absent.
  let section: string;
  if (num > 0) section = sections[0];
  else if (num < 0) section = sections.length > 1 ? sections[1] : sections[0];
  else section = sections.length > 2 ? sections[2] : sections[0];
  const { tokens, numSpec } = tokenizeNumberFormat(section);
  const hasPercent = tokens.some(t => t.kind === 'percent');
  const sciTok = tokens.find(t => t.kind === 'sci') as Extract<FmtToken, { kind: 'sci' }> | undefined;

  let value = num;
  if (hasPercent) value = value * 100;

  let numberText: string;
  let expText = '';
  if (sciTok) {
    const dec = countDecimalPlaces(numSpec);
    const [mantissa, exp] = value.toExponential(dec).split('e');
    numberText = mantissa;
    const e = parseInt(exp, 10);
    const sign = e < 0 ? '-' : (sciTok.expSign ? '+' : '');
    expText = sign + String(Math.abs(e)).padStart(2, '0');
  } else {
    numberText = formatNumberSpec(value, numSpec);
  }

  let result = '';
  let numberEmitted = false;
  for (const t of tokens) {
    if (t.kind === 'lit') result += t.text;
    else if (t.kind === 'percent') result += '%';
    else if (t.kind === 'num') { result += numberText; numberEmitted = true; }
    else if (t.kind === 'sci') result += 'E' + expText;
  }
  if (!numberEmitted && (numSpec.length > 0 || sciTok)) result += numberText;
  return result;
}

function wrapTextLines(ctx: CanvasRenderingContext2D, text: string, maxWidth: number): string[] {
  const lines: string[] = [];
  // Hard line breaks (\n from Alt+Enter) always split regardless of wrapText.
  for (const paragraph of text.split('\n')) {
    lines.push(...wrapParagraphLines(ctx, paragraph, maxWidth));
  }
  return lines;
}

/** Codepoints in the CJK ranges get broken per-character (Excel / JIS X 4051
 *  behaviour). This mirrors the tokenizer in `layoutRichTextLines`. */
function isCJKCodePoint(cp: number): boolean {
  return (cp >= 0x3000 && cp <= 0x9FFF)  // CJK punctuation + CJK Unified Ideographs
      || (cp >= 0xF900 && cp <= 0xFAFF)  // CJK Compatibility Ideographs
      || (cp >= 0xAC00 && cp <= 0xD7AF)  // Hangul Syllables
      || (cp >= 0xFF00 && cp <= 0xFFEF); // Halfwidth/Fullwidth
}

/** Word-wrap a single paragraph (no embedded \n). Unlike a naive
 *  `split(' ')`, CJK characters are treated as individual break opportunities
 *  so that Japanese headings like "夏休みアクティビティ カレンダー 2026"
 *  actually wrap inside a merged cell. ECMA-376 doesn't spec the break
 *  algorithm but this matches what Excel renders on the same input. */
function wrapParagraphLines(ctx: CanvasRenderingContext2D, paragraph: string, maxWidth: number): string[] {
  const lines: string[] = [];
  // Tokenise: runs of non-space non-CJK, single ASCII-space runs, individual
  // CJK characters. Then greedy-fit each token onto the current line.
  const tokens: string[] = [];
  let i = 0;
  while (i < paragraph.length) {
    const ch = paragraph[i];
    const cp = ch.codePointAt(0) ?? 0;
    if (isCJKCodePoint(cp)) {
      tokens.push(ch);
      i += cp > 0xFFFF ? 2 : 1;
    } else if (ch === ' ') {
      let j = i;
      while (j < paragraph.length && paragraph[j] === ' ') j++;
      tokens.push(paragraph.slice(i, j));
      i = j;
    } else {
      let j = i;
      while (j < paragraph.length) {
        const c = paragraph[j];
        const p = c.codePointAt(0) ?? 0;
        if (c === ' ' || isCJKCodePoint(p)) break;
        j += p > 0xFFFF ? 2 : 1;
      }
      tokens.push(paragraph.slice(i, j));
      i = j;
    }
  }
  let current = '';
  for (const tok of tokens) {
    if (current === '') { current = tok; continue; }
    const candidate = current + tok;
    if (ctx.measureText(candidate).width <= maxWidth) {
      current = candidate;
    } else {
      // Token doesn't fit at the end of the current line — break here.
      // Leading spaces at the start of the next line are dropped (matches
      // Excel: wrapped-continuation lines don't preserve the space that
      // caused the break).
      lines.push(current);
      current = tok.replace(/^ +/, '');
      if (current === '') current = tok; // all-space token (preserve width on its own line)
    }
  }
  lines.push(current);
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
  worksheet: Worksheet;
  cellIndex: Map<string, Cell>;
  definedNames: Map<string, DefinedName>;
}

interface CfResult {
  fill?: Fill;
  fontColor?: string;
  fontBold?: boolean;
  fontItalic?: boolean;
  fontUnderline?: boolean;
  fontStrike?: boolean;
  /** Number format override from a matched CF dxf. Higher-priority rules win
   *  (first match through the rule list). Falls back to the cell's own style
   *  numFmt if unset. */
  numFmt?: { numFmtId: number; formatCode: string | null };
  dataBar?: { color: string; ratio: number; gradient: boolean };
  iconSet?: { name: string; index: number };
  /** Per-edge borders from matched CF rules (merged on top of the cell's base
   *  border). Mostly used by `expression` rules whose dxf only sets borders,
   *  e.g. highlighting today's column in a Gantt chart. */
  border?: Border;
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

function cellTextValue(cell: Cell | undefined): string | null {
  if (!cell) return null;
  if (cell.value.type === 'text') return cell.value.text;
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
  const cellIndex = new Map<string, Cell>();
  for (const row of worksheet.rows) {
    for (const c of row.cells) {
      cellIndex.set(`${c.row}:${c.col}`, c);
    }
  }
  const definedNames = new Map<string, DefinedName>();
  for (const dn of worksheet.definedNames ?? []) {
    definedNames.set(dn.name, dn);
  }
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
  // Excel evaluates CF rules in ascending priority (lowest number = highest
  // priority first). For each property (fill/fontColor/border/…) the first
  // matching rule wins, and `stopIfTrue` on a matching rule skips all later
  // rules. Match that here by iterating asc and only setting properties that
  // are still unset.
  compiled.sort((a, b) => {
    const pa = (a.rule as { priority: number }).priority ?? 0;
    const pb = (b.rule as { priority: number }).priority ?? 0;
    return pa - pb;
  });
  return { compiled, worksheet, cellIndex, definedNames };
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

function parseCellIsFormula(f: string): { text?: string; num?: number } {
  const t = f.trim();
  if (t.length >= 2 && t.startsWith('"') && t.endsWith('"')) {
    return { text: t.slice(1, -1).replace(/""/g, '"') };
  }
  const n = parseFloat(t);
  if (!isNaN(n)) return { num: n };
  return { text: t };
}

function cellIsTextMatch(text: string, operator: string, args: string[]): boolean {
  const a = args[0] ?? '';
  const b = args[1] ?? '';
  const ci = (s: string) => s.toLowerCase();
  switch (operator) {
    case 'equal':         return ci(text) === ci(a);
    case 'notEqual':      return ci(text) !== ci(a);
    case 'containsText':  return ci(text).includes(ci(a));
    case 'notContains':   return !ci(text).includes(ci(a));
    case 'beginsWith':    return ci(text).startsWith(ci(a));
    case 'endsWith':      return ci(text).endsWith(ci(a));
    case 'between':       return ci(text) >= ci(a) && ci(text) <= ci(b);
    case 'notBetween':    return ci(text) <  ci(a) || ci(text) >  ci(b);
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
  // First-match-wins (higher priority) for each property. See compileCf.
  // Per ECMA-376 §18.3.1.11, a `<dxf>` is a *differential* format: any child
  // element it contains is an override of the base cell format. So the mere
  // presence of `dxf.fill` means "replace the base fill with this", whatever
  // its patternType / color — including `patternType="none"` (explicit clear)
  // and gradient fills. The paint-site guard (`patternType !== 'none' &&
  // fgColor`) handles whether the result actually paints a color or leaves
  // the cell transparent, so this override stays spec-faithful without
  // second-guessing the fill's shape here.
  if (dxf.fill && !result.fill) result.fill = dxf.fill;
  if (dxf.font?.color && result.fontColor == null) result.fontColor = dxf.font.color;
  if (dxf.font?.bold && result.fontBold == null) result.fontBold = true;
  if (dxf.font?.italic && result.fontItalic == null) result.fontItalic = true;
  if (dxf.font?.underline && result.fontUnderline == null) result.fontUnderline = true;
  if (dxf.font?.strike && result.fontStrike == null) result.fontStrike = true;
  if (dxf.numFmt && result.numFmt == null) {
    result.numFmt = {
      numFmtId: dxf.numFmt.numFmtId,
      formatCode: dxf.numFmt.formatCode || null,
    };
  }
  if (dxf.border) {
    // Merge per-edge — higher-priority edges stay; lower-priority edges fill
    // in unset ones. dxf `border` typically sets only the edges the rule
    // cares about (e.g. left+right for a "today" column marker).
    const existing = result.border ?? {} as Border;
    const merged: Border = {
      left:         existing.left         ?? dxf.border.left,
      right:        existing.right        ?? dxf.border.right,
      top:          existing.top          ?? dxf.border.top,
      bottom:       existing.bottom       ?? dxf.border.bottom,
      diagonalUp:   existing.diagonalUp   ?? dxf.border.diagonalUp,
      diagonalDown: existing.diagonalDown ?? dxf.border.diagonalDown,
    };
    result.border = merged;
  }
}

function evaluateCf(cell: Cell | undefined, row: number, col: number, cfCtx: CfContext, dxfs: Dxf[]): CfResult {
  const result: CfResult = {};
  if (!cfCtx.compiled.length) return result;
  for (const entry of cfCtx.compiled) {
    if (!rangeContains(entry.sqref, row, col)) continue;
    const rule = entry.rule;
    const numVal = cellNumericValue(cell);

    if (rule.type === 'expression') {
      const anchor = entry.sqref[0];
      if (!anchor) continue;
      const matched = evalFormulaToBool(rule.formula, {
        row, col,
        anchorRow: anchor.top, anchorCol: anchor.left,
        cellIndex: cfCtx.cellIndex,
        definedNames: cfCtx.definedNames,
        depth: 0,
      });
      if (matched) {
        applyDxfToResult(result, rule.dxfId != null ? dxfs[rule.dxfId] : null);
        if (rule.stopIfTrue) break;
      }
      continue;
    }

    if (rule.type === 'cellIs') {
      const parsedArgs = rule.formulas.map(parseCellIsFormula);
      const textVal = cellTextValue(cell);
      let matched = false;
      if (numVal != null && parsedArgs.every(a => a.num != null)) {
        matched = cellIsMatch(numVal, rule.operator, parsedArgs.map(a => a.num!));
      } else if (textVal != null && parsedArgs.every(a => a.text != null)) {
        matched = cellIsTextMatch(textVal, rule.operator, parsedArgs.map(a => a.text!));
      }
      if (matched) {
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
      // Custom iconSets (Excel 2010+ x14 extension) override per-threshold icons.
      if (rule.customIcons && rule.customIcons[iconIdx]) {
        const ci = rule.customIcons[iconIdx];
        if (ci.iconSet !== 'NoIcons') {
          result.iconSet = { name: ci.iconSet, index: ci.iconId };
        }
      } else {
        result.iconSet = { name: rule.iconSet, index: iconIdx };
      }
    } else if (rule.type === 'colorScale') {
      if (numVal == null || !entry.scaleStops) continue;
      if (result.fill) continue;
      const color = colorScaleAt(numVal, rule.stops, entry.scaleStops);
      result.fill = { patternType: 'solid', fgColor: color, bgColor: color };
    } else if (rule.type === 'dataBar') {
      if (numVal == null || entry.barMin == null || entry.barMax == null) continue;
      if (result.dataBar) continue;
      const range = entry.barMax - entry.barMin;
      const ratio = range === 0 ? 0 : Math.max(0, Math.min(1, (numVal - entry.barMin) / range));
      result.dataBar = { color: rule.color, ratio, gradient: rule.gradient };
    }
  }
  return result;
}

// ────────────────────────────────────────────────────────────────
// Formula evaluator (conditional-formatting `expression` rules)
//
// Handles the narrow subset of Excel formulas used by CF expression rules:
// numeric/boolean literals, cell references (A1-style, with $ absolute
// markers), defined-name resolution, comparison/arithmetic operators, and
// a handful of functions (AND, OR, NOT, IF, ROUNDDOWN, ROUND, ROUNDUP,
// ISBLANK). Formula strings embed relative references that shift based on
// the evaluation cell's offset from an anchor cell:
//   - CF formulas use the top-left of the rule's `sqref` as anchor
//   - Workbook-level defined names are anchored at A1 (row 1, col 1)
// Column letters outside the defined-name anchor case are also shifted by
// the (col - anchorCol) delta; rows similarly. `$` markers pin the coord.
// ────────────────────────────────────────────────────────────────

interface EvalCtx {
  row: number;
  col: number;
  anchorRow: number;
  anchorCol: number;
  cellIndex: Map<string, Cell>;
  definedNames: Map<string, DefinedName>;
  /** Recursion guard for nested defined-name resolution. */
  depth: number;
}

type EvalScalar = number | boolean | string | null;
type EvalValue = EvalScalar | EvalScalar[];

/** Flatten nested scalars and arrays to a flat list of scalars. */
function flatten(v: EvalValue): EvalScalar[] {
  return Array.isArray(v) ? v : [v];
}

/** Unwrap an array value to its first scalar element (Excel's intersection
 *  behavior is not modeled; we collapse ranges to the first cell when a
 *  scalar is required). */
function toScalar(v: EvalValue): EvalScalar {
  return Array.isArray(v) ? (v[0] ?? 0) : v;
}

const MAX_DEFINED_NAME_DEPTH = 8;

function evalFormulaToBool(formula: string, ctx: EvalCtx): boolean {
  try {
    const v = evalFormula(formula, ctx);
    return toBool(v);
  } catch {
    return false;
  }
}

function toBool(v: EvalValue): boolean {
  const s = toScalar(v);
  if (typeof s === 'boolean') return s;
  if (typeof s === 'number') return s !== 0;
  if (typeof s === 'string') return s.length > 0 && s.toUpperCase() !== 'FALSE';
  return false;
}

function toNum(v: EvalValue): number {
  const s = toScalar(v);
  if (typeof s === 'number') return s;
  if (typeof s === 'boolean') return s ? 1 : 0;
  if (s == null) return 0;
  const n = parseFloat(String(s));
  return isNaN(n) ? 0 : n;
}

function toStr(v: EvalValue): string {
  const s = toScalar(v);
  if (s == null) return '';
  if (typeof s === 'boolean') return s ? 'TRUE' : 'FALSE';
  return String(s);
}

interface Tok {
  kind: 'num' | 'str' | 'op' | 'lparen' | 'rparen' | 'comma' | 'ref' | 'name' | 'bool' | 'colon';
  text: string;
  /** For 'ref': pre-parsed reference. */
  ref?: { colAbs: boolean; col: number; rowAbs: boolean; row: number };
}

const OP_CHARS = new Set(['<', '>', '=', '+', '-', '*', '/', '&', '^', '%']);

function tokenize(formula: string): Tok[] {
  const toks: Tok[] = [];
  let i = 0;
  const s = formula;
  while (i < s.length) {
    const c = s[i];
    if (c === ' ' || c === '\t' || c === '\n' || c === '\r') { i++; continue; }
    if (c === '(') { toks.push({ kind: 'lparen', text: c }); i++; continue; }
    if (c === ')') { toks.push({ kind: 'rparen', text: c }); i++; continue; }
    if (c === ',') { toks.push({ kind: 'comma', text: c }); i++; continue; }
    if (c === ':') { toks.push({ kind: 'colon', text: c }); i++; continue; }
    if (c === '"') {
      let j = i + 1; let buf = '';
      while (j < s.length) {
        if (s[j] === '"' && s[j + 1] === '"') { buf += '"'; j += 2; continue; }
        if (s[j] === '"') break;
        buf += s[j]; j++;
      }
      toks.push({ kind: 'str', text: buf });
      i = j + 1;
      continue;
    }
    if (c >= '0' && c <= '9') {
      let j = i;
      while (j < s.length && ((s[j] >= '0' && s[j] <= '9') || s[j] === '.')) j++;
      toks.push({ kind: 'num', text: s.slice(i, j) });
      i = j;
      continue;
    }
    if (OP_CHARS.has(c)) {
      // Multi-char operators: <=, >=, <>
      if ((c === '<' || c === '>') && (s[i + 1] === '=' || (c === '<' && s[i + 1] === '>'))) {
        toks.push({ kind: 'op', text: s.slice(i, i + 2) });
        i += 2;
      } else {
        toks.push({ kind: 'op', text: c });
        i++;
      }
      continue;
    }
    // Reference or identifier: may start with $, letters, or letters+digits.
    // Defined names allow letters, digits, '_', '.'; cell refs are
    // `$?[A-Z]+\$?[0-9]+` (case-insensitive).
    if (c === '$' || isIdentStart(c)) {
      let j = i;
      while (j < s.length && (s[j] === '$' || isIdentPart(s[j]))) j++;
      const text = s.slice(i, j);
      i = j;
      const ref = tryParseCellRef(text);
      if (ref) {
        toks.push({ kind: 'ref', text, ref });
      } else {
        const up = text.toUpperCase();
        if (up === 'TRUE' || up === 'FALSE') toks.push({ kind: 'bool', text: up });
        else toks.push({ kind: 'name', text });
      }
      continue;
    }
    // Unknown character — skip.
    i++;
  }
  return toks;
}

function isIdentStart(c: string): boolean {
  return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c === '_';
}

function isIdentPart(c: string): boolean {
  return isIdentStart(c) || (c >= '0' && c <= '9') || c === '.';
}

function tryParseCellRef(s: string): { colAbs: boolean; col: number; rowAbs: boolean; row: number } | null {
  // $?[A-Z]+\$?[0-9]+
  let i = 0;
  let colAbs = false, rowAbs = false;
  if (s[i] === '$') { colAbs = true; i++; }
  const colStart = i;
  while (i < s.length && s[i] >= 'A' && s[i].toUpperCase() <= 'Z') {
    if (!(s[i] >= 'A' && s[i] <= 'Z') && !(s[i] >= 'a' && s[i] <= 'z')) break;
    i++;
  }
  if (i === colStart) return null;
  const colLetters = s.slice(colStart, i).toUpperCase();
  if (s[i] === '$') { rowAbs = true; i++; }
  const rowStart = i;
  while (i < s.length && s[i] >= '0' && s[i] <= '9') i++;
  if (i === rowStart) return null;
  if (i !== s.length) return null;
  const rowNum = parseInt(s.slice(rowStart, i), 10);
  let col = 0;
  for (let k = 0; k < colLetters.length; k++) {
    col = col * 26 + (colLetters.charCodeAt(k) - 64);
  }
  return { colAbs, col, rowAbs, row: rowNum };
}

interface Parser {
  toks: Tok[];
  pos: number;
}

function evalFormula(formula: string, ctx: EvalCtx): EvalValue {
  const toks = tokenize(formula);
  const p: Parser = { toks, pos: 0 };
  const v = parseExpr(p, ctx);
  return v;
}

function peek(p: Parser): Tok | undefined { return p.toks[p.pos]; }
function consume(p: Parser): Tok | undefined { return p.toks[p.pos++]; }

function parseExpr(p: Parser, ctx: EvalCtx): EvalValue {
  return parseCmp(p, ctx);
}

function parseCmp(p: Parser, ctx: EvalCtx): EvalValue {
  let left = parseConcat(p, ctx);
  const t = peek(p);
  if (t && t.kind === 'op' && (t.text === '<' || t.text === '>' || t.text === '<=' || t.text === '>=' || t.text === '=' || t.text === '<>')) {
    consume(p);
    const right = parseConcat(p, ctx);
    return applyCmp(t.text, left, right);
  }
  return left;
}

function parseConcat(p: Parser, ctx: EvalCtx): EvalValue {
  let left = parseAdd(p, ctx);
  while (true) {
    const t = peek(p);
    if (!t || t.kind !== 'op' || t.text !== '&') break;
    consume(p);
    const right = parseAdd(p, ctx);
    left = toStr(left) + toStr(right);
  }
  return left;
}

function applyCmp(op: string, a: EvalValue, b: EvalValue): boolean {
  // Numeric-first comparison; fall back to string compare if either side is
  // a non-numeric string. Matches Excel's behavior for dates (stored as
  // serials) and arithmetic operations.
  const an = typeof a === 'string' && isNaN(parseFloat(a)) ? null : toNum(a);
  const bn = typeof b === 'string' && isNaN(parseFloat(b)) ? null : toNum(b);
  if (an !== null && bn !== null) {
    switch (op) {
      case '<':  return an <  bn;
      case '>':  return an >  bn;
      case '<=': return an <= bn;
      case '>=': return an >= bn;
      case '=':  return an === bn;
      case '<>': return an !== bn;
    }
  }
  const sa = String(a ?? ''); const sb = String(b ?? '');
  switch (op) {
    case '<':  return sa <  sb;
    case '>':  return sa >  sb;
    case '<=': return sa <= sb;
    case '>=': return sa >= sb;
    case '=':  return sa === sb;
    case '<>': return sa !== sb;
  }
  return false;
}

function parseAdd(p: Parser, ctx: EvalCtx): EvalValue {
  let left = parseMul(p, ctx);
  while (true) {
    const t = peek(p);
    if (!t || t.kind !== 'op' || (t.text !== '+' && t.text !== '-')) break;
    consume(p);
    const right = parseMul(p, ctx);
    left = t.text === '+' ? toNum(left) + toNum(right) : toNum(left) - toNum(right);
  }
  return left;
}

function parseMul(p: Parser, ctx: EvalCtx): EvalValue {
  let left = parseUnary(p, ctx);
  while (true) {
    const t = peek(p);
    if (!t || t.kind !== 'op' || (t.text !== '*' && t.text !== '/')) break;
    consume(p);
    const right = parseUnary(p, ctx);
    if (t.text === '*') left = toNum(left) * toNum(right);
    else {
      const rn = toNum(right);
      left = rn === 0 ? 0 : toNum(left) / rn;
    }
  }
  return left;
}

function parseUnary(p: Parser, ctx: EvalCtx): EvalValue {
  const t = peek(p);
  if (t && t.kind === 'op' && t.text === '-') { consume(p); return -toNum(parseUnary(p, ctx)); }
  if (t && t.kind === 'op' && t.text === '+') { consume(p); return toNum(parseUnary(p, ctx)); }
  return parsePrimary(p, ctx);
}

function parsePrimary(p: Parser, ctx: EvalCtx): EvalValue {
  const t = consume(p);
  if (!t) return 0;
  if (t.kind === 'num') return parseFloat(t.text);
  if (t.kind === 'str') return t.text;
  if (t.kind === 'bool') return t.text === 'TRUE';
  if (t.kind === 'lparen') {
    const v = parseExpr(p, ctx);
    const next = consume(p);
    if (!next || next.kind !== 'rparen') throw new Error('missing )');
    return v;
  }
  if (t.kind === 'ref') {
    // Range: `A1:B5` — resolve as array of cell values.
    if (peek(p)?.kind === 'colon') {
      consume(p);
      const right = consume(p);
      if (right?.kind !== 'ref' || !right.ref) throw new Error('range: expected ref after :');
      return resolveRange(t.ref!, right.ref, ctx);
    }
    return resolveRef(t.ref!, ctx);
  }
  if (t.kind === 'name') {
    // Function call: NAME(args)
    if (peek(p)?.kind === 'lparen') {
      consume(p);
      const args: EvalValue[] = [];
      if (peek(p)?.kind !== 'rparen') {
        args.push(parseExpr(p, ctx));
        while (peek(p)?.kind === 'comma') {
          consume(p);
          args.push(parseExpr(p, ctx));
        }
      }
      const next = consume(p);
      if (!next || next.kind !== 'rparen') throw new Error('missing )');
      return callFunc(t.text, args, ctx);
    }
    // Defined-name reference: substitute and evaluate.
    const dn = ctx.definedNames.get(t.text);
    if (dn && ctx.depth < MAX_DEFINED_NAME_DEPTH) {
      // Strip `SheetName!` prefix if present; keep just the ref body.
      const body = stripSheetPrefix(dn.formula);
      // Workbook-level defined names anchor at A1 for relative-ref shifts.
      const inner: EvalCtx = {
        ...ctx,
        anchorRow: 1,
        anchorCol: 1,
        depth: ctx.depth + 1,
      };
      return evalFormula(body, inner);
    }
    return 0;
  }
  return 0;
}

function stripSheetPrefix(formula: string): string {
  // Match `'Sheet Name'!ref` or `SheetName!ref`. Only the leading reference
  // prefix is stripped; we don't need cross-sheet lookups because defined
  // names here point to cells on the active sheet.
  const m = formula.match(/^(?:'[^']*'|[A-Za-z_][A-Za-z0-9_.]*)!(.*)$/);
  return m ? m[1] : formula;
}

function resolveRef(
  ref: { colAbs: boolean; col: number; rowAbs: boolean; row: number },
  ctx: EvalCtx,
): EvalScalar {
  const col = ref.colAbs ? ref.col : ref.col + (ctx.col - ctx.anchorCol);
  const row = ref.rowAbs ? ref.row : ref.row + (ctx.row - ctx.anchorRow);
  const cell = ctx.cellIndex.get(`${row}:${col}`);
  return cellValueToEval(cell);
}

function resolveRange(
  a: { colAbs: boolean; col: number; rowAbs: boolean; row: number },
  b: { colAbs: boolean; col: number; rowAbs: boolean; row: number },
  ctx: EvalCtx,
): EvalScalar[] {
  const ac = a.colAbs ? a.col : a.col + (ctx.col - ctx.anchorCol);
  const ar = a.rowAbs ? a.row : a.row + (ctx.row - ctx.anchorRow);
  const bc = b.colAbs ? b.col : b.col + (ctx.col - ctx.anchorCol);
  const br = b.rowAbs ? b.row : b.row + (ctx.row - ctx.anchorRow);
  const c1 = Math.min(ac, bc), c2 = Math.max(ac, bc);
  const r1 = Math.min(ar, br), r2 = Math.max(ar, br);
  const out: EvalScalar[] = [];
  // Cap range size to avoid pathological formulas like A:A (≈1M cells).
  // 4096 cells is plenty for CF use cases.
  const maxCells = 4096;
  for (let r = r1; r <= r2 && out.length < maxCells; r++) {
    for (let c = c1; c <= c2 && out.length < maxCells; c++) {
      out.push(cellValueToEval(ctx.cellIndex.get(`${r}:${c}`)));
    }
  }
  return out;
}

function cellValueToEval(cell: Cell | undefined): EvalScalar {
  // An empty / missing cell is *not* the same as 0. CF expressions like
  // `=$C5=0` or `NOT(ISBLANK($C5))` will match a missing cell if we return
  // 0 here, which is exactly the bug that turned C5-C8 beige on sample-10.
  // Return null and let the arithmetic / comparison operators coerce as
  // needed (null+0 → 0, null="" → true, null=0 → false). See ECMA-376
  // §18.18.62 and the actual Excel evaluation behaviour.
  if (!cell) return null;
  switch (cell.value.type) {
    case 'number': return cell.value.number;
    case 'bool':   return cell.value.bool;
    case 'text':   return cell.value.text;
    case 'error':  return null;
    case 'empty':
    default:       return null;
  }
}

function callFunc(nameRaw: string, args: EvalValue[], ctx: EvalCtx): EvalValue {
  const name = nameRaw.toUpperCase();
  switch (name) {
    // ── Logic ───────────────────────────────────────────────────────────────
    case 'AND':        return args.flatMap(flatten).every(a => toBool(a));
    case 'OR':         return args.flatMap(flatten).some(a => toBool(a));
    case 'NOT':        return !toBool(args[0]);
    case 'IF':         return toBool(args[0]) ? (args[1] ?? true) : (args[2] ?? false);
    case 'IFERROR':    return args[0] == null ? (args[1] ?? 0) : args[0];
    case 'IFS': {
      for (let i = 0; i + 1 < args.length; i += 2) {
        if (toBool(args[i])) return args[i + 1];
      }
      return null;
    }
    case 'TRUE':       return true;
    case 'FALSE':      return false;
    // ── Type checks ─────────────────────────────────────────────────────────
    case 'ISBLANK':    { const s = toScalar(args[0]); return s == null || s === ''; }
    case 'ISNUMBER':   return typeof toScalar(args[0]) === 'number';
    case 'ISTEXT':     return typeof toScalar(args[0]) === 'string';
    case 'ISNONTEXT':  return typeof toScalar(args[0]) !== 'string';
    case 'ISERROR':
    case 'ISERR':
    case 'ISNA':       return toScalar(args[0]) == null;
    case 'ISLOGICAL':  return typeof toScalar(args[0]) === 'boolean';
    // ── Rounding / math ─────────────────────────────────────────────────────
    case 'ROUNDDOWN': {
      const n = toNum(args[0]); const d = toNum(args[1]);
      const p = Math.pow(10, d);
      return (n >= 0 ? Math.floor(n * p) : Math.ceil(n * p)) / p;
    }
    case 'ROUNDUP': {
      const n = toNum(args[0]); const d = toNum(args[1]);
      const p = Math.pow(10, d);
      return (n >= 0 ? Math.ceil(n * p) : Math.floor(n * p)) / p;
    }
    case 'ROUND': {
      const n = toNum(args[0]); const d = toNum(args[1]);
      const p = Math.pow(10, d);
      return Math.round(n * p) / p;
    }
    case 'INT':        return Math.floor(toNum(args[0]));
    case 'TRUNC':      { const n = toNum(args[0]); const d = toNum(args[1] ?? 0); const p = Math.pow(10, d); return (n >= 0 ? Math.floor(n * p) : Math.ceil(n * p)) / p; }
    case 'CEILING':    { const n = toNum(args[0]); const sig = toNum(args[1] ?? 1); return sig === 0 ? 0 : Math.ceil(n / sig) * sig; }
    case 'FLOOR':      { const n = toNum(args[0]); const sig = toNum(args[1] ?? 1); return sig === 0 ? 0 : Math.floor(n / sig) * sig; }
    case 'MOD':        { const a = toNum(args[0]); const b = toNum(args[1]); return b === 0 ? null : a - Math.floor(a / b) * b; }
    case 'POWER':      return Math.pow(toNum(args[0]), toNum(args[1]));
    case 'SQRT':       { const n = toNum(args[0]); return n < 0 ? null : Math.sqrt(n); }
    case 'ABS':        return Math.abs(toNum(args[0]));
    case 'SIGN':       { const n = toNum(args[0]); return n > 0 ? 1 : n < 0 ? -1 : 0; }
    case 'EXP':        return Math.exp(toNum(args[0]));
    case 'LN':         { const n = toNum(args[0]); return n <= 0 ? null : Math.log(n); }
    case 'LOG10':      { const n = toNum(args[0]); return n <= 0 ? null : Math.log10(n); }
    // ── Aggregates ──────────────────────────────────────────────────────────
    case 'MIN':        { const ns = args.flatMap(flatten).filter(v => typeof v === 'number') as number[]; return ns.length ? Math.min(...ns) : 0; }
    case 'MAX':        { const ns = args.flatMap(flatten).filter(v => typeof v === 'number') as number[]; return ns.length ? Math.max(...ns) : 0; }
    case 'SUM':        return args.flatMap(flatten).reduce<number>((s, v) => s + (typeof v === 'number' ? v : 0), 0);
    case 'AVERAGE':    { const ns = args.flatMap(flatten).filter(v => typeof v === 'number') as number[]; return ns.length ? ns.reduce((s, v) => s + v, 0) / ns.length : null; }
    case 'COUNT':      return args.flatMap(flatten).filter(v => typeof v === 'number').length;
    case 'COUNTA':     return args.flatMap(flatten).filter(v => v != null && v !== '').length;
    case 'COUNTBLANK': return args.flatMap(flatten).filter(v => v == null || v === '').length;
    case 'COUNTIF':    return countIf(flatten(args[0]), args[1]);
    case 'SUMIF':      return sumIf(flatten(args[0]), args[1], args[2] !== undefined ? flatten(args[2]) : null);
    case 'AVERAGEIF':  {
      const src = flatten(args[0]);
      const sum = sumIf(src, args[1], args[2] !== undefined ? flatten(args[2]) : null);
      const count = countIf(src, args[1]);
      return count === 0 ? null : toNum(sum) / count;
    }
    // ── Text ────────────────────────────────────────────────────────────────
    case 'LEN':        return toStr(args[0]).length;
    case 'LEFT':       return toStr(args[0]).slice(0, Math.max(0, toNum(args[1] ?? 1)));
    case 'RIGHT':      { const s = toStr(args[0]); const n = Math.max(0, toNum(args[1] ?? 1)); return n >= s.length ? s : s.slice(s.length - n); }
    case 'MID':        { const s = toStr(args[0]); const start = Math.max(1, toNum(args[1])) - 1; const len = Math.max(0, toNum(args[2])); return s.slice(start, start + len); }
    case 'UPPER':      return toStr(args[0]).toUpperCase();
    case 'LOWER':      return toStr(args[0]).toLowerCase();
    case 'TRIM':       return toStr(args[0]).replace(/\s+/g, ' ').trim();
    case 'EXACT':      return toStr(args[0]) === toStr(args[1]);
    case 'FIND':       { const needle = toStr(args[0]); const hay = toStr(args[1]); const start = Math.max(1, toNum(args[2] ?? 1)) - 1; const idx = hay.indexOf(needle, start); return idx < 0 ? null : idx + 1; }
    case 'SEARCH':     { const needle = toStr(args[0]).toLowerCase(); const hay = toStr(args[1]).toLowerCase(); const start = Math.max(1, toNum(args[2] ?? 1)) - 1; const idx = hay.indexOf(needle, start); return idx < 0 ? null : idx + 1; }
    case 'CONCATENATE':
    case 'CONCAT':     return args.flatMap(flatten).map(v => v == null ? '' : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v)).join('');
    case 'T':          { const s = toScalar(args[0]); return typeof s === 'string' ? s : ''; }
    case 'N':          { const s = toScalar(args[0]); return typeof s === 'number' ? s : typeof s === 'boolean' ? (s ? 1 : 0) : 0; }
    case 'VALUE':      return toNum(args[0]);
    // ── Reference ───────────────────────────────────────────────────────────
    case 'ROW':        return ctx.row;    // no-arg form only (current cell row)
    case 'COLUMN':     return ctx.col;    // no-arg form only (current cell col)
    // ── Date / time ─────────────────────────────────────────────────────────
    case 'TODAY':      return todaySerial();
    case 'NOW':        return nowSerial();
    case 'DATE':       return dateToSerial(toNum(args[0]), toNum(args[1]), toNum(args[2]));
    case 'YEAR':       return serialToDate(toNum(args[0])).y;
    case 'MONTH':      return serialToDate(toNum(args[0])).m;
    case 'DAY':        return serialToDate(toNum(args[0])).d;
    case 'WEEKDAY':    {
      // return type 1 (Sun=1..Sat=7) default; type 2 = Mon=1..Sun=7; type 3 = Mon=0..Sun=6.
      const d = serialToJsDate(toNum(args[0]));
      const jsDow = d.getUTCDay(); // Sun=0..Sat=6
      const rt = toNum(args[1] ?? 1);
      if (rt === 2) return jsDow === 0 ? 7 : jsDow;
      if (rt === 3) return jsDow === 0 ? 6 : jsDow - 1;
      return jsDow + 1;
    }
    default:
      return 0;
  }
}

function countIf(source: EvalScalar[], criteria: EvalValue): number {
  const pred = makeCriteriaPredicate(criteria);
  let n = 0;
  for (const v of source) if (pred(v)) n++;
  return n;
}

function sumIf(source: EvalScalar[], criteria: EvalValue, sumRange: EvalScalar[] | null): number {
  const pred = makeCriteriaPredicate(criteria);
  const target = sumRange ?? source;
  let sum = 0;
  for (let i = 0; i < source.length; i++) {
    if (pred(source[i])) {
      const t = target[i];
      if (typeof t === 'number') sum += t;
    }
  }
  return sum;
}

/** Build a predicate matching Excel's COUNTIF/SUMIF criteria syntax:
 *  a bare value (exact match) or a string like ">5", "<>foo", "=100". */
function makeCriteriaPredicate(criteria: EvalValue): (v: EvalScalar) => boolean {
  const raw = toScalar(criteria);
  if (typeof raw !== 'string') {
    const rn = typeof raw === 'number' ? raw : null;
    return (v) => {
      if (rn !== null && typeof v === 'number') return v === rn;
      return v === raw;
    };
  }
  const m = raw.match(/^(<=|>=|<>|<|>|=)(.*)$/);
  const op = m ? m[1] : '=';
  const rhsStr = m ? m[2] : raw;
  const rhsNum = rhsStr.trim() === '' ? NaN : parseFloat(rhsStr);
  const rhsIsNum = !isNaN(rhsNum) && /^-?\d+(\.\d+)?$/.test(rhsStr.trim());
  return (v) => {
    if (rhsIsNum && typeof v === 'number') {
      switch (op) {
        case '<':  return v <  rhsNum;
        case '>':  return v >  rhsNum;
        case '<=': return v <= rhsNum;
        case '>=': return v >= rhsNum;
        case '<>': return v !== rhsNum;
        default:   return v === rhsNum;
      }
    }
    const sv = v == null ? '' : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v);
    switch (op) {
      case '<>': return sv !== rhsStr;
      case '<':  return sv <  rhsStr;
      case '>':  return sv >  rhsStr;
      case '<=': return sv <= rhsStr;
      case '>=': return sv >= rhsStr;
      default:   return sv === rhsStr;
    }
  };
}

// Excel date serial: 1 = 1900-01-01, treats 1900 as leap (serial 60 = fake
// 1900-02-29). For dates ≥ 1900-03-01, offset to Unix epoch is 25569 days.
const EXCEL_EPOCH_OFFSET = 25569;
const MS_PER_DAY = 86400000;

function todaySerial(): number {
  const d = new Date();
  const utcMid = Date.UTC(d.getFullYear(), d.getMonth(), d.getDate());
  return Math.floor(utcMid / MS_PER_DAY) + EXCEL_EPOCH_OFFSET;
}

function nowSerial(): number {
  return Date.now() / MS_PER_DAY + EXCEL_EPOCH_OFFSET;
}

function dateToSerial(y: number, m: number, d: number): number {
  // Excel rolls over out-of-range months/days (e.g. DATE(2019, 13, 1) = Jan 2020).
  const ms = Date.UTC(y, m - 1, d);
  return Math.floor(ms / MS_PER_DAY) + EXCEL_EPOCH_OFFSET;
}

function serialToJsDate(serial: number): Date {
  const ms = (Math.floor(serial) - EXCEL_EPOCH_OFFSET) * MS_PER_DAY;
  return new Date(ms);
}

function serialToDate(serial: number): { y: number; m: number; d: number } {
  const d = serialToJsDate(serial);
  return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate() };
}

// ────────────────────────────────────────────────────────────────
// Shared state for a single renderViewport call
// ────────────────────────────────────────────────────────────────
interface RenderContext {
  worksheet: Worksheet;
  styles: Styles;
  cellMap: Map<string, Cell>;
  mergeAnchorMap: Map<string, { totalW: number; totalH: number; right: number; bottom: number }>;
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
  /** row:col keys for cells that carry a comment; renderer draws a small
   *  red triangle in the top-right corner (ECMA-376 §18.7.3 commentList). */
  commentCells: Set<string>;
  /** row:col → table-style overlay (bold header, banded rows, borders). */
  tableStyleMap: Map<string, TableCellStyle>;
  /** row:col → render-ready SparklineModel for cells that host an
   *  `x14:sparkline`. Built once at viewport start by flattening the
   *  parser's SparklineGroup + per-cell Sparkline pair. */
  sparklineMap: Map<string, SparklineModel>;
  onTextRun?: (info: TextRunInfo) => void;
}

// ────────────────────────────────────────────────────────────────
// Icon Set drawing
// ────────────────────────────────────────────────────────────────
const ICON_COLORS_3 = ['#FF0000', '#FFFF00', '#00B050'];
const ICON_COLORS_4 = ['#FF0000', '#FF6600', '#FFFF00', '#00B050'];
const ICON_COLORS_5 = ['#FF0000', '#FF6600', '#FFFF00', '#92D050', '#00B050'];

function drawCfIcon(ctx: CanvasRenderingContext2D, name: string, index: number, x: number, y: number, sz: number): void {
  if (name === 'NoIcons') return;
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
// Excel Table style overlays (ECMA-376 §18.5)
// ────────────────────────────────────────────────────────────────
// We don't ship the full built-in table-style catalog — instead we derive a
// single "accent" color from the style name and overlay bold header + banded
// fills + horizontal rules so that `TableStyle*` files render with visible
// structure rather than as blank ranges.
export interface TableCellStyle {
  accent: string;
  isHeader: boolean;
  isTotals: boolean;
  /** `true` when this is a banded data row that should get the stripe fill. */
  isBanded: boolean;
  isFirstCol: boolean;
  isLastCol: boolean;
  isTopEdge: boolean;
  isBottomEdge: boolean;
  /** Dxf for the whole-table element of a custom `<tableStyle>`
   *  (ECMA-376 §18.8.40). Border/fill apply to every cell as a base layer. */
  wholeTableDxf?: number;
  /** Dxf for the header-row element of a custom `<tableStyle>`. Provides
   *  header fill, font color/weight, and vertical separators. */
  headerRowDxf?: number;
}

function buildTableStyleMap(worksheet: Worksheet): Map<string, TableCellStyle> {
  const map = new Map<string, TableCellStyle>();
  for (const t of worksheet.tables ?? []) {
    const accent = t.accentColor || '#808080';
    const hdr = Math.max(0, t.headerRowCount ?? 1);
    const tot = Math.max(0, t.totalsRowCount ?? 0);
    const { top, bottom, left, right } = t.range;
    const headerEnd = top + hdr - 1;
    const totalsStart = bottom - tot + 1;
    for (let r = top; r <= bottom; r++) {
      const isHeader = hdr > 0 && r <= headerEnd;
      const isTotals = tot > 0 && r >= totalsStart;
      const dataIdx = (!isHeader && !isTotals) ? (r - headerEnd - 1) : -1;
      for (let c = left; c <= right; c++) {
        map.set(`${r}:${c}`, {
          accent,
          isHeader,
          isTotals,
          isBanded: t.showRowStripes && dataIdx >= 0 && dataIdx % 2 === 1,
          isFirstCol: t.showFirstColumn && c === left,
          isLastCol: t.showLastColumn && c === right,
          isTopEdge: r === top,
          isBottomEdge: r === bottom,
          wholeTableDxf: t.wholeTableDxf,
          headerRowDxf: t.headerRowDxf,
        });
      }
    }
  }
  return map;
}

/** Flatten the worksheet's parsed `sparklineGroups` into a per-cell render
 *  model. Each Sparkline inherits its group's formatting; min/max are
 *  computed from the values when the group's `*AxisType` is `individual`,
 *  shared across the group when `group`, or taken from `manualMin/Max` when
 *  `custom`. The renderer can then look up `row:col` and call core's
 *  `renderSparkline` without further work. */
function buildSparklineMap(worksheet: Worksheet): Map<string, SparklineModel> {
  const map = new Map<string, SparklineModel>();
  for (const g of worksheet.sparklineGroups ?? []) {
    // Group-wide min/max if needed.
    let groupMin = Infinity, groupMax = -Infinity;
    if (g.minAxisType === 'group' || g.maxAxisType === 'group') {
      for (const sl of g.sparklines) {
        for (const v of sl.values) {
          if (typeof v === 'number') {
            if (v < groupMin) groupMin = v;
            if (v > groupMax) groupMax = v;
          }
        }
      }
      if (!isFinite(groupMin) || !isFinite(groupMax)) {
        groupMin = 0; groupMax = 1;
      }
    }
    for (const sl of g.sparklines) {
      const numeric = sl.values.filter((v): v is number => typeof v === 'number');
      const indMin = numeric.length ? Math.min(...numeric) : 0;
      const indMax = numeric.length ? Math.max(...numeric) : 1;
      const min = g.minAxisType === 'custom' && typeof g.manualMin === 'number'
        ? g.manualMin
        : g.minAxisType === 'group' ? groupMin : indMin;
      const max = g.maxAxisType === 'custom' && typeof g.manualMax === 'number'
        ? g.manualMax
        : g.maxAxisType === 'group' ? groupMax : indMax;
      map.set(`${sl.row}:${sl.col}`, {
        kind: g.kind,
        values: sl.values,
        min,
        max,
        displayEmptyCellsAs: (g.displayEmptyCellsAs === 'zero' || g.displayEmptyCellsAs === 'span')
          ? g.displayEmptyCellsAs
          : 'gap',
        displayXAxis: g.displayXAxis,
        lineWeight: g.lineWeight,
        markers: g.markers,
        high: g.high,
        low: g.low,
        first: g.first,
        last: g.last,
        negative: g.negative,
        colorSeries: g.colorSeries,
        colorNegative: g.colorNegative,
        colorAxis: g.colorAxis,
        colorMarkers: g.colorMarkers,
        colorFirst: g.colorFirst,
        colorLast: g.colorLast,
        colorHigh: g.colorHigh,
        colorLow: g.colorLow,
      });
    }
  }
  return map;
}

function stripeColorFor(accent: string): string {
  // Light tint of the accent — mimics TableStyleLight* banded rows.
  const hex = accent.replace('#', '');
  if (hex.length < 6) return '#F2F2F2';
  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  const mix = (ch: number) => Math.round(ch * 0.2 + 255 * 0.8);
  const toHex = (v: number) => v.toString(16).padStart(2, '0').toUpperCase();
  return `#${toHex(mix(r))}${toHex(mix(g))}${toHex(mix(b))}`;
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
      fillDataBar(ctx, cf.dataBar.color, aCx + bInset, aCy + bInset, bW, cH - bInset * 2, cf.dataBar.gradient);
    }
    const mergedBorder = resolveMergeBorder(border, aRow, aCol, info.right, info.bottom, rc.cellMap, styles);
    renderBorder(ctx, mergeBorders(mergedBorder, cf.border), aCx, aCy, cW, cH);

    if (!cell) continue;
    const text = formatCellValue(cell, styles, cf.numFmt);
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
      const tableStyle = rc.tableStyleMap.get(key);
      // Custom `<tableStyle>` dxfs (ECMA-376 §18.8.40). When present, they
      // drive header fill / font color and inter-row borders instead of the
      // built-in accent fallback.
      const tsDxfWhole = (tableStyle?.wholeTableDxf != null)
        ? (styles.dxfs ?? [])[tableStyle.wholeTableDxf] : undefined;
      const tsDxfHeader = (tableStyle?.headerRowDxf != null)
        ? (styles.dxfs ?? [])[tableStyle.headerRowDxf] : undefined;

      // Background fill (base or CF override). ECMA-376 §18.8.22 ST_PatternType.
      // - solid/gray*: blend fgColor with bgColor at the pattern's fg coverage.
      // - directional hatches (dark/light Horizontal/Vertical/Down/Up/Grid/
      //   Trellis): render via a small repeating tile using createPattern so
      //   the hatch actually shows, rather than approximating as a blend.
      if (effectiveFill.gradient && effectiveFill.gradient.stops.length > 0) {
        ctx.fillStyle = buildGradientFill(ctx, effectiveFill.gradient, cx, cy, cellW, cellH);
        ctx.fillRect(cx, cy, cellW, cellH);
      } else if (effectiveFill.patternType && effectiveFill.patternType !== 'none' && effectiveFill.fgColor) {
        const pt = effectiveFill.patternType;
        const bg = effectiveFill.bgColor ?? 'FFFFFF';
        const hatch = hatchPattern(ctx, pt, effectiveFill.fgColor, bg);
        if (hatch) {
          ctx.fillStyle = hatch;
        } else {
          const coverage = patternCoverage(pt);
          ctx.fillStyle = coverage >= 1
            ? hexToRgba(effectiveFill.fgColor)
            : blendHex(effectiveFill.fgColor, bg, coverage);
        }
        ctx.fillRect(cx, cy, cellW, cellH);
      } else if (tableStyle && tableStyle.isHeader && tsDxfHeader?.fill?.fgColor) {
        ctx.fillStyle = hexToRgba(tsDxfHeader.fill.fgColor);
        ctx.fillRect(cx, cy, cellW, cellH);
      } else if (tableStyle && !tableStyle.isHeader && !tableStyle.isTotals && tsDxfWhole?.fill?.fgColor) {
        ctx.fillStyle = hexToRgba(tsDxfWhole.fill.fgColor);
        ctx.fillRect(cx, cy, cellW, cellH);
      } else if (tableStyle && tableStyle.isBanded) {
        ctx.fillStyle = stripeColorFor(tableStyle.accent);
        ctx.fillRect(cx, cy, cellW, cellH);
      }

      // Comment indicator triangle — drawn above fill but below borders so
      // borders still read cleanly around the cell edge.
      if (rc.commentCells.has(key)) {
        drawCommentMarker(ctx, cx, cy, cellW, cellH);
      }

      // DataBar (drawn inside the cell, left-anchored). Excel 2010+ renders
      // these with a horizontal gradient by default.
      if (cf.dataBar && cf.dataBar.ratio > 0) {
        const barInset = 2;
        const barW = Math.max(0, (cellW - barInset * 2) * cf.dataBar.ratio);
        fillDataBar(ctx, cf.dataBar.color, cx + barInset, cy + barInset, barW, cellH - barInset * 2, cf.dataBar.gradient);
      }

      // Sparkline (Office 2010 x14:sparklineGroup). Drawn after the cell
      // background but before borders / text so borders frame the sparkline
      // and any cell text overlays it (matches Excel's z-order, and lets
      // a label like "Trend" share the same cell as the sparkline).
      const sparkModel = rc.sparklineMap.get(key);
      if (sparkModel) {
        renderSparkline(ctx, { x: cx, y: cy, w: cellW, h: cellH }, sparkModel);
      }

      // Grid lines – draw only right + bottom edges once per cell (avoids double-drawing at
      // shared cell boundaries). Half-device-pixel offset (0.5/dpr) aligns each line to the
      // device pixel grid so we get a crisp 1-device-pixel result.
      // Skipped when the sheet has `<sheetView showGridLines="0">` (View →
      // Gridlines unchecked; ECMA-376 §18.3.1.83).
      if (rc.worksheet.showGridlines !== false) {
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

      // Cell borders (base + any CF borders overlaid via per-edge merge).
      // For merged anchors, combine the anchor's border with the right/bottom
      // edges from the constituent cells on those edges (ECMA-376 §18.3.1.55).
      const baseBorder = mergeInfo
        ? resolveMergeBorder(border, rowIndex, colIndex, mergeInfo.right, mergeInfo.bottom, cellMap, styles)
        : border;
      renderBorder(ctx, mergeBorders(baseBorder, cf.border), cx, cy, cellW, cellH);

      // Excel Table style overlay: thin horizontal rules between rows and a
      // thicker bottom edge under the header row (ECMA-376 §18.5). Drawn on
      // top of cell borders so an empty-border data cell still shows table
      // structure.
      if (tableStyle) {
        const horiz = tsDxfWhole?.border?.horizontal;
        const vert  = tsDxfWhole?.border?.vertical;
        const wtTop = tsDxfWhole?.border?.top;
        const wtBot = tsDxfWhole?.border?.bottom;
        const wtLeft = tsDxfWhole?.border?.left;
        const wtRight = tsDxfWhole?.border?.right;
        const hdrBot = tsDxfHeader?.border?.bottom;
        const hdrTop = tsDxfHeader?.border?.top;
        const hasDxfBorder = !!(horiz || vert || wtTop || wtBot || wtLeft || wtRight || hdrBot || hdrTop);
        if (hasDxfBorder) {
          const overlay: Border = { left: null, right: null, top: null, bottom: null };
          if (tableStyle.isTopEdge) overlay.top = wtTop ?? null;
          else if (horiz) overlay.top = horiz;
          if (tableStyle.isHeader && hdrBot) overlay.bottom = hdrBot;
          else if (tableStyle.isBottomEdge) overlay.bottom = wtBot ?? null;
          else if (horiz) overlay.bottom = horiz;
          if (tableStyle.isFirstCol || colIndex === 0) overlay.left = wtLeft ?? null;
          if (tableStyle.isLastCol) overlay.right = wtRight ?? null;
          // Outer table left/right edges
          renderBorder(ctx, overlay, cx, cy, cellW, cellH);
        } else {
          const hp = 0.5 / dpr;
          ctx.strokeStyle = tableStyle.accent;
          ctx.lineWidth = tableStyle.isHeader ? 1.5 : 1;
          ctx.beginPath();
          ctx.moveTo(cx, cy + cellH - hp);
          ctx.lineTo(cx + cellW, cy + cellH - hp);
          if (tableStyle.isTopEdge) {
            ctx.moveTo(cx, cy + hp);
            ctx.lineTo(cx + cellW, cy + hp);
          }
          ctx.stroke();
        }
      }

      // AutoFilter dropdown indicator
      if (rc.autoFilterCells.has(key)) {
        drawAutoFilterArrow(ctx, cx, cy, cw, cellH);
      }

      if (!cell) continue;
      const text = formatCellValue(cell, styles, cf.numFmt);
      if (!text || (text === '0' && rc.worksheet.showZeros === false)) continue;

      const tableBold = !!(tableStyle && (tableStyle.isHeader || tableStyle.isTotals));
      const effectiveBold = font.bold || !!cf.fontBold || tableBold;
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
      // Custom table-style header dxfs can override font color (ECMA-376 §18.8.40).
      const tableFontColor =
        (tableStyle?.isHeader && tsDxfHeader?.font?.color) ? tsDxfHeader.font.color :
        (tableStyle && !tableStyle.isHeader && !tableStyle.isTotals && tsDxfWhole?.font?.color) ? tsDxfWhole.font.color :
        null;
      const textColor = hyperlinkUrl
        ? '#0563C1'
        : (cf.fontColor ?? tableFontColor ?? font.color);
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

      // Text overflow into adjacent empty cells (ECMA-376 §18.3.1.4 "spans"
      // — Excel behavior when `wrapText=false`). Left-aligned text flows
      // rightward, right-aligned flows leftward, and centered splits evenly.
      // We only extend the clip rect; the text itself is still drawn once.
      // Stops at merge-cell boundaries, non-empty cells, and iconSet-left
      // overrun (since an icon sits inside this cell's left padding).
      let drawX = cx;
      let drawW = cellW;
      // Excel only overflows text into adjacent empty cells; numeric values
      // that don't fit are rendered as "####" (they never spill). Cells
      // containing hard line breaks render multi-line in place, so they don't
      // overflow either.
      const hasHardBreak = text.includes('\n');
      if (!mergeInfo && !xf.wrapText && !xf.textRotation && !isNumeric && !hasHardBreak) {
        const textW = ctx.measureText(text).width;
        const textPx = textW + leftPad + paddingX;
        if (textPx > cellW) {
          const overflow = textPx - cellW;
          let extendRight = 0;
          let extendLeft = 0;
          if (alignH === 'right') {
            extendLeft = overflow;
          } else if (alignH === 'center') {
            extendLeft = overflow / 2;
            extendRight = overflow / 2;
          } else {
            extendRight = overflow;
          }
          if (extendRight > 0) {
            let budget = extendRight;
            for (let oci = ci + 1; oci < numCols && budget > 0; oci++) {
              const adjKey = `${rowIndex}:${startCol + oci}`;
              if (mergeSkipSet.has(adjKey) || mergeAnchorMap.has(adjKey)) break;
              const adjCell = cellMap.get(adjKey);
              if (adjCell && adjCell.value.type !== 'empty') break;
              drawW += colWidths[oci];
              budget -= colWidths[oci];
            }
          }
          if (extendLeft > 0) {
            let budget = extendLeft;
            for (let oci = ci - 1; oci >= 0 && budget > 0; oci--) {
              const adjKey = `${rowIndex}:${startCol + oci}`;
              if (mergeSkipSet.has(adjKey) || mergeAnchorMap.has(adjKey)) break;
              const adjCell = cellMap.get(adjKey);
              if (adjCell && adjCell.value.type !== 'empty') break;
              drawX -= colWidths[oci];
              drawW += colWidths[oci];
              budget -= colWidths[oci];
            }
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
      ctx.rect(drawX, cy, drawW, cellH);
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
        // Hard line breaks (\n from Alt+Enter) render as multiple lines even
        // when wrapText is false — this matches Excel's behavior.
        if (text.includes('\n')) {
          const lines = text.split('\n');
          const lineH = Math.round(font.size * ROW_HEIGHT_TO_PX * 1.2);
          const totalTextH = lines.length * lineH;
          let startY: number;
          if (alignV === 'top') { startY = cy + paddingY; ctx.textBaseline = 'top'; }
          else if (alignV === 'center') { startY = cy + (cellH - totalTextH) / 2; ctx.textBaseline = 'top'; }
          else { startY = cy + cellH - totalTextH - paddingY; ctx.textBaseline = 'top'; }
          for (let li = 0; li < lines.length; li++) {
            ctx.fillText(lines[li], textX, startY + li * lineH);
          }
        } else {
          let textY: number;
          if (alignV === 'top') { ctx.textBaseline = 'top'; textY = cy + paddingY; }
          else if (alignV === 'center') { ctx.textBaseline = 'middle'; textY = cy + cellH / 2; }
          else { ctx.textBaseline = 'bottom'; textY = cy + cellH - paddingY; }
          ctx.fillText(text, textX, textY);
        }
      }

      ctx.restore();

      if (text && rc.onTextRun) {
        rc.onTextRun({ text, x: cx, y: cy, width: cellW, height: cellH, row: rowIndex, col: colIndex });
      }
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

  const mergeAnchorMap = new Map<string, { totalW: number; totalH: number; right: number; bottom: number }>();
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
    mergeAnchorMap.set(`${mc.top}:${mc.left}`, { totalW, totalH, right: mc.right, bottom: mc.bottom });
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

  // Build commented-cell lookup. worksheet.commentRefs are A1-style refs
  // ("A1", "B12", "AA3") so we convert each to "row:col" and stash in a Set
  // for O(1) membership checks in the cell loop.
  const commentCells = new Set<string>();
  for (const ref of worksheet.commentRefs ?? []) {
    const parsed = parseA1Ref(ref);
    if (parsed) commentCells.add(`${parsed.row}:${parsed.col}`);
  }

  const tableStyleMap = buildTableStyleMap(worksheet);
  const sparklineMap = buildSparklineMap(worksheet);

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
    commentCells,
    tableStyleMap,
    sparklineMap,
    onTextRun: opts.onTextRun,
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

  // ── Anchored shape groups (custom geometry, incl. embedded images) ────
  if (worksheet.shapeGroups && worksheet.shapeGroups.length > 0) {
    renderShapeGroups(
      ctx, worksheet, cs,
      startRow, startCol,
      scrollOffsetX, scrollOffsetY,
      scrollAreaX, scrollAreaY,
      scrollAreaW, scrollAreaH,
      opts.loadedImages,
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

  // ── Anchored slicers (Office 2010+ pivot/table filter buttons) ──
  if (worksheet.slicers && worksheet.slicers.length > 0) {
    renderSlicers(
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

function renderShapeGroups(
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
  loadedImages?: Map<string, HTMLImageElement>,
): void {
  if (scrollAreaW <= 0 || scrollAreaH <= 0) return;
  const anchors = ws.shapeGroups;
  if (!anchors || anchors.length === 0) return;

  const scrollOriginSheetX = sheetXForCol(ws, startCol, cs);
  const scrollOriginSheetY = sheetYForRow(ws, startRow, cs);

  ctx.save();
  ctx.beginPath();
  ctx.rect(scrollAreaX, scrollAreaY, scrollAreaW, scrollAreaH);
  ctx.clip();

  for (const anchor of anchors) {
    const fromCol1 = anchor.fromCol + 1;
    const fromRow1 = anchor.fromRow + 1;
    const toCol1   = anchor.toCol   + 1;
    const toRow1   = anchor.toRow   + 1;

    const x1 = sheetXForCol(ws, fromCol1, cs) + (anchor.fromColOff * cs) / EMU_PER_PX;
    const y1 = sheetYForRow(ws, fromRow1, cs) + (anchor.fromRowOff * cs) / EMU_PER_PX;
    const x2 = sheetXForCol(ws, toCol1,   cs) + (anchor.toColOff   * cs) / EMU_PER_PX;
    const y2 = sheetYForRow(ws, toRow1,   cs) + (anchor.toRowOff   * cs) / EMU_PER_PX;
    const w = x2 - x1;
    const h = y2 - y1;
    if (w <= 0 || h <= 0) continue;

    const canvasX = scrollAreaX + (x1 - scrollOriginSheetX) - scrollOffsetX;
    const canvasY = scrollAreaY + (y1 - scrollOriginSheetY) - scrollOffsetY;

    if (canvasX + w < scrollAreaX || canvasX > scrollAreaX + scrollAreaW) continue;
    if (canvasY + h < scrollAreaY || canvasY > scrollAreaY + scrollAreaH) continue;

    for (const shape of anchor.shapes) {
      const sx = canvasX + shape.x * w;
      const sy = canvasY + shape.y * h;
      const sw = shape.w * w;
      const sh = shape.h * h;
      if (sw <= 0 || sh <= 0) continue;
      drawShape(ctx, shape, sx, sy, sw, sh, loadedImages);
    }
  }

  ctx.restore();
}

function drawShape(
  ctx: CanvasRenderingContext2D,
  shape: ShapeInfo,
  sx: number, sy: number, sw: number, sh: number,
  loadedImages?: Map<string, HTMLImageElement>,
): void {
  ctx.save();
  if (shape.rot !== 0) {
    ctx.translate(sx + sw / 2, sy + sh / 2);
    ctx.rotate((shape.rot * Math.PI) / 180);
    ctx.translate(-sw / 2, -sh / 2);
  } else {
    ctx.translate(sx, sy);
  }

  if (shape.geom.type === 'custom') {
    for (const path of shape.geom.paths) {
      if (path.w <= 0 || path.h <= 0) continue;
      const kx = sw / path.w;
      const ky = sh / path.h;
      ctx.beginPath();
      // Track pen position for arcTo center computation.
      let penX = 0, penY = 0;
      // Track subpath start for close lineTo.
      let subX = 0, subY = 0;
      for (const cmd of path.commands) {
        switch (cmd.op) {
          case 'moveTo': {
            const px = cmd.x * kx, py = cmd.y * ky;
            ctx.moveTo(px, py);
            penX = subX = px; penY = subY = py;
            break;
          }
          case 'lineTo': {
            const px = cmd.x * kx, py = cmd.y * ky;
            ctx.lineTo(px, py);
            penX = px; penY = py;
            break;
          }
          case 'cubicBezTo': {
            const ex = cmd.x3 * kx, ey = cmd.y3 * ky;
            ctx.bezierCurveTo(
              cmd.x1 * kx, cmd.y1 * ky,
              cmd.x2 * kx, cmd.y2 * ky,
              ex, ey,
            );
            penX = ex; penY = ey;
            break;
          }
          case 'quadBezTo': {
            const ex = cmd.x2 * kx, ey = cmd.y2 * ky;
            ctx.quadraticCurveTo(cmd.x1 * kx, cmd.y1 * ky, ex, ey);
            penX = ex; penY = ey;
            break;
          }
          case 'arcTo': {
            // ECMA-376 §20.1.9.3: pen lies on ellipse at stAng;
            // derive center from pen + stAng, then sweep swAng.
            const rx = cmd.wr * kx, ry = cmd.hr * ky;
            if (rx <= 0 || ry <= 0) break;
            const stRad = (cmd.stAng / 60000) * (Math.PI / 180);
            const swRad = (cmd.swAng / 60000) * (Math.PI / 180);
            const cx = penX - Math.cos(stRad) * rx;
            const cy = penY - Math.sin(stRad) * ry;
            const endRad = stRad + swRad;
            ctx.ellipse(cx, cy, rx, ry, 0, stRad, endRad, swRad < 0);
            penX = cx + Math.cos(endRad) * rx;
            penY = cy + Math.sin(endRad) * ry;
            break;
          }
          case 'close':
            ctx.closePath();
            penX = subX; penY = subY;
            break;
        }
      }
      fillAndStroke(ctx, shape);
    }
  } else if (shape.geom.type === 'preset') {
    ctx.beginPath();
    switch (shape.geom.name) {
      case 'ellipse':
      case 'roundRect': {
        const rx = sw / 2, ry = sh / 2;
        ctx.ellipse(rx, ry, rx, ry, 0, 0, Math.PI * 2);
        break;
      }
      default:
        ctx.rect(0, 0, sw, sh);
    }
    fillAndStroke(ctx, shape);
  } else if (shape.geom.type === 'image') {
    // Image leaf inside a group (e.g. a sun-emoji clip-art nested in the
    // calendar header). The caller pre-decodes every data URL seen in
    // `ws.shapeGroups[*].shapes[*].geom` via XlsxWorkbook.renderViewport,
    // so we should normally have it in `loadedImages`. If not, fall back
    // to a silent skip — drawing an empty rect would look worse.
    const img = loadedImages?.get(shape.geom.dataUrl);
    if (img) {
      ctx.drawImage(img, 0, 0, sw, sh);
    }
  }
  ctx.restore();
}

function fillAndStroke(ctx: CanvasRenderingContext2D, shape: ShapeInfo): void {
  if (shape.fillColor) {
    ctx.fillStyle = shape.fillColor;
    ctx.fill();
  }
  if (shape.strokeColor && shape.strokeWidth > 0) {
    ctx.strokeStyle = shape.strokeColor;
    ctx.lineWidth = Math.max(0.5, shape.strokeWidth / EMU_PER_PX);
    ctx.stroke();
  }
}

// ────────────────────────────────────────────────────────────────
// Border drawing
// ────────────────────────────────────────────────────────────────
/**
 * Overlay any CF-rule border edges on top of the cell's base border. CF
 * borders win per-edge where they set a style (e.g. a red left+right for a
 * "today" column marker replaces the underlying edge only, leaving top/bottom
 * from the base style intact).
 */
/**
 * Resolve the outer border of a merged range. Excel keeps each constituent
 * cell's own style; the merged rectangle's outer edges come from the cells
 * on those edges (e.g. the right edge from the rightmost column's `right`),
 * not from the anchor cell alone. Without this the right border of an
 * `E2:F2` merge — stored on F2 — goes missing.
 */
function resolveMergeBorder(
  anchorBorder: Border,
  anchorRow: number,
  anchorCol: number,
  rightCol: number,
  bottomRow: number,
  cellMap: Map<string, Cell>,
  styles: Styles,
): Border {
  if (rightCol === anchorCol && bottomRow === anchorRow) return anchorBorder;
  const edgeBorder = (r: number, c: number): Border | null => {
    if (r === anchorRow && c === anchorCol) return null;
    const cell = cellMap.get(`${r}:${c}`);
    if (!cell) return null;
    return resolveXf(styles, cell.styleIndex).border;
  };
  const rightB  = edgeBorder(anchorRow, rightCol);
  const bottomB = edgeBorder(bottomRow, anchorCol);
  const cornerB = edgeBorder(bottomRow, rightCol);
  const pick = (primary: BorderEdge | null | undefined, ...rest: Array<BorderEdge | null | undefined>): BorderEdge | null => {
    if (primary?.style) return primary;
    for (const r of rest) if (r?.style) return r;
    return primary ?? null;
  };
  return {
    left:         anchorBorder.left,
    top:          anchorBorder.top,
    right:        pick(rightB?.right,   cornerB?.right,   anchorBorder.right),
    bottom:       pick(bottomB?.bottom, cornerB?.bottom,  anchorBorder.bottom),
    diagonalUp:   anchorBorder.diagonalUp ?? null,
    diagonalDown: anchorBorder.diagonalDown ?? null,
  };
}

function mergeBorders(base: Border, overlay: Border | undefined): Border {
  if (!overlay) return base;
  const pick = (a: BorderEdge | null | undefined, b: BorderEdge | null | undefined): BorderEdge | null =>
    (b && b.style) ? b : (a ?? null);
  return {
    left:         pick(base.left,         overlay.left),
    right:        pick(base.right,        overlay.right),
    top:          pick(base.top,          overlay.top),
    bottom:       pick(base.bottom,       overlay.bottom),
    diagonalUp:   pick(base.diagonalUp,   overlay.diagonalUp),
    diagonalDown: pick(base.diagonalDown, overlay.diagonalDown),
  };
}

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
    catAxisFormatCode: chart.catAxisFormatCode ?? null,
    catAxisMin: chart.catAxisMin ?? null,
    catAxisMax: chart.catAxisMax ?? null,
    titleFontBold: chart.titleFontBold ?? null,
    catAxisFontBold: chart.catAxisFontBold ?? null,
    valAxisFontBold: chart.valAxisFontBold ?? null,
    catAxisCrosses: chart.catAxisCrosses ?? null,
    catAxisCrossesAt: chart.catAxisCrossesAt ?? null,
    valAxisCrosses: chart.valAxisCrosses ?? null,
    valAxisCrossesAt: chart.valAxisCrossesAt ?? null,
    catAxisLineColor: chart.catAxisLineColor ?? null,
    catAxisLineWidthEmu: chart.catAxisLineWidthEmu ?? null,
    valAxisLineColor: chart.valAxisLineColor ?? null,
    valAxisLineWidthEmu: chart.valAxisLineWidthEmu ?? null,
    series: chart.series.map(s => ({
      name: s.name,
      color: s.color ?? null,
      values: s.values,
      seriesType: s.seriesType ?? null,
      categories: s.categories.length > 0 ? s.categories : null,
      showMarker: s.showMarker ?? null,
      valFormatCode: s.valFormatCode ?? null,
      markerSymbol: s.markerSymbol ?? null,
      markerSize: s.markerSize ?? null,
      markerFill: s.markerFill ?? null,
      markerLine: s.markerLine ?? null,
      dataPointOverrides: s.dataPointOverrides ?? null,
      dataLabelOverrides: s.dataLabelOverrides ?? null,
      seriesDataLabels: s.seriesDataLabels ?? null,
      errBars: s.errBars ?? null,
    })),
    showDataLabels: chart.showDataLabels ?? false,
    valMin: chart.valAxisMin ?? null,
    valMax: chart.valAxisMax ?? null,
    catAxisTitle: chart.catAxisTitle ?? null,
    valAxisTitle: chart.valAxisTitle ?? null,
    catAxisHidden: chart.catAxisHidden ?? false,
    valAxisHidden: chart.valAxisHidden ?? false,
    plotAreaBg: null,
    // `<c:chartSpace><c:spPr>` resolution: when the spPr element was present
    // we honor whatever it said (solid hex or `<a:noFill/>` → null =
    // transparent). When spPr was absent the file is relying on the Excel
    // default, which is an opaque white chart area — keep that so legacy
    // charts still get their familiar frame.
    chartBg: chart.hasChartSpPr ? (chart.chartBg ?? null) : 'FFFFFF',
    legendManualLayout: chart.legendManualLayout ?? null,
    // <c:legend> is the authoritative signal: present → show, absent → hide.
    // A single-series bar chart in Excel typically omits <c:legend>, so we
    // must honor that rather than deriving from series count.
    showLegend: chart.showLegend ?? false,
    legendPos: chart.legendPos ?? null,
    catAxisCrossBetween: 'between',
    // Default `out` per ECMA-376 §21.2.2.49 ST_TickMark when the spec
    // didn't say. (We previously hard-coded `cross` which made every
    // chart pretend it had crossing ticks even when the file said
    // none / out.)
    valAxisMajorTickMark: chart.valAxisMajorTickMark ?? 'out',
    catAxisMajorTickMark: chart.catAxisMajorTickMark ?? 'out',
    valAxisMinorTickMark: chart.valAxisMinorTickMark ?? null,
    catAxisMinorTickMark: chart.catAxisMinorTickMark ?? null,
    titleFontSizeHpt: chart.titleFontSizeHpt ?? null,
    titleFontColor: chart.titleFontColor ?? null,
    titleFontFace: chart.titleFontFace ?? null,
    catAxisFontSizeHpt: chart.catAxisFontSizeHpt ?? null,
    valAxisFontSizeHpt: chart.valAxisFontSizeHpt ?? null,
    dataLabelFontSizeHpt: null,
    subtotalIndices: [],
    valAxisFormatCode: chart.valAxisFormatCode ?? null,
    barGapWidth: chart.barGapWidth ?? null,
    barOverlap: chart.barOverlap ?? null,
    dataLabelPosition: chart.dataLabelPosition ?? null,
    dataLabelFontColor: chart.dataLabelFontColor ?? null,
    dataLabelFormatCode: chart.dataLabelFormatCode ?? null,
    titleManualLayout: chart.titleManualLayout ?? null,
    plotAreaManualLayout: chart.plotAreaManualLayout ?? null,
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

    // XLSX natural rendering is device-px at 96 DPI where 1pt = 4/3 px. Scale
    // that by `cs` so OOXML-specified font sizes (title/axes) scale with zoom.
    const ptToPx = (4 / 3) * cs;
    renderChart(ctx, adaptChartData(anchor.chart), { x: cx, y: cy, w: cw, h: ch }, ptToPx);
    ctx.restore();
  }
}

// ── renderSlicers ───────────────────────────────────────────────────────────
//
// Office 2010+ pivot / table slicer. We don't own a slicer engine, so this
// renders a static button bank: header with the slicer caption, then one
// button per saved item using the selection flags from the slicerCache. The
// visual language (pale blue outline, white "selected" buttons on a darker
// background, gray "deselected" buttons) intentionally mirrors Excel's
// default slicer style — the workbook may ship a custom `slicerStyle` but
// rendering that is deferred (the built-in look is already recognisable).

const SLICER_HEADER_FONT = '600 12px "Meiryo UI", "Segoe UI", sans-serif';
const SLICER_ITEM_FONT   = '11px "Meiryo UI", "Segoe UI", sans-serif';
const SLICER_BG           = '#FFFFFF';
const SLICER_BORDER       = '#BFBFBF';
const SLICER_HEADER_BG    = '#F2F2F2';
const SLICER_HEADER_FG    = '#404040';
const SLICER_ITEM_SEL_BG  = '#FFFFFF';
const SLICER_ITEM_SEL_FG  = '#000000';
const SLICER_ITEM_SEL_BD  = '#A5A5A5';
const SLICER_ITEM_OFF_BG  = '#E7E6E6';
const SLICER_ITEM_OFF_FG  = '#A6A6A6';
const SLICER_ITEM_OFF_BD  = '#C6C6C6';

function renderSlicers(
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
  const slicers = ws.slicers;
  if (!slicers) return;

  const scrollOriginSheetX = sheetXForCol(ws, startCol, cs);
  const scrollOriginSheetY = sheetYForRow(ws, startRow, cs);

  for (const anchor of slicers) {
    const fromCol1 = anchor.fromCol + 1;
    const fromRow1 = anchor.fromRow + 1;
    const toCol1   = anchor.toCol   + 1;
    const toRow1   = anchor.toRow   + 1;

    const shX1 = sheetXForCol(ws, fromCol1, cs) + (anchor.fromColOff * cs) / EMU_PER_PX;
    const shY1 = sheetYForRow(ws, fromRow1, cs) + (anchor.fromRowOff * cs) / EMU_PER_PX;
    const shX2 = sheetXForCol(ws, toCol1,   cs) + (anchor.toColOff   * cs) / EMU_PER_PX;
    const shY2 = sheetYForRow(ws, toRow1,   cs) + (anchor.toRowOff   * cs) / EMU_PER_PX;

    const w = shX2 - shX1;
    const h = shY2 - shY1;
    if (w <= 0 || h <= 0) continue;

    const x = scrollAreaX + (shX1 - scrollOriginSheetX) - scrollOffsetX;
    const y = scrollAreaY + (shY1 - scrollOriginSheetY) - scrollOffsetY;

    if (x + w < scrollAreaX || x > scrollAreaX + scrollAreaW) continue;
    if (y + h < scrollAreaY || y > scrollAreaY + scrollAreaH) continue;

    ctx.save();
    ctx.beginPath();
    ctx.rect(scrollAreaX, scrollAreaY, scrollAreaW, scrollAreaH);
    ctx.clip();

    drawSlicerFrame(ctx, anchor.caption, anchor.items, x, y, w, h, cs);

    ctx.restore();
  }
}

function drawSlicerFrame(
  ctx: CanvasRenderingContext2D,
  caption: string,
  items: SlicerItem[],
  x: number,
  y: number,
  w: number,
  h: number,
  cs: number,
): void {
  // Outer frame (white with a soft gray hairline).
  ctx.fillStyle = SLICER_BG;
  ctx.fillRect(x, y, w, h);
  ctx.strokeStyle = SLICER_BORDER;
  ctx.lineWidth = 1;
  ctx.strokeRect(x + 0.5, y + 0.5, w - 1, h - 1);

  // Header band with caption.
  const headerH = Math.max(20 * cs, 14);
  ctx.fillStyle = SLICER_HEADER_BG;
  ctx.fillRect(x + 1, y + 1, w - 2, headerH);
  ctx.fillStyle = SLICER_HEADER_FG;
  ctx.font = scaleFont(SLICER_HEADER_FONT, cs);
  ctx.textBaseline = 'middle';
  ctx.textAlign = 'left';
  const headerPad = 6 * cs;
  drawClippedText(ctx, caption, x + headerPad, y + headerH / 2 + 1, w - 2 * headerPad);

  // Item buttons. Items expand to fill the available height (up to a
  // minimum button size) and are clipped by the slicer rect. We don't
  // implement scroll arrows because this renderer is non-interactive.
  if (items.length === 0) return;
  const gap = Math.max(1, Math.round(2 * cs));
  const innerPad = 4 * cs;
  const listX = x + innerPad;
  const listY = y + headerH + innerPad;
  const listW = w - 2 * innerPad;
  const listH = h - headerH - 2 * innerPad;
  if (listW <= 0 || listH <= 0) return;

  // Prefer Excel's rough row height (~20 sheet-px) but compress if the
  // slicer is shallow so at least the first items fit.
  const preferredItemH = Math.max(18 * cs, 16);
  const maxVisibleByH = Math.max(1, Math.floor((listH + gap) / (preferredItemH + gap)));
  const visible = Math.min(items.length, maxVisibleByH);
  const itemH = Math.min(preferredItemH, (listH - gap * (visible - 1)) / visible);
  if (itemH <= 0) return;

  ctx.font = scaleFont(SLICER_ITEM_FONT, cs);
  const itemPad = 8 * cs;
  for (let i = 0; i < visible; i++) {
    const item = items[i];
    const iy = listY + i * (itemH + gap);
    const selected = item.selected;
    ctx.fillStyle = selected ? SLICER_ITEM_SEL_BG : SLICER_ITEM_OFF_BG;
    ctx.fillRect(listX, iy, listW, itemH);
    ctx.strokeStyle = selected ? SLICER_ITEM_SEL_BD : SLICER_ITEM_OFF_BD;
    ctx.lineWidth = 1;
    ctx.strokeRect(listX + 0.5, iy + 0.5, listW - 1, itemH - 1);
    ctx.fillStyle = selected ? SLICER_ITEM_SEL_FG : SLICER_ITEM_OFF_FG;
    drawClippedText(ctx, item.name, listX + itemPad, iy + itemH / 2 + 1, listW - 2 * itemPad);
  }
}

function scaleFont(css: string, cs: number): string {
  // Re-scale the leading `<size>px` token by `cs`. Safe fallback: leave the
  // string as-is so the slicer remains readable when parsing fails.
  return css.replace(/(\d+(?:\.\d+)?)px/, (_, n) => `${Math.round(Number(n) * cs)}px`);
}

function drawClippedText(
  ctx: CanvasRenderingContext2D,
  text: string,
  x: number,
  y: number,
  maxWidth: number,
): void {
  if (maxWidth <= 0) return;
  let s = text;
  if (ctx.measureText(s).width > maxWidth) {
    const ellipsis = '…';
    while (s.length > 0 && ctx.measureText(s + ellipsis).width > maxWidth) {
      s = s.slice(0, -1);
    }
    s = s.length > 0 ? s + ellipsis : '';
  }
  ctx.fillText(s, x, y);
}
