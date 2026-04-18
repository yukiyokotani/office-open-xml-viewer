import type {
  Worksheet, Styles, Cell, CellValue, Font, Fill, Border, CellXf,
  ViewportRange, RenderViewportOptions,
  CfRule, CellRange, CfStop, CfValue, Dxf,
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

function resolveXf(styles: Styles, styleIndex: number): { font: Font; fill: Fill; border: Border; xf: CellXf } {
  const xf: CellXf = styles.cellXfs[styleIndex] ?? styles.cellXfs[0] ?? {
    fontId: 0, fillId: 0, borderId: 0, numFmtId: 0, alignH: null, alignV: null, wrapText: false,
  };
  const font: Font = styles.fonts[xf.fontId] ?? { bold: false, italic: false, underline: false, size: DEFAULT_FONT_SIZE, color: null, name: null };
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

function applyFormat(num: number, numFmtId: number, formatCode: string | null): string {
  const isDateFmtId = (id: number) => (id >= 14 && id <= 17) || id === 22;
  if (isDateFmtId(numFmtId)) return formatExcelDate(num);
  if (formatCode) return applyFormatCode(num, formatCode);
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

function formatExcelDate(serial: number): string {
  const date = new Date((serial - 25569) * 86400 * 1000);
  return date.toLocaleDateString();
}

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

  const { styles, cellMap, mergeAnchorMap, mergeSkipSet, cfContext, cs } = rc;
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

      // Grid line
      ctx.strokeStyle = '#d0d0d0';
      ctx.lineWidth = 0.5 * cs;
      ctx.strokeRect(cx + 0.5, cy + 0.5, cellW - 1, cellH - 1);

      // Cell borders
      renderBorder(ctx, border, cx, cy, cellW, cellH, cs);

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

      // Text overflow for left-aligned non-merged non-wrap cells
      let drawW = cellW;
      if (!mergeInfo && !xf.wrapText && alignH !== 'right' && alignH !== 'center') {
        const textPx = ctx.measureText(text).width + paddingX * 2;
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
        textX = cx + paddingX;
        textAlign = 'left';
      }

      ctx.textAlign = textAlign;

      ctx.save();
      ctx.beginPath();
      ctx.rect(cx, cy, drawW, cellH);
      ctx.clip();

      if (xf.wrapText) {
        const lines = wrapTextLines(ctx, text, cellW - paddingX * 2);
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
        if (font.underline) {
          const metrics = ctx.measureText(text);
          const tW = Math.min(metrics.width, drawW - paddingX * 2);
          const uy = alignV === 'top'
            ? cy + paddingY + Math.round(font.size * ROW_HEIGHT_TO_PX) + 1
            : alignV === 'center'
              ? cy + cellH / 2 + Math.round(font.size * ROW_HEIGHT_TO_PX * 0.55)
              : cy + cellH - paddingY + 1;
          const ux = alignH === 'right' ? cx + cellW - paddingX - tW
            : alignH === 'center' ? cx + cellW / 2 - tW / 2
            : cx + paddingX;
          ctx.save();
          ctx.strokeStyle = font.color ? hexToRgba(font.color) : '#000000';
          ctx.lineWidth = 1 * cs;
          ctx.beginPath(); ctx.moveTo(ux, uy); ctx.lineTo(ux + tW, uy); ctx.stroke();
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

  // ── Row/col headers (drawn last, always on top) ──────────────
  renderHeaders(ctx, canvasW, canvasH,
    startRow, startCol, numRows, numCols,
    scrollColWidths, scrollRowHeights,
    scrollOffsetX, scrollOffsetY,
    frozenColWidths, frozenRowHeights,
    frozenW, frozenH,
    hw, hh, cs,
  );

  // ── Freeze pane separator lines ──────────────────────────────
  if (freezeRows > 0) {
    ctx.save();
    ctx.strokeStyle = FREEZE_LINE_COLOR;
    ctx.lineWidth = 1 * cs;
    ctx.beginPath();
    ctx.moveTo(hw, scrollAreaY + 0.5);
    ctx.lineTo(canvasW, scrollAreaY + 0.5);
    ctx.stroke();
    ctx.restore();
  }
  if (freezeCols > 0) {
    ctx.save();
    ctx.strokeStyle = FREEZE_LINE_COLOR;
    ctx.lineWidth = 1 * cs;
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
  hw: number, hh: number, cs: number,
): void {
  const HEADER_BG = '#f8f9fa';
  const HEADER_BORDER = '#c8ccd0';
  const HEADER_TEXT = '#444';
  const headerFontSize = Math.max(1, Math.round(11 * cs));
  const HEADER_FONT = `${headerFontSize}px ${DEFAULT_FONT_FAMILY}`;
  const scrollAreaX = hw + frozenW;
  const scrollAreaY = hh + frozenH;

  // Corner
  ctx.fillStyle = HEADER_BG;
  ctx.fillRect(0, 0, hw, hh);
  ctx.strokeStyle = HEADER_BORDER;
  ctx.lineWidth = 1 * cs;
  ctx.strokeRect(0.5, 0.5, hw - 1, hh - 1);

  ctx.font = HEADER_FONT;
  ctx.fillStyle = HEADER_TEXT;

  // Helper: draw one column header cell
  const drawColHeader = (col: number, cx: number, cw: number) => {
    ctx.fillStyle = HEADER_BG;
    ctx.fillRect(cx, 0, cw, hh);
    ctx.strokeStyle = HEADER_BORDER;
    ctx.lineWidth = 0.5 * cs;
    ctx.strokeRect(cx + 0.5, 0.5, cw - 1, hh - 1);
    ctx.fillStyle = HEADER_TEXT;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText(colToLetter(col), cx + cw / 2, hh / 2);
  };

  // Helper: draw one row header cell
  const drawRowHeader = (row: number, cy: number, ch: number) => {
    ctx.fillStyle = HEADER_BG;
    ctx.fillRect(0, cy, hw, ch);
    ctx.strokeStyle = HEADER_BORDER;
    ctx.lineWidth = 0.5 * cs;
    ctx.strokeRect(0.5, cy + 0.5, hw - 1, ch - 1);
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
// Border drawing
// ────────────────────────────────────────────────────────────────
function renderBorder(ctx: CanvasRenderingContext2D, border: Border, x: number, y: number, w: number, h: number, cs = 1): void {
  const edges = [
    { edge: border.top,    x1: x,     y1: y,     x2: x + w, y2: y },
    { edge: border.bottom, x1: x,     y1: y + h, x2: x + w, y2: y + h },
    { edge: border.left,   x1: x,     y1: y,     x2: x,     y2: y + h },
    { edge: border.right,  x1: x + w, y1: y,     x2: x + w, y2: y + h },
  ];
  for (const { edge, x1, y1, x2, y2 } of edges) {
    if (!edge || !edge.style || edge.style === 'none') continue;
    ctx.beginPath();
    ctx.strokeStyle = edge.color ? hexToRgba(edge.color) : '#000000';
    ctx.lineWidth = borderStyleWidth(edge.style) * cs;
    if (edge.style === 'dashed' || edge.style === 'dotted' || edge.style === 'dashDot') {
      ctx.setLineDash(edge.style === 'dotted' ? [2, 2] : [4, 2]);
    } else {
      ctx.setLineDash([]);
    }
    ctx.moveTo(x1, y1);
    ctx.lineTo(x2, y2);
    ctx.stroke();
    ctx.setLineDash([]);
  }
}

function borderStyleWidth(style: string): number {
  switch (style) {
    case 'thick': return 2;
    case 'medium': case 'mediumDashed': case 'mediumDashDot': return 1.5;
    default: return 1;
  }
}
