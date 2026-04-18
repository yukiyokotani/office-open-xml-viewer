import type { Worksheet, Styles, Cell, CellValue, Font, Fill, Border, CellXf, ViewportRange, RenderViewportOptions } from './types.js';

const DEFAULT_FONT_FAMILY = 'Calibri, Arial, sans-serif';
const DEFAULT_FONT_SIZE = 11;
const MDW = 7;
const ROW_HEIGHT_TO_PX = 4 / 3; // pt → px at 96 DPI

// Width of the row-number header column (px)
export const HEADER_W = 50;
// Height of the column-letter header row (px)
export const HEADER_H = 22;

// OOXML spec: pixel = trunc(((256*w + 128/MDW) / 256) * MDW)
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

function buildFont(font: Font): string {
  const style = font.italic ? 'italic ' : '';
  const weight = font.bold ? 'bold ' : '';
  const sizePx = Math.round(font.size * ROW_HEIGHT_TO_PX);
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
  if (cleaned.includes('%')) {
    return (num * 100).toFixed(countDecimalPlaces(cleaned)) + '%';
  }
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

export function renderViewport(
  ctx: CanvasRenderingContext2D,
  worksheet: Worksheet,
  styles: Styles,
  viewport: ViewportRange,
  opts: RenderViewportOptions = {},
): void {
  const canvasW = ctx.canvas.width / (opts.dpr ?? 1);
  const canvasH = ctx.canvas.height / (opts.dpr ?? 1);
  const scrollOffsetX = opts.scrollOffsetX ?? 0;
  const scrollOffsetY = opts.scrollOffsetY ?? 0;

  ctx.clearRect(0, 0, canvasW, canvasH);
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, canvasW, canvasH);

  const { row: startRow, col: startCol, rows: numRows, cols: numCols } = viewport;

  // Column X positions (relative to cell area origin; may start negative due to scroll offset)
  const colXs: number[] = [];
  const colWidths: number[] = [];
  let x = -scrollOffsetX;
  for (let c = startCol; c < startCol + numCols; c++) {
    colXs.push(x);
    const w = colWidthToPx(worksheet.colWidths[c] ?? worksheet.defaultColWidth);
    colWidths.push(w);
    x += w;
  }

  // Row Y positions (relative to cell area origin)
  const rowYs: number[] = [];
  const rowHeights: number[] = [];
  let y = -scrollOffsetY;
  for (let r = startRow; r < startRow + numRows; r++) {
    rowYs.push(y);
    const h = rowHeightToPx(worksheet.rowHeights[r] ?? worksheet.defaultRowHeight);
    rowHeights.push(h);
    y += h;
  }

  // Build cell lookup
  const cellMap = new Map<string, Cell>();
  for (const row of worksheet.rows) {
    for (const cell of row.cells) {
      cellMap.set(`${cell.row}:${cell.col}`, cell);
    }
  }

  // Build merge lookup
  const mergeAnchorMap = new Map<string, { totalW: number; totalH: number }>();
  const mergeSkipSet = new Set<string>();
  for (const mc of worksheet.mergeCells ?? []) {
    let totalW = 0;
    for (let c = mc.left; c <= mc.right; c++) {
      totalW += colWidthToPx(worksheet.colWidths[c] ?? worksheet.defaultColWidth);
    }
    let totalH = 0;
    for (let r = mc.top; r <= mc.bottom; r++) {
      totalH += rowHeightToPx(worksheet.rowHeights[r] ?? worksheet.defaultRowHeight);
    }
    mergeAnchorMap.set(`${mc.top}:${mc.left}`, { totalW, totalH });
    for (let r = mc.top; r <= mc.bottom; r++) {
      for (let c = mc.left; c <= mc.right; c++) {
        if (r === mc.top && c === mc.left) continue;
        mergeSkipSet.add(`${r}:${c}`);
      }
    }
  }

  // --- Render cells (clipped to cell area, excluding headers) ---
  ctx.save();
  ctx.beginPath();
  ctx.rect(HEADER_W, HEADER_H, canvasW - HEADER_W, canvasH - HEADER_H);
  ctx.clip();

  for (let ri = 0; ri < numRows; ri++) {
    const rowIndex = startRow + ri;
    const cy = HEADER_H + rowYs[ri];
    const ch = rowHeights[ri];

    for (let ci = 0; ci < numCols; ci++) {
      const colIndex = startCol + ci;
      const cx = HEADER_W + colXs[ci];
      const cw = colWidths[ci];

      const key = `${rowIndex}:${colIndex}`;
      if (mergeSkipSet.has(key)) continue;

      const mergeInfo = mergeAnchorMap.get(key);
      const cellW = mergeInfo ? mergeInfo.totalW : cw;
      const cellH = mergeInfo ? mergeInfo.totalH : ch;

      const cell = cellMap.get(key);
      const styleIndex = cell?.styleIndex ?? 0;
      const { font, fill, border, xf } = resolveXf(styles, styleIndex);

      // Background fill
      if (fill.patternType !== 'none' && fill.patternType !== '' && fill.fgColor) {
        ctx.fillStyle = hexToRgba(fill.fgColor);
        ctx.fillRect(cx, cy, cellW, cellH);
      }

      // Grid line
      if (!mergeInfo) {
        ctx.strokeStyle = '#d0d0d0';
        ctx.lineWidth = 0.5;
        ctx.strokeRect(cx + 0.5, cy + 0.5, cw - 1, ch - 1);
      } else {
        // Draw grid lines for the bounding box of the merge
        ctx.strokeStyle = '#d0d0d0';
        ctx.lineWidth = 0.5;
        ctx.strokeRect(cx + 0.5, cy + 0.5, cellW - 1, cellH - 1);
      }

      // Border edges
      renderBorder(ctx, border, cx, cy, cellW, cellH);

      if (!cell) continue;
      const text = formatCellValue(cell, styles);
      if (!text) continue;

      ctx.font = buildFont(font);
      ctx.fillStyle = font.color ? hexToRgba(font.color) : '#000000';

      const paddingX = 3;
      const paddingY = 2;
      const alignH = xf.alignH ?? (cell.value.type === 'number' ? 'right' : 'left');
      const alignV = xf.alignV ?? 'bottom';

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
      ctx.rect(cx, cy, cellW, cellH);
      ctx.clip();

      if (xf.wrapText) {
        const lines = wrapTextLines(ctx, text, cellW - paddingX * 2);
        const lineH = Math.round(font.size * ROW_HEIGHT_TO_PX * 1.2);
        const totalTextH = lines.length * lineH;
        let startY: number;
        if (alignV === 'top') {
          startY = cy + paddingY;
          ctx.textBaseline = 'top';
        } else if (alignV === 'center') {
          startY = cy + (cellH - totalTextH) / 2;
          ctx.textBaseline = 'top';
        } else {
          startY = cy + cellH - totalTextH - paddingY;
          ctx.textBaseline = 'top';
        }
        for (let li = 0; li < lines.length; li++) {
          ctx.fillText(lines[li], textX, startY + li * lineH);
        }
      } else {
        // Underline via canvas
        if (font.underline) {
          const metrics = ctx.measureText(text);
          const textW = Math.min(metrics.width, cellW - paddingX * 2);
          const uy = alignV === 'top'
            ? cy + paddingY + Math.round(font.size * ROW_HEIGHT_TO_PX)
            : alignV === 'center'
              ? cy + cellH / 2 + Math.round(font.size * ROW_HEIGHT_TO_PX * 0.5)
              : cy + cellH - paddingY + 1;
          const ux = alignH === 'right' ? cx + cellW - paddingX - textW
            : alignH === 'center' ? cx + cellW / 2 - textW / 2
            : cx + paddingX;
          ctx.save();
          ctx.strokeStyle = font.color ? hexToRgba(font.color) : '#000000';
          ctx.lineWidth = 1;
          ctx.beginPath();
          ctx.moveTo(ux, uy);
          ctx.lineTo(ux + textW, uy);
          ctx.stroke();
          ctx.restore();
        }

        let textY: number;
        if (alignV === 'top') {
          ctx.textBaseline = 'top';
          textY = cy + paddingY;
        } else if (alignV === 'center') {
          ctx.textBaseline = 'middle';
          textY = cy + cellH / 2;
        } else {
          ctx.textBaseline = 'bottom';
          textY = cy + cellH - paddingY;
        }
        ctx.fillText(text, textX, textY);
      }

      ctx.restore();
    }
  }

  ctx.restore(); // end cell area clip

  // --- Render row/column headers (drawn last, always on top) ---
  renderHeaders(ctx, canvasW, canvasH, startRow, startCol, numRows, numCols, colXs, colWidths, rowYs, rowHeights);
}

function renderHeaders(
  ctx: CanvasRenderingContext2D,
  canvasW: number,
  canvasH: number,
  startRow: number,
  startCol: number,
  numRows: number,
  numCols: number,
  colXs: number[],
  colWidths: number[],
  rowYs: number[],
  rowHeights: number[],
): void {
  const HEADER_BG = '#f8f9fa';
  const HEADER_BORDER = '#c8ccd0';
  const HEADER_TEXT = '#444';
  const HEADER_FONT = `11px ${DEFAULT_FONT_FAMILY}`;

  // Corner cell
  ctx.fillStyle = HEADER_BG;
  ctx.fillRect(0, 0, HEADER_W, HEADER_H);
  ctx.strokeStyle = HEADER_BORDER;
  ctx.lineWidth = 1;
  ctx.strokeRect(0.5, 0.5, HEADER_W - 1, HEADER_H - 1);

  ctx.font = HEADER_FONT;
  ctx.fillStyle = HEADER_TEXT;
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';

  // Column letter headers
  ctx.save();
  ctx.beginPath();
  ctx.rect(HEADER_W, 0, canvasW - HEADER_W, HEADER_H);
  ctx.clip();

  for (let ci = 0; ci < numCols; ci++) {
    const cx = HEADER_W + colXs[ci];
    const cw = colWidths[ci];
    if (cx + cw <= HEADER_W || cx >= canvasW) continue;

    ctx.fillStyle = HEADER_BG;
    ctx.fillRect(cx, 0, cw, HEADER_H);
    ctx.strokeStyle = HEADER_BORDER;
    ctx.lineWidth = 0.5;
    ctx.strokeRect(cx + 0.5, 0.5, cw - 1, HEADER_H - 1);

    ctx.fillStyle = HEADER_TEXT;
    ctx.textAlign = 'center';
    ctx.fillText(colToLetter(startCol + ci), cx + cw / 2, HEADER_H / 2);
  }
  ctx.restore();

  // Row number headers
  ctx.save();
  ctx.beginPath();
  ctx.rect(0, HEADER_H, HEADER_W, canvasH - HEADER_H);
  ctx.clip();

  for (let ri = 0; ri < numRows; ri++) {
    const cy = HEADER_H + rowYs[ri];
    const ch = rowHeights[ri];
    if (cy + ch <= HEADER_H || cy >= canvasH) continue;

    ctx.fillStyle = HEADER_BG;
    ctx.fillRect(0, cy, HEADER_W, ch);
    ctx.strokeStyle = HEADER_BORDER;
    ctx.lineWidth = 0.5;
    ctx.strokeRect(0.5, cy + 0.5, HEADER_W - 1, ch - 1);

    ctx.fillStyle = HEADER_TEXT;
    ctx.textAlign = 'right';
    ctx.textBaseline = 'middle';
    ctx.fillText(String(startRow + ri), HEADER_W - 4, cy + ch / 2);
  }
  ctx.restore();
}

function renderBorder(
  ctx: CanvasRenderingContext2D,
  border: Border,
  x: number,
  y: number,
  w: number,
  h: number,
): void {
  const edges = [
    { edge: border.top, x1: x, y1: y, x2: x + w, y2: y },
    { edge: border.bottom, x1: x, y1: y + h, x2: x + w, y2: y + h },
    { edge: border.left, x1: x, y1: y, x2: x, y2: y + h },
    { edge: border.right, x1: x + w, y1: y, x2: x + w, y2: y + h },
  ];
  for (const { edge, x1, y1, x2, y2 } of edges) {
    if (!edge || !edge.style || edge.style === 'none') continue;
    ctx.beginPath();
    ctx.strokeStyle = edge.color ? hexToRgba(edge.color) : '#000000';
    ctx.lineWidth = borderStyleWidth(edge.style);
    ctx.moveTo(x1, y1);
    ctx.lineTo(x2, y2);
    ctx.stroke();
  }
}

function borderStyleWidth(style: string): number {
  switch (style) {
    case 'thick': return 2;
    case 'medium': return 1.5;
    case 'thin': return 1;
    default: return 1;
  }
}
