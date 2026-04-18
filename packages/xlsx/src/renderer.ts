import type { Worksheet, Styles, Cell, CellValue, Font, Fill, Border, CellXf, ViewportRange, RenderViewportOptions } from './types.js';

const DEFAULT_FONT_FAMILY = 'Calibri, Arial, sans-serif';
const DEFAULT_FONT_SIZE = 11;
// Max digit width of the default font at 96 DPI.
// Calibri 11pt ≈ 7px, Meiryo UI 11pt ≈ 8px.
// Excel stores column width in units of this value.
const MDW = 7;
const ROW_HEIGHT_TO_PX = 4 / 3; // pt to px at 96 DPI: 96/72

// OOXML spec: pixel = trunc(((256*w + 128/MDW) / 256) * MDW)
function colWidthToPx(w: number): number {
  return Math.trunc(((256 * w + 128 / MDW) / 256) * MDW);
}

function rowHeightToPx(h: number): number {
  return Math.round(h * ROW_HEIGHT_TO_PX); // h is in pt
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

// ECMA-376 built-in numFmtId → format category
function formatCellValue(cell: Cell, styles: Styles): string {
  if (cell.value.type !== 'number') return cellValueText(cell.value);
  const xf = styles.cellXfs[cell.styleIndex ?? 0];
  const numFmtId = xf?.numFmtId ?? 0;
  const num = cell.value.number;

  // Look up custom format first
  const customFmt = styles.numFmts?.find(f => f.numFmtId === numFmtId);

  return applyFormat(num, numFmtId, customFmt?.formatCode ?? null);
}

function applyFormat(num: number, numFmtId: number, formatCode: string | null): string {
  // Date serial range: 1 to ~50000 roughly
  const isDateFmtId = (id: number) => (id >= 14 && id <= 17) || id === 22;

  if (isDateFmtId(numFmtId)) {
    return formatExcelDate(num);
  }

  // Try to interpret formatCode for custom formats
  if (formatCode) {
    return applyFormatCode(num, formatCode);
  }

  // Built-in IDs (ECMA-376 §18.8.30)
  switch (numFmtId) {
    case 0: return String(num); // General
    case 1: return Math.round(num).toString(); // 0
    case 2: return num.toFixed(2); // 0.00
    case 3: return formatThousands(num, 0); // #,##0
    case 4: return formatThousands(num, 2); // #,##0.00
    case 9: return Math.round(num * 100) + '%'; // 0%
    case 10: return (num * 100).toFixed(2) + '%'; // 0.00%
    case 11: return num.toExponential(2); // 0.00E+00
    case 37: return formatThousands(num, 0); // #,##0 ;(#,##0)
    case 38: return formatThousands(num, 0); // #,##0 ;[Red](#,##0)
    case 39: return formatThousands(num, 2); // #,##0.00;(#,##0.00)
    case 40: return formatThousands(num, 2); // #,##0.00;[Red](#,##0.00)
    case 49: return String(num); // @ (text)
    default: return String(num);
  }
}

function formatThousands(num: number, decimals: number): string {
  return num.toLocaleString('en-US', { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
}

// Excel date serial → locale date string (1900 date system)
function formatExcelDate(serial: number): string {
  // Excel serial 1 = 1900-01-01, but Excel incorrectly treats 1900 as leap year
  // Subtract 25569 to get Unix days from 1970-01-01, then multiply by ms/day
  const date = new Date((serial - 25569) * 86400 * 1000);
  return date.toLocaleDateString();
}

function countDecimalPlaces(fmt: string): number {
  const m = fmt.match(/\.([0#]+)/);
  return m ? m[1].length : 0;
}

function applyFormatCode(num: number, formatCode: string): string {
  // Handle positive;negative;zero sections (take first section for positive/zero)
  const sections = formatCode.split(';');
  const section = num < 0 && sections.length > 1 ? sections[1] : sections[0];

  // Remove color/condition brackets like [Red], [$¥-411]
  const cleaned = section.replace(/\[.*?\]/g, '').replace(/_./g, '').replace(/\*/g, '');

  // Detect percent
  if (cleaned.includes('%')) {
    const dec = countDecimalPlaces(cleaned);
    return (num * 100).toFixed(dec) + '%';
  }

  // Detect thousands separator
  const hasThousands = cleaned.includes(',') && (cleaned.includes('#') || cleaned.includes('0'));
  const dec = countDecimalPlaces(cleaned);

  if (hasThousands) {
    return formatThousands(num, dec);
  }

  // Fixed decimal
  if (cleaned.includes('.')) {
    return num.toFixed(dec);
  }

  // Integer or unknown
  if (cleaned.match(/[#0]/)) {
    return Math.round(num).toString();
  }

  return String(num);
}

// Wrap text into lines that fit within maxWidth
function wrapText(ctx: CanvasRenderingContext2D, text: string, maxWidth: number): string[] {
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

export function renderViewport(
  ctx: CanvasRenderingContext2D,
  worksheet: Worksheet,
  styles: Styles,
  viewport: ViewportRange,
  _opts: RenderViewportOptions = {},
): void {
  const canvasW = ctx.canvas.width;
  const canvasH = ctx.canvas.height;

  ctx.clearRect(0, 0, canvasW, canvasH);
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, canvasW, canvasH);

  const { row: startRow, col: startCol, rows: numRows, cols: numCols } = viewport;

  // Pre-compute column X positions
  const colXs: number[] = [];
  const colWidths: number[] = [];
  let x = 0;
  for (let c = startCol; c < startCol + numCols; c++) {
    colXs.push(x);
    const w = colWidthToPx(worksheet.colWidths[c] ?? worksheet.defaultColWidth);
    colWidths.push(w);
    x += w;
  }

  // Pre-compute row Y positions
  const rowYs: number[] = [];
  const rowHeights: number[] = [];
  let y = 0;
  for (let r = startRow; r < startRow + numRows; r++) {
    rowYs.push(y);
    const h = rowHeightToPx(worksheet.rowHeights[r] ?? worksheet.defaultRowHeight);
    rowHeights.push(h);
    y += h;
  }

  // Build cell lookup from worksheet data
  const cellMap = new Map<string, Cell>();
  for (const row of worksheet.rows) {
    for (const cell of row.cells) {
      cellMap.set(`${cell.row}:${cell.col}`, cell);
    }
  }

  // Render cells
  for (let ri = 0; ri < numRows; ri++) {
    const rowIndex = startRow + ri;
    const cy = rowYs[ri];
    const ch = rowHeights[ri];

    for (let ci = 0; ci < numCols; ci++) {
      const colIndex = startCol + ci;
      const cx = colXs[ci];
      const cw = colWidths[ci];

      const cell = cellMap.get(`${rowIndex}:${colIndex}`);
      const styleIndex = cell?.styleIndex ?? 0;
      const { font, fill, border, xf } = resolveXf(styles, styleIndex);

      // Background fill
      if (fill.patternType !== 'none' && fill.patternType !== '' && fill.fgColor) {
        ctx.fillStyle = hexToRgba(fill.fgColor);
        ctx.fillRect(cx, cy, cw, ch);
      }

      // Grid line (light gray default)
      ctx.strokeStyle = '#d0d0d0';
      ctx.lineWidth = 0.5;
      ctx.strokeRect(cx + 0.5, cy + 0.5, cw - 1, ch - 1);

      // Border edges
      renderBorder(ctx, border, cx, cy, cw, ch);

      if (!cell) continue;
      const text = formatCellValue(cell, styles);
      if (!text) continue;

      // Text setup
      ctx.font = buildFont(font);
      ctx.fillStyle = font.color ? hexToRgba(font.color) : '#000000';

      const paddingX = 2;
      const paddingY = 2;
      const alignH = xf.alignH ?? (cell.value.type === 'number' ? 'right' : 'left');
      const alignV = xf.alignV ?? 'bottom';

      let textX: number;
      let textAlign: CanvasTextAlign;
      if (alignH === 'right') {
        textX = cx + cw - paddingX;
        textAlign = 'right';
      } else if (alignH === 'center') {
        textX = cx + cw / 2;
        textAlign = 'center';
      } else {
        textX = cx + paddingX;
        textAlign = 'left';
      }

      ctx.textAlign = textAlign;

      ctx.save();
      ctx.beginPath();
      ctx.rect(cx, cy, cw, ch);
      ctx.clip();

      if (xf.wrapText) {
        const lines = wrapText(ctx, text, cw - paddingX * 2);
        const lineH = Math.round(font.size * ROW_HEIGHT_TO_PX * 1.2);
        const totalH = lines.length * lineH;
        let startY: number;
        if (alignV === 'top') {
          startY = cy + paddingY;
          ctx.textBaseline = 'top';
        } else if (alignV === 'center') {
          startY = cy + (ch - totalH) / 2;
          ctx.textBaseline = 'top';
        } else {
          startY = cy + ch - totalH - paddingY;
          ctx.textBaseline = 'top';
        }
        for (let li = 0; li < lines.length; li++) {
          ctx.fillText(lines[li], textX, startY + li * lineH);
        }
      } else {
        let textY: number;
        if (alignV === 'top') {
          ctx.textBaseline = 'top';
          textY = cy + paddingY;
        } else if (alignV === 'center') {
          ctx.textBaseline = 'middle';
          textY = cy + ch / 2;
        } else {
          // bottom (Excel default for most cells)
          ctx.textBaseline = 'bottom';
          textY = cy + ch - paddingY;
        }
        ctx.fillText(text, textX, textY);
      }

      ctx.restore();
    }
  }
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
