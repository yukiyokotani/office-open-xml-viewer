import type {
  Document, SectionProps, BodyElement, DocParagraph, DocTable, DocTableRow, DocTableCell,
  DocRun, TextRun, ImageRun, NumberingInfo, LineSpacing, BorderSpec, TableBorders, CellBorders,
} from './types';

// 1pt = 96/72 CSS px at screen
const PT_TO_PX = 96 / 72;

interface RenderState {
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;
  scale: number;    // px per pt
  contentX: number; // left of content area (px)
  contentW: number; // width of content area (px)
  y: number;        // current Y cursor (px)
  pageH: number;    // full page height (px)
  defaultColor: string;
}

export function renderDocumentToCanvas(
  doc: Document,
  canvas: HTMLCanvasElement | OffscreenCanvas,
  pageIndex: number,
  opts: { width?: number; dpr?: number; defaultTextColor?: string } = {},
): void {
  const sec = doc.section;
  const dpr = opts.dpr ?? devicePixelRatio ?? 1;
  const cssWidth = opts.width ?? sec.pageWidth * PT_TO_PX;
  const scale = cssWidth / sec.pageWidth;  // px per pt
  const cssHeight = sec.pageHeight * scale;

  canvas.width = Math.round(cssWidth * dpr);
  canvas.height = Math.round(cssHeight * dpr);

  if (canvas instanceof HTMLCanvasElement) {
    canvas.style.width = `${cssWidth}px`;
    canvas.style.height = `${cssHeight}px`;
  }

  const ctx = canvas.getContext('2d') as CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;
  ctx.scale(dpr, dpr);

  // White background
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, cssWidth, cssHeight);

  const state: RenderState = {
    ctx,
    scale,
    contentX: sec.marginLeft * scale,
    contentW: (sec.pageWidth - sec.marginLeft - sec.marginRight) * scale,
    y: sec.marginTop * scale,
    pageH: cssHeight,
    defaultColor: opts.defaultTextColor ?? '#000000',
  };

  // Split body at page breaks — pick the requested page
  const pages = splitPages(doc.body);
  const elements = pages[pageIndex] ?? pages[0] ?? [];

  for (const el of elements) {
    renderBodyElement(el, state);
  }
}

function splitPages(body: BodyElement[]): BodyElement[][] {
  const pages: BodyElement[][] = [[]];
  for (const el of body) {
    if (el.type === 'pageBreak') {
      pages.push([]);
    } else {
      pages[pages.length - 1].push(el);
    }
  }
  return pages;
}

// ===== Body element dispatch =====

function renderBodyElement(el: BodyElement, state: RenderState): void {
  if (el.type === 'paragraph') {
    renderParagraph(el as unknown as DocParagraph, state);
  } else if (el.type === 'table') {
    renderTable(el as unknown as DocTable, state);
  }
}

// ===== Paragraph rendering =====

function renderParagraph(para: DocParagraph, state: RenderState): void {
  const { ctx, scale, contentX, contentW, defaultColor } = state;

  state.y += para.spaceBefore * scale;

  const indLeft = para.indentLeft * scale;
  const indRight = para.indentRight * scale;
  const indFirst = para.indentFirst * scale;

  // Numbering prefix (indent is already baked into para.indentLeft / para.indentFirst)
  let numPrefix = '';
  let numTab = 0;
  if (para.numbering) {
    numPrefix = para.numbering.text + '\t';
    numTab = para.numbering.tab * scale;
  }

  const paraX = contentX + indLeft;
  const firstLineX = paraX + indFirst;
  const paraW = contentW - indLeft - indRight;

  // Collect all text segments with formatting
  const segments = buildSegments(para.runs);

  if (segments.length === 0) {
    // Empty paragraph still takes up line height
    const fontSize = getDefaultFontSize(para) * scale;
    state.y += fontSize * lineSpacingMultiplier(para.lineSpacing);
    state.y += para.spaceAfter * scale;
    return;
  }

  // Layout lines
  const lines = layoutLines(ctx, segments, paraW, firstLineX - paraX, scale);

  // Render each line
  let firstLine = true;
  for (const line of lines) {
    const lineH = line.height * scale * lineSpacingMultiplier(para.lineSpacing);
    const baseline = state.y + line.ascent;

    let x = firstLine ? firstLineX : paraX;

    // Render numbering prefix on first line
    if (firstLine && numPrefix) {
      const numFontSize = getDefaultFontSize(para) * scale;
      ctx.font = `${numFontSize}px sans-serif`;
      ctx.fillStyle = defaultColor;
      ctx.fillText(para.numbering!.text, x - numTab, baseline);
    }

    // Alignment offset for this line
    const lineWidth = line.segments.reduce((s, seg) => s + seg.measuredWidth, 0);
    let alignOffset = 0;
    if (para.alignment === 'right') alignOffset = paraW - (x - paraX) - lineWidth;
    else if (para.alignment === 'center') alignOffset = (paraW - (x - paraX) - lineWidth) / 2;

    x += alignOffset;

    for (const seg of line.segments) {
      if ('dataUrl' in seg) {
        // Image segment
        renderInlineImage(ctx, seg as LayoutImageSeg, x, baseline, scale);
        x += seg.measuredWidth;
        continue;
      }
      const s = seg as LayoutTextSeg;
      ctx.font = buildFont(s.bold, s.italic, s.fontSize * scale, s.fontFamily);
      ctx.fillStyle = s.color ? `#${s.color}` : defaultColor;

      if (s.strikethrough || s.underline) {
        ctx.save();
      }

      ctx.fillText(s.text, x, baseline);

      if (s.underline) {
        const lw = ctx.measureText(s.text).width;
        ctx.strokeStyle = s.color ? `#${s.color}` : defaultColor;
        ctx.lineWidth = Math.max(0.5, s.fontSize * scale * 0.05);
        ctx.beginPath();
        ctx.moveTo(x, baseline + s.fontSize * scale * 0.12);
        ctx.lineTo(x + lw, baseline + s.fontSize * scale * 0.12);
        ctx.stroke();
      }

      if (s.strikethrough) {
        const lw = ctx.measureText(s.text).width;
        ctx.strokeStyle = s.color ? `#${s.color}` : defaultColor;
        ctx.lineWidth = Math.max(0.5, s.fontSize * scale * 0.05);
        ctx.beginPath();
        ctx.moveTo(x, baseline - s.fontSize * scale * 0.3);
        ctx.lineTo(x + lw, baseline - s.fontSize * scale * 0.3);
        ctx.stroke();
      }

      if (s.strikethrough || s.underline) ctx.restore();

      x += s.measuredWidth;
    }

    state.y += lineH;
    firstLine = false;
  }

  state.y += para.spaceAfter * scale;
}

// ===== Text layout =====

interface LayoutTextSeg {
  text: string;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strikethrough: boolean;
  fontSize: number;  // pt
  color: string | null;
  fontFamily: string | null;
  measuredWidth: number;  // px (set during layout)
}

interface LayoutImageSeg {
  dataUrl: string;
  widthPt: number;
  heightPt: number;
  measuredWidth: number;
}

interface LayoutLine {
  segments: (LayoutTextSeg | LayoutImageSeg)[];
  height: number;  // pt
  ascent: number;  // px
}

function buildSegments(runs: DocRun[]): (LayoutTextSeg | LayoutImageSeg)[] {
  const segs: (LayoutTextSeg | LayoutImageSeg)[] = [];
  for (const run of runs) {
    if (run.type === 'text') {
      const t = run as unknown as TextRun & { type: 'text' };
      // Split on spaces/words for wrapping (keep spaces with preceding word)
      const words = splitTextForLayout(t.text);
      for (const w of words) {
        segs.push({
          text: w,
          bold: t.bold,
          italic: t.italic,
          underline: t.underline,
          strikethrough: t.strikethrough,
          fontSize: t.fontSize,
          color: t.color,
          fontFamily: t.fontFamily,
          measuredWidth: 0,
        });
      }
    } else if (run.type === 'image') {
      const img = run as unknown as ImageRun & { type: 'image' };
      segs.push({ dataUrl: img.dataUrl, widthPt: img.widthPt, heightPt: img.heightPt, measuredWidth: 0 });
    }
    // line breaks handled implicitly as paragraph ends
  }
  return segs;
}

function splitTextForLayout(text: string): string[] {
  // Split text into word+trailing-space chunks
  const result: string[] = [];
  let i = 0;
  while (i < text.length) {
    let j = i;
    while (j < text.length && text[j] !== ' ') j++;
    // Include trailing spaces
    while (j < text.length && text[j] === ' ') j++;
    if (j > i) result.push(text.slice(i, j));
    i = j;
  }
  return result.length ? result : [text];
}

function layoutLines(
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  segs: (LayoutTextSeg | LayoutImageSeg)[],
  maxWidth: number,
  firstIndent: number,
  scale: number,
): LayoutLine[] {
  const lines: LayoutLine[] = [];
  let currentLine: (LayoutTextSeg | LayoutImageSeg)[] = [];
  let currentWidth = 0;
  let lineHeight = 0;   // pt
  let lineAscent = 0;   // px
  let isFirst = true;
  const availW = (w: number) => w - (isFirst ? firstIndent : 0);

  const flush = () => {
    if (currentLine.length > 0) {
      lines.push({ segments: currentLine, height: lineHeight || 10, ascent: lineAscent });
    }
    currentLine = [];
    currentWidth = 0;
    lineHeight = 0;
    lineAscent = 0;
    isFirst = false;
  };

  for (const seg of segs) {
    let w: number;
    let h: number;
    let asc: number;

    if ('dataUrl' in seg) {
      w = seg.widthPt * scale;
      h = seg.heightPt;
      asc = seg.heightPt * scale;
      seg.measuredWidth = w;
    } else {
      const s = seg as LayoutTextSeg;
      ctx.font = buildFont(s.bold, s.italic, s.fontSize * scale, s.fontFamily);
      const m = ctx.measureText(s.text);
      w = m.width;
      h = s.fontSize;
      asc = (m.actualBoundingBoxAscent ?? s.fontSize * scale * 0.75);
      seg.measuredWidth = w;
    }

    // Check if it fits
    if (currentLine.length > 0 && currentWidth + w > availW(maxWidth)) {
      flush();
    }

    currentLine.push(seg);
    currentWidth += w;
    if (h > lineHeight) lineHeight = h;
    if (asc > lineAscent) lineAscent = asc;
  }
  flush();

  return lines;
}

function renderInlineImage(
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  seg: LayoutImageSeg,
  x: number,
  baseline: number,
  scale: number,
): void {
  const img = new Image();
  img.src = seg.dataUrl;
  const w = seg.widthPt * scale;
  const h = seg.heightPt * scale;
  // Images are drawn synchronously only if already loaded; otherwise skipped for now
  try {
    ctx.drawImage(img, x, baseline - h, w, h);
  } catch (_) { /* not loaded */ }
}

// ===== Table rendering =====

function renderTable(table: DocTable, state: RenderState): void {
  const { ctx, scale, contentX, contentW } = state;

  // Calculate column widths — scale to fit content width if needed
  let totalColW = table.colWidths.reduce((s, w) => s + w, 0) * scale;
  const colScale = totalColW > contentW ? contentW / totalColW : 1;
  const colWidths = table.colWidths.map(w => w * scale * colScale);

  const tableX = contentX;
  const tableW = colWidths.reduce((s, w) => s + w, 0);

  // First pass: calculate row heights
  const rowHeights: number[] = [];
  for (const row of table.rows) {
    const rowH = calculateRowHeight(row, table, colWidths, scale, state);
    rowHeights.push(rowH);
  }

  // Second pass: render
  let y = state.y;
  for (let ri = 0; ri < table.rows.length; ri++) {
    const row = table.rows[ri];
    const rowH = rowHeights[ri];
    let x = tableX;
    let ci = 0;

    for (const cell of row.cells) {
      const span = Math.min(cell.colSpan, colWidths.length - ci);
      const cellW = colWidths.slice(ci, ci + span).reduce((s, w) => s + w, 0);

      if (cell.vMerge !== false) { // not a continuation row
        renderCell(cell, table, x, y, cellW, rowH, state);
      }

      x += cellW;
      ci += span;
    }

    // Draw row borders
    y += rowH;
  }

  state.y = y;
}

function calculateRowHeight(
  row: DocTableRow,
  table: DocTable,
  colWidths: number[],
  scale: number,
  state: RenderState,
): number {
  if (row.rowHeight != null) return row.rowHeight * scale;

  // Measure cell content heights
  let maxH = 10 * scale;
  let ci = 0;
  for (const cell of row.cells) {
    const span = Math.min(cell.colSpan, colWidths.length - ci);
    const cellW = colWidths.slice(ci, ci + span).reduce((s, w) => s + w, 0);
    const contentW = cellW - (table.cellMarginLeft + table.cellMarginRight) * scale;

    let h = (table.cellMarginTop + table.cellMarginBottom) * scale;
    for (const para of cell.content) {
      h += measureParaHeight(state.ctx, para, contentW, scale);
      h += (para.spaceBefore + para.spaceAfter) * scale;
    }
    if (h > maxH) maxH = h;
    ci += span;
  }
  return maxH;
}

function measureParaHeight(
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  para: DocParagraph,
  maxWidth: number,
  scale: number,
): number {
  const segs = buildSegments(para.runs);
  if (segs.length === 0) return getDefaultFontSize(para) * scale;
  const lines = layoutLines(ctx, segs, maxWidth, 0, scale);
  const mult = lineSpacingMultiplier(para.lineSpacing);
  return lines.reduce((s, l) => s + l.height * scale * mult, 0);
}

function renderCell(
  cell: DocTableCell,
  table: DocTable,
  x: number,
  y: number,
  w: number,
  h: number,
  state: RenderState,
): void {
  const { ctx, scale } = state;

  // Background
  if (cell.background) {
    ctx.fillStyle = `#${cell.background}`;
    ctx.fillRect(x, y, w, h);
  }

  // Borders
  drawCellBorders(ctx, x, y, w, h, cell.borders, table.borders, scale);

  // Content
  const mt = table.cellMarginTop * scale;
  const mb = table.cellMarginBottom * scale;
  const ml = table.cellMarginLeft * scale;
  const mr = table.cellMarginRight * scale;

  const cellState: RenderState = {
    ctx,
    scale,
    contentX: x + ml,
    contentW: w - ml - mr,
    y: y + mt,
    pageH: state.pageH,
    defaultColor: state.defaultColor,
  };

  // Vertical alignment
  if (cell.vAlign === 'center' || cell.vAlign === 'bottom') {
    const contentH = cell.content.reduce((s, p) =>
      s + measureParaHeight(ctx, p, w - ml - mr, scale) + (p.spaceBefore + p.spaceAfter) * scale, 0);
    if (cell.vAlign === 'center') cellState.y = y + (h - contentH) / 2;
    else cellState.y = y + h - contentH - mb;
  }

  for (const para of cell.content) {
    renderParagraph(para, cellState);
  }
}

function drawCellBorders(
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  x: number, y: number, w: number, h: number,
  cell: CellBorders,
  table: TableBorders,
  scale: number,
): void {
  const top = cell.top ?? table.top;
  const bottom = cell.bottom ?? table.bottom;
  const left = cell.left ?? table.left;
  const right = cell.right ?? table.right;

  if (top && top.style !== 'none') drawBorderLine(ctx, x, y, x + w, y, top, scale);
  if (bottom && bottom.style !== 'none') drawBorderLine(ctx, x, y + h, x + w, y + h, bottom, scale);
  if (left && left.style !== 'none') drawBorderLine(ctx, x, y, x, y + h, left, scale);
  if (right && right.style !== 'none') drawBorderLine(ctx, x + w, y, x + w, y + h, right, scale);
}

function drawBorderLine(
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  x1: number, y1: number, x2: number, y2: number,
  spec: BorderSpec,
  scale: number,
): void {
  ctx.save();
  ctx.strokeStyle = spec.color ? `#${spec.color}` : '#000000';
  ctx.lineWidth = Math.max(0.5, spec.width * scale);
  ctx.beginPath();
  ctx.moveTo(x1, y1);
  ctx.lineTo(x2, y2);
  ctx.stroke();
  ctx.restore();
}

// ===== Utilities =====

function buildFont(bold: boolean, italic: boolean, sizePx: number, family: string | null): string {
  const w = bold ? 'bold' : 'normal';
  const s = italic ? 'italic' : 'normal';
  const f = normalizeFontFamily(family);
  return `${s} ${w} ${sizePx}px ${f}`;
}

function normalizeFontFamily(family: string | null): string {
  if (!family) return '"Noto Sans JP", "Hiragino Sans", "Meiryo", sans-serif';
  const lower = family.toLowerCase();
  if (lower.includes('meiryo') || lower.includes('メイリオ')) {
    return '"Meiryo UI", "Meiryo", "Noto Sans JP", sans-serif';
  }
  if (lower.includes('游') || lower.includes('yu ')) {
    return '"Yu Gothic", "YuGothic", "Noto Sans JP", sans-serif';
  }
  if (lower.includes('ipa')) {
    return '"IPAexGothic", "Noto Sans JP", sans-serif';
  }
  if (lower.includes('segoe')) {
    return '"Segoe UI", sans-serif';
  }
  return `"${family}", sans-serif`;
}

function getDefaultFontSize(para: DocParagraph): number {
  // Find first text run's font size as representative
  for (const run of para.runs) {
    if (run.type === 'text') {
      return (run as unknown as TextRun).fontSize;
    }
  }
  return 10; // pt fallback
}

function lineSpacingMultiplier(ls: LineSpacing | null): number {
  if (!ls) return 1.2;
  if (ls.rule === 'auto') return ls.value * 1.2;
  return 1.2; // for exact/atLeast, use line value directly in pt (handled separately if needed)
}
