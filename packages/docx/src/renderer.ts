import type {
  Document, BodyElement, DocParagraph, DocTable, DocTableRow, DocTableCell,
  DocRun, TextRun, ImageRun, FieldRun, HeaderFooter, LineSpacing, BorderSpec, TableBorders, CellBorders,
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
  /** 0-based page index currently being rendered */
  pageIndex: number;
  /** total page count in the document */
  totalPages: number;
  /** preloaded image bitmaps keyed by dataUrl */
  images: Map<string, ImageBitmap>;
  /** when true, layout is performed but nothing is drawn (used for header/footer height measurement) */
  dryRun: boolean;
  /** section left margin in pt — used to convert margin-relative anchor X to page-absolute */
  marginLeft: number;
}

export interface RenderDocumentOptions {
  width?: number;
  dpr?: number;
  defaultTextColor?: string;
  /** total pages in the document (used to resolve NUMPAGES fields) */
  totalPages?: number;
}

// ===== Image preloading =====

interface ImagePair {
  url: string;
  colorReplaceFrom?: string;
}

/** Returns a stable map key for a (url, colorReplaceFrom) pair. */
function imageKey(url: string, colorReplaceFrom?: string): string {
  return colorReplaceFrom ? `${url}|clr:${colorReplaceFrom}` : url;
}

function collectImagePairs(doc: Document): ImagePair[] {
  const seen = new Map<string, ImagePair>();
  const walk = (runs: DocRun[]) => {
    for (const run of runs) {
      if (run.type === 'image') {
        const img = run as unknown as ImageRun;
        const key = imageKey(img.dataUrl, img.colorReplaceFrom);
        if (!seen.has(key)) seen.set(key, { url: img.dataUrl, colorReplaceFrom: img.colorReplaceFrom });
      }
    }
  };
  const walkBody = (body: BodyElement[]) => {
    for (const el of body) {
      if (el.type === 'paragraph') walk((el as unknown as DocParagraph).runs);
      if (el.type === 'table') {
        const tbl = el as unknown as DocTable;
        for (const row of tbl.rows)
          for (const cell of row.cells)
            for (const para of cell.content) walk(para.runs);
      }
    }
  };
  walkBody(doc.body);
  if (doc.headers.default) walkBody(doc.headers.default.body);
  if (doc.headers.first)   walkBody(doc.headers.first.body);
  if (doc.headers.even)    walkBody(doc.headers.even.body);
  if (doc.footers.default) walkBody(doc.footers.default.body);
  if (doc.footers.first)   walkBody(doc.footers.first.body);
  if (doc.footers.even)    walkBody(doc.footers.even.body);
  return [...seen.values()];
}

/**
 * Apply a:clrChange color replacement: turn every pixel whose (R,G,B) matches colorHex into
 * fully transparent (alpha=0). Returns a new ImageBitmap with the modified pixels.
 */
async function applyColorReplacement(bmp: ImageBitmap, colorHex: string): Promise<ImageBitmap> {
  const r = parseInt(colorHex.slice(0, 2), 16);
  const g = parseInt(colorHex.slice(2, 4), 16);
  const b = parseInt(colorHex.slice(4, 6), 16);

  const offscreen = new OffscreenCanvas(bmp.width, bmp.height);
  const ctx2 = offscreen.getContext('2d')!;
  ctx2.drawImage(bmp, 0, 0);

  const imgData = ctx2.getImageData(0, 0, bmp.width, bmp.height);
  const d = imgData.data;

  for (let i = 0; i < d.length; i += 4) {
    if (d[i] === r && d[i + 1] === g && d[i + 2] === b) {
      d[i + 3] = 0; // make transparent
    }
  }

  ctx2.putImageData(imgData, 0, 0);
  return createImageBitmap(offscreen);
}

async function preloadImages(doc: Document): Promise<Map<string, ImageBitmap>> {
  const pairs = collectImagePairs(doc);
  const entries = await Promise.all(
    pairs.map(async (pair): Promise<[string, ImageBitmap] | null> => {
      try {
        const res = await fetch(pair.url);
        const blob = await res.blob();
        let bmp = await createImageBitmap(blob);
        if (pair.colorReplaceFrom) {
          bmp = await applyColorReplacement(bmp, pair.colorReplaceFrom);
        }
        return [imageKey(pair.url, pair.colorReplaceFrom), bmp];
      } catch {
        return null;
      }
    }),
  );
  return new Map(entries.filter((e): e is [string, ImageBitmap] => e !== null));
}

// ===== Main entry =====

export async function renderDocumentToCanvas(
  doc: Document,
  canvas: HTMLCanvasElement | OffscreenCanvas,
  pageIndex: number,
  opts: RenderDocumentOptions = {},
): Promise<void> {
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

  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, cssWidth, cssHeight);

  const pages = splitPages(doc.body);
  const totalPages = Math.max(opts.totalPages ?? pages.length, pages.length);
  const elements = pages[pageIndex] ?? pages[0] ?? [];

  const images = await preloadImages(doc);

  const baseState: RenderState = {
    ctx,
    scale,
    contentX: sec.marginLeft * scale,
    contentW: (sec.pageWidth - sec.marginLeft - sec.marginRight) * scale,
    y: sec.marginTop * scale,
    pageH: cssHeight,
    defaultColor: opts.defaultTextColor ?? '#000000',
    pageIndex,
    totalPages,
    images,
    dryRun: false,
    marginLeft: sec.marginLeft,
  };

  // Header: top of page, starting at headerDistance
  const header = pickHeaderFooter(doc.headers, pageIndex, totalPages, doc.section.titlePage, doc.section.evenAndOddHeaders);
  if (header) {
    renderHeaderFooter(header, sec.headerDistance * scale, baseState);
  }

  // Footer: anchored from bottom, rising by its measured height
  const footer = pickHeaderFooter(doc.footers, pageIndex, totalPages, doc.section.titlePage, doc.section.evenAndOddHeaders);
  if (footer) {
    const footerHeight = measureHeaderFooterHeight(footer, baseState);
    const footerTopY = cssHeight - sec.footerDistance * scale - footerHeight;
    renderHeaderFooter(footer, footerTopY, baseState);
  }

  // Body
  const bodyState: RenderState = { ...baseState, y: sec.marginTop * scale };
  for (const el of elements) {
    renderBodyElement(el, bodyState);
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

function pickHeaderFooter(
  set: Document['headers'],
  pageIndex: number,
  _totalPages: number,
  titlePage: boolean,
  evenAndOdd: boolean,
): HeaderFooter | null {
  if (titlePage && pageIndex === 0 && set.first) return set.first;
  if (evenAndOdd && pageIndex % 2 === 1 && set.even) return set.even;
  return set.default ?? null;
}

function renderHeaderFooter(hf: HeaderFooter, topY: number, base: RenderState): void {
  const state: RenderState = { ...base, y: topY };
  for (const el of hf.body) renderBodyElement(el, state);
}

function measureHeaderFooterHeight(hf: HeaderFooter, base: RenderState): number {
  const state: RenderState = { ...base, y: 0, dryRun: true };
  for (const el of hf.body) renderBodyElement(el, state);
  return state.y;
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
  const { ctx, scale, contentX, contentW, defaultColor, dryRun } = state;
  // Capture Y before spaceBefore — used for paragraph-relative anchor image positioning
  const paragraphStartY = state.y;

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

  // Collect all text segments with formatting (resolving field runs against page context)
  const segments = buildSegments(para.runs, state);

  if (segments.length === 0) {
    const fontSize = getDefaultFontSize(para) * scale;
    state.y += fontSize * lineSpacingMultiplier(para.lineSpacing);
    state.y += para.spaceAfter * scale;
    return;
  }

  const lines = layoutLines(ctx, segments, paraW, firstLineX - paraX, scale);

  let firstLine = true;
  for (const line of lines) {
    const lineH = line.height * scale * lineSpacingMultiplier(para.lineSpacing);
    const baseline = state.y + line.ascent;

    let x = firstLine ? firstLineX : paraX;

    if (firstLine && numPrefix && !dryRun) {
      const numFontSize = getDefaultFontSize(para) * scale;
      ctx.font = `${numFontSize}px sans-serif`;
      ctx.fillStyle = defaultColor;
      ctx.fillText(para.numbering!.text, x - numTab, baseline);
    }

    const lineWidth = line.segments.reduce((s, seg) => s + seg.measuredWidth, 0);
    let alignOffset = 0;
    if (para.alignment === 'right') alignOffset = paraW - (x - paraX) - lineWidth;
    else if (para.alignment === 'center') alignOffset = (paraW - (x - paraX) - lineWidth) / 2;

    x += alignOffset;

    for (const seg of line.segments) {
      if ('dataUrl' in seg) {
        if (!dryRun) renderInlineImage(ctx, seg as LayoutImageSeg, x, baseline, scale, state.images);
        x += seg.measuredWidth;
        continue;
      }
      const s = seg as LayoutTextSeg;
      if (!dryRun) {
        ctx.font = buildFont(s.bold, s.italic, s.fontSize * scale, s.fontFamily);
        ctx.fillStyle = s.color ? `#${s.color}` : defaultColor;

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
      }

      x += s.measuredWidth;
    }

    state.y += lineH;
    firstLine = false;
  }

  state.y += para.spaceAfter * scale;

  // Anchor images are absolutely positioned — draw after inline flow
  renderAnchorImages(para, state, paragraphStartY);
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
  /** true = wp:anchor: skip inline flow, draw at absolute page coords */
  anchor: boolean;
  anchorXPt: number;
  anchorYPt: number;
  anchorXFromMargin: boolean;
  anchorYFromPara: boolean;
  /** When set, pixels matching this hex color are replaced with alpha=0 before drawing. */
  colorReplaceFrom?: string;
  measuredWidth: number;
}

/** Sentinel that forces a new line when encountered in layoutLines. */
interface LayoutLineBreak {
  lineBreak: true;
  fontSize: number;  // pt — used to set line height on empty lines
  measuredWidth: 0;
}

type LayoutSeg = LayoutTextSeg | LayoutImageSeg | LayoutLineBreak;

interface LayoutLine {
  segments: (LayoutTextSeg | LayoutImageSeg)[];
  height: number;  // pt
  ascent: number;  // px
}

function buildSegments(runs: DocRun[], state: RenderState): LayoutSeg[] {
  const segs: LayoutSeg[] = [];
  for (const run of runs) {
    if (run.type === 'text') {
      const t = run as unknown as TextRun & { type: 'text' };
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
      segs.push({
        dataUrl: img.dataUrl,
        widthPt: img.widthPt,
        heightPt: img.heightPt,
        anchor: img.anchor ?? false,
        anchorXPt: img.anchorXPt ?? 0,
        anchorYPt: img.anchorYPt ?? 0,
        anchorXFromMargin: img.anchorXFromMargin ?? false,
        anchorYFromPara: img.anchorYFromPara ?? false,
        colorReplaceFrom: img.colorReplaceFrom,
        measuredWidth: 0,
      });
    } else if (run.type === 'break') {
      if (run.breakType === 'line') {
        // Determine font size for the line break height from surrounding text runs
        const fontSize = findNearbyFontSize(runs, runs.indexOf(run));
        segs.push({ lineBreak: true, fontSize, measuredWidth: 0 });
      }
      // page/column breaks handled at the document level (splitPages)
    } else if (run.type === 'field') {
      const f = run as unknown as FieldRun & { type: 'field' };
      const text = resolveFieldText(f, state);
      if (text) {
        segs.push({
          text,
          bold: f.bold,
          italic: f.italic,
          underline: f.underline,
          strikethrough: f.strikethrough,
          fontSize: f.fontSize,
          color: f.color,
          fontFamily: f.fontFamily,
          measuredWidth: 0,
        });
      }
    }
  }
  return segs;
}

function findNearbyFontSize(runs: DocRun[], idx: number): number {
  // Look backwards then forwards for a text or field run to get font size
  for (let i = idx - 1; i >= 0; i--) {
    const r = runs[i];
    if (r.type === 'text') return (r as unknown as TextRun).fontSize;
    if (r.type === 'field') return (r as unknown as FieldRun).fontSize;
  }
  for (let i = idx + 1; i < runs.length; i++) {
    const r = runs[i];
    if (r.type === 'text') return (r as unknown as TextRun).fontSize;
    if (r.type === 'field') return (r as unknown as FieldRun).fontSize;
  }
  return 10; // pt fallback
}

function resolveFieldText(f: FieldRun, state: RenderState): string {
  if (f.fieldType === 'page') return String(state.pageIndex + 1);
  if (f.fieldType === 'numPages') return String(state.totalPages);
  return f.fallbackText;
}

/** Returns true for code-points that permit line-break between adjacent characters (CJK). */
function hasCJKBreakOpportunity(text: string): boolean {
  for (let i = 0; i < text.length; ) {
    const cp = text.codePointAt(i)!;
    if (
      (cp >= 0x3000 && cp <= 0x9FFF)  ||
      (cp >= 0xF900 && cp <= 0xFAFF)  ||
      (cp >= 0xAC00 && cp <= 0xD7AF)  ||
      (cp >= 0xFF00 && cp <= 0xFFEF)
    ) return true;
    i += cp > 0xFFFF ? 2 : 1;
  }
  return false;
}

/**
 * Binary-search the longest prefix of `text` whose rendered width fits in `maxWidth`.
 * Used for CJK overflow splitting.
 */
function fitCJKPrefix(
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  text: string,
  maxWidth: number,
): string {
  const chars = [...text]; // spread handles surrogate pairs
  let lo = 0, hi = chars.length;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (ctx.measureText(chars.slice(0, mid).join('')).width <= maxWidth) lo = mid;
    else hi = mid - 1;
  }
  return chars.slice(0, lo).join('');
}

/**
 * Split a text run into layout-segment strings.
 * Each segment is an atomic unit for word-level fitting; CJK overflow is handled in layoutLines.
 */
function splitTextForLayout(text: string): string[] {
  const result: string[] = [];
  let i = 0;
  while (i < text.length) {
    let j = i;
    while (j < text.length && text[j] !== ' ') j++;
    while (j < text.length && text[j] === ' ') j++;
    if (j > i) result.push(text.slice(i, j));
    i = j;
  }
  return result.length ? result : [text];
}

function layoutLines(
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  segs: LayoutSeg[],
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
  const availW = () => maxWidth - (isFirst ? firstIndent : 0);

  const flush = (forceHeight?: number) => {
    lines.push({ segments: currentLine, height: forceHeight !== undefined ? forceHeight : (lineHeight || 10), ascent: lineAscent });
    currentLine = [];
    currentWidth = 0;
    lineHeight = 0;
    lineAscent = 0;
    isFirst = false;
  };

  const addToLine = (s: LayoutTextSeg | LayoutImageSeg, w: number, h: number, asc: number) => {
    currentLine.push(s);
    currentWidth += w;
    if (h > lineHeight) lineHeight = h;
    if (asc > lineAscent) lineAscent = asc;
  };

  // Use an explicit queue so CJK split-tails can be re-queued
  const queue: LayoutSeg[] = [...segs];

  while (queue.length > 0) {
    const seg = queue.shift()!;

    // ── Line-break sentinel ──────────────────────────────
    if ('lineBreak' in seg) {
      flush(seg.fontSize);
      continue;
    }

    // ── Image segment ────────────────────────────────────
    if ('dataUrl' in seg) {
      if (seg.anchor) { seg.measuredWidth = 0; continue; }
      const w = seg.widthPt * scale;
      const h = seg.heightPt;
      const asc = seg.heightPt * scale;
      seg.measuredWidth = w;
      if (currentLine.length > 0 && currentWidth + w > availW()) flush();
      addToLine(seg, w, h, asc);
      continue;
    }

    // ── Text segment ─────────────────────────────────────
    const s = seg as LayoutTextSeg;
    ctx.font = buildFont(s.bold, s.italic, s.fontSize * scale, s.fontFamily);
    const m = ctx.measureText(s.text);
    const w = m.width;
    const h = s.fontSize;
    const asc = m.actualBoundingBoxAscent ?? s.fontSize * scale * 0.75;

    if (currentWidth + w <= availW()) {
      // Fits on current line as-is
      s.measuredWidth = w;
      addToLine(s, w, h, asc);
    } else if (currentLine.length === 0) {
      // Nothing on the line yet — force-fit (word is wider than the whole column)
      s.measuredWidth = w;
      addToLine(s, w, h, asc);
    } else if (hasCJKBreakOpportunity(s.text)) {
      // CJK overflow: split at the maximum prefix that fits, re-queue the tail
      const available = availW() - currentWidth;
      const prefix = fitCJKPrefix(ctx, s.text, available);
      if (prefix.length > 0) {
        const pm = ctx.measureText(prefix);
        const headSeg: LayoutTextSeg = { ...s, text: prefix, measuredWidth: pm.width };
        addToLine(headSeg, pm.width, h, pm.actualBoundingBoxAscent ?? asc);
        const tail = s.text.slice(prefix.length);
        if (tail) queue.unshift({ ...s, text: tail, measuredWidth: 0 });
      } else {
        // No prefix fits — start a new line, re-queue the whole segment
        flush();
        queue.unshift(s);
      }
    } else {
      // Latin word wrap: flush and put this word on the next line
      flush();
      s.measuredWidth = w;
      addToLine(s, w, h, asc);
    }
  }

  if (currentLine.length > 0) flush();

  return lines;
}

function renderInlineImage(
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  seg: LayoutImageSeg,
  x: number,
  baseline: number,
  scale: number,
  images: Map<string, ImageBitmap>,
): void {
  // Anchor images are skipped during layout (measuredWidth=0, not added to line.segments)
  // and are drawn later by renderAnchorImages — so this function only handles inline images.
  if (seg.anchor) return;
  const bmp = images.get(imageKey(seg.dataUrl, seg.colorReplaceFrom));
  if (!bmp) return;
  const w = seg.widthPt * scale;
  const h = seg.heightPt * scale;
  ctx.drawImage(bmp, x, baseline - h, w, h);
}

/** Collect and draw all anchor images from a paragraph (called after inline flow). */
function renderAnchorImages(
  para: DocParagraph,
  state: RenderState,
  paragraphTopPx: number,
): void {
  if (state.dryRun) return;
  for (const run of para.runs) {
    if (run.type !== 'image') continue;
    const img = run as unknown as ImageRun;
    if (!img.anchor) continue;
    const bmp = state.images.get(imageKey(img.dataUrl, img.colorReplaceFrom));
    if (!bmp) continue;
    const w = img.widthPt * state.scale;
    const h = img.heightPt * state.scale;

    // Resolve X: margin-relative offsets need section.marginLeft added
    const pageX = img.anchorXFromMargin
      ? (state.marginLeft + (img.anchorXPt ?? 0)) * state.scale
      : (img.anchorXPt ?? 0) * state.scale;

    // Resolve Y: paragraph-relative offsets use the paragraph's top Y in canvas px
    const pageY = img.anchorYFromPara
      ? paragraphTopPx + (img.anchorYPt ?? 0) * state.scale
      : (img.anchorYPt ?? 0) * state.scale;

    state.ctx.drawImage(bmp, pageX, pageY, w, h);
  }
}

// ===== Table rendering =====

function renderTable(table: DocTable, state: RenderState): void {
  const { ctx, scale, contentX, contentW, dryRun } = state;

  const totalColW = table.colWidths.reduce((s, w) => s + w, 0) * scale;
  const colScale = totalColW > contentW ? contentW / totalColW : 1;
  const colWidths = table.colWidths.map(w => w * scale * colScale);

  const tableX = contentX;

  const rowHeights: number[] = [];
  for (const row of table.rows) {
    const rowH = calculateRowHeight(row, table, colWidths, scale, state);
    rowHeights.push(rowH);
  }

  let y = state.y;
  for (let ri = 0; ri < table.rows.length; ri++) {
    const row = table.rows[ri];
    const rowH = rowHeights[ri];
    let x = tableX;
    let ci = 0;

    for (const cell of row.cells) {
      const span = Math.min(cell.colSpan, colWidths.length - ci);
      const cellW = colWidths.slice(ci, ci + span).reduce((s, w) => s + w, 0);

      if (cell.vMerge !== false) {
        if (!dryRun) renderCell(cell, table, x, y, cellW, rowH, state);
        else measureCellContent(cell, table, cellW, scale, state);
      }

      x += cellW;
      ci += span;
    }

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

  let maxH = 10 * scale;
  let ci = 0;
  for (const cell of row.cells) {
    const span = Math.min(cell.colSpan, colWidths.length - ci);
    const cellW = colWidths.slice(ci, ci + span).reduce((s, w) => s + w, 0);
    const contentW = cellW - (table.cellMarginLeft + table.cellMarginRight) * scale;

    let h = (table.cellMarginTop + table.cellMarginBottom) * scale;
    for (const para of cell.content) {
      h += measureParaHeight(state, para, contentW, scale);
      h += (para.spaceBefore + para.spaceAfter) * scale;
    }
    if (h > maxH) maxH = h;
    ci += span;
  }
  return maxH;
}

function measureParaHeight(
  state: RenderState,
  para: DocParagraph,
  maxWidth: number,
  scale: number,
): number {
  const segs = buildSegments(para.runs, state);
  if (segs.length === 0) return getDefaultFontSize(para) * scale;
  const lines = layoutLines(state.ctx, segs, maxWidth, 0, scale);
  const mult = lineSpacingMultiplier(para.lineSpacing);
  return lines.reduce((s, l) => s + l.height * scale * mult, 0);
}

function measureCellContent(
  cell: DocTableCell,
  table: DocTable,
  cellW: number,
  scale: number,
  state: RenderState,
): void {
  const ml = table.cellMarginLeft * scale;
  const mr = table.cellMarginRight * scale;
  const innerW = cellW - ml - mr;
  for (const para of cell.content) {
    measureParaHeight(state, para, innerW, scale);
  }
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

  if (cell.background) {
    ctx.fillStyle = `#${cell.background}`;
    ctx.fillRect(x, y, w, h);
  }

  drawCellBorders(ctx, x, y, w, h, cell.borders, table.borders, scale);

  const mt = table.cellMarginTop * scale;
  const mb = table.cellMarginBottom * scale;
  const ml = table.cellMarginLeft * scale;
  const mr = table.cellMarginRight * scale;

  const cellState: RenderState = {
    ...state,
    contentX: x + ml,
    contentW: w - ml - mr,
    y: y + mt,
  };

  if (cell.vAlign === 'center' || cell.vAlign === 'bottom') {
    const contentH = cell.content.reduce((s, p) =>
      s + measureParaHeight(state, p, w - ml - mr, scale) + (p.spaceBefore + p.spaceAfter) * scale, 0);
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
  for (const run of para.runs) {
    if (run.type === 'text') {
      return (run as unknown as TextRun).fontSize;
    }
    if (run.type === 'field') {
      return (run as unknown as FieldRun).fontSize;
    }
  }
  return 10; // pt fallback
}

function lineSpacingMultiplier(ls: LineSpacing | null): number {
  if (!ls) return 1.2;
  if (ls.rule === 'auto') return ls.value * 1.2;
  return 1.2; // for exact/atLeast, use line value directly in pt (handled separately if needed)
}
