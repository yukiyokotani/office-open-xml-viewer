import type {
  Document, BodyElement, DocParagraph, DocTable, DocTableRow, DocTableCell,
  DocRun, TextRun, ImageRun, ShapeRun, FieldRun, HeaderFooter, LineSpacing, BorderSpec, TableBorders, CellBorders,
  TabStop, ParagraphBorders, ParaBorderEdge, SectionProps,
} from './types';
import { buildCustomPath, hexToRgba } from '@silurus/ooxml-core';

const HIGHLIGHT_COLORS: Record<string, string> = {
  yellow: '#FFFF00', cyan: '#00FFFF', green: '#00FF00', magenta: '#FF00FF',
  blue: '#0000FF', red: '#FF0000', darkBlue: '#000080', darkCyan: '#008080',
  darkGreen: '#008000', darkMagenta: '#800080', darkRed: '#800000',
  darkYellow: '#808000', darkGray: '#808080', lightGray: '#C0C0C0',
  black: '#000000', white: '#FFFFFF',
};

// 1pt = 96/72 CSS px at screen
const PT_TO_PX = 96 / 72;

/** Anchor image float that affects text wrap on the current page. */
interface FloatRect {
  mode: 'square' | 'topAndBottom';
  /** Hex key of the image bitmap (used to defer drawing until final Y is known). */
  imageKey: string;
  /** Absolute canvas X of the image box (without dist padding). */
  imageX: number;
  imageY: number;
  imageW: number;
  imageH: number;
  /** Padded exclusion rectangle for text wrap. */
  xLeft: number;
  xRight: number;
  yTop: number;
  yBottom: number;
  /** wrapText: "bothSides" | "left" | "right" | "largest" (only square uses this). */
  side: string;
  /** true once the image itself has been drawn (drawn after its paragraph lays out). */
  drawn: boolean;
}

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
  /** Active anchor-image floats that constrain text layout on the current page. */
  floats: FloatRect[];
  /** ECMA-376 §17.6.5 docGrid (type + pitch), applied to auto line spacing. */
  docGrid: DocGridCtx;
}

export interface RenderDocumentOptions {
  width?: number;
  dpr?: number;
  defaultTextColor?: string;
  /** total pages in the document (used to resolve NUMPAGES fields) */
  totalPages?: number;
  /** Pre-computed page splits (from computePages). When provided, skips internal pagination. */
  prebuiltPages?: BodyElement[][];
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
    if (!canvas.style.display) canvas.style.display = 'block';
  }

  const ctx = canvas.getContext('2d') as CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;
  ctx.scale(dpr, dpr);

  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, cssWidth, cssHeight);

  const pages = opts.prebuiltPages ?? computePages(doc.body, sec, ctx);
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
    floats: [],
    docGrid: { type: sec.docGridType ?? null, linePitchPt: sec.docGridLinePitch ?? null },
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
  renderBodyElements(elements, bodyState);
}

/**
 * Split body into pages, honoring explicit page breaks AND measuring content
 * overflow for automatic pagination. All measurements are done in pt (scale=1).
 */
export function computePages(
  body: BodyElement[],
  section: SectionProps,
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
): BodyElement[][] {
  const contentH = section.pageHeight - section.marginTop - section.marginBottom;
  const contentW = section.pageWidth - section.marginLeft - section.marginRight;
  const measureState = buildMeasureState(ctx, section);

  const pages: BodyElement[][] = [[]];
  let y = 0;
  let prevPara: DocParagraph | null = null;
  // Keep measureState.y in sync with the current page's content Y so that
  // registerAnchorFloats/WrapLayoutCtx anchor relative to where we actually
  // are on the page. Anchor floats are registered on the measureState as
  // paragraphs are processed and cleared when we flip to a new page, exactly
  // like the real renderer does.
  measureState.y = section.marginTop;
  measureState.floats = [];
  const newPage = () => {
    if (pages[pages.length - 1].length > 0) {
      pages.push([]);
      y = 0;
      prevPara = null;
      measureState.y = section.marginTop;
      measureState.floats = [];
    }
  };

  // ECMA-376 §17.3.1.15: keepNext means this paragraph must stay on the same
  // page as the next paragraph. The simplest interpretation, and what Word
  // appears to do in practice, is "treat the keepNext chain as a single unit
  // for page-break purposes" — so here we look ahead and add the next
  // paragraph's (or first line's) height to the break decision.
  const estimateNextBlockHeight = (startIdx: number): number => {
    const nxt = body[startIdx];
    if (!nxt) return 0;
    if (nxt.type === 'paragraph') {
      // We only need enough room for the first line so that "keepNext" avoids
      // orphaning the current paragraph at the bottom of a page while the
      // next begins on a new page. Using the full paragraph is safer for a
      // single-line next; for multi-line we rely on that paragraph's own
      // break logic after placing.
      return estimateParagraphHeight(measureState, nxt as unknown as DocParagraph, contentW, false);
    }
    if (nxt.type === 'table') {
      return estimateTableHeight(measureState, nxt as unknown as DocTable, contentW);
    }
    return 0;
  };

  for (let i = 0; i < body.length; i++) {
    const el = body[i];
    if (el.type === 'pageBreak') {
      pages.push([]);
      y = 0;
      prevPara = null;
      continue;
    }
    if (el.type === 'paragraph') {
      const para = el as unknown as DocParagraph;
      if (para.pageBreakBefore) newPage();
      const suppressBefore = contextualSuppressed(prevPara, para);

      // Register this paragraph's anchor-image floats on the measureState so
      // subsequent paragraphs estimate around them (text-wrap around images
      // adds lines that the float-unaware estimate would otherwise miss,
      // which caused page 2 of demo/sample-1 to spill past the bottom
      // margin). The renderer runs the same registerAnchorFloats call; by
      // mirroring it here the paginator sees the same layout.
      const paragraphAnchorY = measureState.y + (suppressBefore ? 0 : para.spaceBefore);
      registerAnchorFloats(para, measureState, paragraphAnchorY);

      const h = estimateParagraphHeight(measureState, para, contentW, suppressBefore, section.marginLeft);
      // Break if this paragraph alone doesn't fit, OR if keepNext is set and
      // placing it would leave no room for the next block on the same page.
      const needNext = para.keepNext ? estimateNextBlockHeight(i + 1) : 0;
      const needed = h + needNext;
      if (y > 0 && y + needed > contentH) newPage();
      pages[pages.length - 1].push(el);
      y += h;
      measureState.y += h;
      prevPara = para;
    } else if (el.type === 'table') {
      const tbl = el as unknown as DocTable;
      const h = estimateTableHeight(measureState, tbl, contentW);
      if (y + h > contentH) newPage();
      pages[pages.length - 1].push(el);
      y += h;
      measureState.y += h;
      prevPara = null;
    }
  }
  return pages;
}

function buildMeasureState(
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  section: SectionProps,
): RenderState {
  return {
    ctx,
    scale: 1,
    contentX: 0,
    contentW: section.pageWidth - section.marginLeft - section.marginRight,
    y: 0,
    pageH: section.pageHeight,
    defaultColor: '#000000',
    pageIndex: 0,
    totalPages: 1,
    images: new Map(),
    dryRun: true,
    marginLeft: section.marginLeft,
    floats: [],
    docGrid: { type: section.docGridType ?? null, linePitchPt: section.docGridLinePitch ?? null },
  };
}

function estimateParagraphHeight(
  state: RenderState,
  para: DocParagraph,
  contentWPt: number,
  suppressSpaceBefore = false,
  paraXPt = 0,
): number {
  const indLeft = para.indentLeft;
  const indRight = para.indentRight;
  const paraW = Math.max(1, contentWPt - indLeft - indRight);
  const segs = buildSegments(para.runs, state);
  let textH: number;
  if (segs.length === 0) {
    const fs = getDefaultFontSize(para);
    const { asc, desc } = emptyLineNaturalPx(fs, 1);
    textH = lineBoxHeight(para.lineSpacing, asc, desc, 1, state.docGrid);
  } else {
    // When anchor-image floats are active on the current page the paragraph
    // wraps around them, adding lines compared to a full-width layout. Use
    // the same WrapLayoutCtx the renderer uses so estimate and render agree.
    const wrapCtx: WrapLayoutCtx | undefined = state.floats.length > 0 ? {
      startPageY: state.y,
      paraX: paraXPt,
      floats: state.floats,
      lineBoxH: (a, d) => lineBoxHeight(para.lineSpacing, a, d, 1, state.docGrid),
      pageH: state.pageH,
    } : undefined;
    const lines = layoutLines(state.ctx, segs, paraW, para.indentFirst, 1, para.tabStops, wrapCtx);
    textH = lines.reduce((s, l) => s + lineBoxHeight(para.lineSpacing, l.ascent, l.descent, 1, state.docGrid), 0);
  }
  return textH + (suppressSpaceBefore ? 0 : para.spaceBefore) + para.spaceAfter;
}

function estimateTableHeight(state: RenderState, table: DocTable, contentWPt: number): number {
  const totalColW = table.colWidths.reduce((s, w) => s + w, 0);
  const colScale = totalColW > contentWPt ? contentWPt / totalColW : 1;
  const colWidths = table.colWidths.map((w) => w * colScale);
  let h = 0;
  for (const row of table.rows) {
    if (row.rowHeight != null) {
      h += row.rowHeight;
      continue;
    }
    let rowH = 10;
    let ci = 0;
    for (const cell of row.cells) {
      const span = Math.min(cell.colSpan, colWidths.length - ci);
      const cellW = colWidths.slice(ci, ci + span).reduce((s, w) => s + w, 0);
      const innerW = Math.max(1, cellW - table.cellMarginLeft - table.cellMarginRight);
      let ch = table.cellMarginTop + table.cellMarginBottom;
      for (const para of cell.content) {
        ch += estimateParagraphHeight(state, para, innerW);
      }
      if (ch > rowH) rowH = ch;
      ci += span;
    }
    h += rowH;
  }
  return h;
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
  renderBodyElements(hf.body, state);
}

function measureHeaderFooterHeight(hf: HeaderFooter, base: RenderState): number {
  const state: RenderState = { ...base, y: 0, dryRun: true, floats: [] };
  renderBodyElements(hf.body, state);
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

function contextualSuppressed(prev: DocParagraph | null, curr: DocParagraph): boolean {
  return !!(prev?.contextualSpacing && curr.contextualSpacing && prev.styleId && prev.styleId === curr.styleId);
}

function renderBodyElements(elements: BodyElement[], state: RenderState): void {
  let prevPara: DocParagraph | null = null;
  for (const el of elements) {
    if (el.type === 'paragraph') {
      const para = el as unknown as DocParagraph;
      renderParagraph(para, state, contextualSuppressed(prevPara, para));
      prevPara = para;
    } else if (el.type === 'table') {
      renderTable(el as unknown as DocTable, state);
      prevPara = null;
    }
  }
}

function renderParaList(paras: DocParagraph[], state: RenderState): void {
  let prevPara: DocParagraph | null = null;
  for (const para of paras) {
    renderParagraph(para, state, contextualSuppressed(prevPara, para));
    prevPara = para;
  }
}

// ===== Paragraph rendering =====

function renderParagraph(para: DocParagraph, state: RenderState, suppressSpaceBefore = false): void {
  const { ctx, scale, contentX, contentW, defaultColor, dryRun } = state;
  // Capture Y before spaceBefore — used for paragraph-relative anchor image positioning
  const paragraphStartY = state.y;

  if (!suppressSpaceBefore) state.y += para.spaceBefore * scale;

  // Register anchor floats from this paragraph (must happen after spaceBefore so that
  // paragraph-relative Y resolves against the textAreaTop, matching Word).
  registerAnchorFloats(para, state, state.y);

  // behindDoc shapes must render before text so they appear behind it.
  renderAnchorImages(para, state, paragraphStartY, 'behind');

  // If any topAndBottom float already extends past state.y, skip past it before text starts.
  state.y = skipPastTopAndBottom(state.y, state.floats);

  const textAreaTopY = state.y;

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
    const fontSizePt = getDefaultFontSize(para);
    const { asc, desc } = emptyLineNaturalPx(fontSizePt, scale);
    const emptyH = lineBoxHeight(para.lineSpacing, asc, desc, scale, state.docGrid);
    if (para.shading && !dryRun) {
      ctx.fillStyle = `#${para.shading}`;
      ctx.fillRect(contentX + indLeft, textAreaTopY, paraW, emptyH);
    }
    state.y += emptyH;
    if (para.borders && !dryRun) {
      drawParaBorders(ctx, contentX + indLeft, textAreaTopY, paraW, emptyH, para.borders, scale);
    }
    state.y += para.spaceAfter * scale;
    renderAnchorImages(para, state, paragraphStartY);
    return;
  }

  const wrapCtx: WrapLayoutCtx | undefined = state.floats.length > 0 ? {
    startPageY: state.y,
    paraX,
    floats: state.floats,
    lineBoxH: (a, d) => lineBoxHeight(para.lineSpacing, a, d, scale, state.docGrid),
    pageH: state.pageH,
  } : undefined;

  const lines = layoutLines(ctx, segments, paraW, firstLineX - paraX, scale, para.tabStops, wrapCtx);

  if (para.shading && !dryRun) {
    const totalTextH = lines.reduce((s, l) => s + lineBoxHeight(para.lineSpacing, l.ascent, l.descent, scale, state.docGrid), 0);
    ctx.fillStyle = `#${para.shading}`;
    ctx.fillRect(contentX + indLeft, textAreaTopY, paraW, totalTextH);
  }

  // ECMA-376 §17.18.44 ST_Jc: "both" and "distribute" fully justify the line
  // by expanding inter-word spaces. The last line of a "both" paragraph is
  // traditionally left-aligned (not stretched); "distribute" also stretches
  // the last line. We count whitespace chars in trailing positions of each
  // segment and divide the slack proportionally across them.
  const isJustified =
    para.alignment === 'justify' ||
    para.alignment === 'both' ||
    para.alignment === 'distribute';
  const stretchLastLine = para.alignment === 'distribute';

  const countTrailingSpaces = (s: string) => {
    let c = 0;
    for (let i = s.length - 1; i >= 0 && s[i] === ' '; i--) c++;
    return c;
  };

  for (let li = 0; li < lines.length; li++) {
    const line = lines[li];
    const firstLine = li === 0;
    const isLastLine = li === lines.length - 1;

    // Honor wrap-computed line topY (may push past topAndBottom floats).
    if (line.topY !== undefined && line.topY > state.y) state.y = line.topY;

    // Word centers the font's natural line (ascent+descent) within the expanded
    // line box — extra space from auto/exact/atLeast goes half above and half
    // below the glyphs. Baseline = top + halfExtra + ascent.
    const lineH = lineBoxHeight(para.lineSpacing, line.ascent, line.descent, scale, state.docGrid);
    const naturalLineH = line.ascent + line.descent;
    const baseline = state.y + (lineH - naturalLineH) / 2 + line.ascent;

    // Per-line X range (may be narrower than paraW when wrapping around floats).
    const lineLeft = paraX + line.xOffset;
    const lineAvailW = line.availWidth;
    let x = firstLine ? lineLeft + indFirst : lineLeft;

    if (firstLine && numPrefix && !dryRun) {
      const numFontSize = getDefaultFontSize(para) * scale;
      ctx.font = `${numFontSize}px sans-serif`;
      ctx.fillStyle = defaultColor;
      ctx.fillText(para.numbering!.text, x - numTab, baseline);
    }

    const lineWidth = line.segments.reduce((s, seg) => s + seg.measuredWidth, 0);
    let alignOffset = 0;
    if (para.alignment === 'right' || para.alignment === 'end') {
      alignOffset = lineAvailW - (x - lineLeft) - lineWidth;
    } else if (para.alignment === 'center') {
      alignOffset = (lineAvailW - (x - lineLeft) - lineWidth) / 2;
    }
    x += alignOffset;

    // Inter-word adjustment per whitespace char on this line. Positive slack
    // (lineWidth < availW) expands spaces to fill; negative slack (lineWidth >
    // availW, typically from canvas measuring ~1 px wider than Word) compresses
    // spaces so the final glyph lands on the right margin instead of overflowing.
    // Compression is capped so we never eat more than the natural width of a
    // space, and is only applied when the line is a candidate for justification
    // (jc=both/distribute, not the last line unless distribute).
    let extraPerSpace = 0;
    const applyJustify = isJustified && (!isLastLine || stretchLastLine);
    if (applyJustify) {
      let totalTrailingSpaces = 0;
      for (let si = 0; si < line.segments.length; si++) {
        const seg = line.segments[si];
        if (si === line.segments.length - 1) break; // trailing spaces on final seg don't stretch
        if ('text' in seg) totalTrailingSpaces += countTrailingSpaces((seg as LayoutTextSeg).text);
      }
      const slack = lineAvailW - (x - lineLeft) - lineWidth;
      if (totalTrailingSpaces > 0) {
        extraPerSpace = slack / totalTrailingSpaces;
        // Don't compress past zero-width spaces — limit compression to at most
        // half the widest space on the line. Estimated from default font size.
        const minExtra = -line.ascent * 0.25;
        if (extraPerSpace < minExtra) extraPerSpace = minExtra;
      }
    }

    for (let si = 0; si < line.segments.length; si++) {
      const seg = line.segments[si];
      const isLastSeg = si === line.segments.length - 1;
      if ('isTab' in seg) {
        // Tabs render as blank space; width was resolved during layout.
        x += seg.measuredWidth;
        continue;
      }
      if ('dataUrl' in seg) {
        if (!dryRun) renderInlineImage(ctx, seg as LayoutImageSeg, x, baseline, scale, state.images);
        x += seg.measuredWidth;
        continue;
      }
      const s = seg as LayoutTextSeg;
      if (!dryRun) {
        const effSizePx = calcEffectiveFontPx(s, scale);
        const yOffset = s.vertAlign === 'super'
          ? -s.fontSize * scale * 0.35
          : s.vertAlign === 'sub'
            ? s.fontSize * scale * 0.15
            : 0;
        ctx.font = buildFont(s.bold, s.italic, effSizePx, s.fontFamily);

        if (s.highlight) {
          ctx.fillStyle = HIGHLIGHT_COLORS[s.highlight] ?? '#FFFF00';
          ctx.fillRect(x, baseline + yOffset - effSizePx * 0.85, s.measuredWidth, effSizePx * 1.1);
        }

        ctx.fillStyle = s.color ? `#${s.color}` : defaultColor;
        ctx.fillText(s.text, x, baseline + yOffset);

        const lineColor = s.color ? `#${s.color}` : defaultColor;
        const lineW = Math.max(0.5, effSizePx * 0.05);
        const textW = ctx.measureText(s.text).width;

        if (s.underline) {
          ctx.strokeStyle = lineColor;
          ctx.lineWidth = lineW;
          const uy = baseline + yOffset + effSizePx * 0.12;
          ctx.beginPath(); ctx.moveTo(x, uy); ctx.lineTo(x + textW, uy); ctx.stroke();
        }

        if (s.strikethrough) {
          ctx.strokeStyle = lineColor;
          ctx.lineWidth = lineW;
          const sy = baseline + yOffset - effSizePx * 0.3;
          ctx.beginPath(); ctx.moveTo(x, sy); ctx.lineTo(x + textW, sy); ctx.stroke();
        }

        if (s.doubleStrikethrough) {
          ctx.strokeStyle = lineColor;
          ctx.lineWidth = lineW;
          const sy1 = baseline + yOffset - effSizePx * 0.35;
          const sy2 = baseline + yOffset - effSizePx * 0.22;
          ctx.beginPath(); ctx.moveTo(x, sy1); ctx.lineTo(x + textW, sy1); ctx.stroke();
          ctx.beginPath(); ctx.moveTo(x, sy2); ctx.lineTo(x + textW, sy2); ctx.stroke();
        }
      }

      x += s.measuredWidth;
      // Inter-word justification slack (applied AFTER the segment so the next
      // segment starts at a shifted baseline). Skip on the final segment —
      // trailing spaces at line end don't participate in stretching.
      if (extraPerSpace > 0 && !isLastSeg) {
        const trailing = countTrailingSpaces(s.text);
        if (trailing > 0) x += trailing * extraPerSpace;
      }
    }

    state.y += lineH;
  }

  if (para.borders && !dryRun) {
    const textH = state.y - textAreaTopY;
    drawParaBorders(ctx, contentX + indLeft, textAreaTopY, paraW, textH, para.borders, scale);
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
  vertAlign: 'super' | 'sub' | null;
  measuredWidth: number;  // px (set during layout)
  smallCaps?: boolean;
  doubleStrikethrough?: boolean;
  highlight?: string | null;
}

/**
 * Horizontal tab. Width is resolved during layout against paragraph tab stops
 * (or the default 36pt interval if no explicit stop is configured).
 */
interface LayoutTabSeg {
  isTab: true;
  fontSize: number;  // pt — for line-height purposes
  measuredWidth: number;
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

type LayoutSeg = LayoutTextSeg | LayoutImageSeg | LayoutLineBreak | LayoutTabSeg;

interface LayoutLine {
  segments: (LayoutTextSeg | LayoutImageSeg | LayoutTabSeg)[];
  height: number;  // pt — max fontSize on line (for empty-line sizing fallback)
  ascent: number;  // px — fontBoundingBoxAscent (font-metric, stable per font+size)
  descent: number; // px — fontBoundingBoxDescent
  /** Additional horizontal offset (px) from paraX, caused by wrap-around floats. */
  xOffset: number;
  /** Effective available width (px) for this line after float exclusion. */
  availWidth: number;
  /** When wrap context is active, the absolute canvas Y where this line begins. */
  topY?: number;
}

/** Additional context passed to layoutLines so it can honor floats on the current page. */
interface WrapLayoutCtx {
  startPageY: number;   // absolute canvas Y where the first line should start
  paraX: number;        // absolute canvas X of the paragraph's content left edge
  floats: FloatRect[];  // floats active on the current page
  /** Per-line box-height resolver (line natural ascent+descent → total px box height). */
  lineBoxH: (ascentPx: number, descentPx: number) => number;
  /** Hard cap on Y to keep layout from running past the page. */
  pageH: number;
}

function buildSegments(runs: DocRun[], state: RenderState): LayoutSeg[] {
  const segs: LayoutSeg[] = [];
  const pushTextPiece = (
    text: string,
    base: TextRun | FieldRun,
    vertAlign: 'super' | 'sub' | null,
  ) => {
    const displayText = (base.allCaps || base.smallCaps) ? text.toUpperCase() : text;
    for (const word of splitTextForLayout(displayText)) {
      segs.push({
        text: word,
        bold: base.bold,
        italic: base.italic,
        underline: base.underline,
        strikethrough: base.strikethrough,
        fontSize: base.fontSize,
        color: base.color,
        fontFamily: base.fontFamily,
        vertAlign,
        measuredWidth: 0,
        smallCaps: base.smallCaps ?? false,
        doubleStrikethrough: base.doubleStrikethrough ?? false,
        highlight: base.highlight ?? null,
      });
    }
  };

  for (const run of runs) {
    if (run.type === 'text') {
      const t = run as unknown as TextRun & { type: 'text' };
      // Split on tab chars so tab alignment can be resolved during layout.
      const parts = t.text.split('\t');
      for (let i = 0; i < parts.length; i++) {
        if (parts[i].length > 0) pushTextPiece(parts[i], t, t.vertAlign);
        if (i < parts.length - 1) {
          segs.push({ isTab: true, fontSize: t.fontSize, measuredWidth: 0 });
        }
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
      if (text) pushTextPiece(text, f, f.vertAlign);
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
  tabStops: TabStop[] = [],
  wrapCtx?: WrapLayoutCtx,
): LayoutLine[] {
  const lines: LayoutLine[] = [];
  let currentLine: (LayoutTextSeg | LayoutImageSeg | LayoutTabSeg)[] = [];
  let currentWidth = 0;
  // Width of trailing spaces on the last added text token. Those spaces
  // collapse if the next word wraps and the current word becomes the last on
  // the line — so we subtract them during fit checks to avoid premature wraps.
  let lastTokenTrailingSpaceW = 0;
  let lineHeight = 0;   // pt
  let lineAscent = 0;   // px
  let lineDescent = 0;  // px
  let isFirst = true;
  // Effective width/offset for the current line after float exclusion.
  let lineMaxWidth = maxWidth;
  let lineXOffset = 0;
  let currentLineTopY = wrapCtx?.startPageY ?? 0;

  // Compute wrap constraints for a new line about to start. Mutates lineXOffset/lineMaxWidth/currentLineTopY.
  const startLine = (): void => {
    lineXOffset = 0;
    lineMaxWidth = maxWidth;
    if (!wrapCtx) return;
    // Probe height: the smallest plausible line height; good enough for float intersection check.
    const probeH = 10 * scale;
    // Keep pushing past any topAndBottom block we sit inside.
    for (let guard = 0; guard < 16; guard++) {
      const lineBot = currentLineTopY + probeH;
      let skip: number | null = null;
      for (const f of wrapCtx.floats) {
        if (f.mode !== 'topAndBottom') continue;
        if (lineBot > f.yTop && currentLineTopY < f.yBottom) {
          skip = skip === null ? f.yBottom : Math.max(skip, f.yBottom);
        }
      }
      if (skip === null) break;
      currentLineTopY = skip;
    }
    // Now compute horizontal constraint from square floats.
    const paraXLeft = wrapCtx.paraX;
    const paraXRight = wrapCtx.paraX + maxWidth;
    let left = paraXLeft;
    let right = paraXRight;
    const lineBot = currentLineTopY + probeH;
    for (const f of wrapCtx.floats) {
      if (f.mode !== 'square') continue;
      if (lineBot <= f.yTop || currentLineTopY >= f.yBottom) continue;
      // Decide which side text should flow on. "left"/"right" refer to the side TEXT occupies.
      const spaceLeft = f.xLeft - paraXLeft;
      const spaceRight = paraXRight - f.xRight;
      let textOnLeft: boolean;
      switch (f.side) {
        case 'left':    textOnLeft = true;  break;
        case 'right':   textOnLeft = false; break;
        case 'largest':
        case 'bothSides':
        default:        textOnLeft = spaceLeft >= spaceRight; break;
      }
      if (textOnLeft) {
        if (f.xLeft < right) right = Math.max(left, f.xLeft);
      } else {
        if (f.xRight > left) left = Math.min(right, f.xRight);
      }
    }
    const eff = Math.max(0, right - left);
    lineXOffset = Math.max(0, left - paraXLeft);
    lineMaxWidth = Math.min(maxWidth - lineXOffset, eff);
    if (lineMaxWidth < 0) lineMaxWidth = 0;
  };
  startLine();

  const availW = () => lineMaxWidth - (isFirst ? firstIndent : 0);

  // Default tab interval when no matching explicit stop exists (Word's default is 720 twips = 36pt)
  const DEFAULT_TAB_PT = 36;

  const flush = (forceHeight?: number) => {
    const h = forceHeight !== undefined ? forceHeight : (lineHeight || 10);
    // If the line has no measured content (empty/line-break line), synthesize
    // stable ascent/descent from the effective font size so wrap/baseline math
    // stays consistent with non-empty lines.
    const hasContent = lineAscent > 0 || lineDescent > 0;
    const asc = hasContent ? lineAscent : h * scale * 0.8;
    const desc = hasContent ? lineDescent : h * scale * 0.2;
    lines.push({
      segments: currentLine,
      height: h,
      ascent: asc,
      descent: desc,
      xOffset: lineXOffset,
      availWidth: lineMaxWidth,
      topY: wrapCtx ? currentLineTopY : undefined,
    });
    if (wrapCtx) {
      currentLineTopY += wrapCtx.lineBoxH(asc, desc);
    }
    currentLine = [];
    currentWidth = 0;
    lastTokenTrailingSpaceW = 0;
    lineHeight = 0;
    lineAscent = 0;
    lineDescent = 0;
    isFirst = false;
    startLine();
  };

  const addToLine = (
    s: LayoutTextSeg | LayoutImageSeg | LayoutTabSeg,
    w: number,
    h: number,
    asc: number,
    desc: number,
    trailingSpaceW: number = 0,
  ) => {
    currentLine.push(s);
    currentWidth += w;
    lastTokenTrailingSpaceW = trailingSpaceW;
    if (h > lineHeight) lineHeight = h;
    if (asc > lineAscent) lineAscent = asc;
    if (desc > lineDescent) lineDescent = desc;
  };

  const effectiveFontPx = (s: LayoutTextSeg): number => calcEffectiveFontPx(s, scale);

  const measureText = (s: LayoutTextSeg): TextMetrics => {
    ctx.font = buildFont(s.bold, s.italic, effectiveFontPx(s), s.fontFamily);
    return ctx.measureText(s.text);
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

    // ── Tab segment ──────────────────────────────────────
    if ('isTab' in seg) {
      // Absolute position on the line measured from paraX (line origin for continuation lines)
      const absFromParaX = currentWidth + (isFirst ? firstIndent : 0);
      // Find the next tab stop strictly greater than the current position
      const stop = tabStops.find((t) => t.pos * scale > absFromParaX);
      let tabWidth: number;
      if (stop) {
        tabWidth = stop.pos * scale - absFromParaX;
      } else {
        // Round up to the next DEFAULT_TAB_PT boundary
        const nextDefault = Math.ceil((absFromParaX + 0.01) / (DEFAULT_TAB_PT * scale)) * (DEFAULT_TAB_PT * scale);
        tabWidth = nextDefault - absFromParaX;
      }
      // Clamp to avoid negative widths; if tab would overflow the line, wrap instead
      if (tabWidth <= 0) {
        flush();
        queue.unshift(seg);
        continue;
      }
      if (currentWidth + tabWidth > availW() && currentLine.length > 0) {
        flush();
        queue.unshift(seg);
        continue;
      }
      seg.measuredWidth = tabWidth;
      addToLine(seg, tabWidth, seg.fontSize, seg.fontSize * scale * 0.8, seg.fontSize * scale * 0.2);
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
      addToLine(seg, w, h, asc, 0);
      continue;
    }

    // ── Text segment ─────────────────────────────────────
    const s = seg as LayoutTextSeg;
    const m = measureText(s);
    const w = m.width;
    // Line-height tracks the un-scaled pt font so super/sub don't shrink the line.
    const h = s.fontSize;
    // Prefer font-metric ascent/descent (stable per font+size) so baselines and
    // line boxes do not jitter based on the specific characters on each line.
    const asc = m.fontBoundingBoxAscent ?? m.actualBoundingBoxAscent ?? s.fontSize * scale * 0.8;
    const desc = m.fontBoundingBoxDescent ?? m.actualBoundingBoxDescent ?? s.fontSize * scale * 0.2;
    // Trailing spaces collapse at line breaks in Word, so a wrap-fit check
    // should treat both (a) the candidate word's own trailing space AND
    // (b) the current line's last token trailing space as collapsible. That
    // keeps wrap decisions spec-accurate and prevents premature wraps that
    // would otherwise bloat justify slack (e.g. dropping "narrows" to line 2
    // when it still fits after "trail "'s trailing space collapses).
    const trimmed = s.text.replace(/ +$/, '');
    const trailingSpaceW = s.text.endsWith(' ')
      ? w - ctx.measureText(trimmed).width
      : 0;
    const wForFit = w - trailingSpaceW;
    const currentWidthNoTail = currentWidth - lastTokenTrailingSpaceW;

    if (currentWidthNoTail + wForFit <= availW()) {
      // Fits on current line as-is
      s.measuredWidth = w;
      addToLine(s, w, h, asc, desc, trailingSpaceW);
    } else if (hasCJKBreakOpportunity(s.text)) {
      // CJK overflow: split at the maximum prefix that fits, re-queue the tail
      const available = availW() - currentWidth;
      ctx.font = buildFont(s.bold, s.italic, effectiveFontPx(s), s.fontFamily);
      const prefix = available > 0 ? fitCJKPrefix(ctx, s.text, available) : '';
      if (prefix.length > 0) {
        const pm = ctx.measureText(prefix);
        const headSeg: LayoutTextSeg = { ...s, text: prefix, measuredWidth: pm.width };
        addToLine(headSeg, pm.width, h, asc, desc);
        const tail = s.text.slice(prefix.length);
        if (tail) queue.unshift({ ...s, text: tail, measuredWidth: 0 });
      } else if (currentLine.length > 0) {
        // No prefix fits but line has content — flush and retry on a fresh line
        flush();
        queue.unshift(s);
      } else {
        // Empty line and not even one char fits — force-fit one char to guarantee progress
        const firstChar = [...s.text][0] ?? '';
        if (firstChar) {
          const fm = ctx.measureText(firstChar);
          const headSeg: LayoutTextSeg = { ...s, text: firstChar, measuredWidth: fm.width };
          addToLine(headSeg, fm.width, h, asc, desc);
          const tail = s.text.slice(firstChar.length);
          if (tail) queue.unshift({ ...s, text: tail, measuredWidth: 0 });
        }
      }
    } else if (currentLine.length === 0) {
      // Nothing on the line yet and no CJK break — force-fit (word wider than column)
      s.measuredWidth = w;
      addToLine(s, w, h, asc, desc);
    } else {
      // Latin word wrap: flush and put this word on the next line
      flush();
      s.measuredWidth = w;
      addToLine(s, w, h, asc, desc);
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

/** Collect and draw anchor images with wrapMode='none' (or unspecified).
 * Wrap floats (square/topAndBottom/tight/through) are drawn by registerAnchorFloats.
 *
 * `phase` = 'behind' draws only shapes with behindDoc=true (sorted by zOrder asc);
 * `phase` = 'front' draws shapes without behindDoc + all anchor images. */
function renderAnchorImages(
  para: DocParagraph,
  state: RenderState,
  paragraphTopPx: number,
  phase: 'behind' | 'front' = 'front',
): void {
  if (state.dryRun) return;
  if (phase === 'behind') {
    const shapes = para.runs
      .filter((r): r is ShapeRun & { type: 'shape' } =>
        r.type === 'shape' && !!(r as unknown as ShapeRun).behindDoc)
      .slice()
      .sort((a, b) =>
        ((a as unknown as ShapeRun).zOrder ?? 0) - ((b as unknown as ShapeRun).zOrder ?? 0));
    for (const s of shapes) renderAnchorShape(s as unknown as ShapeRun, state, paragraphTopPx);
    return;
  }
  for (const run of para.runs) {
    if (run.type === 'shape') {
      const s = run as unknown as ShapeRun;
      if (s.behindDoc) continue;
      renderAnchorShape(s, state, paragraphTopPx);
      continue;
    }
    if (run.type !== 'image') continue;
    const img = run as unknown as ImageRun;
    if (!img.anchor) continue;
    if (isWrapFloat(img.wrapMode)) continue;  // drawn as a float
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

/** Draw a wps:wsp shape via core's custGeom primitive. */
function renderAnchorShape(shape: ShapeRun, state: RenderState, paragraphTopPx: number): void {
  const { ctx, scale } = state;
  const w = shape.widthPt * scale;
  const h = shape.heightPt * scale;
  if (w <= 0 || h <= 0) return;
  const x = shape.anchorXFromMargin
    ? (state.marginLeft + shape.anchorXPt) * scale
    : shape.anchorXPt * scale;
  const y = shape.anchorYFromPara
    ? paragraphTopPx + shape.anchorYPt * scale
    : shape.anchorYPt * scale;

  const rot = shape.rotation ?? 0;
  ctx.save();
  if (rot !== 0) {
    ctx.translate(x + w / 2, y + h / 2);
    ctx.rotate((rot * Math.PI) / 180);
    ctx.translate(-(x + w / 2), -(y + h / 2));
  }
  ctx.beginPath();
  buildCustomPath(ctx as CanvasRenderingContext2D, shape.subpaths, x, y, w, h);
  const fillStyle = resolveShapeFillStyle(shape.fill, ctx as CanvasRenderingContext2D, x, y, w, h);
  if (fillStyle) {
    ctx.fillStyle = fillStyle;
    ctx.fill();
  }
  if (shape.stroke && (shape.strokeWidth ?? 0) > 0) {
    ctx.strokeStyle = hexToRgba(shape.stroke);
    ctx.lineWidth = Math.max(0.5, (shape.strokeWidth ?? 0) * scale);
    ctx.stroke();
  }
  ctx.restore();
}

function resolveShapeFillStyle(
  fill: ShapeRun['fill'],
  ctx: CanvasRenderingContext2D,
  x: number, y: number, w: number, h: number,
): string | CanvasGradient | null {
  if (!fill) return null;
  if (fill.fillType === 'solid') return hexToRgba(fill.color);
  if (fill.fillType === 'gradient') {
    if (fill.stops.length === 0) return null;
    if (fill.stops.length === 1) return hexToRgba(fill.stops[0].color);
    let gradient: CanvasGradient;
    if (fill.gradType === 'radial') {
      const cx = x + w / 2, cy = y + h / 2;
      const r = Math.sqrt(w * w + h * h) / 2;
      gradient = ctx.createRadialGradient(cx, cy, 0, cx, cy, r);
    } else {
      const rad = (fill.angle * Math.PI) / 180;
      const cx = x + w / 2, cy = y + h / 2;
      const gradLen = (Math.abs(Math.cos(rad)) * w + Math.abs(Math.sin(rad)) * h) / 2;
      gradient = ctx.createLinearGradient(
        cx - Math.cos(rad) * gradLen, cy - Math.sin(rad) * gradLen,
        cx + Math.cos(rad) * gradLen, cy + Math.sin(rad) * gradLen,
      );
    }
    for (const s of fill.stops) {
      gradient.addColorStop(Math.min(1, Math.max(0, s.position)), hexToRgba(s.color));
    }
    return gradient;
  }
  return null;
}

function isWrapFloat(mode?: string): boolean {
  return mode === 'square' || mode === 'topAndBottom' || mode === 'tight' || mode === 'through';
}

/** Register floats from a paragraph's anchor images and draw the image bitmap immediately. */
function registerAnchorFloats(para: DocParagraph, state: RenderState, paragraphAnchorY: number): void {
  for (const run of para.runs) {
    if (run.type !== 'image') continue;
    const img = run as unknown as ImageRun;
    if (!img.anchor) continue;
    if (!isWrapFloat(img.wrapMode)) continue;

    const mode: 'square' | 'topAndBottom' =
      img.wrapMode === 'topAndBottom' ? 'topAndBottom' : 'square';

    const scale = state.scale;
    const w = img.widthPt * scale;
    const h = img.heightPt * scale;
    const pageX = img.anchorXFromMargin
      ? (state.marginLeft + (img.anchorXPt ?? 0)) * scale
      : (img.anchorXPt ?? 0) * scale;
    const pageY = img.anchorYFromPara
      ? paragraphAnchorY + (img.anchorYPt ?? 0) * scale
      : (img.anchorYPt ?? 0) * scale;
    const dt = (img.distTop    ?? 0) * scale;
    const db = (img.distBottom ?? 0) * scale;
    const dl = (img.distLeft   ?? 0) * scale;
    const dr = (img.distRight  ?? 0) * scale;

    const key = imageKey(img.dataUrl, img.colorReplaceFrom);
    const rect: FloatRect = {
      mode,
      imageKey: key,
      imageX: pageX,
      imageY: pageY,
      imageW: w,
      imageH: h,
      xLeft: pageX - dl,
      xRight: pageX + w + dr,
      yTop: pageY - dt,
      yBottom: pageY + h + db,
      side: img.wrapSide ?? 'bothSides',
      drawn: false,
    };
    state.floats.push(rect);

    if (!state.dryRun) {
      const bmp = state.images.get(key);
      if (bmp) state.ctx.drawImage(bmp, rect.imageX, rect.imageY, rect.imageW, rect.imageH);
      rect.drawn = true;
    }
  }
}

/** If y is inside a topAndBottom float, return the float bottom; otherwise return y. */
function skipPastTopAndBottom(y: number, floats: FloatRect[]): number {
  for (let guard = 0; guard < 16; guard++) {
    let next = y;
    for (const f of floats) {
      if (f.mode !== 'topAndBottom') continue;
      if (y >= f.yTop && y < f.yBottom) next = Math.max(next, f.yBottom);
    }
    if (next === y) return y;
    y = next;
  }
  return y;
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
  if (segs.length === 0) {
    const fs = getDefaultFontSize(para);
    const { asc, desc } = emptyLineNaturalPx(fs, scale);
    return lineBoxHeight(para.lineSpacing, asc, desc, scale, state.docGrid);
  }
  const lines = layoutLines(state.ctx, segs, maxWidth, 0, scale, para.tabStops);
  return lines.reduce((s, l) => s + lineBoxHeight(para.lineSpacing, l.ascent, l.descent, scale, state.docGrid), 0);
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

  renderParaList(cell.content, cellState);
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

function drawParaBorders(
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D,
  x: number, y: number, w: number, h: number,
  borders: ParagraphBorders,
  scale: number,
): void {
  const drawEdge = (edge: ParaBorderEdge | null, x1: number, y1: number, x2: number, y2: number) => {
    if (!edge || edge.style === 'none') return;
    const spec: BorderSpec = { width: edge.width, color: edge.color, style: edge.style };
    drawBorderLine(ctx, x1, y1, x2, y2, spec, scale);
  };
  const sp = (edge: ParaBorderEdge | null) => (edge?.space ?? 0) * scale;
  drawEdge(borders.top,    x, y - sp(borders.top),         x + w, y - sp(borders.top));
  drawEdge(borders.bottom, x, y + h + sp(borders.bottom),  x + w, y + h + sp(borders.bottom));
  drawEdge(borders.left,   x - sp(borders.left), y,        x - sp(borders.left), y + h);
  drawEdge(borders.right,  x + w + sp(borders.right), y,   x + w + sp(borders.right), y + h);
}

// ===== Utilities =====

function calcEffectiveFontPx(s: LayoutTextSeg, scale: number): number {
  let size = s.fontSize * scale;
  if (s.smallCaps) size *= 0.8;
  if (s.vertAlign) size *= 0.65;
  return size;
}

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
  if (typeof para.defaultFontSize === 'number') return para.defaultFontSize;
  return 10; // pt fallback
}

/** Document-grid context passed to line-box computation.  When the section's
 *  `w:docGrid` is "lines"/"linesAndChars" with a positive pitch (ECMA-376
 *  §17.6.5), auto line spacing multiplies against the grid pitch instead of
 *  the font's natural line height. Without this, a 56-pt heading with
 *  lineRule="auto" value=4.33 would claim 56×1.25×4.33 ≈ 303pt of vertical
 *  space; with this, it claims max(natural, 18pt × 4.33) ≈ 78pt — matching
 *  Word's rendering on grids typical of Japanese/Chinese templates. */
interface DocGridCtx {
  /** "default" | "lines" | "linesAndChars" | "snapToChars" */
  type: string | null | undefined;
  /** Grid pitch in pt (already converted from twips in the parser). */
  linePitchPt: number | null | undefined;
}

function isGridLineRule(ctx: DocGridCtx | undefined): boolean {
  if (!ctx || !ctx.linePitchPt || ctx.linePitchPt <= 0) return false;
  return ctx.type === 'lines' || ctx.type === 'linesAndChars';
}

/**
 * Compute the total line-box height in px from a line's natural font metrics
 * (fontBoundingBoxAscent + fontBoundingBoxDescent) per ECMA-376 §17.3.1.33.
 *
 *   auto    → natural × value ("single" = 1 natural line, "double" = 2).
 *             When docGrid type=lines|linesAndChars is active, the
 *             multiplier applies against the grid pitch instead, with a
 *             floor of the natural line height.
 *   exact   → value in pt, converted to px (ignores font and grid).
 *   atLeast → max(natural, value in pt × scale).
 *   null    → natural, or grid pitch if the section defines one.
 */
function lineBoxHeight(
  ls: LineSpacing | null,
  ascentPx: number,
  descentPx: number,
  scale: number,
  grid?: DocGridCtx,
): number {
  const natural = ascentPx + descentPx;
  const hasGrid = isGridLineRule(grid);
  if (!ls) {
    return hasGrid ? Math.max(natural, grid!.linePitchPt! * scale) : natural;
  }
  if (ls.rule === 'auto') {
    if (hasGrid) {
      return Math.max(natural, grid!.linePitchPt! * ls.value * scale);
    }
    return natural * ls.value;
  }
  if (ls.rule === 'exact') return ls.value * scale;
  if (ls.rule === 'atLeast') return Math.max(natural, ls.value * scale);
  return natural;
}

/** Natural single-line height in px for an empty paragraph (no rendered text). */
function emptyLineNaturalPx(fontSizePt: number, scale: number): { asc: number; desc: number } {
  return { asc: fontSizePt * scale * 0.8, desc: fontSizePt * scale * 0.2 };
}
