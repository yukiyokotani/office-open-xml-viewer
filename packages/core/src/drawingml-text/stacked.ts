import { breakDrawingMlText, type DrawingMlLineSegment, type DrawingMlInputRun } from './break.js';
import {
  drawingMlParagraphSpacing,
  drawingMlSpacedLineBox,
  type DrawingMlLineSpacing,
  type DrawingMlTextRect,
} from './metrics.js';
import { drawingMlLineShouldJustify } from './align.js';

/**
 * Stacked vertical text: `a:bodyPr@vert` `wordArtVert` ("one letter on top of
 * another") and `wordArtVertRtl` (the same, columns right to left), ECMA-376
 * §20.1.10.83. The standard names the modes only; the layout below is
 * PowerPoint's, measured with the issue #1626 control decks (reference
 * "electronic distribution" PDF export; the export scales text y by 1.0015,
 * which the measurements remove).
 *
 * - Glyphs stay upright and advance down a column by their cell (host
 *   supplied: see `stacked-faces.ts`), plus rPr@spc. Line breaking is the
 *   shared DrawingML breaker with the column length as the line width.
 * - Every column is a line box rotated +90°: thickness = its largest cell,
 *   split half and half around the glyph axis, spaced by `a:lnSpc` with the
 *   shared {@link drawingMlSpacedLineBox} rule. Its ascent lies on the page
 *   right in both directions. When spacing makes a line shorter than its
 *   natural box, the body's last line keeps its natural descent.
 * - In the rotated frame a paragraph is [spcBef][lines][spcAft] from right
 *   to left; `wordArtVertRtl` stacks these items right to left. The same,
 *   unmirrored items stack left to right for `wordArtVert`, so there spcAft
 *   lies left of a paragraph's columns and spcBef right. spcFirstLastPara
 *   suppresses the first spcBef and last spcAft as for horizontal text.
 * - `anchor` places the column block across the columns (t = the start
 *   side: left for wordArtVert, right for wordArtVertRtl); `anchorCtr` centres
 *   the block's longest column along the columns. Insets stay physical.
 * - `algn` works along each column: l / ctr / r = top / centre / bottom; just
 *   stretches the spaces of a line that is not a paragraph's last or ended by
 *   a:br; dist spreads the glyphs of every line (one glyph is centred).
 *   Trailing spaces take no part in alignment.
 */

export type StackedDirection = 'wordArtVert' | 'wordArtVertRtl';

/** How one glyph stands in the column (see `stacked-faces.ts`). */
export type StackedGlyphKind = 'upright' | 'verticalGlyph' | 'rotated';

/** One glyph as the host measured it, in canvas px. */
export interface StackedGlyph<T> {
  text: string;
  style: T;
  kind: StackedGlyphKind;
  /** Advance down the column, including rPr@spc. */
  advance: number;
  /** Thickness this glyph gives its column (its face cell). */
  thickness: number;
  /** An ordinary space: stretched by `just`, ignored at a column end. */
  space: boolean;
}

export interface StackedParagraph<T> {
  runs: readonly DrawingMlInputRun<T>[];
  alignment?: string;
  /** Paragraph margins and indent in the column reading direction (px). */
  marginStart?: number;
  marginEnd?: number;
  firstLineIndent?: number;
  /** Atomic paragraph marker in the reading direction. It precedes text on
   * the first column only; its advance participates in wrapping/alignment. */
  firstLinePrefix?: { advance: number; thickness: number };
  lineSpacing?: DrawingMlLineSpacing;
  spaceBefore?: DrawingMlLineSpacing;
  spaceAfter?: DrawingMlLineSpacing;
  /** Column thickness of a line without glyphs (the paragraph mark's cell). */
  emptyThickness: number;
  eastAsianLineBreak?: boolean;
}

export interface StackedLayoutOptions<T, G extends StackedGlyph<T> = StackedGlyph<T>> {
  direction: StackedDirection;
  /** The text rectangle after insets, in canvas px. */
  rect: DrawingMlTextRect;
  anchor?: string;
  anchorCtr?: boolean;
  /** `wrap="square"`; false never breaks a column. */
  wrap: boolean;
  spcFirstLastPara?: boolean;
  /** `a:normAutofit@lnSpcReduction` as a fraction. */
  lnSpcReduction?: number;
  pxPerPt: number;
  /** Split a text segment into measured glyphs (one grapheme each). */
  glyphs(text: string, style: T): G[];
  /** Atomic inline object in column coordinates. Its advance is its upright
   * rendered height and its thickness is its rendered width. OMML remains one
   * math zone (§22.1.2.77); this is library layout policy, not an Office pitch
   * inference from the letter-stacking rule in §20.1.10.83. */
  objectGlyph?(segment: Extract<DrawingMlLineSegment<T>, { type: 'object' }>): G;
  sameStyle?(a: T, b: T): boolean;
}

/**
 * A host glyph with its place: column axis x and the top of its cell. `span`
 * is the index of the uninterrupted span of authored text the glyph belongs
 * to: it advances at a paragraph end, a manual break, a tab, an inline object
 * and a wrap that dropped spaces, and not at a wrap inside a word.
 */
export type StackedPlacedGlyph<G> = G & { axisX: number; cellTop: number; column: number; span: number };

export interface StackedColumn {
  axisX: number;
  /** Natural thickness (largest cell). */
  thickness: number;
  paragraph: number;
  /** Start of an authored first-column marker, after indent and alignment. */
  prefixTop?: number;
}

export interface StackedLayout<G> {
  glyphs: StackedPlacedGlyph<G>[];
  columns: StackedColumn[];
  /** Width of the column block (across the columns). */
  blockWidth: number;
}

interface Line<G> {
  /** The line's glyphs, each with its text span (see {@link StackedPlacedGlyph}). */
  glyphs: (G & { span: number })[];
  paragraph: number;
  lastInParagraph: boolean;
  endsWithBreak: boolean;
  thickness: number;
  firstInParagraph: boolean;
  ascent: number;
  descent: number;
}

const sumAdvance = <T>(glyphs: readonly StackedGlyph<T>[]): number =>
  glyphs.reduce((sum, g) => sum + g.advance, 0);

function contentGlyphs<T>(glyphs: readonly StackedGlyph<T>[]): number {
  let end = glyphs.length;
  while (end > 0 && glyphs[end - 1].space) end--;
  return end;
}

export function layoutStackedText<T, G extends StackedGlyph<T> = StackedGlyph<T>>(
  paragraphs: readonly StackedParagraph<T>[],
  options: StackedLayoutOptions<T, G>,
): StackedLayout<G> {
  const { rect, pxPerPt } = options;
  const reduction = options.lnSpcReduction ?? 0;
  const lines: Line<G>[] = [];
  // Items along the progression, in the rotated frame's own order
  // ([spcBef][lines][spcAft] per paragraph); spacing is a plain gap.
  const items: ({ type: 'gap'; size: number } | { type: 'line'; line: Line<G> })[][] = [];

  let span = 0;
  paragraphs.forEach((para, index) => {
    const broken = breakDrawingMlText(para.runs, {
      maxWidth: options.wrap ? Math.max(0, rect.height - (para.marginStart ?? 0) - (para.marginEnd ?? 0)) : Number.POSITIVE_INFINITY,
      firstLineIndent: (para.firstLineIndent ?? 0) + (para.firstLinePrefix?.advance ?? 0),
      measureText: (text, style) => sumAdvance(options.glyphs(text, style)),
      sameStyle: options.sameStyle,
      eastAsianLineBreak: para.eastAsianLineBreak,
    });
    const paraLines: Line<G>[] = broken.map((line, i) => {
      const glyphs = line.segments.flatMap((seg) => {
        let made: G[] = [];
        if (seg.type === 'text') made = options.glyphs(seg.text, seg.style);
        else {
          // A tab or an inline object separates the text around it: the text
          // after it is two spans on, and an object is the span in between.
          span += 2;
          if (seg.type === 'object') {
            if (!options.objectGlyph) throw new Error('Stacked inline objects require a host object adapter');
            const glyph = options.objectGlyph(seg);
            // Library fallback policy: a zero-size object (unavailable OMML in
            // PPTX) contributes no inline glyph. Keep the breaker's display
            // boundaries so an otherwise empty column uses the paragraph mark,
            // just as the host's empty horizontal line does (§22.1.2.77–78).
            if (glyph.advance !== 0 || glyph.thickness !== 0) made = [glyph];
          }
        }
        return made.map((glyph) => ({ ...glyph, span: seg.type === 'text' ? span : span - 1 }));
      });
      if (i === broken.length - 1 || line.endsWithBreak || line.dropsSpace) span++;
      // Iterative maxima throughout: a single run can hold far more glyphs
      // than a call's argument list (no spread into Math.max).
      let thickness = glyphs.length > 0 ? 0 : para.emptyThickness;
      for (const g of glyphs) if (g.thickness > thickness) thickness = g.thickness;
      if (i === 0 && para.firstLinePrefix) thickness = Math.max(thickness, para.firstLinePrefix.thickness);
      const natural = { ascent: thickness / 2, descent: thickness / 2 };
      const box = drawingMlSpacedLineBox(natural, para.lineSpacing, pxPerPt, reduction);
      return {
        glyphs, paragraph: index, firstInParagraph: i === 0, lastInParagraph: i === broken.length - 1,
        endsWithBreak: line.endsWithBreak === true, thickness,
        ascent: box.ascent, descent: box.descent,
      };
    });
    if (paraLines.length === 0) {
      const t = para.emptyThickness;
      const box = drawingMlSpacedLineBox({ ascent: t / 2, descent: t / 2 }, para.lineSpacing, pxPerPt, reduction);
      paraLines.push({ glyphs: [], paragraph: index, firstInParagraph: true, lastInParagraph: true, endsWithBreak: false, thickness: t,
        ascent: box.ascent, descent: box.descent });
    }
    const first = index === 0;
    const last = index === paragraphs.length - 1;
    const edges = options.spcFirstLastPara === true;
    const before = first && !edges ? 0
      : drawingMlParagraphSpacing(para.spaceBefore, paraLines[0].thickness, pxPerPt);
    const after = last && !edges ? 0
      : drawingMlParagraphSpacing(para.spaceAfter, paraLines[paraLines.length - 1].thickness, pxPerPt);
    const paraItems: ({ type: 'gap'; size: number } | { type: 'line'; line: Line<G> })[] = [];
    if (before) paraItems.push({ type: 'gap', size: before });
    for (const line of paraLines) {
      paraItems.push({ type: 'line', line });
      lines.push(line);
    }
    if (after) paraItems.push({ type: 'gap', size: after });
    items.push(paraItems);
  });

  // A spaced line shorter than its natural box: the body's last line keeps
  // its natural descent (wordartvert-supp2 LS/LR sweeps, both directions).
  const lastLine = lines[lines.length - 1];
  if (lastLine && lastLine.descent < lastLine.thickness / 2) lastLine.descent = lastLine.thickness / 2;

  const rtl = options.direction === 'wordArtVertRtl';
  // Progression coordinate u from the block's start side. wordArtVertRtl
  // walks each paragraph's items in order; wordArtVert walks them reversed
  // but keeps each line box unmirrored (descent on its left).
  const axisU: number[] = [];
  let u = 0;
  for (const paraItems of items) {
    const ordered = rtl ? paraItems : paraItems.slice().reverse();
    // Lines inside a paragraph keep their order in both directions.
    const lineItems = paraItems.filter((item) => item.type === 'line');
    let lineCursor = 0;
    for (const item of ordered) {
      if (item.type === 'gap') {
        u += item.size;
        continue;
      }
      const line = (lineItems[lineCursor++] as { type: 'line'; line: Line<G> }).line;
      axisU.push(u + (rtl ? line.ascent : line.descent));
      u += line.ascent + line.descent;
    }
  }
  const blockWidth = u;
  const slack = rect.width - blockWidth;
  const offset = options.anchor === 'ctr' ? slack / 2 : options.anchor === 'b' ? Math.max(0, slack) : 0;

  // Along the columns: the region the alignment works in.
  let regionTop = rect.top;
  let regionLength = rect.height;
  if (options.anchorCtr) {
    let longest = 0;
    for (const l of lines) longest = Math.max(longest,
      sumAdvance(l.glyphs.slice(0, contentGlyphs(l.glyphs)))
        + (l.firstInParagraph ? paragraphs[l.paragraph].firstLinePrefix?.advance ?? 0 : 0));
    regionTop = rect.top + (rect.height - longest) / 2;
    regionLength = longest;
  }

  const glyphs: StackedPlacedGlyph<G>[] = [];
  const columns: StackedColumn[] = [];
  lines.forEach((line, column) => {
    const axisX = rtl ? rect.left + rect.width - offset - axisU[column] : rect.left + offset + axisU[column];
    const paragraph = paragraphs[line.paragraph];
    const alignment = paragraph.alignment;
    const prefix = line.firstInParagraph ? paragraph.firstLinePrefix : undefined;
    const prefixAdvance = prefix?.advance ?? 0;
    const leading = (paragraph.marginStart ?? 0)
      + (line.firstInParagraph ? paragraph.firstLineIndent ?? 0 : 0) + prefixAdvance;
    const trailing = paragraph.marginEnd ?? 0;
    const columnTop = regionTop + leading;
    const count = contentGlyphs(line.glyphs);
    const length = sumAdvance(line.glyphs.slice(0, count));
    const extra = regionLength - leading - trailing - length;
    let start = columnTop;
    let gapAfter: (i: number) => number = () => 0;
    const lastInPara = line.lastInParagraph;
    const distribute = alignment === 'dist' || alignment === 'thaiDist';
    if (distribute) {
      if (count <= 1) start = columnTop + extra / 2;
      else gapAfter = (i) => (i < count - 1 ? Math.max(0, extra) / (count - 1) : 0);
    } else if (drawingMlLineShouldJustify(alignment, lastInPara, line.endsWithBreak)) {
      const spaces = line.glyphs.slice(0, count).filter((g) => g.space).length;
      if (spaces > 0) gapAfter = (i) => (i < count && line.glyphs[i].space ? Math.max(0, extra) / spaces : 0);
    } else if (alignment === 'ctr') {
      start = columnTop + extra / 2;
    } else if (alignment === 'r') {
      start = columnTop + extra;
    }
    columns.push({ axisX, thickness: line.thickness, paragraph: line.paragraph,
      ...(prefix ? { prefixTop: start - prefixAdvance } : {}),
    });
    let y = start;
    line.glyphs.forEach((g, i) => {
      glyphs.push({ ...g, axisX, cellTop: y, column });
      y += g.advance + gapAfter(i);
    });
  });
  return { glyphs, columns, blockWidth };
}
