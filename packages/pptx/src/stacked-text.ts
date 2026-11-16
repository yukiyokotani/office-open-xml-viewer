// Stacked vertical text, `a:bodyPr@vert` `wordArtVert` / `wordArtVertRtl`
// (ECMA-376 §20.1.10.83). Layout lives in core (`layoutStackedText`, with the
// face and character data in `stacked-faces.ts`); this module turns PowerPoint
// segments into measured glyphs and paints the placed glyphs.
//
// Glyph classes (issue #1626 controls, PowerPoint reference PDF export):
//   upright        — every character outside the sideways set: stands upright,
//                    centred on the column axis, in a cell of 7/6 × face box ×
//                    size; its baseline sits usWinDescent × size (the face
//                    box descent) above the cell bottom.
//   verticalGlyph  — a sideways character the drawing face has a vertical
//                    glyph for: a 1 em cell, baseline face-box ascent below
//                    the cell top, the face box centred on the axis.
//   rotated        — a sideways character without one: turned 90° clockwise,
//                    advancing by its horizontal advance, face box centred on
//                    the axis (ascent to the page right).
// Canvas cannot select a face's `vert` glyph portably, so a vertical glyph is
// drawn as its Unicode vertical presentation form where one exists (、。 and
// the fullwidth brackets) and otherwise as the character turned 90° clockwise
// inside its 1 em cell, which is how those faces design the remaining forms
// (ー ～ ＝ ＿ ￣ ｜ … ‥ ‐ arrows ― ∥). A vertical glyph from vrt2-only coverage
// (`nominal`, e.g. Yu Gothic / Meiryo ∥) stands unturned as its horizontal
// glyph in the presentation-form frame (#1827 controls; see stacked-faces.ts).
//
// §20.1.10.82 names decoration styles but does not place them in stacked text.
// PowerPoint electronic PDF controls (Arial 18/24/32pt, both modes, mixed sizes
// and wrapped columns) put one underline on the LEFT of each column, independently
// of its progression; strikes cross each upright glyph in its own frame.
// Character/number markers are upright stacked prefixes (including punctuation),
// consuming cells only at paragraph start. marL + indent moves the whole prefix;
// continuation columns start at marL (§21.1.2.2.7). These controls establish the
// frames, not new per-face decoration metrics: offsets/weights and dash/wave
// shapes reuse the horizontal painter's policy in the font-derived cell frame.

import {
  layoutStackedText,
  paintStackedRotated,
  paintStackedVerticalForm,
  STACKED_CELL_FACTOR,
  stackedClusters,
  stackedFaceBox,
  stackedUprightBaseline,
  stackedVerticalGlyphText,
  type DrawingMlInputRun,
  type DrawingMlLineSpacing,
  type DrawingMlTextRect,
  type StackedGlyph,
  type StackedParagraph,
} from '@silurus/ooxml-core/internal/drawingml-text';
import type { HyperlinkTarget } from '@silurus/ooxml-core';

/** The segment fields stacked layout and paint read (renderer's LayoutSegment). */
export interface StackedSegmentStyle {
  math?: { width: number; ascent: number; descent: number };
  font: string;
  sizePx: number;
  color: string;
  noFill?: boolean;
  strikethrough?: boolean;
  letterSpacingPx?: number;
  faceFamily?: string;
  faceFamilyLatin?: string;
  /** Face box (em) of the drawing face: usWin, or typo under USE_TYPO_METRICS. */
  lineMetric?: { readonly glyph: { ascent: number; descent: number } | undefined };
  /** Face box of the run's latin face. */
  lineMetricLatin?: { readonly glyph: { ascent: number; descent: number } | undefined } | null;
  /** The run's resolved hyperlink, handed to the selection overlay per glyph. */
  hyperlink?: HyperlinkTarget;
}

export interface StackedParagraphInput<T extends StackedSegmentStyle> {
  runs: readonly DrawingMlInputRun<T>[];
  alignment?: string;
  marginStart?: number;
  marginEnd?: number;
  firstLineIndent?: number;
  lineSpacing?: DrawingMlLineSpacing;
  spaceBefore?: DrawingMlLineSpacing;
  spaceAfter?: DrawingMlLineSpacing;
  /** Style of the paragraph mark (sizes an empty column). */
  markStyle: T;
  /** Explicit endParaRPr; absent inherits the last visible run's decoration. */
  endMarkStyle?: T;
  eastAsianLineBreak?: boolean;
  marker?: { text: string; style: T } | { advance: number; thickness: number };
}

export interface StackedBodyInput<T extends StackedSegmentStyle> {
  vert: 'wordArtVert' | 'wordArtVertRtl';
  rect: DrawingMlTextRect;
  anchor?: string;
  anchorCtr?: boolean;
  wrap: boolean;
  spcFirstLastPara?: boolean;
  lnSpcReduction?: number;
  pxPerPt: number;
  paragraphs: readonly StackedParagraphInput<T>[];
  /** Existing equation painter, supplied by the renderer's math-resource owner. */
  drawObject?(style: T, x: number, y: number): void;
  /** Underlines use the common +90° column frame and column font size. */
  drawUnderline?(style: T, width: number, columnSizePx: number): void;
  sameUnderline?(a: T, b: T): boolean;
  /** Strikes use each glyph's own frame, baseline and measured width. */
  drawStrike?(style: T, x: number, baseline: number, width: number): void;
  drawMarker?(paragraph: number, axisX: number, top: number): void;
  markerPaint?(paragraph: number, axisX: number, baseline: number): string | CanvasGradient | CanvasPattern;
  sameStyle?(a: T, b: T): boolean;
}

export interface StackedGlyphRun {
  text: string;
  font: string;
  fontSize: number;
  /** Glyph cell in body coordinates. */
  x: number;
  y: number;
  w: number;
  h: number;
  /** Zero-based index of the authored line (a paragraph, or the part of one
   *  between manual line breaks) the glyph belongs to. A line that wraps into
   *  further columns keeps its index. */
  logicalLine: number;
  hyperlink?: HyperlinkTarget;
}

type Ctx2D = CanvasRenderingContext2D;

/**
 * Face box of a segment in em: the #1610 glyph box of its face, else the
 * catalogued box of that family at any weight, else the resolved Canvas font
 * box (an unknown face).
 */
function faceBox(ctx: Ctx2D, style: StackedSegmentStyle): { ascent: number; descent: number } {
  const known = style.lineMetric?.glyph
    ?? (style.faceFamily
      ? stackedFaceBox(style.faceFamily, /\bbold\b/.test(style.font), /\bitalic\b/.test(style.font)) : undefined);
  if (known) return known;
  ctx.font = style.font;
  const m = ctx.measureText('M');
  const size = style.sizePx || 1;
  const ascent = m.fontBoundingBoxAscent;
  const descent = m.fontBoundingBoxDescent;
  if (Number.isFinite(ascent) && Number.isFinite(descent) && ascent + descent > 0) {
    return { ascent: ascent / size, descent: descent / size };
  }
  return { ascent: 0.9, descent: 0.25 };
}

interface Measured<T> extends StackedGlyph<T> {
  box: { ascent: number; descent: number };
  cell: number;
  /** A vertical glyph drawn unturned as its horizontal glyph. */
  nominal?: true;
}

function makeMeasurer<T extends StackedSegmentStyle>(ctx: Ctx2D) {
  const boxes = new Map<T, { ascent: number; descent: number }>();
  return (text: string, style: T, upright = false): Measured<T>[] => {
    let box = boxes.get(style);
    if (!box) {
      box = faceBox(ctx, style);
      boxes.set(style, box);
    }
    const family = style.faceFamily ?? '';
    const cellEm = box.ascent + box.descent;
    const cell = STACKED_CELL_FACTOR * cellEm * style.sizePx;
    // The run's latin face also sizes the column (MS Mincho glyphs of a run
    // whose latin face is Arial sit in an Arial-thick column, wordartvert
    // slides 10-11), as it sizes a horizontal line (#1610).
    const latinFamily = style.faceFamilyLatin ?? family;
    const latinBox = style.lineMetricLatin?.glyph
      ?? stackedFaceBox(latinFamily, /\bbold\b/.test(style.font), /\bitalic\b/.test(style.font));
    const latinEm = latinFamily === family ? cellEm
      : latinBox ? latinBox.ascent + latinBox.descent : cellEm;
    const thickness = Math.max(cell, STACKED_CELL_FACTOR * latinEm * style.sizePx);
    const spc = style.letterSpacingPx ?? 0;
    const face = { family, bold: /\bbold\b/.test(style.font), italic: /\bitalic\b/.test(style.font) };
    return stackedClusters(text, face, {
      upright: cell,
      verticalGlyph: style.sizePx,
      rotated: (ch) => {
        ctx.font = style.font;
        return ctx.measureText(ch).width;
      },
    }, upright).map(({ text: ch, kind, advance, nominal }) => (
      { text: ch, style, kind, advance: advance + spc, thickness, space: ch === ' ', box, cell, ...(nominal ? { nominal } : {}) }));
  };
}

/**
 * Lay out and paint a stacked body. `ctx` is in slide coordinates; the body
 * rectangle has already been placed by the caller (shape rotation and flips
 * are applied outside, as for horizontal text). Returns one run per glyph for
 * the selection overlay.
 */
export function renderStackedText<T extends StackedSegmentStyle>(
  ctx: Ctx2D,
  body: StackedBodyInput<T>,
): StackedGlyphRun[] {
  const measure = makeMeasurer<T>(ctx);
  const markers = body.paragraphs.map(p => p.marker && 'text' in p.marker
    ? measure(p.marker.text, p.marker.style, true) : []);
  const paragraphs: StackedParagraph<T>[] = body.paragraphs.map((p, index) => ({
    // Upright OMML's height is the advance along a stacked column; the
    // horizontal equation width remains its cross-column extent.
    runs: p.runs.map((run) => run.type === 'object' && run.style.math
      ? { ...run, width: run.style.math.ascent + run.style.math.descent } : run),
    alignment: p.alignment,
    marginStart: p.marginStart,
    marginEnd: p.marginEnd,
    firstLineIndent: p.firstLineIndent,
    firstLinePrefix: p.marker && ('text' in p.marker ? {
      advance: markers[index].reduce((sum, g) => sum + g.advance, 0),
      thickness: markers[index].reduce((max, g) => Math.max(max, g.thickness), 0),
    } : p.marker),
    lineSpacing: p.lineSpacing,
    spaceBefore: p.spaceBefore,
    spaceAfter: p.spaceAfter,
    emptyThickness: measure('M', p.markStyle)[0].thickness,
    eastAsianLineBreak: p.eastAsianLineBreak,
  }));
  const layout = layoutStackedText<T, Measured<T>>(paragraphs, {
    direction: body.vert,
    rect: body.rect,
    anchor: body.anchor,
    anchorCtr: body.anchorCtr,
    wrap: body.wrap,
    spcFirstLastPara: body.spcFirstLastPara,
    lnSpcReduction: body.lnSpcReduction,
    pxPerPt: body.pxPerPt,
    glyphs: measure,
    sameStyle: body.sameStyle,
    objectGlyph: ({ style }) => {
      if (!style.math || !body.drawObject) throw new Error('Missing stacked equation adapter');
      const height = style.math.ascent + style.math.descent;
      return { text: '', style, kind: 'upright', advance: height, thickness: style.math.width,
        space: false, box: { ascent: 0, descent: 0 }, cell: height };
    },
  });
  // One pass gathers column metrics; painting below never scans the document
  // again per column. Keep the stacked cell's extra leading on the ascent side:
  // rotated baseline = column left + face descent, not the raw face-box centre.
  // Omitting that leading moves uniform underlines right by half (cell-faceBox).
  const columnMetrics = layout.columns.map(() => ({ descent: 0, size: 0 }));
  for (const g of layout.glyphs) {
    if (g.style.math) continue;
    const metric = columnMetrics[g.column];
    metric.descent = Math.max(metric.descent, g.box.descent * g.style.sizePx);
    metric.size = Math.max(metric.size, g.style.sizePx);
  }
  if (body.drawUnderline) {
    let span: { style: T; column: number; top: number; bottom: number } | undefined;
    const flush = () => {
      if (!span) return;
      const column = layout.columns[span.column];
      const metric = columnMetrics[span.column];
      ctx.save();
      ctx.translate(column.axisX - column.thickness / 2 + metric.descent, span.top);
      ctx.rotate(Math.PI / 2);
      body.drawUnderline!(span.style, span.bottom - span.top, metric.size);
      ctx.restore();
    };
    for (const g of layout.glyphs) {
      if (g.style.math) { flush(); span = undefined; continue; }
      if (span && span.column === g.column && (body.sameUnderline?.(span.style, g.style) ?? span.style === g.style)) {
        span.bottom = g.cellTop + g.advance;
      } else {
        flush();
        span = { style: g.style, column: g.column, top: g.cellTop, bottom: g.cellTop + g.advance };
      }
    }
    flush();
  }
  const runs: StackedGlyphRun[] = [];
  const prevAlign = ctx.textAlign;
  const prevBaseline = ctx.textBaseline;
  const paintGlyph = (
    g: Measured<T>, axisX: number, cellTop: number, markerParagraph?: number, logicalLine = 0,
  ) => {
    const marker = markerParagraph !== undefined;
    const { style, box } = g;
    if (style.math) {
      body.drawObject!(style, axisX - g.thickness / 2, cellTop);
      return;
    }
    const size = style.sizePx;
    if (!marker) runs.push({
      text: g.text, font: style.font, fontSize: size,
      x: axisX - g.thickness / 2, y: cellTop, w: g.thickness, h: g.advance, logicalLine,
      ...(style.hyperlink ? { hyperlink: style.hyperlink } : {}),
    });
    ctx.font = style.font;
    const strike = (text: string, x: number, baseline: number, centered = false) => {
      if (marker || !style.strikethrough || !body.drawStrike) return;
      const width = ctx.measureText(text).width;
      body.drawStrike(style, centered ? x - width / 2 : x, baseline, width);
    };
    if (style.noFill) return;
    ctx.fillStyle = marker && body.markerPaint
      ? body.markerPaint(markerParagraph, axisX, cellTop + g.cell - box.descent * size) : style.color;
    if (g.kind === 'upright') {
      ctx.textAlign = 'center';
      ctx.textBaseline = 'alphabetic';
      const baseline = stackedUprightBaseline(cellTop, g.cell, box, size);
      strike(g.text, axisX, baseline, true);
      if (!g.space) ctx.fillText(g.text, axisX, baseline);
    } else if (g.kind === 'verticalGlyph') {
      // A `vert` cluster with a mark has no presentation form and takes the
      // turned fallback.
      const form = stackedVerticalGlyphText(g);
      if (form !== null) {
        paintStackedVerticalForm(ctx, form, axisX, cellTop, box, size, (left, baseline) => strike(form, left, baseline));
      } else {
        const left = axisX - (box.ascent + box.descent) * size / 2;
        ctx.save();
        ctx.translate(left + size / 2, cellTop + size / 2);
        ctx.rotate(Math.PI / 2);
        ctx.textAlign = 'center';
        if (!marker && style.strikethrough && body.drawStrike) {
          ctx.textBaseline = 'alphabetic';
          const alphabeticAscent = ctx.measureText(g.text).fontBoundingBoxAscent;
          ctx.textBaseline = 'middle';
          const middleAscent = ctx.measureText(g.text).fontBoundingBoxAscent;
          // TextMetrics ascent is relative to the active baseline. Subtract
          // to locate the alphabetic strike frame without moving the glyph;
          // engines without these metrics use the segment's face-box centre.
          const baseline = Number.isFinite(alphabeticAscent - middleAscent)
            ? alphabeticAscent - middleAscent : (box.ascent - box.descent) * size / 2;
          strike(g.text, 0, baseline, true);
        }
        ctx.textBaseline = 'middle';
        ctx.fillText(g.text, 0, 0);
        ctx.restore();
      }
    } else {
      paintStackedRotated(ctx, g.text, axisX, cellTop, box, size, () => strike(g.text, 0, 0));
    }
  };
  let glyphIndex = 0;
  let logicalLine = 0;
  for (const [columnIndex, column] of layout.columns.entries()) {
    if (column.prefixTop !== undefined) {
      let top = column.prefixTop;
      for (const g of markers[column.paragraph]) {
        paintGlyph(g, column.axisX, top, column.paragraph);
        top += g.advance;
      }
      body.drawMarker?.(column.paragraph, column.axisX, column.prefixTop);
    }
    let last: typeof layout.glyphs[number] | undefined;
    while (glyphIndex < layout.glyphs.length && layout.glyphs[glyphIndex].column === columnIndex) {
      const g = layout.glyphs[glyphIndex++];
      paintGlyph(g, g.axisX, g.cellTop, undefined, logicalLine);
      last = g;
    }
    // The same uniform/mixed-size Office controls strike the invisible final
    // paragraph mark too. Use endParaRPr when authored, otherwise the last
    // run's face; a measured space supplies the invisible mark's width.
    // It paints after the last cell but adds no glyph, wrap or selection cell.
    const endStyle = body.paragraphs[column.paragraph].endMarkStyle ?? last?.style;
    if (last && endStyle && !endStyle.math && endStyle.strikethrough && body.drawStrike
      && layout.columns[columnIndex + 1]?.paragraph !== column.paragraph) {
      const mark = measure(' ', endStyle)[0];
      ctx.font = endStyle.font;
      const width = ctx.measureText(' ').width;
      body.drawStrike(endStyle, column.axisX - width / 2,
        last.cellTop + last.advance + mark.cell - mark.box.descent * endStyle.sizePx, width);
    }
    if (column.endsLine) logicalLine++;
  }
  ctx.textAlign = prevAlign;
  ctx.textBaseline = prevBaseline;
  return runs;
}
