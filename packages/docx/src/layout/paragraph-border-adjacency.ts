import type { ParagraphBorders, ParaBorderEdge } from '../types.js';
import type { ParagraphLayoutSource } from './text.js';
import { WORD_PARAGRAPH_BORDER_FLOW_RESERVATION } from './paint-compatibility.js';

export type ParagraphBorderEdges = Readonly<{
  top: 'top' | 'between' | 'none';
  bottom: 'bottom' | 'none';
}>;

export interface ParagraphBorderMerge {
  suppressTop?: boolean;
  suppressBottom?: boolean;
}

export interface ParagraphBorderSegment {
  side: 'top' | 'bottom' | 'left' | 'right';
  edge: ParaBorderEdge;
  x1: number;
  y1: number;
  x2: number;
  y2: number;
}

/** WORD_PARAGRAPH_BORDER_FLOW_RESERVATION: the total thickness of a
 * paragraph border of ST_Border style `style` (§17.18.2) and authored stroke
 * width `width` (w:sz / 8): the sum of its strokes and the gaps between them. */
export function wordParagraphBorderThicknessPt(style: string, width: number): number {
  void WORD_PARAGRAPH_BORDER_FLOW_RESERVATION;
  const w = Math.max(0, width);
  switch (style) {
    case 'double': return 3 * w;
    case 'triple': return 5 * w;
    case 'thinThickSmallGap': case 'thickThinSmallGap': return w + 1.5;
    case 'thinThickThinSmallGap': return w + 3;
    case 'thinThickMediumGap': case 'thickThinMediumGap': return 2 * w;
    case 'thinThickThinMediumGap': return 3 * w;
    case 'thinThickLargeGap': case 'thickThinLargeGap': return w + 2.25;
    case 'thinThickThinLargeGap': return 2 * w + 3;
    case 'threeDEmboss': case 'threeDEngrave': return w + 2 * Math.min(1.5, Math.max(0.75, w / 2));
    case 'wave': return w === 1.5 ? 3.75 : 3;
    case 'doubleWave': return 5.25;
    default: return w;
  }
}

/** How far a paragraph border's strokes lie outside the border box (text ±
 * `w:space`): its whole thickness. */
export function paragraphBorderStrokeOutsetPt(edge: ParaBorderEdge): number {
  return wordParagraphBorderThicknessPt(edge.style, edge.width ?? 0);
}

/** WORD_PARAGRAPH_BORDER_FLOW_RESERVATION: a wave paragraph border paints
 * hairline zigzags filling its band, whatever its w:sz: one for `wave`, two
 * stacked for `doubleWave`. */
export function wordParagraphWaveBorder(
  style: string,
  width: number,
): Readonly<{ bandPt: number; rows: 1 | 2; hairlinePt: number }> | undefined {
  void WORD_PARAGRAPH_BORDER_FLOW_RESERVATION;
  if (style !== 'wave' && style !== 'doubleWave') return undefined;
  return Object.freeze({
    bandPt: wordParagraphBorderThicknessPt(style, width),
    rows: style === 'wave' ? 1 : 2,
    hairlinePt: WORD_PARAGRAPH_WAVE_HAIRLINE_PT,
  });
}

const WORD_PARAGRAPH_WAVE_HAIRLINE_PT = 0.375;

/** Half the block extent a retained paragraph border segment paints about its
 * centre line: half its whole thickness (the stroke pattern is centred in it). */
export function paragraphBorderSegmentHalfExtentPt(
  border: Readonly<{ authoredStyle: string; widthPt: number }>,
): number {
  return wordParagraphBorderThicknessPt(border.authoredStyle, border.widthPt) / 2;
}

/** Vertical extent reserved below text for a visible bottom paragraph border
 * (WORD_PARAGRAPH_BORDER_FLOW_RESERVATION). */
export function bottomBorderExtentPt(
  borders: ParagraphBorders | null | undefined,
  merge?: ParagraphBorderMerge,
): number {
  if (!borders || merge?.suppressBottom) return 0;
  const bottom = borders.bottom;
  if (!bottom || bottom.style === 'none') return 0;
  return (bottom.space ?? 0) + paragraphBorderStrokeOutsetPt(bottom);
}

/** Reserve the painted outer edge of a visible top paragraph border above its
 * text (WORD_PARAGRAPH_BORDER_FLOW_RESERVATION). A between border separating
 * two paragraphs of one border group takes its spacing on both sides of the
 * stroke: the previous paragraph's bottom edge is suppressed in the group.
 * A grouped paragraph whose top edge is suppressed reserves nothing. */
export function topBorderExtentPt(
  borders: ParagraphBorders | null | undefined,
  edge: ParagraphBorderEdges['top'],
): number {
  if (!borders || edge === 'none') return 0;
  const top = borders[edge];
  if (!top || top.style === 'none' || top.style === 'nil') return 0;
  const space = top.space ?? 0;
  return (edge === 'between' ? 2 * space : space) + paragraphBorderStrokeOutsetPt(top);
}

function effectiveEdge(edge: ParaBorderEdge | null | undefined): ParaBorderEdge | null {
  return edge == null || edge.style === 'none' ? null : edge;
}

function sameParagraphEdge(a: ParaBorderEdge | null | undefined, b: ParaBorderEdge | null | undefined): boolean {
  const left = effectiveEdge(a);
  const right = effectiveEdge(b);
  if (left == null || right == null) return left == null && right == null;
  return left.style === right.style
    && left.width === right.width
    && (left.space ?? 0) === (right.space ?? 0)
    && (left.color ?? null) === (right.color ?? null);
}

function sameParagraphBorders(
  a: ParagraphBorders | null | undefined,
  b: ParagraphBorders | null | undefined,
): boolean {
  if (!a || !b) return false;
  return sameParagraphEdge(a.top, b.top)
    && sameParagraphEdge(a.bottom, b.bottom)
    && sameParagraphEdge(a.left, b.left)
    && sameParagraphEdge(a.right, b.right)
    && sameParagraphEdge(a.between, b.between);
}

export function hasVisibleParagraphBorder(
  borders: ParagraphBorders | null | undefined,
): boolean {
  if (!borders) return false;
  return [borders.top, borders.right, borders.bottom, borders.left, borders.between]
    .some((edge) => edge != null && edge.style !== 'none');
}

/** Pure §17.3.1.7 matching predicate; callers supply actual flow adjacency. */
export function paragraphsShareBorderBox(
  previous: ParagraphLayoutSource | null,
  current: ParagraphLayoutSource | null,
): boolean {
  if (!previous || !current || previous.framePr || current.framePr) return false;
  return hasVisibleParagraphBorder(previous.borders)
    && hasVisibleParagraphBorder(current.borders)
    && sameParagraphBorders(previous.borders, current.borders);
}

/** Resolves edge ownership once for any retained paragraph flow container. */
export function resolveParagraphBorderEdges(
  previous: ParagraphLayoutSource | null,
  current: ParagraphLayoutSource,
  next: ParagraphLayoutSource | null,
  groupedFrameFlow = false,
): ParagraphBorderEdges {
  const shares = (left: ParagraphLayoutSource | null, right: ParagraphLayoutSource | null): boolean =>
    groupedFrameFlow
      ? !!left && !!right
        && !!left.framePr && !!right.framePr
        && hasVisibleParagraphBorder(left.borders)
        && hasVisibleParagraphBorder(right.borders)
        && sameParagraphBorders(left.borders, right.borders)
      : paragraphsShareBorderBox(left, right);
  const joinsPrevious = shares(previous, current);
  const joinsNext = shares(current, next);
  const between = current.borders?.between;
  return Object.freeze({
    top: joinsPrevious
      ? between && between.style !== 'none' ? 'between' : 'none'
      : 'top',
    bottom: joinsNext ? 'none' : 'bottom',
  });
}
