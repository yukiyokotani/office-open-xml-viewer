import type { Slide, SlideElement, SlideElementOrigin, TextBody } from './types.js';
import type { TextSelectionContextOptions } from '@silurus/ooxml-core';
import { boundedChartContextText } from '@silurus/ooxml-core/internal/chart-context';

/** Bounds for a PPTX text-or-element selection-context snapshot. */
export type PptxSelectionContextOptions = TextSelectionContextOptions;

export interface PptxSlidePoint {
  readonly x: number;
  readonly y: number;
}

export interface PptxElementContextOptions {
  /** Extra hit radius in slide EMU for line-like shapes. Default 0. */
  readonly tolerance?: number;
  /** Maximum returned element text. Default 16,384; hard maximum 65,536. */
  readonly maxTextCharacters?: number;
}

export interface PptxElementContext {
  readonly format: 'pptx';
  readonly kind: 'element';
  readonly slideIndex: number;
  /** Paint-order index in this rendered slide snapshot; not an editor tree id. */
  readonly elementIndex: number;
  readonly origin: SlideElementOrigin | 'unknown';
  readonly elementType: SlideElement['type'];
  readonly point: PptxSlidePoint;
  readonly bounds: Readonly<{
    x: number;
    y: number;
    width: number;
    height: number;
    rotation: number;
    flipH: boolean;
    flipV: boolean;
  }>;
  readonly shapeId?: string;
  readonly name?: string;
  readonly geometry?: string;
  readonly text?: string;
  readonly mimeType?: string;
  readonly mediaKind?: 'audio' | 'video';
  readonly rowCount?: number;
  readonly columnCount?: number;
  readonly seriesCount?: number;
  readonly truncated: boolean;
  readonly truncationReasons: readonly ('text')[];
  readonly textCharacters: number;
  readonly maxTextCharacters: number;
}

/** Geometry resolved from a DrawingML element id. Used by modern-comment
 * anchors and available to primitive-UI consumers without exposing the mutable
 * slide model. */
export interface PptxElementBounds {
  readonly elementId: string;
  readonly elementIndex: number;
  readonly origin: SlideElementOrigin | 'unknown';
  readonly elementType: SlideElement['type'];
  readonly bounds: PptxElementContext['bounds'];
}

export type PptxSelectionContext =
  | import('./selection-context.js').PptxTextSelectionContext
  | import('./selection-context.js').PptxCommentSelectionContext
  | PptxElementContext;

export const MAX_ELEMENT_TEXT_CHARACTERS = 65_536;
const DEFAULT_ELEMENT_TEXT_CHARACTERS = 16_384;
const LINE_GEOMETRIES = new Set(['line', 'straightconnector1']);

function inverseFramePoint(element: SlideElement, point: PptxSlidePoint): PptxSlidePoint {
  const cx = element.x + element.width / 2;
  const cy = element.y + element.height / 2;
  const radians = (-element.rotation * Math.PI) / 180;
  const cos = Math.cos(radians);
  const sin = Math.sin(radians);
  const dx = point.x - cx;
  const dy = point.y - cy;
  let x = cx + cos * dx - sin * dy;
  let y = cy + sin * dx + cos * dy;
  if (element.flipH) x = 2 * cx - x;
  if (element.flipV) y = 2 * cy - y;
  return { x, y };
}

function distanceToSegment(point: PptxSlidePoint, start: PptxSlidePoint, end: PptxSlidePoint): number {
  const dx = end.x - start.x;
  const dy = end.y - start.y;
  const lengthSquared = dx * dx + dy * dy;
  if (lengthSquared === 0) return Math.hypot(point.x - start.x, point.y - start.y);
  const projection = Math.max(0, Math.min(
    1,
    ((point.x - start.x) * dx + (point.y - start.y) * dy) / lengthSquared,
  ));
  return Math.hypot(
    point.x - (start.x + projection * dx),
    point.y - (start.y + projection * dy),
  );
}

export function hitTestPptxElement(
  element: SlideElement,
  point: PptxSlidePoint,
  tolerance = 0,
): boolean {
  if (!Number.isFinite(point.x) || !Number.isFinite(point.y)) return false;
  const local = inverseFramePoint(element, point);
  const safeTolerance = Number.isFinite(tolerance) && tolerance > 0 ? tolerance : 0;
  if (element.type === 'shape' &&
      (LINE_GEOMETRIES.has(element.geometry.toLowerCase()) || element.width === 0 || element.height === 0)) {
    return distanceToSegment(
      local,
      { x: element.x, y: element.y },
      { x: element.x + element.width, y: element.y + element.height },
    ) <= safeTolerance;
  }
  const minX = Math.min(element.x, element.x + element.width);
  const maxX = Math.max(element.x, element.x + element.width);
  const minY = Math.min(element.y, element.y + element.height);
  const maxY = Math.max(element.y, element.y + element.height);
  return local.x >= minX && local.x <= maxX && local.y >= minY && local.y <= maxY;
}

interface ElementTextPiece {
  readonly text: string;
  readonly beginsPart: boolean;
}

function* textBodyParts(body: TextBody | null): Iterable<ElementTextPiece> {
  if (!body) return;
  for (const paragraph of body.paragraphs) {
    let beginsPart = true;
    for (const run of paragraph.runs) {
      const text = run.type === 'text' ? run.text : run.type === 'break' ? '\n' : '[equation]';
      if (!text) continue;
      yield { text, beginsPart };
      beginsPart = false;
    }
  }
}

function* elementTextParts(element: SlideElement): Iterable<ElementTextPiece> {
  if (element.type === 'shape') {
    yield* textBodyParts(element.textBody);
    return;
  }
  if (element.type === 'table') {
    for (const row of element.rows) {
      for (const cell of row.cells) yield* textBodyParts(cell.textBody);
    }
    return;
  }
}

function safeUtf16Prefix(value: string, maxCodeUnits: number): string {
  let end = Math.min(value.length, maxCodeUnits);
  if (end > 0 && end < value.length) {
    const previous = value.charCodeAt(end - 1);
    const next = value.charCodeAt(end);
    if (previous >= 0xD800 && previous <= 0xDBFF && next >= 0xDC00 && next <= 0xDFFF) end--;
  }
  return value.slice(0, end);
}

export function limitPptxElementContext(
  context: PptxElementContext,
  requestedMaxTextCharacters = DEFAULT_ELEMENT_TEXT_CHARACTERS,
): PptxElementContext {
  if (!Number.isFinite(requestedMaxTextCharacters) || requestedMaxTextCharacters < 0) {
    throw new RangeError('maxTextCharacters must be a finite non-negative number.');
  }
  const maxTextCharacters = Math.min(
    MAX_ELEMENT_TEXT_CHARACTERS,
    Math.floor(requestedMaxTextCharacters),
  );
  const sourceText = context.text;
  const text = sourceText === undefined
    ? undefined
    : safeUtf16Prefix(sourceText, maxTextCharacters);
  const newlyTruncated = sourceText !== undefined && text!.length < sourceText.length;
  return {
    ...structuredClone(context),
    ...(text === undefined ? {} : { text }),
    truncated: context.truncated || newlyTruncated,
    truncationReasons: context.truncated || newlyTruncated ? ['text'] : [],
    textCharacters: text?.length ?? 0,
    maxTextCharacters,
  };
}

function boundedJoinedText(parts: Iterable<ElementTextPiece>, maxCharacters: number): {
  text?: string;
  truncated: boolean;
  textCharacters: number;
} {
  const chunks: string[] = [];
  let length = 0;
  let truncated = false;
  let found = false;
  for (const part of parts) {
    found = true;
    if (part.beginsPart && chunks.length > 0) {
      if (length >= maxCharacters) { truncated = true; break; }
      chunks.push('\n');
      length++;
    }
    const allowed = Math.max(0, maxCharacters - length);
    const chunk = safeUtf16Prefix(part.text, allowed);
    chunks.push(chunk);
    length += chunk.length;
    if (chunk.length < part.text.length) { truncated = true; break; }
  }
  return found
    ? { text: chunks.join(''), truncated, textCharacters: length }
    : { truncated: false, textCharacters: 0 };
}

export function hitTestPptxSlideContext(
  slideIndex: number,
  slide: Slide,
  point: PptxSlidePoint,
  options: PptxElementContextOptions = {},
): PptxElementContext | null {
  if (!Number.isFinite(point.x) || !Number.isFinite(point.y)) {
    throw new RangeError('PPTX hit-test point must contain finite coordinates.');
  }
  const requestedTextMax = options.maxTextCharacters ?? DEFAULT_ELEMENT_TEXT_CHARACTERS;
  if (!Number.isFinite(requestedTextMax) || requestedTextMax < 0) {
    throw new RangeError('maxTextCharacters must be a finite non-negative number.');
  }
  const maxTextCharacters = Math.min(MAX_ELEMENT_TEXT_CHARACTERS, Math.floor(requestedTextMax));
  const tolerance = options.tolerance ?? 0;
  if (!Number.isFinite(tolerance) || tolerance < 0) {
    throw new RangeError('tolerance must be a finite non-negative number.');
  }
  for (let elementIndex = slide.elements.length - 1; elementIndex >= 0; elementIndex--) {
    const element = slide.elements[elementIndex];
    if (!hitTestPptxElement(element, point, tolerance)) continue;
    const boundedText = element.type === 'chart'
      ? boundedChartContextText(element.chart, maxTextCharacters)
      : boundedJoinedText(elementTextParts(element), maxTextCharacters);
    return {
      format: 'pptx',
      kind: 'element',
      slideIndex,
      elementIndex,
      origin: slide.elementSources?.[elementIndex]?.origin ?? 'unknown',
      elementType: element.type,
      point: { ...point },
      bounds: {
        x: element.x, y: element.y, width: element.width, height: element.height,
        rotation: element.rotation, flipH: element.flipH, flipV: element.flipV,
      },
      ...(element.type === 'shape' ? {
        ...(element.id === undefined ? {} : { shapeId: element.id }),
        ...(element.name === undefined ? {} : { name: element.name }),
        geometry: element.geometry,
      } : {}),
      ...(boundedText.text === undefined ? {} : { text: boundedText.text }),
      ...(element.type === 'picture' ? { mimeType: element.mimeType } : {}),
      ...(element.type === 'media' ? { mimeType: element.mimeType, mediaKind: element.mediaKind } : {}),
      ...(element.type === 'table' ? {
        rowCount: element.rows.length,
        columnCount: element.cols.length,
      } : {}),
      ...(element.type === 'chart' ? { seriesCount: element.chart.series.length } : {}),
      truncated: boundedText.truncated,
      truncationReasons: boundedText.truncated ? ['text'] : [],
      textCharacters: boundedText.textCharacters,
      maxTextCharacters,
    };
  }
  return null;
}

export function findPptxElementBoundsByIds(
  slide: Slide,
  elementIds: readonly string[],
): readonly PptxElementBounds[] {
  const requested = new Set(elementIds.filter((id) => id.length > 0));
  if (requested.size === 0) return Object.freeze([]);
  const found = new Map<string, PptxElementBounds>();
  for (const [elementIndex, element] of slide.elements.entries()) {
    const elementId = element.id;
    if (!elementId || !requested.has(elementId)) continue;
    const origin = slide.elementSources?.[elementIndex]?.origin ?? 'unknown';
    const rank = origin === 'slide' ? 3 : origin === 'layout' ? 2 : origin === 'master' ? 1 : 0;
    const previous = found.get(elementId);
    const previousRank = previous?.origin === 'slide'
      ? 3
      : previous?.origin === 'layout'
        ? 2
        : previous?.origin === 'master'
          ? 1
          : 0;
    // Drawing ids are scoped to a slide part. Master/layout decorations may
    // reuse the same numeric id after their models are composed into the slide;
    // the comment's slide moniker therefore resolves the slide-authored element.
    if (previous && previousRank > rank) continue;
    found.set(elementId, Object.freeze({
      elementId,
      elementIndex,
      origin,
      elementType: element.type,
      bounds: Object.freeze({
        x: element.x,
        y: element.y,
        width: element.width,
        height: element.height,
        rotation: element.rotation,
        flipH: element.flipH,
        flipV: element.flipV,
      }),
    }));
  }
  return Object.freeze(elementIds.flatMap((id) => {
    const bounds = found.get(id);
    return bounds ? [bounds] : [];
  }));
}
