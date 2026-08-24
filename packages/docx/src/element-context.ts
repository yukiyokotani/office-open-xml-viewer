import {
  boundedChartContextText,
  MAX_CHART_CONTEXT_TEXT_CHARACTERS,
} from '@silurus/ooxml-core/internal/chart-context';
import { inverseMapAffinePoint, mapAffinePoint } from './layout/affine.js';
import { selectDocumentLayoutPage } from './layout/document-layout-variants.js';
import {
  elementGeometryForPage,
  type DrawingGeometry,
  type ElementGeometry,
} from './layout/text-index.js';
import { paintResourceRegistryOf } from './layout/runtime-state.js';
import type {
  DeepReadonly,
  DrawingLayout,
  LayoutRect,
  LayoutServices,
  PaintNode,
  PaintResourceRegistry,
  PointPt,
  SourceRef,
} from './layout/types.js';
import type {
  DocxElementContext,
  DocxPagePoint,
  DocxSelectionSourceLocator,
} from './selection-context.js';

export interface DocxElementContextOptions {
  readonly maxTextCharacters?: number;
  /** Select the same DATE/TIME layout variant used for rendering. */
  readonly currentDate?: Date | number;
}

export interface SelectedDocxElementContextOptions extends DocxElementContextOptions {
  readonly defaultCurrentDateMs: number;
}

export const MAX_DOCX_ELEMENT_TEXT_CHARACTERS = MAX_CHART_CONTEXT_TEXT_CHARACTERS;
const DEFAULT_DOCX_ELEMENT_TEXT_CHARACTERS = 16_384;

function safeUtf16Prefix(value: string, maxCodeUnits: number): string {
  let end = Math.min(value.length, maxCodeUnits);
  if (end > 0 && end < value.length) {
    const previous = value.charCodeAt(end - 1);
    const next = value.charCodeAt(end);
    if (previous >= 0xD800 && previous <= 0xDBFF && next >= 0xDC00 && next <= 0xDFFF) end--;
  }
  return value.slice(0, end);
}

function boundedText(parts: Iterable<string>, maxTextCharacters: number): {
  readonly text?: string;
  readonly truncated: boolean;
} {
  const chunks: string[] = [];
  let length = 0;
  let truncated = false;
  for (const value of parts) {
    if (value.length === 0) continue;
    if (length > 0) {
      if (length >= maxTextCharacters) { truncated = true; break; }
      chunks.push('\n');
      length++;
    }
    const part = safeUtf16Prefix(value, Math.max(0, maxTextCharacters - length));
    chunks.push(part);
    length += part.length;
    if (part.length < value.length) { truncated = true; break; }
  }
  const text = chunks.join('');
  return { ...(text.length === 0 ? {} : { text }), truncated };
}

function* nodeText(node: DeepReadonly<PaintNode>): Iterable<string> {
  if (node.kind === 'paragraph') {
    for (const line of node.lines) {
      const text = line.placements.flatMap((placement) =>
        placement.kind === 'text' ? [placement.text] : []).join('');
      if (text) yield text;
    }
    return;
  }
  if (node.kind === 'table') {
    for (const row of node.rows) {
      for (const cell of row.cells) {
        for (const block of cell.blocks) yield* nodeText(block.layout);
      }
    }
    return;
  }
  if (node.kind === 'note' || node.kind === 'textbox') {
    for (const block of node.story.blocks) yield* nodeText(block);
  }
}

function* shapeTextParts(geometry: DrawingGeometry): Iterable<string> {
  for (const command of geometry.drawing.commands) {
    if (command.kind === 'text' || command.kind === 'watermark-text') yield command.text;
  }
  for (const textBox of geometry.textBoxes) yield* nodeText(textBox);
}

function normalizedMax(value: number | undefined): number {
  const requested = value ?? DEFAULT_DOCX_ELEMENT_TEXT_CHARACTERS;
  if (!Number.isFinite(requested) || requested < 0) {
    throw new RangeError('maxTextCharacters must be a finite non-negative number.');
  }
  return Math.min(MAX_DOCX_ELEMENT_TEXT_CHARACTERS, Math.floor(requested));
}

function sourceLocator(source: SourceRef): DocxSelectionSourceLocator {
  return {
    story: source.story,
    storyInstance: source.storyInstance,
    path: [...source.path],
  };
}

function geometryBounds(geometry: ElementGeometry): LayoutRect {
  return 'drawing' in geometry ? geometry.drawing.inkBounds : geometry.placement.bounds;
}

function mappedBounds(geometry: ElementGeometry): DocxElementContext['bounds'] {
  const bounds = geometryBounds(geometry);
  const corners = [
    mapAffinePoint(geometry.pointToPage, bounds),
    mapAffinePoint(geometry.pointToPage, { xPt: bounds.xPt + bounds.widthPt, yPt: bounds.yPt }),
    mapAffinePoint(geometry.pointToPage, { xPt: bounds.xPt, yPt: bounds.yPt + bounds.heightPt }),
    mapAffinePoint(geometry.pointToPage, {
      xPt: bounds.xPt + bounds.widthPt,
      yPt: bounds.yPt + bounds.heightPt,
    }),
  ];
  const left = Math.min(...corners.map((corner) => corner.xPt));
  const top = Math.min(...corners.map((corner) => corner.yPt));
  const right = Math.max(...corners.map((corner) => corner.xPt));
  const bottom = Math.max(...corners.map((corner) => corner.yPt));
  return { xPt: left, yPt: top, widthPt: right - left, heightPt: bottom - top };
}

function contains(rect: LayoutRect, point: PointPt): boolean {
  return point.xPt >= rect.xPt && point.xPt <= rect.xPt + rect.widthPt &&
    point.yPt >= rect.yPt && point.yPt <= rect.yPt + rect.heightPt;
}

function pointInPolygon(point: PointPt, vertices: readonly PointPt[]): boolean {
  let inside = false;
  for (let i = 0, j = vertices.length - 1; i < vertices.length; j = i++) {
    const a = vertices[i];
    const b = vertices[j];
    if (((a.yPt > point.yPt) !== (b.yPt > point.yPt)) &&
      point.xPt < (b.xPt - a.xPt) * (point.yPt - a.yPt) / (b.yPt - a.yPt) + a.xPt) {
      inside = !inside;
    }
  }
  return inside;
}

function containsDrawing(drawing: DrawingLayout, local: PointPt): boolean {
  if (!contains(drawing.inkBounds, local)) return false;
  if (!drawing.clip) return true;
  return drawing.clip.kind === 'rect'
    ? contains(drawing.clip.rect, local)
    : pointInPolygon(local, drawing.clip.points);
}

function containsElement(
  geometry: ElementGeometry,
  pagePoint: PointPt,
  local: PointPt,
): boolean {
  for (const clip of geometry.clips) {
    const clipPoint = inverseMapAffinePoint(clip.pointToPage, pagePoint);
    if (!clipPoint || !contains(clip.bounds, clipPoint)) return false;
  }
  return 'drawing' in geometry
    ? containsDrawing(geometry.drawing, local)
    : contains(geometry.placement.bounds, local);
}

function resourceCommand(
  drawing: DrawingLayout,
  kind: 'chart' | 'image',
): Extract<DrawingLayout['commands'][number], { kind: 'resource' }> | undefined {
  return drawing.commands.find((command): command is Extract<
    DrawingLayout['commands'][number], { kind: 'resource' }
  > => command.kind === 'resource' && command.resourceKind === kind);
}

function drawingType(drawing: DrawingLayout): DocxElementContext['elementType'] | null {
  if (resourceCommand(drawing, 'chart')) return 'chart';
  if (drawing.commands.some((command) => command.kind === 'drawingml-shape' ||
    command.kind === 'drawingml-image-fill' || command.kind === 'fill-rect' ||
    command.kind === 'stroke-rect' || command.kind === 'text' ||
    command.kind === 'watermark-text')) return 'shape';
  if (resourceCommand(drawing, 'image')) return 'image';
  return null;
}

function contextForElement(
  geometry: ElementGeometry,
  pageIndex: number,
  elementIndex: number,
  point: DocxPagePoint,
  registry: PaintResourceRegistry,
  maxTextCharacters: number,
): DocxElementContext | null {
  const elementType = 'drawing' in geometry
    ? drawingType(geometry.drawing)
    : geometry.placement.resourceKind;
  if (!elementType) return null;
  let text: string | undefined;
  let truncated = false;
  let mimeType: string | undefined;
  let seriesCount: number | undefined;
  if (elementType === 'chart') {
    const resourceKey = 'drawing' in geometry
      ? resourceCommand(geometry.drawing, 'chart')!.resourceKey
      : geometry.placement.resourceKey;
    const descriptor = registry.resolve(resourceKey, 'chart');
    const bounded = boundedChartContextText(descriptor.model, maxTextCharacters);
    text = bounded.text;
    truncated = bounded.truncated;
    seriesCount = descriptor.model.series.length;
  } else if (elementType === 'image') {
    const resourceKey = 'drawing' in geometry
      ? resourceCommand(geometry.drawing, 'image')!.resourceKey
      : geometry.placement.resourceKey;
    const descriptor = registry.descriptors.find((candidate) =>
      candidate.resourceKey === resourceKey && candidate.kind === 'image' &&
      'mimeType' in candidate);
    if (!descriptor) throw new Error(`Unknown image paint resource: ${resourceKey}`);
    mimeType = (descriptor as { readonly mimeType: string }).mimeType;
  } else {
    const bounded = boundedText(shapeTextParts(geometry as DrawingGeometry), maxTextCharacters);
    text = bounded.text;
    truncated = bounded.truncated;
  }
  return {
    format: 'docx',
    kind: 'element',
    pageIndex,
    elementIndex,
    elementType,
    point: { ...point },
    bounds: mappedBounds(geometry),
    source: sourceLocator('drawing' in geometry ? geometry.drawing.source : geometry.source),
    ...(text === undefined ? {} : { text }),
    ...(mimeType === undefined ? {} : { mimeType }),
    ...(seriesCount === undefined ? {} : { seriesCount }),
    truncated,
    truncationReasons: truncated ? ['text'] : [],
    textCharacters: text?.length ?? 0,
    maxTextCharacters,
  };
}

export function hitTestDocxElementContext(
  layout: Parameters<typeof elementGeometryForPage>[0],
  pageIndex: number,
  point: DocxPagePoint,
  registry: PaintResourceRegistry,
  options: DocxElementContextOptions = {},
): DocxElementContext | null {
  if (!Number.isFinite(point.xPt) || !Number.isFinite(point.yPt)) {
    throw new RangeError('DOCX hit-test point must contain finite page coordinates.');
  }
  const maxTextCharacters = normalizedMax(options.maxTextCharacters);
  const elements = elementGeometryForPage(layout, pageIndex);
  for (let index = elements.length - 1; index >= 0; index--) {
    const geometry = elements[index];
    const local = inverseMapAffinePoint(geometry.pointToPage, point);
    if (!local || !containsElement(geometry, point, local)) continue;
    const context = contextForElement(
      geometry, pageIndex, index, point, registry, maxTextCharacters,
    );
    if (context) return context;
  }
  return null;
}

/** Select the exact layout variant used by paint, then hit-test its drawings. */
export function hitTestSelectedDocxElementContext(
  services: LayoutServices,
  pageIndex: number,
  point: DocxPagePoint,
  options: SelectedDocxElementContextOptions,
): DocxElementContext | null {
  const selected = selectDocumentLayoutPage(services, {
    currentDate: options.currentDate,
    defaultCurrentDateMs: options.defaultCurrentDateMs,
  }, pageIndex);
  return hitTestDocxElementContext(
    selected.layout,
    pageIndex,
    point,
    paintResourceRegistryOf(services),
    options,
  );
}

export function limitDocxElementContext(
  context: DocxElementContext,
  requestedMaxTextCharacters: number | undefined,
): DocxElementContext {
  const maxTextCharacters = normalizedMax(requestedMaxTextCharacters);
  const text = context.text === undefined
    ? undefined
    : safeUtf16Prefix(context.text, maxTextCharacters);
  const truncated = context.truncated ||
    (context.text !== undefined && text!.length < context.text.length);
  return {
    ...structuredClone(context),
    ...(text === undefined ? {} : { text }),
    truncated,
    truncationReasons: truncated ? ['text'] : [],
    textCharacters: text?.length ?? 0,
    maxTextCharacters,
  };
}
