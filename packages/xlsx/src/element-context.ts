import { EMU_PER_PX } from '@silurus/ooxml-core';
import {
  boundedChartContextText,
  MAX_CHART_CONTEXT_TEXT_CHARACTERS,
} from '@silurus/ooxml-core/internal/chart-context';
import type {
  ChartAnchor,
  ImageAnchor,
  ShapeAnchor,
  ShapeInfo,
  ViewportRange,
  Worksheet,
} from './types.js';
import {
  HEADER_H,
  HEADER_W,
  getGridGeometryForWorksheet,
  sheetAnchoredRectX,
} from './renderer.js';
import { usesNativeOneCellExtent } from './internal/cell-anchor-geometry.js';
import { inverseImageTransformPoint, rotatedImageBounds } from './internal/image-anchor-transform.js';
import type { GridAxisGeometry } from './internal/grid-axis-geometry.js';
import type { XlsxElementContext } from './selection.js';

interface CellAnchorLike {
  readonly fromCol: number;
  readonly fromColOff: number;
  readonly fromRow: number;
  readonly fromRowOff: number;
  readonly toCol: number;
  readonly toColOff: number;
  readonly toRow: number;
  readonly toRowOff: number;
  readonly editAs?: string;
  readonly nativeExtCx?: number;
  readonly nativeExtCy?: number;
}

export interface XlsxElementHitViewport {
  readonly width: number;
  readonly height: number;
  readonly cellScale: number;
  readonly viewport: ViewportRange;
  /** Render-local offsets inside the first visible row/column, in logical px. */
  readonly scrollOffsetX: number;
  readonly scrollOffsetY: number;
  readonly freezeRows: number;
  readonly freezeCols: number;
}

interface AnchoredRectContext {
  readonly colAxis: GridAxisGeometry;
  readonly rowAxis: GridAxisGeometry;
  readonly scale: number;
  readonly startRow: number;
  readonly startCol: number;
  readonly scrollOffsetX: number;
  readonly scrollOffsetY: number;
  readonly scrollAreaX: number;
  readonly scrollAreaY: number;
  readonly scrollAreaW: number;
  readonly scrollAreaH: number;
  readonly rtl: boolean;
  readonly canvasWidth: number;
}

interface CanvasRect {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

export interface XlsxElementOutlineProjection {
  readonly rect: CanvasRect;
  readonly clip: CanvasRect;
  readonly rotation: number;
}

function anchoredCanvasRect(anchor: CellAnchorLike, context: AnchoredRectContext): CanvasRect | null {
  const x1 = context.colAxis.offsetOf(anchor.fromCol + 1) +
    (anchor.fromColOff * context.scale) / EMU_PER_PX;
  const y1 = context.rowAxis.offsetOf(anchor.fromRow + 1) +
    (anchor.fromRowOff * context.scale) / EMU_PER_PX;
  let width: number;
  let height: number;
  if (usesNativeOneCellExtent(anchor)) {
    width = (anchor.nativeExtCx! * context.scale) / EMU_PER_PX;
    height = (anchor.nativeExtCy! * context.scale) / EMU_PER_PX;
  } else {
    const x2 = context.colAxis.offsetOf(anchor.toCol + 1) +
      (anchor.toColOff * context.scale) / EMU_PER_PX;
    const y2 = context.rowAxis.offsetOf(anchor.toRow + 1) +
      (anchor.toRowOff * context.scale) / EMU_PER_PX;
    width = x2 - x1;
    height = y2 - y1;
  }
  if (width <= 0 || height <= 0) return null;
  const scrollOriginX = context.colAxis.offsetOf(context.startCol);
  const scrollOriginY = context.rowAxis.offsetOf(context.startRow);
  const logicalX = context.scrollAreaX + (x1 - scrollOriginX) - context.scrollOffsetX;
  return {
    x: sheetAnchoredRectX(logicalX, width, context.canvasWidth, context.rtl),
    y: context.scrollAreaY + (y1 - scrollOriginY) - context.scrollOffsetY,
    width,
    height,
  };
}

function pointInRect(point: Readonly<{ x: number; y: number }>, rect: CanvasRect): boolean {
  return point.x >= rect.x && point.x <= rect.x + rect.width &&
    point.y >= rect.y && point.y <= rect.y + rect.height;
}

function intersectsClip(rect: CanvasRect, context: AnchoredRectContext): boolean {
  const clipX = sheetAnchoredRectX(
    context.scrollAreaX,
    context.scrollAreaW,
    context.canvasWidth,
    context.rtl,
  );
  return rect.x + rect.width >= clipX && rect.x <= clipX + context.scrollAreaW &&
    rect.y + rect.height >= context.scrollAreaY && rect.y <= context.scrollAreaY + context.scrollAreaH;
}

function pointInClip(point: Readonly<{ x: number; y: number }>, context: AnchoredRectContext): boolean {
  const clipX = sheetAnchoredRectX(
    context.scrollAreaX,
    context.scrollAreaW,
    context.canvasWidth,
    context.rtl,
  );
  return point.x >= clipX && point.x <= clipX + context.scrollAreaW &&
    point.y >= context.scrollAreaY && point.y <= context.scrollAreaY + context.scrollAreaH;
}

function anchoredContextForViewport(
  worksheet: Worksheet,
  viewport: XlsxElementHitViewport,
): AnchoredRectContext {
  const geometry = getGridGeometryForWorksheet(worksheet);
  const scale = viewport.cellScale;
  const { col: colAxis, row: rowAxis } = geometry.axesAtScale(scale);
  const headerW = Math.round(HEADER_W * scale);
  const headerH = Math.round(HEADER_H * scale);
  const frozenColBands = colAxis.bandsToCover(
    1, viewport.freezeCols, Math.max(0, viewport.width - headerW),
  );
  const frozenRowBands = rowAxis.bandsToCover(
    1, viewport.freezeRows, Math.max(0, viewport.height - headerH),
  );
  const frozenW = frozenColBands.reduce((sum, band) => sum + band.size, 0);
  const frozenH = frozenRowBands.reduce((sum, band) => sum + band.size, 0);
  return {
    colAxis,
    rowAxis,
    scale,
    startRow: viewport.viewport.row,
    startCol: viewport.viewport.col,
    scrollOffsetX: viewport.scrollOffsetX * scale,
    scrollOffsetY: viewport.scrollOffsetY * scale,
    scrollAreaX: headerW + frozenW,
    scrollAreaY: headerH + frozenH,
    scrollAreaW: Math.max(0, viewport.width - headerW - frozenW),
    scrollAreaH: Math.max(0, viewport.height - headerH - frozenH),
    rtl: worksheet.rightToLeft === true,
    canvasWidth: viewport.width,
  };
}

/** Project retained object focus into the same clipped canvas geometry used by hit testing. */
export function projectXlsxElementContext(
  worksheet: Worksheet,
  context: XlsxElementContext,
  viewport: XlsxElementHitViewport,
): XlsxElementOutlineProjection | null {
  const geometry = anchoredContextForViewport(worksheet, viewport);
  let anchor: CellAnchorLike | undefined;
  let rotation = 0;
  if (context.elementType === 'chart') anchor = worksheet.charts[context.elementIndex];
  else if (context.elementType === 'image') anchor = worksheet.images[context.elementIndex];
  else {
    anchor = (worksheet.shapeGroups ?? [])[context.elementIndex];
  }
  if (!anchor) return null;
  let rect = anchoredCanvasRect(anchor, geometry);
  if (!rect || !intersectsClip(rect, geometry)) return null;
  if (context.elementType === 'shape') {
    const group = (worksheet.shapeGroups ?? [])[context.elementIndex];
    const shape = context.shapeIndex === undefined ? undefined : group?.shapes[context.shapeIndex];
    if (shape) {
      rect = {
        x: rect.x + shape.x * rect.width,
        y: rect.y + shape.y * rect.height,
        width: shape.w * rect.width,
        height: shape.h * rect.height,
      };
      rotation = shape.rot;
    }
  }
  const clipX = sheetAnchoredRectX(
    geometry.scrollAreaX,
    geometry.scrollAreaW,
    geometry.canvasWidth,
    geometry.rtl,
  );
  return {
    rect,
    clip: {
      x: clipX,
      y: geometry.scrollAreaY,
      width: geometry.scrollAreaW,
      height: geometry.scrollAreaH,
    },
    rotation,
  };
}

function anchorLocator(anchor: CellAnchorLike): XlsxElementContext['anchor'] {
  return {
    from: {
      row: anchor.fromRow + 1,
      col: anchor.fromCol + 1,
      offsetX: anchor.fromColOff,
      offsetY: anchor.fromRowOff,
    },
    to: {
      row: anchor.toRow + 1,
      col: anchor.toCol + 1,
      offsetX: anchor.toColOff,
      offsetY: anchor.toRowOff,
    },
  };
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

function shapeText(shape: ShapeInfo, maxTextCharacters: number): {
  text?: string;
  truncated: boolean;
  textCharacters: number;
} {
  if (!shape.text) return { truncated: false, textCharacters: 0 };
  const chunks: string[] = [];
  let length = 0;
  let truncated = false;
  for (const [paragraphIndex, paragraph] of shape.text.paragraphs.entries()) {
    if (paragraphIndex > 0) {
      if (length >= maxTextCharacters) { truncated = true; break; }
      chunks.push('\n');
      length++;
    }
    for (const run of paragraph.runs) {
      const value = run.type === 'text' ? run.text : run.type === 'break' ? '\n' : '[equation]';
      const part = safeUtf16Prefix(value, Math.max(0, maxTextCharacters - length));
      chunks.push(part);
      length += part.length;
      if (part.length < value.length) { truncated = true; break; }
    }
    if (truncated) break;
  }
  return { text: chunks.join(''), truncated, textCharacters: length };
}

function baseContext(
  worksheet: Worksheet,
  sheetIndex: number,
  elementType: XlsxElementContext['elementType'],
  elementIndex: number,
  anchor: CellAnchorLike,
  text: string | undefined,
  truncated: boolean,
  maxTextCharacters: number,
): XlsxElementContext {
  return {
    format: 'xlsx',
    kind: 'element',
    sheetIndex,
    sheetName: worksheet.name,
    elementType,
    elementIndex,
    anchor: anchorLocator(anchor),
    ...(text === undefined ? {} : { text }),
    truncated,
    truncationReasons: truncated ? ['text'] : [],
    textCharacters: text?.length ?? 0,
    maxTextCharacters,
  };
}

function hitChart(
  worksheet: Worksheet,
  sheetIndex: number,
  point: Readonly<{ x: number; y: number }>,
  context: AnchoredRectContext,
  maxTextCharacters: number,
): XlsxElementContext | null {
  for (let index = worksheet.charts.length - 1; index >= 0; index--) {
    const anchor: ChartAnchor = worksheet.charts[index];
    const rect = anchoredCanvasRect(anchor, context);
    if (!rect || !intersectsClip(rect, context) || !pointInRect(point, rect)) continue;
    const bounded = boundedChartContextText(anchor.chart, maxTextCharacters);
    return {
      ...baseContext(
        worksheet, sheetIndex, 'chart', index, anchor,
        bounded.text, bounded.truncated, bounded.maxTextCharacters,
      ),
      seriesCount: anchor.chart.series.length,
    };
  }
  return null;
}

function inverseShapePoint(
  point: Readonly<{ x: number; y: number }>,
  rect: CanvasRect,
  shape: ShapeInfo,
): Readonly<{ x: number; y: number }> {
  const shapeRect = {
    x: rect.x + shape.x * rect.width,
    y: rect.y + shape.y * rect.height,
    width: shape.w * rect.width,
    height: shape.h * rect.height,
  };
  const cx = shapeRect.x + shapeRect.width / 2;
  const cy = shapeRect.y + shapeRect.height / 2;
  const radians = (-shape.rot * Math.PI) / 180;
  const dx = point.x - cx;
  const dy = point.y - cy;
  let x = cx + Math.cos(radians) * dx - Math.sin(radians) * dy;
  let y = cy + Math.sin(radians) * dx + Math.cos(radians) * dy;
  if (shape.flipH) x = 2 * cx - x;
  if (shape.flipV) y = 2 * cy - y;
  return { x, y };
}

function hitShape(
  worksheet: Worksheet,
  sheetIndex: number,
  point: Readonly<{ x: number; y: number }>,
  context: AnchoredRectContext,
  maxTextCharacters: number,
): XlsxElementContext | null {
  const groups = worksheet.shapeGroups ?? [];
  for (let groupIndex = groups.length - 1; groupIndex >= 0; groupIndex--) {
    const anchor: ShapeAnchor = groups[groupIndex];
    const rect = anchoredCanvasRect(anchor, context);
    if (!rect || !intersectsClip(rect, context)) continue;
    for (let shapeIndex = anchor.shapes.length - 1; shapeIndex >= 0; shapeIndex--) {
      const shape = anchor.shapes[shapeIndex];
      const local = inverseShapePoint(point, rect, shape);
      const shapeRect = {
        x: rect.x + shape.x * rect.width,
        y: rect.y + shape.y * rect.height,
        width: shape.w * rect.width,
        height: shape.h * rect.height,
      };
      if (!pointInRect(local, shapeRect)) continue;
      const bounded = shapeText(shape, maxTextCharacters);
      return {
        ...baseContext(
          worksheet, sheetIndex, 'shape', groupIndex, anchor,
          bounded.text, bounded.truncated, maxTextCharacters,
        ),
        shapeIndex,
        shapeCount: anchor.shapes.length,
        ...(shape.geom.type === 'image' ? { mimeType: shape.geom.mimeType } : {}),
      };
    }
  }
  return null;
}

function hitImage(
  worksheet: Worksheet,
  sheetIndex: number,
  point: Readonly<{ x: number; y: number }>,
  context: AnchoredRectContext,
  maxTextCharacters: number,
): XlsxElementContext | null {
  for (let index = worksheet.images.length - 1; index >= 0; index--) {
    const anchor: ImageAnchor = worksheet.images[index];
    const rect = anchoredCanvasRect(anchor, context);
    if (!rect) continue;
    const bounds = rotatedImageBounds(rect, anchor.rotation);
    if (!intersectsClip(bounds, context)) continue;
    const local = inverseImageTransformPoint(
      point, rect, anchor.rotation, anchor.flipH, anchor.flipV,
    );
    if (!pointInRect(local, rect)) continue;
    return {
      ...baseContext(worksheet, sheetIndex, 'image', index, anchor, undefined, false, maxTextCharacters),
      mimeType: anchor.svgImagePath ? 'image/svg+xml' : anchor.mimeType,
    };
  }
  return null;
}

/** Hit-test only on demand (normally pointerup): no per-frame spatial index. */
export function hitTestXlsxElementContext(
  worksheet: Worksheet,
  sheetIndex: number,
  point: Readonly<{ x: number; y: number }>,
  viewport: XlsxElementHitViewport,
  requestedMaxTextCharacters = MAX_CHART_CONTEXT_TEXT_CHARACTERS,
): XlsxElementContext | null {
  if (!Number.isFinite(point.x) || !Number.isFinite(point.y)) {
    throw new RangeError('XLSX hit-test point must contain finite coordinates.');
  }
  if (!Number.isFinite(requestedMaxTextCharacters) || requestedMaxTextCharacters < 0) {
    throw new RangeError('maxTextCharacters must be a finite non-negative number.');
  }
  const maxTextCharacters = Math.min(
    MAX_CHART_CONTEXT_TEXT_CHARACTERS,
    Math.floor(requestedMaxTextCharacters),
  );
  const context = anchoredContextForViewport(worksheet, viewport);
  if (!pointInClip(point, context)) return null;
  // Matches renderer paint order: images, shape groups, then charts.
  return hitChart(worksheet, sheetIndex, point, context, maxTextCharacters)
    ?? hitShape(worksheet, sheetIndex, point, context, maxTextCharacters)
    ?? hitImage(worksheet, sheetIndex, point, context, maxTextCharacters);
}

export function limitXlsxElementContext(
  context: XlsxElementContext,
  requestedMaxTextCharacters: number | undefined,
): XlsxElementContext {
  const requested = requestedMaxTextCharacters ?? context.maxTextCharacters;
  if (!Number.isFinite(requested) || requested < 0) {
    throw new RangeError('maxTextCharacters must be a finite non-negative number.');
  }
  const maxTextCharacters = Math.min(
    MAX_CHART_CONTEXT_TEXT_CHARACTERS,
    Math.floor(requested),
  );
  const text = context.text === undefined
    ? undefined
    : safeUtf16Prefix(context.text, maxTextCharacters);
  const truncated = context.truncated || (context.text !== undefined && text!.length < context.text.length);
  return {
    ...structuredClone(context),
    ...(text === undefined ? {} : { text }),
    truncated,
    truncationReasons: truncated ? ['text'] : [],
    textCharacters: text?.length ?? 0,
    maxTextCharacters,
  };
}
