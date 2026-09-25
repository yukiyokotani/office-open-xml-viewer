import type { ArrowEnd, DrawingMLShapePaintPlan, Stroke } from '@silurus/ooxml-core';
import type { ShapeRun } from '../types.js';
import type {
  DrawingPaintCommand,
  DeepReadonly,
  LayoutDiagnostic,
  LayoutRect,
  VmlTextPathAcquisitionInput,
} from './types.js';
import { snapshotPlainData } from './plain-data.js';
import type { TextLayoutService } from './text.js';

type ShapeDrawingCommand = Extract<
  DrawingPaintCommand,
  { kind: 'noop' | 'drawingml-shape' | 'drawingml-image-fill' | 'watermark-text' }
>;

export type ShapeDrawingPlanResult =
  | Readonly<{
      status: 'planned';
      command: ShapeDrawingCommand;
    }>
  | Readonly<{
      status: 'unsupported';
      command: Extract<DrawingPaintCommand, { kind: 'noop' }>;
      diagnostics: readonly Readonly<Omit<LayoutDiagnostic, 'source'>>[];
    }>;

// When fitshape is active, every shaped dimension is multiplied by the same
// target/source ratio, so a 1pt internal reference cancels out exactly. This is
// used only when VML omitted font-size; an authored size is never replaced.
const SCALE_NEUTRAL_REFERENCE_FONT_SIZE_PT = 1;

function unsupportedTextPathPlan(
  message: string,
): Extract<ShapeDrawingPlanResult, { status: 'unsupported' }> {
  return Object.freeze({
    status: 'unsupported',
    command: Object.freeze({ kind: 'noop' }),
    diagnostics: Object.freeze([Object.freeze({
      code: 'UNSUPPORTED_FEATURE',
      severity: 'error',
      message,
    })]),
  });
}

function arrowEnd(end: ShapeRun['headEnd']): ArrowEnd | undefined {
  return end ? { type: end.type, w: end.w, len: end.len } : undefined;
}

function strokeFill(fill: DeepReadonly<ShapeRun['strokeFill']>): Stroke['fill'] | undefined {
  if (!fill) return undefined;
  if (fill.fillType === 'gradient') {
    return {
      fillType: 'gradient',
      stops: fill.stops.map((stop) => ({ position: stop.position, color: stop.color })),
      angle: fill.angle,
      gradType: fill.gradType,
      ...(fill.scaled === undefined ? {} : { scaled: fill.scaled }),
      ...(fill.path === undefined ? {} : { path: fill.path }),
      ...(fill.fillToRect === undefined ? {} : { fillToRect: { ...fill.fillToRect } }),
      ...(fill.tileRect === undefined ? {} : { tileRect: { ...fill.tileRect } }),
      ...(fill.flip === undefined ? {} : { flip: fill.flip }),
      ...(fill.rotWithShape === undefined ? {} : { rotWithShape: fill.rotWithShape }),
    };
  }
  return { fillType: 'pattern', fg: fill.fg, bg: fill.bg, preset: fill.preset };
}

function shapeStroke(shape: DeepReadonly<ShapeRun>): Stroke | null {
  if (!shape.stroke || !shape.strokeWidth || shape.strokeWidth <= 0) return null;
  const fill = strokeFill(shape.strokeFill);
  return {
    color: shape.stroke,
    width: shape.strokeWidth,
    ...(fill ? { fill } : {}),
    ...(shape.strokeDash ? { dashStyle: shape.strokeDash } : {}),
    ...(shape.strokeCustomDash?.length ? { customDash: shape.strokeCustomDash } : {}),
    ...(shape.strokeCap ? { lineCap: shape.strokeCap } : {}),
    ...(shape.strokeJoin ? { lineJoin: shape.strokeJoin } : {}),
    ...(shape.strokeMiterLimit !== undefined && shape.strokeMiterLimit !== null
      ? { miterLimit: shape.strokeMiterLimit }
      : {}),
    ...(shape.strokeAlignment ? { alignment: shape.strokeAlignment } : {}),
    ...(shape.strokeCompound ? { cmpd: shape.strokeCompound } : {}),
    ...(arrowEnd(shape.headEnd) ? { headEnd: arrowEnd(shape.headEnd) } : {}),
    ...(arrowEnd(shape.tailEnd) ? { tailEnd: arrowEnd(shape.tailEnd) } : {}),
  };
}

export function planShapeDrawing(
  shape: DeepReadonly<ShapeRun>,
  bounds: LayoutRect,
  text?: TextLayoutService,
  textPath?: Readonly<VmlTextPathAcquisitionInput>,
  imageFillResourceKey?: string,
): ShapeDrawingPlanResult {
  const parserControlled = textPath !== undefined && (
    textPath.textPathOk !== undefined
    || textPath.on !== undefined
    || textPath.fitShape !== undefined
    || textPath.fitPath !== undefined
    || textPath.trim !== undefined
    || textPath.xScale !== undefined
  );
  // The stable public ShapeRun predates the private VML control projection.
  // Absence therefore keeps its historical visible+fitshape interpretation;
  // parser-created values always carry explicit false defaults and follow VML.
  const textPathEnabled = textPath !== undefined && (
    parserControlled
      ? textPath.textPathOk === true && textPath.on === true
      : true
  );
  if (textPathEnabled) {
    if (shape.fill?.fillType === 'image') {
      return unsupportedTextPathPlan('VML textPath with a DrawingML image fill is not rendered');
    }
    if (textPath.fitPath === true) {
      return unsupportedTextPathPlan('VML textPath fitPath=true is not rendered');
    }
    if (textPath.xScale === true) {
      return unsupportedTextPathPlan('VML textPath xScale=true is not rendered');
    }
    if (textPath.string.trim().length === 0) {
      return Object.freeze({ status: 'planned', command: Object.freeze({ kind: 'noop' }) });
    }
    if (!text) throw new Error('Shape textPath acquisition requires TextLayoutService');
    const fitShape = parserControlled ? textPath.fitShape === true : true;
    if (textPath.fontSizePt !== undefined
      && (!Number.isFinite(textPath.fontSizePt) || textPath.fontSizePt < 0)) {
      throw new RangeError('VML textPath fontSizePt must be finite and non-negative');
    }
    if (!fitShape && textPath.fontSizePt === undefined) {
      return unsupportedTextPathPlan(
        'VML textPath fitShape=false requires an authored font-size',
      );
    }
    if (textPath.fontSizePt === 0) {
      return Object.freeze({ status: 'planned', command: Object.freeze({ kind: 'noop' }) });
    }
    const fontSizePt = textPath.fontSizePt ?? SCALE_NEUTRAL_REFERENCE_FONT_SIZE_PT;
    const family = textPath.fontFamily ?? undefined;
    const shaped = text.shape({
      text: textPath.string,
      fontSizePt,
      fonts: {
        ascii: family,
        highAnsi: family,
        eastAsia: family,
        complexScript: family,
      },
      weight: textPath.bold ? 700 : 400,
      style: textPath.italic ? 'italic' : 'normal',
      measure: true,
    });
    if (textPath.trim === true && !shaped.inkBounds) {
      return unsupportedTextPathPlan(
        'VML textPath trim=true requires glyph ink bounds',
      );
    }
    // The retained source rectangle must choose one coherent metric domain:
    // trim removes the font's reserved box and uses tight actual ink, while
    // non-trim keeps the typographic advance/ascent/descent rectangle.
    const xMinPt = textPath.trim === true ? shaped.inkBounds?.xMinPt ?? 0 : 0;
    const xMaxPt = textPath.trim === true
      ? shaped.inkBounds?.xMaxPt ?? 0
      : shaped.advancePt;
    const ascentPt = textPath.trim === true ? shaped.inkBounds?.ascentPt ?? 0 : shaped.ascentPt;
    const descentPt = textPath.trim === true ? shaped.inkBounds?.descentPt ?? 0 : shaped.descentPt;
    const sourceBounds = {
      xPt: xMinPt,
      yPt: -ascentPt,
      widthPt: xMaxPt - xMinPt,
      heightPt: ascentPt + descentPt,
    };
    if (
      !Number.isFinite(shaped.advancePt)
      || Object.values(sourceBounds).some((value) => !Number.isFinite(value))
      || shaped.spans.some((span) => !Number.isFinite(span.advancePt))
    ) {
      throw new Error('Shape textPath acquisition produced non-finite metrics');
    }
    if (
      sourceBounds.widthPt <= 0
      || sourceBounds.heightPt <= 0
      || shaped.spans.length === 0
    ) {
      return unsupportedTextPathPlan('VML textPath produced empty glyph metrics');
    }
    return Object.freeze({
      status: 'planned',
      command: snapshotPlainData({
        kind: 'watermark-text' as const,
        rect: { ...bounds },
        text: textPath.string,
        fill: shape.fill ? { ...shape.fill, ...(shape.fill.fillType === 'gradient'
          ? { stops: shape.fill.stops.map((stop) => ({ ...stop })) }
          : {}) } : null,
        opacity: Math.max(0, Math.min(1, shape.fillOpacity ?? 1)),
        rotationDeg: shape.rotation ?? 0,
        fitShape,
        fontSizePt,
        sourceBounds,
        spans: shaped.spans.map((span) => ({
          text: span.text,
          advancePt: span.advancePt,
          fontRoute: span.fontRoute,
          fontWeight: span.font.weight,
          fontStyle: span.font.style,
        })),
      }, 'VML textPath command'),
    });
  }
  const plan: DrawingMLShapePaintPlan = {
    rect: { x: bounds.xPt, y: bounds.yPt, w: bounds.widthPt, h: bounds.heightPt },
    geometry: shape.presetGeometry
      ? {
          kind: 'preset',
          name: shape.presetGeometry,
          adjustments: [...(shape.adjValues ?? [])],
        }
      : {
          kind: 'custom',
          subpaths: shape.subpaths.map((subpath) => subpath.map((command) => ({ ...command }))),
          // ECMA-376 §20.1.9.15 per-path fill/stroke, kept only when it
          // lines up with the subpaths.
          ...(shape.subpathPaint?.length && shape.subpathPaint.length === shape.subpaths.length
            ? { paint: shape.subpathPaint.map((paint) => ({ ...paint })) }
            : {}),
        },
    fill: shape.fill && shape.fill.fillType !== 'image' ? { ...shape.fill, ...(shape.fill.fillType === 'gradient'
      ? { stops: shape.fill.stops.map((stop) => ({ ...stop })) }
      : {}) } : null,
    stroke: shapeStroke(shape),
    transform: {
      rotationDeg: shape.rotation ?? 0,
      flipH: shape.flipH ?? false,
      flipV: shape.flipV ?? false,
    },
  };
  if (shape.fill?.fillType === 'image') {
    if (shape.fill.tile !== undefined) {
      return Object.freeze({
        status: 'unsupported',
        command: Object.freeze({ kind: 'noop' }),
        diagnostics: Object.freeze([Object.freeze({
          code: 'UNSUPPORTED_FEATURE',
          severity: 'error',
          message: 'Tiled DrawingML shape image fills are not rendered',
        })]),
      });
    }
    if (!imageFillResourceKey) {
      throw new Error('DrawingML shape image fill requires a retained image resource key');
    }
    return Object.freeze({
      status: 'planned',
      command: snapshotPlainData({
        kind: 'drawingml-image-fill' as const,
        plan,
        resourceKey: imageFillResourceKey,
        ...(shape.fill.fillRect === undefined ? {} : {
          fillRect: {
            l: shape.fill.fillRect.l ?? 0,
            t: shape.fill.fillRect.t ?? 0,
            r: shape.fill.fillRect.r ?? 0,
            b: shape.fill.fillRect.b ?? 0,
          },
        }),
      }, 'DrawingML shape image-fill command'),
    });
  }
  return Object.freeze({
    status: 'planned',
    command: snapshotPlainData(
      { kind: 'drawingml-shape', plan } as const,
      'DrawingML shape command',
    ),
  });
}
