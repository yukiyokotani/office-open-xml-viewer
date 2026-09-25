import type { Glow, Reflection, Shadow, SoftEdge } from '../types/common.js';
import type { ChartExElementStyle } from '../types/chart.js';
import { MAX_CANVAS_AREA } from '../canvas/clamp.js';
import { hexToRgba } from '../shape/paint.js';
import {
  applyInnerShadow,
  applyOuterShadow,
  applyReflection,
  applySoftEdge,
  type EffectBBox,
  type PaintShape,
} from '../shape/effects.js';
import { compactStyleIndex } from './sparse-style-index.js';

const EMU_PER_PT = 12_700;

/**
 * One chart render may rasterize at most one maximum-size Canvas worth of
 * effect pixels. Effect surfaces are temporary and cropped per mark, but a
 * chart can contain 10,000 points; an aggregate budget prevents a validly
 * bounded point list from multiplying auxiliary-canvas work without bound.
 */
export const MAX_CHART_EFFECT_RASTER_PIXELS = MAX_CANVAS_AREA;

interface Matrix2D {
  a: number;
  b: number;
  c: number;
  d: number;
  e: number;
  f: number;
}

interface ChartEffectBudget {
  perConsumerMaximum: number;
}

const activeBudgets = new WeakMap<CanvasRenderingContext2D, ChartEffectBudget>();

/** Scope aggregate effect work to one synchronous chart render. */
export function withChartEffectBudget<T>(
  ctx: CanvasRenderingContext2D,
  paint: () => T,
  maximumPixels = MAX_CHART_EFFECT_RASTER_PIXELS,
  consumerUpperBound = 1,
): T {
  const previous = activeBudgets.get(ctx);
  const consumers = Number.isSafeInteger(consumerUpperBound) && consumerUpperBound > 0
    ? consumerUpperBound : 1;
  activeBudgets.set(ctx, {
    perConsumerMaximum: Math.floor(Math.max(0, maximumPixels) / consumers),
  });
  try {
    return paint();
  } finally {
    if (previous) activeBudgets.set(ctx, previous);
    else activeBudgets.delete(ctx);
  }
}

export interface ChartStyleEffectRecipe {
  shadow?: Shadow;
  innerShadow?: Shadow;
  glow?: Glow;
  softEdge?: SoftEdge;
  reflection?: Reflection;
}

function styleOwnsEffect(style: ChartExElementStyle | null | undefined): boolean {
  return style?.effectAuthored === true
    || style?.effectUnsupported === true
    || style?.effectNoStyle === true
    || style?.shadows != null
    || style?.innerShadows != null
    || style?.glows != null
    || style?.softEdges != null
    || style?.reflections != null;
}

/** Pick the first style layer that actually owns the effect component. */
export function chartStyleEffectOwner(
  ...styles: Array<ChartExElementStyle | null | undefined>
): ChartExElementStyle | undefined {
  return styles.find(styleOwnsEffect) ?? undefined;
}

function paletteEntry<T>(values: Array<T | null> | null | undefined, index: number): T | undefined {
  if (!values?.length) return undefined;
  return values[index % values.length] ?? undefined;
}

/**
 * Resolve the whole DrawingML effect component. Direct shape properties own
 * the component whenever they authored one, including an empty or unsupported
 * list. That deliberately suppresses the less-specific linked/numeric style;
 * effect-list children do not inherit independently.
 */
export function chartStyleEffectRecipe(
  direct: ChartExElementStyle | null | undefined,
  fallback: ChartExElementStyle | null | undefined,
  index: number,
  fallbackIndex = index,
): ChartStyleEffectRecipe | undefined {
  const directOwns = styleOwnsEffect(direct);
  const style = directOwns ? direct : fallback;
  if (!style || style.effectNoStyle === true || style.effectUnsupported === true) return undefined;
  const sourceIndex = directOwns ? index : fallbackIndex;
  const compactIndex = style.effectFormattingIndices
    ? compactStyleIndex(style.effectFormattingIndices, sourceIndex)
    : -1;
  const paletteIndex = style.effectColorIndex ?? (compactIndex >= 0 ? compactIndex : sourceIndex);
  const recipe = {
    shadow: paletteEntry(style.shadows, paletteIndex),
    innerShadow: paletteEntry(style.innerShadows, paletteIndex),
    glow: paletteEntry(style.glows, paletteIndex),
    softEdge: paletteEntry(style.softEdges, paletteIndex),
    reflection: paletteEntry(style.reflections, paletteIndex),
  };
  return Object.values(recipe).some(Boolean) ? recipe : undefined;
}

function currentTransform(ctx: CanvasRenderingContext2D): Matrix2D {
  if (typeof ctx.getTransform === 'function') {
    const matrix = ctx.getTransform();
    if ([matrix.a, matrix.b, matrix.c, matrix.d, matrix.e, matrix.f].every(Number.isFinite)) {
      return matrix;
    }
  }
  return { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };
}

function transformPoint(matrix: Matrix2D, x: number, y: number): readonly [number, number] {
  return [
    matrix.a * x + matrix.c * y + matrix.e,
    matrix.b * x + matrix.d * y + matrix.f,
  ];
}

function deviceBBox(matrix: Matrix2D, bbox: EffectBBox): EffectBBox {
  const points = [
    transformPoint(matrix, bbox.x, bbox.y),
    transformPoint(matrix, bbox.x + bbox.w, bbox.y),
    transformPoint(matrix, bbox.x, bbox.y + bbox.h),
    transformPoint(matrix, bbox.x + bbox.w, bbox.y + bbox.h),
  ];
  const xs = points.map(point => point[0]);
  const ys = points.map(point => point[1]);
  const minX = Math.min(...xs);
  const maxX = Math.max(...xs);
  const minY = Math.min(...ys);
  const maxY = Math.max(...ys);
  return { x: minX, y: minY, w: maxX - minX, h: maxY - minY };
}

function rasterCost(
  bbox: EffectBBox,
  margin: number,
  surfaces: number,
): number {
  const width = Math.max(1, Math.ceil(bbox.w + margin * 2));
  const height = Math.max(1, Math.ceil(bbox.h + margin * 2));
  const cost = width * height * surfaces;
  return Number.isSafeInteger(cost) ? cost : Number.POSITIVE_INFINITY;
}

/** Conservative auxiliary-pixel work for one effect-bearing mark. */
export function chartEffectRasterWorkUpperBound(
  effects: ChartStyleEffectRecipe,
  bbox: EffectBBox,
  effectScale: number,
  deviceW: number,
  deviceH: number,
): number {
  const softReach = effects.softEdge
    ? Math.ceil(effects.softEdge.radius * 3 * effectScale) + 2 : 0;
  const innerReach = effects.innerShadow
    ? Math.ceil(
        (effects.innerShadow.blur * 3 + Math.abs(effects.innerShadow.dist)) * effectScale,
      ) + 2
    : 0;
  const outerReach = effects.shadow
    ? Math.ceil(
        (effects.shadow.blur * 3 + Math.abs(effects.shadow.dist)) * effectScale,
      ) + 2
    : 0;
  const reflectionCost = effects.reflection ? deviceW * deviceH : 0;
  return reflectionCost
    + (effects.shadow ? rasterCost(bbox, outerReach, 1) : 0)
    + (effects.softEdge ? rasterCost(bbox, softReach, 2) : 0)
    + (effects.innerShadow ? rasterCost(bbox, innerReach, 2) : 0);
}

function nativeShadow(
  ctx: CanvasRenderingContext2D,
  shadow: Shadow,
  ptToPx: number,
): void {
  const direction = shadow.dir * Math.PI / 180;
  const distance = shadow.dist / EMU_PER_PT * ptToPx;
  ctx.shadowColor = hexToRgba(shadow.color, shadow.alpha);
  ctx.shadowBlur = Math.max(0, shadow.blur / EMU_PER_PT * ptToPx);
  ctx.shadowOffsetX = Math.cos(direction) * distance;
  ctx.shadowOffsetY = Math.sin(direction) * distance;
}

function nativeGlow(
  ctx: CanvasRenderingContext2D,
  glow: Glow,
  ptToPx: number,
): void {
  ctx.shadowColor = hexToRgba(glow.color, glow.alpha);
  ctx.shadowBlur = Math.max(0, glow.radius / EMU_PER_PT * ptToPx);
  ctx.shadowOffsetX = 0;
  ctx.shadowOffsetY = 0;
}

/**
 * Paint the allocation-free outer part of a 3-D datum effect behind the
 * already depth-sorted scene. Repainting the datum with `destination-over`
 * preserves inter-datum occlusion while allowing Canvas to emit its shadow or
 * glow outside the final silhouette.
 *
 * Inner shadow, soft edge, and reflection deliberately do not run here. Their
 * shared primitives composite a complete 2-D bitmap and cannot preserve the
 * per-face depth ordering of an interleaved 3-D scene. The resolved style stays
 * in the model; this fail-closed paint boundary can be removed once effects are
 * composited per depth-sorted datum rather than per full Canvas.
 */
export function paintChartStyleOuterEffectBehind(
  ctx: CanvasRenderingContext2D,
  direct: ChartExElementStyle | null | undefined,
  fallback: ChartExElementStyle | null | undefined,
  index: number,
  ptToPx: number,
  paintBody: (target: CanvasRenderingContext2D) => void,
  fallbackIndex = index,
): void {
  const effects = chartStyleEffectRecipe(direct, fallback, index, fallbackIndex);
  const nativeEffect: Shadow | Glow | undefined = effects?.shadow ?? effects?.glow;
  if (!nativeEffect) return;
  ctx.save();
  ctx.globalCompositeOperation = 'destination-over';
  if ('radius' in nativeEffect) nativeGlow(ctx, nativeEffect, ptToPx);
  else nativeShadow(ctx, nativeEffect, ptToPx);
  paintBody(ctx);
  ctx.restore();
}

function canvasExtent(ctx: CanvasRenderingContext2D): readonly [number, number] {
  const canvas = ctx.canvas as { width?: number; height?: number } | undefined;
  return [canvas?.width ?? 0, canvas?.height ?? 0];
}

/**
 * Paint one chart mark with its resolved DrawingML effects. Outer shadow is
 * always represented through the allocation-free native Canvas shadow slot,
 * so a large styled series cannot lose its later shadows as a raster budget is
 * consumed. Inner shadow, soft edge, and reflection are admitted all-or-none
 * per mark against the chart-wide remaining budget. That makes uncommon raster
 * combinations are admitted by a render-wide per-consumer ceiling. The caller
 * supplies a conservative consumer-count upper bound, so aggregate temporary
 * work stays bounded without making later marks lose effects merely because
 * equivalent earlier marks painted first. Reflection uses the shared full-
 * canvas primitive only when it fits that same order-independent allowance.
 */
export function paintChartStyleEffects(
  ctx: CanvasRenderingContext2D,
  direct: ChartExElementStyle | null | undefined,
  fallback: ChartExElementStyle | null | undefined,
  index: number,
  bbox: EffectBBox,
  ptToPx: number,
  paintBody: (target: CanvasRenderingContext2D) => void,
  fallbackIndex = index,
): void {
  const effects = chartStyleEffectRecipe(direct, fallback, index, fallbackIndex);
  if (!effects) {
    paintBody(ctx);
    return;
  }

  const transform = currentTransform(ctx);
  const transformedBBox = deviceBBox(transform, bbox);
  const transformScale = Math.max(
    Math.hypot(transform.a, transform.b),
    Math.hypot(transform.c, transform.d),
  );
  const effectScale = Math.max(0, ptToPx * transformScale / EMU_PER_PT);
  const [deviceW, deviceH] = canvasExtent(ctx);
  const haveDeviceCanvas = deviceW > 0 && deviceH > 0;
  const devicePaint: PaintShape = target => {
    target.setTransform(transform as DOMMatrix2DInit);
    paintBody(target as CanvasRenderingContext2D);
  };
  const resetToDevice = () => ctx.setTransform(1, 0, 0, 1, 0, 0);

  const requiredRasterWork = chartEffectRasterWorkUpperBound(
    effects, transformedBBox, effectScale, deviceW, deviceH,
  );
  const budget = activeBudgets.get(ctx);
  const rasterAllowed = haveDeviceCanvas
    && Number.isSafeInteger(requiredRasterWork)
    && requiredRasterWork <= (budget?.perConsumerMaximum ?? 0);

  let shadowRasterized = false;
  if (rasterAllowed && effects.shadow) {
    ctx.save();
    resetToDevice();
    shadowRasterized = applyOuterShadow(
      ctx as never,
      devicePaint,
      transformedBBox,
      effects.shadow,
      effectScale,
      deviceW,
      deviceH,
      Math.atan2(transform.b, transform.a) * 180 / Math.PI,
    );
    ctx.restore();
  }

  // Prefer the composed-silhouette outer-shadow primitive above. If its
  // bounded auxiliary surface is unavailable, retain the normal Office theme
  // effect through Canvas's allocation-free shadow slot for every mark. Canvas
  // exposes only one such slot: when that fallback shadow and a glow coexist,
  // the shadow remains authoritative and the glow is the explicit limitation.
  const nativeEffect: Shadow | Glow | undefined = shadowRasterized
    ? effects.glow
    : effects.shadow ?? effects.glow;

  if (rasterAllowed && effects.reflection) {
    ctx.save();
    resetToDevice();
    applyReflection(
      ctx as never,
      devicePaint,
      transformedBBox,
      effects.reflection,
      effectScale,
      deviceW,
      deviceH,
    );
    ctx.restore();
  }

  let bodyPainted = false;
  // A native shadow/glow must be emitted by the live body paint. Canvas cannot
  // attach that native effect to the already-composited soft-edge bitmap, so
  // keep the visible outer effect and omit softEdge for this uncommon combined
  // recipe rather than silently losing the style's shadow.
  if (!nativeEffect && rasterAllowed && effects.softEdge) {
      ctx.save();
      resetToDevice();
      applySoftEdge(
        ctx as never,
        devicePaint,
        transformedBBox,
        effects.softEdge,
        effectScale,
        deviceW,
        deviceH,
      );
      ctx.restore();
      bodyPainted = true;
  }
  if (!bodyPainted) {
    ctx.save();
    if (nativeEffect) {
      if ('radius' in nativeEffect) nativeGlow(ctx, nativeEffect, ptToPx);
      else nativeShadow(ctx, nativeEffect, ptToPx);
    }
    paintBody(ctx);
    ctx.restore();
  }

  if (rasterAllowed && effects.innerShadow) {
      ctx.save();
      resetToDevice();
      applyInnerShadow(
        ctx as never,
        devicePaint,
        transformedBBox,
        effects.innerShadow,
        effectScale,
        deviceW,
        deviceH,
      );
      ctx.restore();
  }
}
