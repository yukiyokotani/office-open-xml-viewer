import {
  HARD_MAX_DECODED_IMAGE_BYTES,
  MAX_DECODED_IMAGE_BYTES,
  MAX_RASTER_PIXELS,
  OoxmlDecodedImageLimitError,
} from './pixel-budget.js';
import { decodedBitmapRetainedTarget } from './raster-target.js';

export type DecodedImageBudgetStrategy = 'adaptive' | 'strict';

/** Format-neutral raster-memory policy used by DOCX, PPTX and XLSX paints. */
export interface ImageResourceOptions {
  /**
   * Soft aggregate budget for decoded RGBA surfaces owned by one document
   * paint. Adaptive paints lower raster resolution to remain within it; strict
   * paints preserve requested targets and surface the typed quota error.
   * Defaults to 128 MiB and cannot exceed the internal 512 MiB hard ceiling.
   */
  decodedByteBudget?: number;
  /** Default `adaptive`: preserve the document by reducing raster resolution. */
  strategy?: DecodedImageBudgetStrategy;
}

export interface NormalizedImageResourceOptions {
  readonly decodedByteBudget: number;
  readonly strategy: DecodedImageBudgetStrategy;
}

export function normalizeImageResourceOptions(
  options: ImageResourceOptions | undefined = undefined,
): Readonly<NormalizedImageResourceOptions> {
  if (options !== undefined
    && (options === null || typeof options !== 'object' || Array.isArray(options))) {
    throw new TypeError('imageResources must be an object when provided');
  }
  const decodedByteBudget = options?.decodedByteBudget ?? MAX_DECODED_IMAGE_BYTES;
  if (!Number.isSafeInteger(decodedByteBudget)
    || decodedByteBudget < 4
    || decodedByteBudget > HARD_MAX_DECODED_IMAGE_BYTES) {
    throw new RangeError(
      `imageResources.decodedByteBudget must be a safe integer from 4 to ${HARD_MAX_DECODED_IMAGE_BYTES} bytes`,
    );
  }
  const strategy = options?.strategy ?? 'adaptive';
  if (strategy !== 'adaptive' && strategy !== 'strict') {
    throw new TypeError("imageResources.strategy must be 'adaptive' or 'strict'");
  }
  return Object.freeze({ decodedByteBudget, strategy });
}

export interface DecodedImageTargetDemand {
  /** Cache identity including any pixel effect which produces another surface. */
  readonly key: string;
  /** Full-source decode target after crop expansion, in device pixels. */
  readonly targetWidthPx: number;
  readonly targetHeightPx: number;
  /** Declared source grid for a browser raster that supports decoder resize.
   * When every native surface fits, preserving this grid keeps the established
   * renderer output byte-for-byte. The display target is used only when native
   * ownership would cross an aggregate or per-surface boundary. */
  readonly sourceWidthPx?: number;
  readonly sourceHeightPx?: number;
  /** Simultaneously retained owned surfaces for this identity. Default 1. */
  readonly retainedSurfaceCount?: number;
}

export interface DecodedImageTargetPlan {
  readonly targets: ReadonlyMap<string, Readonly<{
    /** Decoder request axes. Browsers retain source aspect ratio from these bounds. */
    width: number;
    height: number;
    /** Actual aspect-preserving grid charged to the decoded-surface budget. */
    retainedWidth: number;
    retainedHeight: number;
    retainedPixels: number;
  }>>;
  readonly idealBytes: number;
  readonly plannedBytes: number;
  readonly qualityScale: number;
  readonly degraded: boolean;
}

interface MutableDemand {
  targetWidth: number;
  targetHeight: number;
  sourceWidth?: number;
  sourceHeight?: number;
  surfaces: number;
}

function positivePixelTarget(value: number): number {
  return Number.isFinite(value) && value > 0
    ? Math.max(1, Math.ceil(value))
    : 1;
}

function retainedSurfaceCount(value: number | undefined): number {
  return Number.isSafeInteger(value) && (value ?? 0) > 0 ? value as number : 1;
}

function byteCost(
  demands: ReadonlyMap<string, MutableDemand>,
  size: (demand: MutableDemand) => Readonly<{ width: number; height: number }>,
): number {
  let bytes = 0;
  for (const demand of demands.values()) {
    const selected = size(demand);
    const next = selected.width * selected.height * demand.surfaces * 4;
    bytes = Math.min(Number.MAX_SAFE_INTEGER, bytes + next);
  }
  return bytes;
}

/**
 * Allocate one uniform pixels-per-display-pixel scale across a render pass.
 * This is deliberately geometry-driven rather than sample-driven: every image
 * receives the same quality ratio, repeated cache identities are charged once,
 * and integer flooring guarantees the resulting plan never rounds above the
 * configured byte budget when at least one RGBA pixel per demand can fit.
 */
export function planDecodedImageTargets(
  input: readonly DecodedImageTargetDemand[],
  options: Readonly<NormalizedImageResourceOptions>,
): Readonly<DecodedImageTargetPlan> {
  const demands = new Map<string, MutableDemand>();
  for (const item of input) {
    const targetWidth = positivePixelTarget(item.targetWidthPx);
    const targetHeight = positivePixelTarget(item.targetHeightPx);
    const hasSource = item.sourceWidthPx !== undefined && item.sourceHeightPx !== undefined;
    const sourceWidth = hasSource ? positivePixelTarget(item.sourceWidthPx as number) : undefined;
    const sourceHeight = hasSource ? positivePixelTarget(item.sourceHeightPx as number) : undefined;
    const surfaces = retainedSurfaceCount(item.retainedSurfaceCount);
    const existing = demands.get(item.key);
    if (existing) {
      existing.targetWidth = Math.max(existing.targetWidth, targetWidth);
      existing.targetHeight = Math.max(existing.targetHeight, targetHeight);
      if (sourceWidth !== undefined && sourceHeight !== undefined) {
        existing.sourceWidth = Math.max(existing.sourceWidth ?? 0, sourceWidth);
        existing.sourceHeight = Math.max(existing.sourceHeight ?? 0, sourceHeight);
      }
      existing.surfaces = Math.max(existing.surfaces, surfaces);
    } else {
      demands.set(item.key, { targetWidth, targetHeight, sourceWidth, sourceHeight, surfaces });
    }
  }

  const nativeSize = (demand: MutableDemand) => ({
    width: demand.sourceWidth ?? demand.targetWidth,
    height: demand.sourceHeight ?? demand.targetHeight,
  });
  const scaledRequestSize = (demand: MutableDemand, scale: number) => ({
    width: Math.max(1, Math.floor(demand.targetWidth * scale)),
    height: Math.max(1, Math.floor(demand.targetHeight * scale)),
  });
  const retainedSize = (demand: MutableDemand, scale: number) => {
    const requested = scaledRequestSize(demand, scale);
    if (demand.sourceWidth === undefined || demand.sourceHeight === undefined) {
      return requested;
    }
    return decodedBitmapRetainedTarget(
      { width: demand.sourceWidth, height: demand.sourceHeight },
      requested.width,
      requested.height,
    ) ?? { width: demand.sourceWidth, height: demand.sourceHeight };
  };
  const plannedTarget = (demand: MutableDemand, scale: number) => {
    const requested = scaledRequestSize(demand, scale);
    const retained = retainedSize(demand, scale);
    return Object.freeze({
      ...requested,
      retainedWidth: retained.width,
      retainedHeight: retained.height,
      retainedPixels: retained.width * retained.height,
    });
  };
  const nativeBytes = byteCost(demands, nativeSize);
  const nativeSurfacesFit = [...demands.values()].every((demand) => {
    const native = nativeSize(demand);
    return native.width * native.height <= MAX_RASTER_PIXELS;
  });
  const preserveNative = nativeSurfacesFit && nativeBytes <= options.decodedByteBudget;
  const idealBytes = byteCost(demands, demand => retainedSize(demand, 1));
  if (preserveNative) {
    const targets = new Map<string, ReturnType<typeof plannedTarget>>();
    for (const [key, demand] of demands) {
      const native = nativeSize(demand);
      targets.set(key, Object.freeze({
        width: native.width,
        height: native.height,
        retainedWidth: native.width,
        retainedHeight: native.height,
        retainedPixels: native.width * native.height,
      }));
    }
    return Object.freeze({
      targets,
      idealBytes,
      plannedBytes: nativeBytes,
      qualityScale: 1,
      degraded: false,
    });
  }
  if (options.strategy === 'strict' && idealBytes > options.decodedByteBudget) {
    throw new OoxmlDecodedImageLimitError(
      'active-decoded-bytes',
      options.decodedByteBudget,
      idealBytes,
    );
  }
  if (options.strategy === 'adaptive') {
    const minimumBytes = byteCost(demands, demand => retainedSize(demand, 0));
    if (minimumBytes > options.decodedByteBudget) {
      throw new OoxmlDecodedImageLimitError(
        'active-decoded-bytes',
        options.decodedByteBudget,
        minimumBytes,
      );
    }
  }
  let adaptiveScale = 1;
  if (options.strategy === 'adaptive' && idealBytes > options.decodedByteBudget) {
    // Aspect-preserving decoder targets round the non-dominant axis upward.
    // Binary-search the largest uniform requested scale whose actual retained
    // grids stay within the byte budget; a closed-form square-root estimate can
    // otherwise exceed the limit by a few rows after integer projection.
    let lower = 0;
    let upper = 1;
    for (let iteration = 0; iteration < 40; iteration++) {
      const candidate = (lower + upper) / 2;
      const bytes = byteCost(demands, demand => retainedSize(demand, candidate));
      if (bytes <= options.decodedByteBudget) lower = candidate;
      else upper = candidate;
    }
    adaptiveScale = lower;
  }
  const targets = new Map<string, ReturnType<typeof plannedTarget>>();
  for (const [key, demand] of demands) {
    targets.set(key, plannedTarget(demand, adaptiveScale));
  }
  const plannedBytes = byteCost(demands, demand => retainedSize(demand, adaptiveScale));
  const effectiveQualityScale = adaptiveScale < 1 && demands.size > 0
    ? Math.min(...[...demands.entries()].flatMap(([key, demand]) => {
        const target = targets.get(key) as NonNullable<ReturnType<typeof targets.get>>;
        return [target.width / demand.targetWidth, target.height / demand.targetHeight];
      }))
    : adaptiveScale;
  return Object.freeze({
    targets,
    idealBytes,
    plannedBytes,
    qualityScale: effectiveQualityScale,
    degraded: adaptiveScale < 1,
  });
}
