// ── Shared raster pixel-dimension budget (DoS / decode-bomb guard) ───────────
//
// One source of truth for the caps that bound encoded raster inputs and retained
// decoded surfaces. Metafile-embedded DIBs cannot request decoder-side resizing,
// so they use the retained-surface limits directly; standalone browser rasters
// may use the larger source limits only when decoded to a bounded display target.

/**
 * Implementation hard ceiling for one raster axis. This rejects obviously huge
 * inputs early; it is not a portable browser capability claim. A runtime or
 * device may impose a lower canvas/decode limit, whose ordinary decode/draw
 * failure remains possible below this ceiling.
 */
export const MAX_RASTER_DIMENSION = 32767;

/**
 * Pixel budget for one decoded raster: 32 MP (2^25 px). A decoded surface is
 * `width × height × 4` bytes of RGBA, so this bounds one bitmap to 128 MiB.
 * A crafted 60000×60000 header (~3.6e9 px → ~14 GB RGBA)
 * is refused before any allocation. With both
 * axes ≤ MAX_RASTER_DIMENSION the product stays ≤ ~1.07e9 — exact in an
 * IEEE-754 double — so a plain numeric comparison suffices (no BigInt).
 */
export const MAX_RASTER_PIXELS = 1 << 25; // 33_554_432 px = 32 MP / 128 MiB RGBA

/**
 * Hard ceiling for the encoded raster's declared source grid. Display-sized
 * decoding may safely retain far fewer pixels than the source contains, but a
 * decoder is still allowed to do implementation-specific intermediate work.
 * Keep obviously hostile headers away from it even when a tiny output was
 * requested. The source ceiling is four times the retained-surface ceiling,
 * bounding a conventional 4-byte full-frame intermediate to 512 MiB while
 * leaving useful headroom for high-resolution authored assets.
 */
export const MAX_RASTER_SOURCE_PIXELS = 1 << 27;

/**
 * Derived source-axis ceiling. A positive grid within the source-pixel budget
 * cannot exceed this value on either axis, so this adds no format-specific
 * magic number (in particular, JPEG's 16-bit SOF bound is not imposed on PNG
 * or TIFF). It keeps error classification and unsafe-number checks explicit.
 */
export const MAX_RASTER_SOURCE_DIMENSION = MAX_RASTER_SOURCE_PIXELS;

/** Maximum decoded RGBA ownership retained or leased per document. */
export const MAX_DECODED_IMAGE_BYTES = MAX_RASTER_PIXELS * 4;

/**
 * Non-disableable aggregate ceiling for caller-configured decoded-image
 * budgets. The ordinary 128 MiB value above is the adaptive default; trusted
 * desktop integrations may raise it, but never beyond the same 512 MiB RGBA
 * envelope used to bound an encoded source grid before browser decoding.
 */
export const HARD_MAX_DECODED_IMAGE_BYTES = MAX_RASTER_SOURCE_PIXELS * 4;

/** Keep simultaneous browser decoders bounded even before exact pixels exist. */
export const MAX_CONCURRENT_IMAGE_DECODES = 2;

export type OoxmlDecodedImageLimitMetric =
  | 'image-dimension'
  | 'image-pixels'
  | 'active-decoded-bytes';

export interface OoxmlDecodedImageLimitDetails {
  readonly metric: OoxmlDecodedImageLimitMetric;
  readonly limit: number;
  readonly observed: number;
}

function isDecodedImageLimitMetric(value: unknown): value is OoxmlDecodedImageLimitMetric {
  return value === 'image-dimension'
    || value === 'image-pixels'
    || value === 'active-decoded-bytes';
}

function isNonNegativeSafeInteger(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0;
}

/** Catchable hard-quota crossing for decoded image surfaces. */
export class OoxmlDecodedImageLimitError extends RangeError {
  readonly code = 'ooxml-decoded-image-limit' as const;

  constructor(
    readonly metric: OoxmlDecodedImageLimitMetric,
    readonly limit: number,
    readonly observed: number,
  ) {
    super(`OOXML decoded image limit exceeded: ${metric} ${observed} > ${limit}`);
    this.name = 'OoxmlDecodedImageLimitError';
    Object.setPrototypeOf(this, OoxmlDecodedImageLimitError.prototype);
  }
}

/**
 * Snapshot a decoded-image quota crossing from this or another bundled core
 * realm. Every caller-controlled property read stays inside the guard so a
 * hostile Proxy or accessor cannot escape, and only clone-safe scalar fields
 * leave this boundary.
 */
export function getOoxmlDecodedImageLimitDetails(
  error: unknown,
): OoxmlDecodedImageLimitDetails | undefined {
  if (!error || typeof error !== 'object') return undefined;
  try {
    const candidate = error as {
      readonly code?: unknown;
      readonly metric?: unknown;
      readonly limit?: unknown;
      readonly observed?: unknown;
    };
    const code = candidate.code;
    const metric = candidate.metric;
    const limit = candidate.limit;
    const observed = candidate.observed;
    if (
      code !== 'ooxml-decoded-image-limit'
      || !isDecodedImageLimitMetric(metric)
      || !isNonNegativeSafeInteger(limit)
      || !isNonNegativeSafeInteger(observed)
      || observed <= limit
    ) {
      return undefined;
    }
    return { metric, limit, observed };
  } catch {
    return undefined;
  }
}

export function isOoxmlDecodedImageLimitError(
  error: unknown,
): error is OoxmlDecodedImageLimitError {
  if (!getOoxmlDecodedImageLimitDetails(error)) return false;
  try {
    const candidate = error as { readonly name?: unknown; readonly message?: unknown };
    return typeof candidate.name === 'string' && typeof candidate.message === 'string';
  } catch {
    return false;
  }
}
