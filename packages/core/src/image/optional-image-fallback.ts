import { imagePaintProjection, withImagePaintProjection, type ImagePaintProjectionOptions } from './crop';
/** Optional image codecs that can be omitted from the base viewer bundle. */
export type OptionalImageCodec = 'tiff';

/**
 * Internal acquisition signal: the image is recognized and its geometry is
 * valid, but the application did not opt in to the codec needed to decode it.
 * Renderers contain this signal at the authored image bounds. It must never be
 * used for malformed input, codec failures, or resource-budget violations.
 */
export class OptionalImageCodecUnavailableError extends Error {
  readonly code = 'ooxml-optional-image-codec-unavailable' as const;

  constructor(readonly codec: OptionalImageCodec) {
    super(`${codec.toUpperCase()} image requires an optional codec`);
    this.name = 'OptionalImageCodecUnavailableError';
    Object.setPrototypeOf(this, OptionalImageCodecUnavailableError.prototype);
  }
}

/** Recognize the internal signal without relying on realm-specific instanceof. */
export function isOptionalImageCodecUnavailableError(
  error: unknown,
  codec?: OptionalImageCodec,
): error is OptionalImageCodecUnavailableError {
  if (typeof error !== 'object' || error === null) return false;
  try {
    const candidate = error as { readonly code?: unknown; readonly codec?: unknown };
    return candidate.code === 'ooxml-optional-image-codec-unavailable'
      && candidate.codec === 'tiff'
      && (codec === undefined || candidate.codec === codec);
  } catch {
    return false;
  }
}

export interface OptionalImagePlaceholderBounds {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

const OPTIONAL_IMAGE_LABELS: Readonly<Record<OptionalImageCodec | 'pict', string>> = {
  tiff: 'TIFF image unavailable',
  pict: 'PICT image unsupported',
};

/**
 * Paint the same bounded, non-throwing capability placeholder in every format.
 * PICT is an opaque image resource, not an optional codec: callers may
 * label that unsupported format without claiming that its pixels were decoded.
 * The fixed label avoids shaping package-controlled or attacker-controlled text.
 */
export function paintOptionalImagePlaceholder(
  ctx: CanvasRenderingContext2D,
  codec: OptionalImageCodec | 'pict',
  bounds: OptionalImagePlaceholderBounds,
  options?: Readonly<ImagePaintProjectionOptions>,
): void {
  if (![bounds.x, bounds.y, bounds.width, bounds.height].every(Number.isFinite)
    || bounds.width <= 0 || bounds.height <= 0) return;
  const paint = (x: number, y: number, width: number, height: number) => {
    ctx.save();
    try {
      ctx.fillStyle = '#888'; ctx.font = '11px sans-serif';
      ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
      ctx.fillText(OPTIONAL_IMAGE_LABELS[codec], x + width / 2, y + height / 2, width);
    } finally { ctx.restore(); }
  };
  if (!options) { paint(bounds.x, bounds.y, bounds.width, bounds.height); return; }
  const plan = options.projection ?? imagePaintProjection(bounds.width, bounds.height, null, options.transform, false);
  if (plan.width !== bounds.width || plan.height !== bounds.height) throw new RangeError('Image placeholder projection destination mismatch');
  withImagePaintProjection(ctx, bounds.x, bounds.y, plan, paint);
}
