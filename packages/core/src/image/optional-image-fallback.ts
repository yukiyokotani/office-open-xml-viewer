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

const OPTIONAL_IMAGE_LABELS: Readonly<Record<OptionalImageCodec, string>> = {
  tiff: 'TIFF image unavailable',
};

/**
 * Paint the same bounded, non-throwing capability placeholder in every format.
 * The fixed label avoids shaping package-controlled or attacker-controlled text.
 */
export function paintOptionalImagePlaceholder(
  ctx: CanvasRenderingContext2D,
  codec: OptionalImageCodec,
  bounds: OptionalImagePlaceholderBounds,
): void {
  if (![bounds.x, bounds.y, bounds.width, bounds.height].every(Number.isFinite)
    || bounds.width <= 0 || bounds.height <= 0) return;
  ctx.save();
  try {
    ctx.fillStyle = '#888';
    ctx.font = '11px sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    const label = OPTIONAL_IMAGE_LABELS[codec];
    ctx.fillText(
      label,
      bounds.x + bounds.width / 2,
      bounds.y + bounds.height / 2,
      bounds.width,
    );
  } finally {
    ctx.restore();
  }
}
