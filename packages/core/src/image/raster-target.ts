import type { RasterDimensions } from './raster-dimensions.js';

function positiveFinite(value: number | undefined): number | undefined {
  return typeof value === 'number' && Number.isFinite(value) && value > 0
    ? value
    : undefined;
}

/** Smallest integer, aspect-preserving source grid that covers the requested
 * axes. The scale-dominating axis is assigned directly before the other axis is
 * derived, avoiding floating-point roundoff such as 25 * (7 / 25) > 7. */
export function aspectPreservingRasterTarget(
  source: Readonly<RasterDimensions>,
  targetWidthPx: number | undefined,
  targetHeightPx: number | undefined,
  allowSingleAxis = false,
): RasterDimensions | null {
  if (!Number.isFinite(source.width) || !Number.isFinite(source.height)
    || !(source.width > 0) || !(source.height > 0)) return null;
  const targetWidth = positiveFinite(targetWidthPx);
  const targetHeight = positiveFinite(targetHeightPx);
  if (allowSingleAxis ? targetWidth === undefined && targetHeight === undefined
    : targetWidth === undefined || targetHeight === undefined) return null;
  if ((targetWidth !== undefined && !(targetWidth < source.width))
    || (targetHeight !== undefined && !(targetHeight < source.height))) return null;

  const widthScale = targetWidth === undefined ? 0 : targetWidth / source.width;
  const heightScale = targetHeight === undefined ? 0 : targetHeight / source.height;
  let width: number;
  let height: number;
  if (targetWidth !== undefined && widthScale >= heightScale) {
    width = Math.max(1, Math.ceil(targetWidth));
    height = Math.max(1, Math.ceil(source.height * width / source.width));
  } else if (targetHeight !== undefined) {
    height = Math.max(1, Math.ceil(targetHeight));
    width = Math.max(1, Math.ceil(source.width * height / source.height));
  } else {
    return null;
  }
  return width < source.width && height < source.height ? { width, height } : null;
}

/** Pixel grid retained by the browser when the HTML decode request supplies
 * only `resizeWidth`. The one-axis request is intentional: it preserves the
 * decoder's natural (including EXIF-oriented) aspect ratio. This helper is the
 * single accounting counterpart for that browser behavior. */
export function decodedBitmapRetainedTarget(
  source: Readonly<RasterDimensions>,
  targetWidthPx: number | undefined,
  targetHeightPx: number | undefined,
  allowSingleAxis = false,
): RasterDimensions | null {
  const target = aspectPreservingRasterTarget(
    source,
    targetWidthPx,
    targetHeightPx,
    allowSingleAxis,
  );
  if (!target) return null;
  const width = target.width;
  const height = Math.min(
    source.height,
    Math.max(1, Math.ceil(source.height * width / source.width)),
  );
  return { width, height };
}

/** Browser one-axis resize request for an already-decoded pixel surface. */
export function decodedBitmapTargetResizeOptions(
  sourceWidth: number,
  sourceHeight: number,
  targetWidthPx?: number,
  targetHeightPx?: number,
): ImageBitmapOptions | undefined {
  const target = decodedBitmapRetainedTarget(
    { width: sourceWidth, height: sourceHeight },
    targetWidthPx,
    targetHeightPx,
    true,
  );
  return target
    ? { resizeWidth: target.width, resizeQuality: 'high' }
    : undefined;
}
