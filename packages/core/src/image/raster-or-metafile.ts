// Shared admission and decode boundary for DrawingML raster/metafile blips.
// Format players remain in their format-specific modules; this module owns the
// cross-format safety policy, decoder-side resizing, and final surface check.

import { renderEmfToBitmap } from './emf.js';
import { closeImageBitmapIfSupported } from './image-bitmap-lifecycle.js';
import {
  MAX_RASTER_DIMENSION,
  MAX_RASTER_PIXELS,
  MAX_RASTER_SOURCE_DIMENSION,
  MAX_RASTER_SOURCE_PIXELS,
  OoxmlDecodedImageLimitError,
  isOoxmlDecodedImageLimitError,
} from './pixel-budget.js';
import { inspectRasterBlob, type RasterBlobInspection } from './raster-blob-inspection.js';
import {
  sourceRasterExceedsBudget,
  type RasterDimensions,
} from './raster-dimensions.js';
import { decodedBitmapRetainedTarget } from './raster-target.js';
import {
  isTiff,
  isTiffDecodeError,
  TiffDecodeError,
  type TiffRenderer,
} from './tiff-contract.js';
import { isEmf, isWmf, renderWmfToBitmap, wmfRasterTarget } from './wmf.js';

export interface DecodeRasterOptions {
  widthPt?: number;
  heightPt?: number;
  suppressBoundaryFrame?: boolean;
  tiff?: TiffRenderer;
  targetWidthPx?: number;
  targetHeightPx?: number;
  /** Retained base-surface ceiling. Effect pipelines lower this so their base
   * and derived surfaces fit the aggregate decoded-byte budget. */
  maxRetainedPixels?: number;
}

function exceedsRetainedBudget(source: RasterDimensions, pixelLimit: number): boolean {
  return source.width <= 0 || source.height <= 0
    || source.width > MAX_RASTER_DIMENSION || source.height > MAX_RASTER_DIMENSION
    || source.width * source.height > pixelLimit;
}

function hasTiffMimeType(value: string): boolean {
  const mime = value.split(';', 1)[0]?.trim().toLowerCase();
  return mime === 'image/tiff' || mime === 'image/x-tiff';
}

interface RasterDecodePlan {
  readonly retainedDimensions: RasterDimensions;
  readonly resizeOptions: ImageBitmapOptions | null;
}

function decodePlan(
  source: RasterDimensions,
  targetWidthPx: number | undefined,
  targetHeightPx: number | undefined,
  allowSingleAxis = false,
): RasterDecodePlan {
  const native = { retainedDimensions: source, resizeOptions: null };
  const targetWidth = typeof targetWidthPx === 'number'
    && Number.isFinite(targetWidthPx) && targetWidthPx > 0 ? targetWidthPx : undefined;
  const targetHeight = typeof targetHeightPx === 'number'
    && Number.isFinite(targetHeightPx) && targetHeightPx > 0 ? targetHeightPx : undefined;
  // A target is a sufficient downsample request only when neither source axis
  // needs native resolution. If one target axis reaches/exceeds the source,
  // retaining the source grid is genuinely required; quota checks below reject
  // it rather than silently substituting a smaller, insufficient surface.
  const target = decodedBitmapRetainedTarget(
    source,
    targetWidth,
    targetHeight,
    allowSingleAxis,
  );
  if (!target) return native;
  const resizeWidth = target.width;
  // One axis lets the HTML algorithm preserve the oriented source aspect ratio
  // (including EXIF rotation) instead of imposing the coded header's W×H.
  return {
    retainedDimensions: target,
    resizeOptions: { resizeWidth, resizeQuality: 'high' },
  };
}

function rasterLimitError(
  dimensions: RasterDimensions,
  dimensionLimit: number,
  pixelLimit: number,
): OoxmlDecodedImageLimitError {
  const observedDimension = Math.max(dimensions.width, dimensions.height);
  if (!Number.isFinite(observedDimension) || observedDimension > dimensionLimit) {
    return new OoxmlDecodedImageLimitError(
      'image-dimension',
      dimensionLimit,
      Number.isFinite(observedDimension) ? observedDimension : Number.MAX_SAFE_INTEGER,
    );
  }
  const observedPixels = dimensions.width * dimensions.height;
  return new OoxmlDecodedImageLimitError(
    'image-pixels',
    pixelLimit,
    Number.isSafeInteger(observedPixels) && observedPixels >= 0
      ? observedPixels
      : Number.MAX_SAFE_INTEGER,
  );
}

export async function decodeRasterOrMetafile(
  data: Blob,
  opts: DecodeRasterOptions = {},
): Promise<ImageBitmap | null> {
  return decodeRasterOrMetafileWithInspection(data, opts);
}

/** Cache entry point for reusing metadata already inspected before key choice. */
export async function decodeRasterOrMetafileWithInspection(
  data: Blob,
  opts: DecodeRasterOptions = {},
  knownInspection?: RasterBlobInspection,
): Promise<ImageBitmap | null> {
  const {
    widthPt = 0,
    heightPt = 0,
    suppressBoundaryFrame = false,
    tiff,
    targetWidthPx,
    targetHeightPx,
    maxRetainedPixels = MAX_RASTER_PIXELS,
  } = opts;
  const retainedPixelLimit = Number.isSafeInteger(maxRetainedPixels) && maxRetainedPixels > 0
    ? Math.min(maxRetainedPixels, MAX_RASTER_PIXELS)
    : MAX_RASTER_PIXELS;
  const head = new Uint8Array(await data.slice(0, 64 * 1024).arrayBuffer());

  if (isWmf(head)) {
    const { w, h } = wmfRasterTarget(widthPt, heightPt);
    return enforceDecodedBitmapBudget(
      await renderWmfToBitmap(new Uint8Array(await data.arrayBuffer()), w, h, suppressBoundaryFrame),
      retainedPixelLimit,
    );
  }
  if (isEmf(head)) {
    const { w, h } = wmfRasterTarget(widthPt, heightPt);
    return enforceDecodedBitmapBudget(
      await renderEmfToBitmap(new Uint8Array(await data.arrayBuffer()), w, h),
      retainedPixelLimit,
    );
  }

  const inspection = knownInspection ?? await inspectRasterBlob(data, head);
  const rasterDimensions = inspection.dimensions;
  if (rasterDimensions && sourceRasterExceedsBudget(rasterDimensions)) {
    throw rasterLimitError(rasterDimensions, MAX_RASTER_SOURCE_DIMENSION, MAX_RASTER_SOURCE_PIXELS);
  }
  // Prefer content sniffing because OOXML producers sometimes label TIFF parts
  // as application/octet-stream. MIME remains a second recognition signal so
  // valid-but-unsupported TIFF container versions and damaged TIFF parts fail
  // through the diagnostic codec path instead of the browser's generic decoder.
  const tiffInput = isTiff(head) || hasTiffMimeType(data.type);
  const plan = rasterDimensions
    ? decodePlan(rasterDimensions, targetWidthPx, targetHeightPx, tiffInput)
    : null;
  if (plan && exceedsRetainedBudget(plan.retainedDimensions, retainedPixelLimit)) {
    throw rasterLimitError(plan.retainedDimensions, MAX_RASTER_DIMENSION, retainedPixelLimit);
  }
  if (tiffInput) {
    if (!tiff) throw new TiffDecodeError('TIFF image requires an opt-in TIFF codec');
    let bitmap: ImageBitmap | null;
    try {
      bitmap = await tiff.render(new Uint8Array(await data.arrayBuffer()), {
        targetWidthPx,
        targetHeightPx,
        maxRetainedPixels: retainedPixelLimit,
      });
    } catch (error) {
      if (isOoxmlDecodedImageLimitError(error) || isTiffDecodeError(error)) throw error;
      throw new TiffDecodeError('TIFF codec failed to decode the image', { cause: error });
    }
    if (!bitmap) throw new TiffDecodeError('TIFF codec failed to decode the image');
    return enforceDecodedBitmapBudget(bitmap, retainedPixelLimit);
  }
  return enforceDecodedBitmapBudget(
    plan?.resizeOptions
      ? await createImageBitmap(data, plan.resizeOptions)
      : await createImageBitmap(data),
    retainedPixelLimit,
  );
}

function enforceDecodedBitmapBudget(
  bitmap: ImageBitmap | null,
  pixelLimit = MAX_RASTER_PIXELS,
): ImageBitmap | null {
  if (!bitmap) return null;
  const dimensions = { width: Number(bitmap.width), height: Number(bitmap.height) };
  if (!exceedsRetainedBudget(dimensions, pixelLimit)) return bitmap;
  closeImageBitmapIfSupported(bitmap);
  throw rasterLimitError(dimensions, MAX_RASTER_DIMENSION, pixelLimit);
}
