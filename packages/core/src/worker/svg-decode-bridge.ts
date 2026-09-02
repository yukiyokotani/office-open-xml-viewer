import {
  MAX_RASTER_DIMENSION,
  MAX_RASTER_PIXELS,
  OoxmlDecodedImageLimitError,
} from '../image/pixel-budget.js';
import { closeImageBitmapIfSupported } from '../image/image-bitmap-lifecycle.js';

export interface SvgDecodeTarget {
  /** Minimum retained source-grid width in device pixels. */
  readonly targetWidthPx?: number;
  /** Minimum retained source-grid height in device pixels. */
  readonly targetHeightPx?: number;
}

export type SvgBlobDecoder = (
  blob: Blob,
  target?: SvgDecodeTarget,
) => Promise<ImageBitmap>;

export interface WorkerSvgDecodeRequest extends SvgDecodeTarget {
  readonly kind: 'ooxmlDecodeSvg';
  readonly decodeId: number;
  readonly bytes: ArrayBuffer;
}

export type WorkerSvgDecodeResponse =
  | { readonly kind: 'ooxmlSvgDecoded'; readonly decodeId: number; readonly bitmap: ImageBitmap }
  | { readonly kind: 'ooxmlSvgDecodeFailed'; readonly decodeId: number; readonly message: string };

type WorkerPost = (message: unknown, transfer?: Transferable[]) => void;

function finitePositive(value: number | undefined): number | undefined {
  return typeof value === 'number' && Number.isFinite(value) && value > 0
    ? Math.ceil(value)
    : undefined;
}

/** Bound a requested SVG raster surface to the same per-image hard quota used
 * by ordinary decoded OOXML images. Both axes scale together, preserving the
 * authored vector aspect ratio when a host asks for an oversized target. */
export function boundedSvgRasterSize(width: number, height: number): { width: number; height: number } {
  if (!Number.isFinite(width) || !Number.isFinite(height) || !(width > 0) || !(height > 0)) {
    throw new Error(`invalid SVG raster size: ${width}x${height}`);
  }
  // Divide the square roots separately: `width * height` can overflow to
  // Infinity for finite hostile inputs, which would collapse the scale to zero
  // and produce NaN dimensions after multiplication.
  const pixelScale = Math.sqrt(MAX_RASTER_PIXELS) / Math.sqrt(width) / Math.sqrt(height);
  const scale = Math.min(
    1,
    MAX_RASTER_DIMENSION / width,
    MAX_RASTER_DIMENSION / height,
    pixelScale,
  );
  return {
    width: Math.min(MAX_RASTER_DIMENSION, Math.max(1, Math.floor(width * scale))),
    height: Math.min(MAX_RASTER_DIMENSION, Math.max(1, Math.floor(height * scale))),
  };
}

/** Smallest bounded source grid that covers the requested axes without changing
 * the SVG's intrinsic aspect ratio. Target axes are coverage requirements, as
 * they are for ordinary raster decode; they are not a replacement aspect ratio. */
function svgRasterCoverageSize(
  sourceWidth: number,
  sourceHeight: number,
  targetWidth: number | undefined,
  targetHeight: number | undefined,
): { width: number; height: number } {
  const widthScale = targetWidth === undefined ? 0 : targetWidth / sourceWidth;
  const heightScale = targetHeight === undefined ? 0 : targetHeight / sourceHeight;
  const requiredScale = targetWidth === undefined && targetHeight === undefined
    ? 1
    : Math.max(widthScale, heightScale);
  const quotaScale = Math.min(
    MAX_RASTER_DIMENSION / sourceWidth,
    MAX_RASTER_DIMENSION / sourceHeight,
    Math.sqrt(MAX_RASTER_PIXELS) / Math.sqrt(sourceWidth) / Math.sqrt(sourceHeight),
  );
  if (!Number.isFinite(requiredScale) || !(requiredScale > 0)
    || !Number.isFinite(quotaScale) || !(quotaScale > 0)) {
    throw new Error(`invalid SVG raster scale: ${Math.min(requiredScale, quotaScale)}`);
  }
  if (quotaScale < requiredScale) {
    // The quota, rather than a caller axis, dominates. Round inward so the
    // retained surface cannot cross that quota.
    return boundedSvgRasterSize(
      Math.max(1, Math.floor(sourceWidth * quotaScale)),
      Math.max(1, Math.floor(sourceHeight * quotaScale)),
    );
  }
  // Assign the scale-dominating requested axis directly before deriving the
  // other. This avoids floating roundoff growing a 400px request to 401px.
  if (targetWidth !== undefined && widthScale >= heightScale) {
    return boundedSvgRasterSize(
      targetWidth,
      Math.max(1, Math.ceil(sourceHeight * targetWidth / sourceWidth)),
    );
  }
  if (targetHeight !== undefined) {
    return boundedSvgRasterSize(
      Math.max(1, Math.ceil(sourceWidth * targetHeight / sourceHeight)),
      targetHeight,
    );
  }
  return boundedSvgRasterSize(Math.ceil(sourceWidth), Math.ceil(sourceHeight));
}

/** Decode SVG with the Window's standards-compliant image pipeline, then
 * transfer a bounded, display-sized ImageBitmap back to the render worker.
 * Dedicated workers cannot rely on `createImageBitmap(svgBlob)` in Chromium. */
export async function decodeSvgBlobOnMainThread(
  blob: Blob,
  target: SvgDecodeTarget = {},
): Promise<ImageBitmap> {
  if (typeof Image === 'undefined') throw new Error('SVG host decode requires HTMLImageElement');
  const url = URL.createObjectURL(blob);
  try {
    const image = new Image();
    const targetWidth = finitePositive(target.targetWidthPx);
    const targetHeight = finitePositive(target.targetHeightPx);
    await new Promise<void>((resolve, reject) => {
      image.onload = () => {
        if (typeof image.decode === 'function') image.decode().then(resolve).catch(resolve);
        else resolve();
      };
      image.onerror = () => reject(new Error('SVG host decode failed'));
      image.src = url;
    });
    // A viewBox-only SVG receives the HTML default object size (300×150), with
    // its viewBox ratio reflected in naturalWidth/naturalHeight. Truly
    // dimensionless sources can still report zero; use that same default box so
    // the explicit Canvas target below remains drawable and deterministic.
    const hasNaturalSize = Number.isFinite(image.naturalWidth) && image.naturalWidth > 0
      && Number.isFinite(image.naturalHeight) && image.naturalHeight > 0;
    const sourceWidth = hasNaturalSize ? image.naturalWidth : 300;
    const sourceHeight = hasNaturalSize ? image.naturalHeight : 150;
    const size = svgRasterCoverageSize(sourceWidth, sourceHeight, targetWidth, targetHeight);
    image.width = size.width;
    image.height = size.height;
    // Rasterize explicitly instead of createImageBitmap(image): Chromium can
    // return a transparent bitmap for a viewBox-only SVG even after the image
    // loaded successfully. Worker render mode already requires OffscreenCanvas;
    // transferToImageBitmap moves that backing store without retaining a second
    // full-size surface on Window.
    let bitmap: ImageBitmap;
    if (typeof OffscreenCanvas !== 'undefined') {
      const canvas = new OffscreenCanvas(size.width, size.height);
      const context = canvas.getContext('2d');
      if (!context) throw new Error('SVG host raster target has no 2-D context');
      context.drawImage(image, 0, 0, size.width, size.height);
      bitmap = canvas.transferToImageBitmap();
    } else {
      const canvas = document.createElement('canvas');
      canvas.width = size.width;
      canvas.height = size.height;
      const context = canvas.getContext('2d');
      if (!context) throw new Error('SVG host raster target has no 2-D context');
      context.drawImage(image, 0, 0, size.width, size.height);
      bitmap = await createImageBitmap(canvas);
    }
    const pixels = bitmap.width * bitmap.height;
    if (
      bitmap.width > MAX_RASTER_DIMENSION
      || bitmap.height > MAX_RASTER_DIMENSION
      || pixels > MAX_RASTER_PIXELS
    ) {
      closeImageBitmapIfSupported(bitmap);
      throw new OoxmlDecodedImageLimitError('image-pixels', MAX_RASTER_PIXELS, pixels);
    }
    return bitmap;
  } finally {
    URL.revokeObjectURL(url);
  }
}

export function isWorkerSvgDecodeRequest(value: unknown): value is WorkerSvgDecodeRequest {
  return !!value && typeof value === 'object'
    && (value as { kind?: unknown }).kind === 'ooxmlDecodeSvg';
}

export function isWorkerSvgDecodeResponse(value: unknown): value is WorkerSvgDecodeResponse {
  if (!value || typeof value !== 'object') return false;
  const kind = (value as { kind?: unknown }).kind;
  return kind === 'ooxmlSvgDecoded' || kind === 'ooxmlSvgDecodeFailed';
}

/** Main-thread half of the SVG decode bridge. Returns true when `message` was
 * consumed. The async response owns the transferred bitmap; a failed post
 * closes it locally so teardown races cannot leak GPU memory. */
export function respondToWorkerSvgDecodeRequest(
  post: WorkerPost,
  message: unknown,
): boolean {
  if (!isWorkerSvgDecodeRequest(message)) return false;
  void decodeSvgBlobOnMainThread(
    new Blob([message.bytes], { type: 'image/svg+xml' }),
    message,
  ).then((bitmap) => {
    try {
      post({ kind: 'ooxmlSvgDecoded', decodeId: message.decodeId, bitmap }, [bitmap]);
    } catch {
      closeImageBitmapIfSupported(bitmap);
    }
  }).catch((error: unknown) => {
    try {
      post({
        kind: 'ooxmlSvgDecodeFailed',
        decodeId: message.decodeId,
        message: error instanceof Error ? error.message : String(error),
      });
    } catch {
      // The worker has already terminated; no remote waiter remains.
    }
  });
  return true;
}

/** Worker-thread half. It transfers SVG bytes to Window only on a cache miss,
 * correlates the returned ImageBitmap, and exposes a decoder compatible with
 * the shared per-document decoded-bitmap cache. */
export class WorkerSvgDecodeClient {
  private nextId = 1;
  private readonly pending = new Map<
    number,
    { resolve: (bitmap: ImageBitmap) => void; reject: (error: Error) => void }
  >();

  constructor(private readonly post: WorkerPost) {}

  readonly decode: SvgBlobDecoder = async (blob, target = {}) => {
    const decodeId = this.nextId++;
    const bytes = await blob.arrayBuffer();
    return await new Promise<ImageBitmap>((resolve, reject) => {
      this.pending.set(decodeId, { resolve, reject });
      try {
        this.post({ kind: 'ooxmlDecodeSvg', decodeId, bytes, ...target }, [bytes]);
      } catch (error) {
        this.pending.delete(decodeId);
        reject(error instanceof Error ? error : new Error(String(error)));
      }
    });
  };

  accept(message: unknown): boolean {
    if (!isWorkerSvgDecodeResponse(message)) return false;
    const waiter = this.pending.get(message.decodeId);
    if (!waiter) {
      if (message.kind === 'ooxmlSvgDecoded') closeImageBitmapIfSupported(message.bitmap);
      return true;
    }
    this.pending.delete(message.decodeId);
    if (message.kind === 'ooxmlSvgDecoded') waiter.resolve(message.bitmap);
    else waiter.reject(new Error(message.message));
    return true;
  }

  dispose(error = new Error('SVG decode worker disposed')): void {
    for (const waiter of this.pending.values()) waiter.reject(error);
    this.pending.clear();
  }
}
