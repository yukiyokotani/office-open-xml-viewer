// Decoder for embedded SVG images (Microsoft's `asvg:svgBlip` extension,
// MS-ODRAWXML) shared by the DOCX, XLSX and PPTX lazy image pipelines.
//
// The cache is scoped by the stable per-document fetch function, then by OPC
// part path. An object URL exists only while one image loads/decodes and is
// always revoked in the producer's `finally`; cached HTMLImageElements never
// retain a library-owned raw-Blob handle. Browser-managed SVG parse/decode
// storage has no portable byte measurement or explicit close operation, so it
// is bounded by count while fetch/decode work is bounded to two concurrent
// operations per document.

import { getCachedBitmapByPath } from './bitmap-image-by-path.js';
import { withDecodedImageSlot } from './decode-gate.js';
import type { SvgBlobDecoder } from '../worker/svg-decode-bridge.js';

type FetchImage = (path: string, mimeType: string) => Promise<Blob>;
export type SvgImageSource = HTMLImageElement | ImageBitmap;

export interface SvgImageDecodeOptions {
  readonly targetWidthPx?: number;
  readonly targetHeightPx?: number;
  /** Variant/accounting ceiling for worker-rasterized SVG surfaces. */
  readonly maxRetainedPixels?: number;
  readonly workerDecoder?: SvgBlobDecoder;
}

interface SvgCacheEntry {
  promise: Promise<SvgImageSource>;
}

interface DocCache {
  readonly imgs: Map<string, SvgCacheEntry>;
}

const byFetch = new WeakMap<FetchImage, DocCache>();
const MAX_SVG_CACHE_ENTRIES = 256;

function docCacheFor(fetchImage: FetchImage): DocCache {
  let cache = byFetch.get(fetchImage);
  if (!cache) {
    cache = { imgs: new Map() };
    byFetch.set(fetchImage, cache);
  }
  return cache;
}

/** Decode one SVG part, deduped and count-bounded per document. The transient
 * object URL is revoked immediately after load/decode settles. */
export function getCachedSvgImageByPath(
  svgImagePath: string,
  fetchImage: FetchImage,
  options: SvgImageDecodeOptions = {},
): Promise<SvgImageSource> {
  // Dedicated Workers and Node have no HTMLImageElement. Route SVG bytes
  // through the same decoded owner as raster blips; browser/Node
  // createImageBitmap support (or the injected Node factory) determines
  // decodability, and the shared decoder validates the resulting dimensions.
  if (typeof Image === 'undefined') {
    return getCachedBitmapByPath(svgImagePath, 'image/svg+xml', fetchImage, {
      targetWidthPx: options.targetWidthPx,
      targetHeightPx: options.targetHeightPx,
      maxRetainedPixels: options.maxRetainedPixels,
      svgDecoder: options.workerDecoder,
    }).then((bitmap) => {
      if (!bitmap) throw new Error(`svg decode failed: ${svgImagePath}`);
      return bitmap;
    });
  }
  const cache = docCacheFor(fetchImage);
  const hit = cache.imgs.get(svgImagePath);
  if (hit) {
    cache.imgs.delete(svgImagePath);
    cache.imgs.set(svgImagePath, hit);
    return hit.promise;
  }

  const entry = {} as SvgCacheEntry;
  entry.promise = withDecodedImageSlot(fetchImage, async () => {
    if (cache.imgs.get(svgImagePath) !== entry) {
      throw new Error('SVG decode was superseded before it started');
    }
    const blob = await fetchImage(svgImagePath, 'image/svg+xml');
    const url = URL.createObjectURL(blob);
    try {
      const image = new Image();
      await new Promise<void>((resolve, reject) => {
        image.onload = () => {
          if (typeof image.decode === 'function') {
            image.decode().then(resolve).catch(resolve);
          } else {
            resolve();
          }
        };
        image.onerror = () => reject(new Error(`svg load failed: ${svgImagePath}`));
        image.src = url;
      });
      return image;
    } finally {
      URL.revokeObjectURL(url);
    }
  });

  // A rejected old entry must not remove a newer retry at the same path.
  void entry.promise.catch(() => {
    if (cache.imgs.get(svgImagePath) === entry) cache.imgs.delete(svgImagePath);
  });
  cache.imgs.set(svgImagePath, entry);
  if (cache.imgs.size > MAX_SVG_CACHE_ENTRIES) {
    const oldest = cache.imgs.keys().next().value as string;
    cache.imgs.delete(oldest);
  }
  return entry.promise;
}

/** Forget one document's loaded SVG references. Pending queued work checks the
 * map before fetching, and every already-started loader owns URL cleanup. */
export function dropSvgImageCache(fetchImage: FetchImage): void {
  const cache = byFetch.get(fetchImage);
  if (!cache) return;
  cache.imgs.clear();
  byFetch.delete(fetchImage);
}
