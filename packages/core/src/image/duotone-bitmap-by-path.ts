// DrawingML `<a:duotone>` (§20.1.8.23) adapter shared by all three formats.
// Base and recoloured surfaces live in one per-document weighted decoded owner,
// so the expensive transform is memoized without a parallel lifecycle or byte-
// accounting system.

import {
  captureDecodedBitmapCacheEpoch,
  dropCachedDerivedBitmapNamespace,
  getCachedBitmapByPath,
  getCachedDerivedBitmap,
  resolvedCachedBitmapVariantKey,
  type CachedBitmapOptions,
} from './bitmap-image-by-path';
import { applyDuotone, type Duotone, type OffscreenFactory } from './duotone';
import { imageNaturalSize } from './crop';
import { MAX_RASTER_PIXELS } from './pixel-budget.js';
import { decodedBitmapTargetResizeOptions } from './raster-target.js';

type FetchImage = (path: string, mime: string) => Promise<Blob>;

/** Cache key for a decoded bitmap that may carry a `<a:duotone>` recolour. A
 *  plain picture is keyed by its zip `imagePath`; a duotone picture is keyed by
 *  the path PLUS both resolved endpoint colours, so the recoloured bitmap is
 *  cached and looked up separately from the raw blip. Callers compute this both
 *  when warming the cache and when drawing, so the two agree without sharing a
 *  cache reference. Mirrors xlsx's `imageCacheKey` and docx's former
 *  `imageKey(path, colorReplaceFrom)`. */
export function duotoneCacheKey(imagePath: string, duotone?: Duotone | null): string {
  return duotone ? `${imagePath}|duo:${duotone.clr1}:${duotone.clr2}` : imagePath;
}

const DUOTONE_CACHE_NAMESPACE = 'duotone';

/**
 * Decode a raster/metafile blip at `imagePath` and, when `duotone` is set,
 * recolour it along the `clr1`→`clr2` luminance ramp (§20.1.8.23), returning a
 * drawable source cached per document then by `duotoneCacheKey`.
 *
 * With NO duotone this is a thin pass-through to {@link getCachedBitmapByPath}
 * (the shared base cache) — no second-layer entry is created, so a non-duotone
 * picture behaves byte-for-byte as before. With a duotone the base bitmap is
 * decoded (and cached) once, then the recolour runs once per colour pair and is
 * memoized here.
 *
 * The recolour needs a readable pixel grid, so it only applies to a decoded
 * raster/metafile bitmap; a `null` base (an unsupported metafile — true EMF /
 * geometry-less WMF) propagates as `null` and the draw site skips it. When the
 * offscreen pixel pipeline is unavailable, {@link applyDuotone} returns the base
 * unchanged, so the picture still draws (just without the recolour).
 *
 * @param opts.offscreenFactory optional surface factory for environments without
 *   a global `OffscreenCanvas` (node); forwarded to {@link applyDuotone}.
 */
export async function getCachedDuotoneBitmapByPath(
  imagePath: string,
  mimeType: string,
  duotone: Duotone | null | undefined,
  fetchImage: FetchImage,
  opts: CachedBitmapOptions & {
    offscreenFactory?: OffscreenFactory;
    /** When true, an authored duotone that cannot be applied returns null
     * instead of silently drawing the original pixels. Chart picture markers
     * use this fail-closed mode; established shape consumers retain their
     * legacy fallback unless they opt in. */
    failClosedOnDuotoneFailure?: boolean;
  } = {},
): Promise<ImageBitmap | null> {
  const { offscreenFactory, failClosedOnDuotoneFailure = false, ...requestedBitmapOpts } = opts;
  const bitmapOpts = duotone
    ? {
        ...requestedBitmapOpts,
        // Peak pipeline: source ImageBitmap + offscreen backing + ImageData +
        // result ImageBitmap. Bound the base to one quarter of the document
        // byte ceiling so transient pixel work cannot silently double it.
        maxRetainedPixels: Math.min(
          requestedBitmapOpts.maxRetainedPixels ?? MAX_RASTER_PIXELS,
          Math.floor(MAX_RASTER_PIXELS / 4),
        ),
      }
    : requestedBitmapOpts;
  // DrawingML effects consume the authored source pixels. Keep display targets
  // off the base decode and resample only while baking the transformed output;
  // an over-budget native effect grid is rejected rather than approximated.
  const sourceBitmapOpts = duotone
    ? { ...bitmapOpts, targetWidthPx: undefined, targetHeightPx: undefined }
    : bitmapOpts;
  const epoch = duotone
    ? captureDecodedBitmapCacheEpoch(fetchImage, DUOTONE_CACHE_NAMESPACE)
    : undefined;
  // Base, colour-free bitmap from the shared path-keyed cache.
  const base = await getCachedBitmapByPath(imagePath, mimeType, fetchImage, sourceBitmapOpts);
  // No duotone → return the base directly (no second-layer entry). A `null`
  // (unsupported metafile) propagates unchanged.
  if (!duotone || !base) return base;
  // Strict and compatibility callers must not share a derived cache entry: a
  // compatibility pass-through must never make a later strict lookup succeed.
  const resolvedBaseKey = await resolvedCachedBitmapVariantKey(
    imagePath,
    mimeType,
    fetchImage,
    sourceBitmapOpts,
    epoch,
    base,
  );
  const resizeOptions = decodedBitmapTargetResizeOptions(
    Number(base.width),
    Number(base.height),
    requestedBitmapOpts.targetWidthPx,
    requestedBitmapOpts.targetHeightPx,
  );
  const key = `${duotoneCacheKey(
    resolvedBaseKey,
    duotone,
  )}${resizeOptions ? `|resize-width:${resizeOptions.resizeWidth}` : ''}${failClosedOnDuotoneFailure ? '|strict' : ''}`;
  return getCachedDerivedBitmap(
    DUOTONE_CACHE_NAMESPACE,
    key,
    fetchImage,
    async () => {
      const { w, h } = imageNaturalSize(base);
      if (w <= 0 || h <= 0) {
        return { bitmap: failClosedOnDuotoneFailure ? null : base, owned: false };
      }
      const recoloured = await applyDuotone(base, duotone, {
        width: w,
        height: h,
        offscreenFactory,
        targetWidthPx: requestedBitmapOpts.targetWidthPx,
        targetHeightPx: requestedBitmapOpts.targetHeightPx,
      });
      // `applyDuotone` returns a CanvasImageSource; when the pixel pipeline ran
      // it is a fresh ImageBitmap, otherwise it is the unchanged current
      // source. Strict callers fail closed. Compatibility callers still bake a
      // display-sized copy so an unavailable effect surface cannot silently
      // defeat the caller's bounded-resolution request.
      if (recoloured === base) {
        if (failClosedOnDuotoneFailure) return { bitmap: null, owned: false };
        if (!resizeOptions) return { bitmap: base, owned: false };
        if (typeof createImageBitmap === 'undefined') {
          throw new Error('createImageBitmap is unavailable for duotone fallback resampling');
        }
        const resized = await createImageBitmap(base, resizeOptions);
        return { bitmap: resized, owned: resized !== base };
      }
      const bitmap = recoloured as ImageBitmap;
      return { bitmap, owned: bitmap !== base };
    },
    epoch,
  );
}

/** Drop only the duotone namespace. Full document teardown calls the shared
 * decoded-owner drop once. */
export function dropDuotoneBitmapCache(fetchImage: FetchImage): void {
  dropCachedDerivedBitmapNamespace(fetchImage, DUOTONE_CACHE_NAMESPACE);
}
