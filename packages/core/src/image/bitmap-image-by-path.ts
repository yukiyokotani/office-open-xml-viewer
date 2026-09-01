// Decoded-bitmap cache for raster / metafile blips shared by the docx, pptx and
// xlsx renderers, for the lazy byte-on-demand image pipeline. The sibling of
// `svg-image-by-path.ts`: same per-document (per-`fetchImage`) shape, keyed by
// path plus any raster resolution band, but the drawable here is an
// `ImageBitmap` (GPU-backed) decoded via the
// shared `decodeRasterOrMetafile` rather than an `<img>`. Decoding an inlined
// base64 image to an ImageBitmap is expensive, and the same picture is otherwise
// re-decoded on every render (each scroll / resize / interaction / page revisit);
// this caches the decode — the Promise, so concurrent first-renders dedupe — and
// reuses it. Bounded LRU per document.
//
// The cache is keyed FIRST by the document's `fetchImage` closure, then by the
// embedded zip path. Different documents reuse the same internal paths
// (ppt/media/image1.png, word/media/image1.png, xl/media/image1.png), so a
// module-global path→bitmap map would paint document A's image for document B's
// identically-named blip when both are open on the main thread. Keying by
// `fetchImage` (one stable closure per document/deck/workbook instance) scopes
// the cache per byte source; the outer WeakMap also lets a document's whole
// bitmap cache be reclaimed with the document.
//
// GPU-lifecycle discipline (learned in PR #658): three `.catch` sites keep a
// failed or evicted decode from leaking or crashing —
//   1. the RECORD side-chain (`void promise.then(...).catch(() => {})`) that
//      copies the resolved bitmap onto the entry for the synchronous peek path
//      must swallow a decode rejection, or an empty/undecodable blob surfaces as
//      an UNHANDLED rejection (the real caller still sees it via the returned
//      promise);
//   2. the SELF-EVICT (`promise.catch(() => cache.delete(...))`) that removes a
//      transiently-failed entry so the next call retries fresh rather than
//      serving a poisoned rejection;
//   3. the EVICTION close (`oldest?.promise.then((b) => b?.close()).catch(...)`)
//      that releases the GPU backing of an LRU-evicted bitmap THROUGH its promise
//      — never by holding a raw bitmap reference — so a still-in-flight decode is
//      closed only once it resolves, and a draw already in progress is never
//      handed a closed bitmap.
// `dropDecodedBitmapCache` closes every base and derived surface the same way
// (through the promise), for prompt release on the owning viewer's `destroy()`.

import { decodeRasterOrMetafile } from './wmf';
import { closeImageBitmapIfSupported } from './image-bitmap-lifecycle.js';
import { withDecodedImageSlot } from './decode-gate.js';
import {
  MAX_DECODED_IMAGE_BYTES,
  OoxmlDecodedImageLimitError,
} from './pixel-budget.js';
import type { TiffRenderer } from './tiff-contract.js';

type FetchImage = (path: string, mime: string) => Promise<Blob>;
export type DecodedBitmapCacheOwner = object;

const IMAGE_BITMAP_CACHE_MAX = 256;

// Each entry pairs the in-flight/settled decode promise with its resolved bitmap.
// `bitmap` is populated once the promise resolves (see getCachedBitmapByPath),
// giving the synchronous draw sites (picture bullets, §21.1.2.4.2) a settled
// value to read via peekCachedBitmapByPath without awaiting — no separate
// parallel cache to keep in sync, so eviction/teardown only ever drop the whole
// entry.
//
// The decode can resolve to `null` for a metafile we can't rasterize (a true
// EMF, or a WMF with no drawable geometry); the null is cached (avoiding a
// re-fetch+re-sniff every frame) and the draw sites skip a null bitmap.
type BitmapCacheEntry = {
  promise: Promise<ImageBitmap | null>;
  /** Resolves only the surface this entry owns; pass-through results resolve null. */
  ownedPromise: Promise<ImageBitmap | null>;
  bitmap?: ImageBitmap | null;
  weight: number;
};

interface BitmapCacheState {
  readonly entries: Map<string, BitmapCacheEntry>;
  retainedBytes: number;
}

const bitmapCacheByFetch = new WeakMap<DecodedBitmapCacheOwner, BitmapCacheState>();

function bitmapCacheFor(owner: DecodedBitmapCacheOwner): BitmapCacheState {
  let state = bitmapCacheByFetch.get(owner);
  if (!state) {
    state = { entries: new Map(), retainedBytes: 0 };
    bitmapCacheByFetch.set(owner, state);
  }
  return state;
}

// ── Render-pass leases ────────────────────────────────────────────────────────
// The renderers resolve every image a page/slide/sheet references through this
// cache and then DRAW from those references — either synchronously from a
// non-owning lookup map (docx `preloadImages`, xlsx `prefetchImages`) or right
// after a per-element await (pptx). The LRU cap, however, is oblivious to that
// pass: resolving MORE THAN the cap's worth of images in one pass (or a
// concurrent pass on the same document) evicts — and GPU-closes — bitmaps the
// in-flight pass still holds, so the draw would paint a closed bitmap.
//
// A lease makes the pass's liveness need explicit and structural: while at least
// one lease is active for a document's `fetchImage`, any close this module (or a
// sibling per-document cache, via {@link deferBitmapCloseWhileLeased}) would
// perform — LRU eviction or an explicit drop — is DEFERRED and executed when the
// last lease is released. Eviction still removes the cache ENTRY immediately
// (the cache stays size-bounded and the next resolve re-decodes); only the GPU
// release is deferred, so every reference a leased pass obtained stays drawable
// for the duration of the pass. Callers MUST release in a `finally` — an
// unreleased lease keeps its deferred bitmaps alive until the document itself is
// reclaimed. The SVG cache needs no lease: its eviction revokes an object URL,
// which does not invalidate an already-decoded HTMLImageElement.
interface BitmapCacheLeaseState {
  /** Active (unreleased) leases for this document. */
  count: number;
  /** Closes deferred while leased; executed at the last release. */
  deferred: Array<Promise<ImageBitmap | null>>;
  activeBytes: number;
  activeBitmaps: WeakSet<ImageBitmap>;
}

const leasesByFetch = new WeakMap<DecodedBitmapCacheOwner, BitmapCacheLeaseState>();

// Every GPU close this module (and the sibling per-document caches routing
// through {@link deferBitmapCloseWhileLeased}) performs is funneled through
// here. The WeakSet deduplicates closes PER BITMAP: two cache layers can
// resolve to the same bitmap (a second-layer pass-through entry still in its
// in-flight window when both caches are dropped resolves to the base bitmap
// the base cache also closes), and the dedup removes any reliance on
// `ImageBitmap.close()` idempotence across engines.
const closedBitmaps = new WeakSet<ImageBitmap>();

function closeBitmapOnce(bmp: ImageBitmap | null | undefined): void {
  if (!bmp || closedBitmaps.has(bmp)) return;
  closedBitmaps.add(bmp);
  closeImageBitmapIfSupported(bmp);
}

/** Release a document-owned decoded surface once. Browser ImageBitmap exposes
 * `close()`; Node canvas backends may rely on native GC and omit it. */
export function releaseOwnedBitmap(bitmap: ImageBitmap | null | undefined): void {
  closeBitmapOnce(bitmap);
}

/**
 * Hold every decoded bitmap of one document (keyed by `fetchImage`) alive for
 * the duration of a render pass: while the returned release function has not
 * been called, LRU evictions and cache drops defer their GPU `.close()` until
 * the last outstanding lease is released. Acquire before resolving the pass's
 * images and release in a `finally` after the draw that uses them. Leases nest
 * (concurrent passes over the same document each take one); the release
 * function is idempotent.
 */
export function acquireBitmapCacheLease(owner: DecodedBitmapCacheOwner): () => void {
  let state = leasesByFetch.get(owner);
  if (!state) {
    state = { count: 0, deferred: [], activeBytes: 0, activeBitmaps: new WeakSet() };
    leasesByFetch.set(owner, state);
  }
  const s = state;
  s.count++;
  let released = false;
  return () => {
    if (released) return;
    released = true;
    s.count--;
    if (s.count > 0) return;
    // Last lease out: run the deferred closes, through each promise (never a raw
    // bitmap reference) so a still-in-flight decode closes only once it resolves.
    for (const p of s.deferred) p.then((b) => closeBitmapOnce(b)).catch(() => {});
    s.deferred = [];
    s.activeBytes = 0;
    s.activeBitmaps = new WeakSet();
    leasesByFetch.delete(owner);
  };
}

function bitmapWeight(bitmap: ImageBitmap | null): number {
  if (!bitmap) return 0;
  const width = Number(bitmap.width);
  const height = Number(bitmap.height);
  return Number.isSafeInteger(width) && width > 0
    && Number.isSafeInteger(height) && height > 0
    ? width * height * 4
    : 0;
}

function registerActiveBitmap(owner: DecodedBitmapCacheOwner, bitmap: ImageBitmap | null): void {
  if (!bitmap) return;
  const lease = leasesByFetch.get(owner);
  if (!lease || lease.count === 0 || lease.activeBitmaps.has(bitmap)) return;
  const observed = lease.activeBytes + bitmapWeight(bitmap);
  if (observed > MAX_DECODED_IMAGE_BYTES) {
    throw new OoxmlDecodedImageLimitError(
      'active-decoded-bytes',
      MAX_DECODED_IMAGE_BYTES,
      observed,
    );
  }
  lease.activeBitmaps.add(bitmap);
  lease.activeBytes = observed;
}

function evictOldest(
  owner: DecodedBitmapCacheOwner,
  state: BitmapCacheState,
  protectedKey?: string,
): boolean {
  const candidate = [...state.entries].find(([key]) => key !== protectedKey);
  if (!candidate) return false;
  const [key, entry] = candidate;
  state.entries.delete(key);
  state.retainedBytes -= entry.weight;
  deferBitmapCloseWhileLeased(owner, entry.ownedPromise);
  return true;
}

/**
 * Close a document-owned bitmap through its decode promise — or, when a render
 * pass currently holds a lease on the document (see
 * {@link acquireBitmapCacheLease}), defer the close to the last lease release so
 * the pass never draws a closed bitmap. Shared by this module's LRU eviction and
 * drop paths and by every derived namespace sharing the same owner. Closes are
 * deduplicated per bitmap (see {@link closeBitmapOnce}), so two layers that
 * resolve to the same bitmap close it exactly once.
 *
 * @deprecated Compatibility helper for former sibling caches. New decoded
 * surfaces belong in `getCachedDecodedBitmap` / `getCachedDerivedBitmap`.
 * Scheduled for removal in a future breaking release.
 */
export function deferBitmapCloseWhileLeased(
  owner: DecodedBitmapCacheOwner,
  promise: Promise<ImageBitmap | null>,
): void {
  const lease = leasesByFetch.get(owner);
  if (lease && lease.count > 0) {
    lease.deferred.push(promise);
    return;
  }
  promise.then((b) => closeBitmapOnce(b)).catch(() => {});
}

/** Options for {@link getCachedBitmapByPath}. `widthPt`/`heightPt` size a
 * metafile raster; `targetWidthPx`/`targetHeightPx` select a display-resolution
 * variant for ordinary rasters. `suppressBoundaryFrame` is the docx-only WMF
 * window/device-boundary edge suppression (spec-clean default OFF;
 * pptx/xlsx leave it unset). */
export interface CachedBitmapOptions {
  /** Intended draw width in points; sizes any metafile raster target. */
  widthPt?: number;
  /** Intended draw height in points; see `widthPt`. */
  heightPt?: number;
  /** Enable the docx cosmetic window/device-frame suppression heuristic. Default
   *  false = spec-clean. Only docx opts in. */
  suppressBoundaryFrame?: boolean;
  /** Optional TIFF codec retained by the owning document. */
  tiff?: TiffRenderer;
  /** Desired width of the full raster source in device pixels. Requests are
   * quantized into stable bands so nearby zoom levels reuse one decode. */
  targetWidthPx?: number;
  /** Desired height of the full raster source in device pixels. */
  targetHeightPx?: number;
}

const SMALL_RASTER_TARGET_MAX = 64;
const RASTER_TARGET_QUANTUM = 64;

function rasterTargetBand(value: number | undefined): number | undefined {
  if (typeof value !== 'number' || !Number.isFinite(value) || !(value > 0)) return undefined;
  const rounded = Math.ceil(value);
  if (rounded <= SMALL_RASTER_TARGET_MAX) {
    return 2 ** Math.ceil(Math.log2(rounded));
  }
  return Math.ceil(rounded / RASTER_TARGET_QUANTUM) * RASTER_TARGET_QUANTUM;
}

function normalizedBitmapOptions(opts: CachedBitmapOptions): CachedBitmapOptions {
  return {
    ...opts,
    targetWidthPx: rasterTargetBand(opts.targetWidthPx),
    targetHeightPx: rasterTargetBand(opts.targetHeightPx),
  };
}

/** Stable base key shared by the base and derived decoded-surface caches. */
export function cachedBitmapVariantKey(
  imagePath: string,
  opts: CachedBitmapOptions = {},
): string {
  const width = rasterTargetBand(opts.targetWidthPx);
  const height = rasterTargetBand(opts.targetHeightPx);
  const pathKey = `${imagePath.length}:${imagePath}`;
  return width && height ? `raster:${pathKey}:${width}x${height}` : `native:${pathKey}`;
}

interface ProducedBitmap {
  readonly bitmap: ImageBitmap | null;
  /** False when the result is a borrowed pass-through surface. */
  readonly owned: boolean;
}

function getCachedOwnedBitmap(
  key: string,
  owner: DecodedBitmapCacheOwner,
  produce: () => Promise<ProducedBitmap>,
): Promise<ImageBitmap | null> {
  const state = bitmapCacheFor(owner);
  const cache = state.entries;
  const existing = cache.get(key);
  if (existing) {
    cache.delete(key);
    cache.set(key, existing);
    return existing.promise.then((bitmap) => {
      registerActiveBitmap(owner, bitmap);
      return bitmap;
    });
  }

  const produced = withDecodedImageSlot(owner, produce);
  const ownedPromise = produced.then(({ bitmap, owned }) => (owned ? bitmap : null));
  const promise = produced.then(({ bitmap, owned }) => {
    try {
      registerActiveBitmap(owner, bitmap);
      return bitmap;
    } catch (error) {
      if (owned) closeBitmapOnce(bitmap);
      throw error;
    }
  });
  const entry: BitmapCacheEntry = { promise, ownedPromise, weight: 0 };

  void produced
    .then(({ bitmap, owned }) => {
      if (cache.get(key) !== entry) return;
      entry.bitmap = bitmap;
      if (!owned) {
        // A borrowed base bitmap remains owned by its base entry. Keeping a
        // second entry would outlive base eviction and could serve a closed
        // surface, so only concurrent in-flight callers share it.
        cache.delete(key);
        return;
      }
      entry.weight = bitmapWeight(bitmap);
      state.retainedBytes += entry.weight;
      while (state.retainedBytes > MAX_DECODED_IMAGE_BYTES) {
        if (!evictOldest(owner, state, key)) break;
      }
    })
    .catch(() => {});
  promise.catch(() => {
    if (cache.get(key) !== entry) return;
    cache.delete(key);
    state.retainedBytes -= entry.weight;
    deferBitmapCloseWhileLeased(owner, ownedPromise);
  });
  cache.set(key, entry);
  while (cache.size > IMAGE_BITMAP_CACHE_MAX) {
    if (!evictOldest(owner, state, key)) break;
  }
  return promise;
}

const BASE_CACHE_NAMESPACE = 'base';
const BASE_CACHE_PREFIX = `${BASE_CACHE_NAMESPACE}:`;
const DERIVED_CACHE_PREFIX = 'derived:';

/** General document-owned decoded-surface primitive. Loader choice is separate
 * from the owner token; callers sharing a namespace/key dedupe regardless of
 * whether bytes came from an image or media extraction API. */
export function getCachedDecodedBitmap(
  namespace: string,
  cacheKey: string,
  owner: DecodedBitmapCacheOwner,
  create: () => Promise<{ bitmap: ImageBitmap | null; owned: boolean }>,
): Promise<ImageBitmap | null> {
  return getCachedOwnedBitmap(`${namespace}:${cacheKey}`, owner, create);
}

/**
 * Cache a document-owned derived bitmap under the same weighted LRU,
 * decode-concurrency gate, render-pass lease and aggregate byte ceiling as its
 * source bitmap. `create` may return the source bitmap unchanged; such a
 * borrowed result is shared only while in flight and is never closed here.
 */
export function getCachedDerivedBitmap(
  namespace: string,
  cacheKey: string,
  owner: DecodedBitmapCacheOwner,
  create: () => Promise<{ bitmap: ImageBitmap | null; owned: boolean }>,
): Promise<ImageBitmap | null> {
  return getCachedDecodedBitmap(
    `${DERIVED_CACHE_PREFIX}${namespace}`,
    cacheKey,
    owner,
    create,
  );
}

/** Drop one derived-surface namespace without disturbing base blips or sibling
 * transformations. Used by format teardown methods retained for compatibility. */
export function dropCachedDerivedBitmapNamespace(
  owner: DecodedBitmapCacheOwner,
  namespace: string,
): void {
  const state = bitmapCacheByFetch.get(owner);
  if (!state) return;
  const prefix = `${DERIVED_CACHE_PREFIX}${namespace}:`;
  for (const [key, entry] of state.entries) {
    if (!key.startsWith(prefix)) continue;
    state.entries.delete(key);
    state.retainedBytes -= entry.weight;
    deferBitmapCloseWhileLeased(owner, entry.ownedPromise);
  }
}

/**
 * Decode a raster-or-metafile blip at `imagePath` to an `ImageBitmap`, cached per
 * document (keyed by `fetchImage`) then by path. The bytes are fetched lazily
 * through `fetchImage(imagePath, mimeType)` (twin of the audio/video `fetchMedia`
 * path) rather than `fetch`-ing an inlined data URL. The returned bitmap is
 * drawable with `ctx.drawImage`.
 *
 * Decoding goes through core's {@link decodeRasterOrMetafile}, which content-
 * sniffs the bytes: a WMF (which `createImageBitmap` can't decode) is rasterized
 * by the shared minimal player at a size derived from `widthPt`/`heightPt`; a
 * true EMF (or a WMF with no geometry) resolves to `null` so the draw site skips
 * the picture instead of crashing — the `null` is cached too, so the draw skips
 * it without a re-fetch+re-sniff every frame.
 *
 * The cache is bounded by count and decoded RGBA weight. Decodes are also
 * concurrency-limited, and one render-pass lease cannot accumulate more than
 * the shared active decoded-byte ceiling. Quota crossings reject with
 * `OoxmlDecodedImageLimitError`; they are never converted to a silent omission.
 */
export function getCachedBitmapByPath(
  imagePath: string,
  mimeType: string,
  fetchImage: FetchImage,
  opts: CachedBitmapOptions = {},
): Promise<ImageBitmap | null> {
  const normalized = normalizedBitmapOptions(opts);
  const {
    widthPt = 0,
    heightPt = 0,
    suppressBoundaryFrame = false,
    tiff,
    targetWidthPx,
    targetHeightPx,
  } = normalized;
  return getCachedDecodedBitmap(
    BASE_CACHE_NAMESPACE,
    cachedBitmapVariantKey(imagePath, normalized),
    fetchImage,
    async () => {
      const blob = await fetchImage(imagePath, mimeType);
      const bitmap = await decodeRasterOrMetafile(blob, {
        widthPt,
        heightPt,
        suppressBoundaryFrame,
        tiff,
        targetWidthPx,
        targetHeightPx,
      });
      return { bitmap, owned: true };
    },
  );
}

/**
 * Synchronously return a blip's decoded bitmap if its decode has already
 * resolved (warmed by {@link getCachedBitmapByPath}), else `undefined`. Used by
 * the synchronous text-body draw to paint picture bullets (`<a:buBlip>`,
 * §21.1.2.4.2) without awaiting. A still-loading image has no `bitmap` on its
 * entry yet, so it's simply skipped.
 */
export function peekCachedBitmapByPath(
  imagePath: string,
  fetchImage: FetchImage,
): ImageBitmap | null | undefined {
  return bitmapCacheByFetch.get(fetchImage)?.entries
    .get(`${BASE_CACHE_PREFIX}${cachedBitmapVariantKey(imagePath)}`)?.bitmap;
}

/**
 * Close every base and derived decoded bitmap for one document owner and forget
 * that owner. Call from viewer/session teardown so GPU-backed ImageBitmaps are
 * released promptly rather than waiting for GC. When a render pass holds a lease (see
 * {@link acquireBitmapCacheLease} — e.g. a destroy or re-parse racing an
 * in-flight render), the cache is forgotten immediately but the GPU closes are
 * deferred to the last lease release, so the pass never draws a closed bitmap.
 */
export function dropDecodedBitmapCache(owner: DecodedBitmapCacheOwner): void {
  const state = bitmapCacheByFetch.get(owner);
  if (!state) return;
  for (const entry of state.entries.values()) {
    deferBitmapCloseWhileLeased(owner, entry.ownedPromise);
  }
  state.entries.clear();
  state.retainedBytes = 0;
  bitmapCacheByFetch.delete(owner);
}

/** @deprecated Use {@link dropDecodedBitmapCache}; it also owns derived
 * surfaces. Scheduled for removal in a future breaking release. */
export const dropBitmapCacheByPath = dropDecodedBitmapCache;
