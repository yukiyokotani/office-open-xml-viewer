// Decoded-bitmap cache for raster / metafile blips shared by the docx, pptx and
// xlsx renderers, for the lazy byte-on-demand image pipeline. The sibling of
// `svg-image-by-path.ts`: same per-document (per-`fetchImage`) shape, keyed by
// path plus any required raster resolution variant, but the drawable here is an
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

import {
  decodeRasterOrMetafile,
  decodeRasterOrMetafileWithInspection,
} from './raster-or-metafile.js';
import { inspectRasterBlob, type RasterBlobInspection } from './raster-blob-inspection.js';
import { rasterExceedsBudget } from './raster-dimensions.js';
import { decodedBitmapRetainedTarget } from './raster-target.js';
import { closeImageBitmapIfSupported } from './image-bitmap-lifecycle.js';
import { withDecodedImageSlot } from './decode-gate.js';
import {
  MAX_DECODED_IMAGE_BYTES,
  MAX_RASTER_DIMENSION,
  OoxmlDecodedImageLimitError,
} from './pixel-budget.js';
import {
  normalizeImageResourceOptions,
  type ImageResourceOptions,
} from './adaptive-image-budget.js';
import type { TiffRenderer } from './tiff-contract.js';
import { wmfRasterTarget } from './wmf.js';
import type { SvgBlobDecoder } from '../worker/svg-decode-bridge.js';

type FetchImage = (path: string, mime: string) => Promise<Blob>;
export type DecodedBitmapCacheOwner = object;

function isDecodedBitmapCacheOwner(value: unknown): value is DecodedBitmapCacheOwner {
  return (typeof value === 'object' && value !== null) || typeof value === 'function';
}

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
  retainedLimit: number;
}

const bitmapCacheByFetch = new WeakMap<DecodedBitmapCacheOwner, BitmapCacheState>();

interface RasterSourceProfile {
  readonly inspection: RasterBlobInspection;
  initialBlob?: Blob;
}

const rasterProfilesByOwner = new WeakMap<DecodedBitmapCacheOwner, Map<string, Promise<RasterSourceProfile>>>();
const cacheGenerationByOwner = new WeakMap<DecodedBitmapCacheOwner, number>();
const derivedNamespaceGenerationByOwner = new WeakMap<
  DecodedBitmapCacheOwner,
  Map<string, number>
>();

declare const decodedBitmapCacheEpochBrand: unique symbol;

/**
 * Internal operation token spanning the awaits between resolving a base blip
 * and inserting its derived surface. A full owner drop or a targeted namespace
 * drop invalidates the token, so a pre-drop continuation cannot recreate cache
 * state after teardown.
 *
 * @internal
 */
export interface DecodedBitmapCacheEpoch {
  readonly owner: DecodedBitmapCacheOwner;
  readonly ownerGeneration: number;
  readonly namespace: string;
  readonly namespaceGeneration: number;
  readonly [decodedBitmapCacheEpochBrand]: true;
}

function cacheGeneration(owner: DecodedBitmapCacheOwner): number {
  return cacheGenerationByOwner.get(owner) ?? 0;
}

function derivedNamespaceGeneration(
  owner: DecodedBitmapCacheOwner,
  namespace: string,
): number {
  return derivedNamespaceGenerationByOwner.get(owner)?.get(namespace) ?? 0;
}

function advanceDerivedNamespaceGeneration(
  owner: DecodedBitmapCacheOwner,
  namespace: string,
): void {
  let generations = derivedNamespaceGenerationByOwner.get(owner);
  if (!generations) {
    generations = new Map();
    derivedNamespaceGenerationByOwner.set(owner, generations);
  }
  generations.set(namespace, derivedNamespaceGeneration(owner, namespace) + 1);
}

function assertCurrentCacheGeneration(
  owner: DecodedBitmapCacheOwner,
  expected: number,
): void {
  if (cacheGeneration(owner) !== expected) {
    throw new Error('Decoded bitmap cache was dropped during raster inspection');
  }
}

/** @internal Capture before the first await of a base-to-derived operation. */
export function captureDecodedBitmapCacheEpoch(
  owner: DecodedBitmapCacheOwner,
  namespace: string,
): DecodedBitmapCacheEpoch {
  return {
    owner,
    ownerGeneration: cacheGeneration(owner),
    namespace,
    namespaceGeneration: derivedNamespaceGeneration(owner, namespace),
  } as DecodedBitmapCacheEpoch;
}

function assertCurrentOwnerEpoch(
  owner: DecodedBitmapCacheOwner,
  epoch: DecodedBitmapCacheEpoch,
): void {
  if (epoch.owner !== owner || cacheGeneration(owner) !== epoch.ownerGeneration) {
    throw new Error('Decoded bitmap cache was dropped during a derived image operation');
  }
}

function assertCurrentDerivedEpoch(
  owner: DecodedBitmapCacheOwner,
  namespace: string,
  epoch: DecodedBitmapCacheEpoch,
): void {
  assertCurrentOwnerEpoch(owner, epoch);
  if (epoch.namespace !== namespace
    || derivedNamespaceGeneration(owner, namespace) !== epoch.namespaceGeneration) {
    throw new Error('Decoded bitmap cache namespace was dropped during a derived image operation');
  }
}

function releaseInitialRasterBlobs(owner: DecodedBitmapCacheOwner): void {
  const profiles = rasterProfilesByOwner.get(owner);
  if (!profiles) return;
  for (const promise of profiles.values()) {
    void promise.then((profile) => { profile.initialBlob = undefined; }).catch(() => {});
  }
}

function rasterProfileFor(
  owner: DecodedBitmapCacheOwner,
  imagePath: string,
  mimeType: string,
  fetchImage: FetchImage,
): Promise<RasterSourceProfile> {
  let profiles = rasterProfilesByOwner.get(owner);
  if (!profiles) {
    profiles = new Map();
    rasterProfilesByOwner.set(owner, profiles);
  }
  const existing = profiles.get(imagePath);
  if (existing) return existing;
  // Source inspection fetches and retains the initial blob before the decoder
  // cache entry exists. Charge that work to the same per-document gate as SVG
  // loading and bitmap decode so metafile classification cannot bypass the
  // shared concurrency ceiling.
  const promise = withDecodedImageSlot(owner, async () => {
    const blob = await fetchImage(imagePath, mimeType);
    return { inspection: await inspectRasterBlob(blob), initialBlob: blob };
  })
    .catch((error) => {
      profiles?.delete(imagePath);
      throw error;
    });
  profiles.set(imagePath, promise);
  return promise;
}

/** Inspect a path through the same per-document profile cache used by bitmap
 * decoding. Renderers use this only when authored geometry depends on the
 * source grid (notably DrawingML tile fills); the retained Blob is consumed by
 * the subsequent decode, so inspection does not introduce a second fetch. */
export async function inspectCachedRasterSource(
  imagePath: string,
  mimeType: string,
  fetchImage: FetchImage,
  owner: DecodedBitmapCacheOwner = fetchImage,
): Promise<RasterBlobInspection> {
  return (await rasterProfileFor(owner, imagePath, mimeType, fetchImage)).inspection;
}

function bitmapCacheFor(owner: DecodedBitmapCacheOwner): BitmapCacheState {
  let state = bitmapCacheByFetch.get(owner);
  if (!state) {
    state = {
      entries: new Map(),
      retainedBytes: 0,
      retainedLimit: MAX_DECODED_IMAGE_BYTES,
    };
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
  readonly limits: Map<symbol, number>;
  activeLimit: number;
}

const leasesByFetch = new WeakMap<DecodedBitmapCacheOwner, BitmapCacheLeaseState>();

interface BitmapRenderQueueState {
  tail: Promise<void>;
  pending: number;
}

const renderQueuesByOwner = new WeakMap<DecodedBitmapCacheOwner, BitmapRenderQueueState>();

/**
 * Run one image-bearing paint at a time for a document owner. A render pass can
 * resolve many bitmaps concurrently internally, but overlapping page/slide/
 * sheet passes cannot each consume the full document budget at once. This is
 * the same serialized-admission pattern used by native viewers: memory pressure
 * controls overlap while decode parallelism remains inside the admitted job.
 */
export async function withBitmapCacheLease<T>(
  owner: DecodedBitmapCacheOwner,
  options: ImageResourceOptions | undefined,
  paint: () => Promise<T>,
): Promise<T> {
  let queue = renderQueuesByOwner.get(owner);
  if (!queue) {
    let open!: () => void;
    const gate = new Promise<void>((resolve) => { open = resolve; });
    queue = { tail: gate, pending: 1 };
    renderQueuesByOwner.set(owner, queue);
    let releaseLease: (() => void) | undefined;
    try {
      releaseLease = acquireBitmapCacheLease(owner, options);
      return await paint();
    } finally {
      releaseInitialRasterBlobs(owner);
      releaseLease?.();
      open();
      queue.pending--;
      if (queue.pending === 0) renderQueuesByOwner.delete(owner);
    }
  }
  const state = queue;
  const previous = state.tail.catch(() => {});
  let open!: () => void;
  const gate = new Promise<void>((resolve) => { open = resolve; });
  state.pending++;
  state.tail = previous.then(() => gate);
  await previous;
  let releaseLease: (() => void) | undefined;
  try {
    releaseLease = acquireBitmapCacheLease(owner, options);
    return await paint();
  } finally {
    releaseInitialRasterBlobs(owner);
    releaseLease?.();
    open();
    state.pending--;
    if (state.pending === 0) renderQueuesByOwner.delete(owner);
  }
}

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

/** Release a document-owned decoded surface once. Cleanup is best-effort and
 * never replaces the operation failure that led here. Browser ImageBitmap
 * exposes `close()`; Node canvas backends may rely on native GC and omit it. */
export function releaseOwnedBitmap(bitmap: ImageBitmap | null | undefined): void {
  try {
    closeBitmapOnce(bitmap);
  } catch {
    // Resource release must not mask a post, callback, or draw failure.
  }
}

/**
 * Hold every decoded bitmap of one document owner alive for
 * the duration of a render pass: while the returned release function has not
 * been called, LRU evictions and cache drops defer their GPU `.close()` until
 * the last outstanding lease is released. Acquire before resolving the pass's
 * images and release in a `finally` after the draw that uses them. Leases nest
 * (concurrent passes over the same document each take one); the release
 * function is idempotent.
 */
export function acquireBitmapCacheLease(
  owner: DecodedBitmapCacheOwner,
  options?: ImageResourceOptions,
): () => void {
  const policy = normalizeImageResourceOptions(options);
  let state = leasesByFetch.get(owner);
  if (!state) {
    state = {
      count: 0,
      deferred: [],
      activeBytes: 0,
      activeBitmaps: new WeakSet(),
      limits: new Map(),
      activeLimit: policy.decodedByteBudget,
    };
    leasesByFetch.set(owner, state);
  }
  const s = state;
  const token = Symbol('decoded-image-lease');
  s.limits.set(token, policy.decodedByteBudget);
  s.activeLimit = Math.min(...s.limits.values());
  const cache = bitmapCacheFor(owner);
  cache.retainedLimit = s.activeLimit;
  while (cache.retainedBytes > cache.retainedLimit) {
    if (!evictOldest(owner, cache)) break;
  }
  s.count++;
  let released = false;
  return () => {
    if (released) return;
    released = true;
    s.count--;
    s.limits.delete(token);
    if (s.count > 0) {
      s.activeLimit = Math.min(...s.limits.values());
      const remainingCache = bitmapCacheByFetch.get(owner);
      if (remainingCache) remainingCache.retainedLimit = s.activeLimit;
      return;
    }
    // Last lease out: run the deferred closes, through each promise (never a raw
    // bitmap reference) so a still-in-flight decode closes only once it resolves.
    for (const p of s.deferred) p.then((b) => closeBitmapOnce(b)).catch(() => {});
    s.deferred = [];
    s.activeBytes = 0;
    s.activeBitmaps = new WeakSet();
    leasesByFetch.delete(owner);
    const remainingCache = bitmapCacheByFetch.get(owner);
    if (remainingCache) remainingCache.retainedLimit = MAX_DECODED_IMAGE_BYTES;
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
  if (observed > lease.activeLimit) {
    throw new OoxmlDecodedImageLimitError(
      'active-decoded-bytes',
      lease.activeLimit,
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
  /** Desired width of the full raster source in device pixels. */
  targetWidthPx?: number;
  /** Desired height of the full raster source in device pixels. */
  targetHeightPx?: number;
  /** Retained base-surface pixel ceiling for multi-surface effect pipelines. */
  maxRetainedPixels?: number;
  /** Worker-only SVG decoder. Window renderers use HTMLImageElement instead. */
  svgDecoder?: SvgBlobDecoder;
}

function normalizedRasterTarget(value: number | undefined): number | undefined {
  if (typeof value !== 'number' || !Number.isFinite(value) || !(value > 0)) return undefined;
  return Math.ceil(value);
}

function normalizedBitmapOptions(opts: CachedBitmapOptions): CachedBitmapOptions {
  return {
    ...opts,
    targetWidthPx: normalizedRasterTarget(opts.targetWidthPx),
    targetHeightPx: normalizedRasterTarget(opts.targetHeightPx),
  };
}

function retainedPixelLimit(opts: CachedBitmapOptions): number {
  const requested = opts.maxRetainedPixels;
  return typeof requested === 'number' && Number.isSafeInteger(requested) && requested > 0
    ? Math.min(requested, MAX_DECODED_IMAGE_BYTES / 4)
    : MAX_DECODED_IMAGE_BYTES / 4;
}

function hasRestrictedPixelLimit(opts: CachedBitmapOptions): boolean {
  return retainedPixelLimit(opts) < MAX_DECODED_IMAGE_BYTES / 4;
}

/** Re-check a cache hit against the current caller's retained-surface budget.
 * A native bitmap may have been admitted by an earlier, more permissive call;
 * reject that borrowed hit without closing it because the cache still owns it. */
function enforceCachedRetainedBudget(
  bitmap: ImageBitmap | null,
  opts: CachedBitmapOptions,
): ImageBitmap | null {
  if (!bitmap) return null;
  const pixelLimit = retainedPixelLimit(opts);
  const width = Number(bitmap.width);
  const height = Number(bitmap.height);
  const observedDimension = Math.max(width, height);
  if (!Number.isFinite(observedDimension) || observedDimension > MAX_RASTER_DIMENSION) {
    throw new OoxmlDecodedImageLimitError(
      'image-dimension',
      MAX_RASTER_DIMENSION,
      Number.isFinite(observedDimension) ? observedDimension : Number.MAX_SAFE_INTEGER,
    );
  }
  const observedPixels = width * height;
  if (!(width > 0) || !(height > 0) || observedPixels > pixelLimit) {
    throw new OoxmlDecodedImageLimitError(
      'image-pixels',
      pixelLimit,
      Number.isSafeInteger(observedPixels) && observedPixels >= 0
        ? observedPixels
        : Number.MAX_SAFE_INTEGER,
    );
  }
  return bitmap;
}

/** Stable base key shared by the base and derived decoded-surface caches. */
export function cachedBitmapVariantKey(
  imagePath: string,
  opts: CachedBitmapOptions = {},
): string {
  const width = normalizedRasterTarget(opts.targetWidthPx);
  const height = normalizedRasterTarget(opts.targetHeightPx);
  const pathKey = `${imagePath.length}:${imagePath}`;
  const pixelLimit = retainedPixelLimit(opts);
  return width || height
    ? `raster:${pathKey}:${width ?? 0}x${height ?? 0}:p${pixelLimit}`
    : `native:${pathKey}`;
}

interface MetafileVariant {
  readonly width: number;
  readonly height: number;
  readonly suppressBoundaryFrame: boolean;
  readonly pixelLimit: number;
}

function profileIsMetafile(profile: RasterSourceProfile): boolean {
  return profile.inspection.format === 'wmf' || profile.inspection.format === 'emf';
}

function metafileVariant(opts: CachedBitmapOptions): MetafileVariant {
  const target = wmfRasterTarget(opts.widthPt ?? 0, opts.heightPt ?? 0);
  return {
    width: target.w,
    height: target.h,
    suppressBoundaryFrame: opts.suppressBoundaryFrame === true,
    pixelLimit: retainedPixelLimit(opts),
  };
}

function metafileVariantPrefix(imagePath: string, variant: MetafileVariant): string {
  const pathKey = `${imagePath.length}:${imagePath}`;
  return `metafile:${pathKey}:s${variant.suppressBoundaryFrame ? 1 : 0}:p${variant.pixelLimit}:`;
}

function metafileVariantKey(
  imagePath: string,
  opts: CachedBitmapOptions,
  dimensions?: Readonly<{ width: number; height: number }>,
): string {
  const variant = metafileVariant(opts);
  const width = normalizedRasterTarget(dimensions?.width) ?? variant.width;
  const height = normalizedRasterTarget(dimensions?.height) ?? variant.height;
  return `${metafileVariantPrefix(imagePath, variant)}${width}x${height}`;
}

function reusableMetafileVariantKey(
  owner: DecodedBitmapCacheOwner,
  imagePath: string,
  opts: CachedBitmapOptions,
): string | undefined {
  const requested = metafileVariant(opts);
  const prefix = `${BASE_CACHE_PREFIX}${metafileVariantPrefix(imagePath, requested)}`;
  let best: { key: string; pixels: number } | undefined;
  for (const key of bitmapCacheByFetch.get(owner)?.entries.keys() ?? []) {
    if (!key.startsWith(prefix)) continue;
    const match = /^(\d+)x(\d+)$/.exec(key.slice(prefix.length));
    if (!match) continue;
    const width = Number(match[1]);
    const height = Number(match[2]);
    if (width < requested.width || height < requested.height) continue;
    const pixels = width * height;
    if (!best || pixels < best.pixels) {
      best = { key: key.slice(BASE_CACHE_PREFIX.length), pixels };
    }
  }
  return best?.key;
}

function reusableResolutionVariantKey(
  owner: DecodedBitmapCacheOwner,
  imagePath: string,
  opts: CachedBitmapOptions,
  sourceDimensions?: Readonly<{ width: number; height: number }> | null,
): string | undefined {
  const requestedWidth = normalizedRasterTarget(opts.targetWidthPx);
  const requestedHeight = normalizedRasterTarget(opts.targetHeightPx);
  if (!requestedWidth && !requestedHeight) return undefined;
  const pixelLimit = retainedPixelLimit(opts);
  const pathKey = `${imagePath.length}:${imagePath}`;
  const prefix = `${BASE_CACHE_PREFIX}raster:${pathKey}:`;
  let best: { key: string; pixels: number } | undefined;
  for (const [key, entry] of bitmapCacheByFetch.get(owner)?.entries ?? []) {
    if (!key.startsWith(prefix)) continue;
    const match = /^(\d+)x(\d+):p(\d+)$/.exec(key.slice(prefix.length));
    if (!match) continue;
    const width = Number(match[1]);
    const height = Number(match[2]);
    if (Number(match[3]) !== pixelLimit
      || width < (requestedWidth ?? 0)
      || height < (requestedHeight ?? 0)) continue;
    const actualWidth = Number(entry.bitmap?.width);
    const actualHeight = Number(entry.bitmap?.height);
    const hasActualGrid = Number.isFinite(actualWidth) && actualWidth > 0
      && Number.isFinite(actualHeight) && actualHeight > 0;
    const retained = sourceDimensions
      ? decodedBitmapRetainedTarget(sourceDimensions, width, height)
        ?? sourceDimensions
      : undefined;
    const pixels = hasActualGrid
      ? actualWidth * actualHeight
      : retained
        ? retained.width * retained.height
        : width * height;
    if (!best || pixels < best.pixels) best = { key: key.slice(BASE_CACHE_PREFIX.length), pixels };
  }
  return best?.key;
}

function profileNeedsResolutionVariant(
  profile: RasterSourceProfile,
  opts: CachedBitmapOptions = {},
): boolean {
  const pixelLimit = retainedPixelLimit(opts);
  const dimensions = profile.inspection.dimensions;
  if (!dimensions) return false;
  const sourcePixels = dimensions.width * dimensions.height;
  // A render plan's explicit pixel limit is the source's share of the complete
  // paint budget. Preserve the native surface whenever it fits that share,
  // even if its aspect differs from the geometry-derived fallback grid.
  // Display-mode plans set the limit to the display grid itself, so a larger
  // source still takes the bounded variant.
  if (opts.maxRetainedPixels !== undefined
    && !rasterExceedsBudget(dimensions)
    && sourcePixels <= pixelLimit) return false;
  const targetWidth = normalizedRasterTarget(opts.targetWidthPx);
  const targetHeight = normalizedRasterTarget(opts.targetHeightPx);
  const hasSmallerDisplayTarget = profile.inspection.format === 'tiff'
    ? (targetWidth !== undefined || targetHeight !== undefined)
      && (targetWidth === undefined || targetWidth < dimensions.width)
      && (targetHeight === undefined || targetHeight < dimensions.height)
    : targetWidth !== undefined
      && targetHeight !== undefined
      && (targetWidth < dimensions.width || targetHeight < dimensions.height);
  return hasSmallerDisplayTarget
    || rasterExceedsBudget(dimensions)
    || sourcePixels > pixelLimit;
}

/** Resolve the actual base-cache key after source inspection. A sufficient
 * downsample target receives a display-resolution variant (TIFF also supports
 * a single-axis request); targets that require the source grid share one
 * path-native entry. When `resolvedBitmap` is supplied, key the exact retained
 * surface so concurrent cache evolution cannot relabel a derived effect. */
export async function resolvedCachedBitmapVariantKey(
  imagePath: string,
  mimeType: string,
  fetchImage: FetchImage,
  opts: CachedBitmapOptions = {},
  epoch?: DecodedBitmapCacheEpoch,
  resolvedBitmap?: ImageBitmap,
): Promise<string> {
  if (epoch) assertCurrentOwnerEpoch(fetchImage, epoch);
  const normalized = normalizedBitmapOptions(opts);
  const generation = epoch?.ownerGeneration ?? cacheGeneration(fetchImage);
  const profile = await rasterProfileFor(fetchImage, imagePath, mimeType, fetchImage);
  if (epoch) assertCurrentOwnerEpoch(fetchImage, epoch);
  else assertCurrentCacheGeneration(fetchImage, generation);
  const metafile = profileIsMetafile(profile);
  if (resolvedBitmap) {
    // Derived transforms must follow the surface actually returned, not a
    // reusable-cache choice that can change while the caller awaits. The
    // retained grid is a stable content identity for one source path.
    if (metafile) {
      return metafileVariantKey(imagePath, normalized, {
        width: Number(resolvedBitmap.width),
        height: Number(resolvedBitmap.height),
      });
    }
    return cachedBitmapVariantKey(imagePath, {
      targetWidthPx: Number(resolvedBitmap.width),
      targetHeightPx: Number(resolvedBitmap.height),
      maxRetainedPixels: normalized.maxRetainedPixels,
    });
  }
  if (metafile) {
    return reusableMetafileVariantKey(fetchImage, imagePath, normalized)
      ?? metafileVariantKey(imagePath, normalized);
  }
  if (!normalized.targetWidthPx && !normalized.targetHeightPx) {
    return cachedBitmapVariantKey(imagePath);
  }
  const nativeKey = `${BASE_CACHE_PREFIX}${cachedBitmapVariantKey(imagePath)}`;
  if (!hasRestrictedPixelLimit(normalized)
    && bitmapCacheByFetch.get(fetchImage)?.entries.has(nativeKey)) {
    return cachedBitmapVariantKey(imagePath);
  }
  return profileNeedsResolutionVariant(profile, normalized)
    ? reusableResolutionVariantKey(
        fetchImage,
        imagePath,
        normalized,
        profile.inspection.dimensions,
      )
      ?? cachedBitmapVariantKey(imagePath, normalized)
    : cachedBitmapVariantKey(imagePath);
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
      while (state.retainedBytes > state.retainedLimit) {
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
  epoch?: DecodedBitmapCacheEpoch,
): Promise<ImageBitmap | null> {
  if (epoch) assertCurrentDerivedEpoch(owner, namespace, epoch);
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
  // Compatibility teardown adapters can run before their image owner has been
  // initialized. WeakMap.get historically made that a no-op; the generation
  // advance uses WeakMap.set, so retain the old behavior explicitly.
  if (!isDecodedBitmapCacheOwner(owner)) return;
  // Advance even when no entry exists: an operation may have resolved its base
  // but not inserted the derived entry yet, and this drop must still win.
  advanceDerivedNamespaceGeneration(owner, namespace);
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
 * document (keyed by `owner`) then by path. The bytes are fetched lazily
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
 * `owner` defaults to the loader identity. Supplying it explicitly lets a
 * document unite resources obtained through different byte APIs (for example
 * PPTX picture parts and media-poster parts) under one cache, queue and budget.
 */
export function getCachedBitmapByPath(
  imagePath: string,
  mimeType: string,
  fetchImage: FetchImage,
  opts: CachedBitmapOptions = {},
  owner: DecodedBitmapCacheOwner = fetchImage,
): Promise<ImageBitmap | null> {
  const normalized = normalizedBitmapOptions(opts);
  const {
    widthPt = 0,
    heightPt = 0,
    suppressBoundaryFrame = false,
    tiff,
    targetWidthPx,
    targetHeightPx,
    maxRetainedPixels,
    svgDecoder,
  } = normalized;
  if (mimeType === 'image/svg+xml' && svgDecoder) {
    return getCachedDecodedBitmap(
      BASE_CACHE_NAMESPACE,
      cachedBitmapVariantKey(imagePath, normalized),
      owner,
      async () => ({
        bitmap: await svgDecoder(await fetchImage(imagePath, mimeType), {
          targetWidthPx,
          targetHeightPx,
        }),
        owned: true,
      }),
    );
  }
  const decode = (
    cacheKey: string,
    initial?: RasterSourceProfile,
    resizeToTarget = true,
  ) => {
    const initialBlob = initial?.initialBlob;
    if (initial) initial.initialBlob = undefined;
    return getCachedDecodedBitmap(
      BASE_CACHE_NAMESPACE,
      cacheKey,
      owner,
      async () => {
        const blob = initialBlob ?? await fetchImage(imagePath, mimeType);
        const decodeOpts = {
          widthPt,
          heightPt,
          suppressBoundaryFrame,
          tiff,
          targetWidthPx: resizeToTarget ? targetWidthPx : undefined,
          targetHeightPx: resizeToTarget ? targetHeightPx : undefined,
          maxRetainedPixels,
        };
        const bitmap = initial
          ? await decodeRasterOrMetafileWithInspection(blob, decodeOpts, initial.inspection)
          : await decodeRasterOrMetafile(blob, decodeOpts);
        return { bitmap, owned: true };
      },
    ).then((bitmap) => enforceCachedRetainedBudget(bitmap, normalized));
  };
  const generation = cacheGeneration(owner);
  return rasterProfileFor(owner, imagePath, mimeType, fetchImage).then((profile) => {
    assertCurrentCacheGeneration(owner, generation);
    if (profileIsMetafile(profile)) {
      return decode(
        reusableMetafileVariantKey(owner, imagePath, normalized)
          ?? metafileVariantKey(imagePath, normalized),
        profile,
      );
    }
    if (!targetWidthPx && !targetHeightPx) {
      return decode(cachedBitmapVariantKey(imagePath), profile, false);
    }
    const nativeKey = `${BASE_CACHE_PREFIX}${cachedBitmapVariantKey(imagePath)}`;
    if (!hasRestrictedPixelLimit(normalized)
      && bitmapCacheByFetch.get(owner)?.entries.has(nativeKey)) {
      return decode(cachedBitmapVariantKey(imagePath), profile, false);
    }
    const needsResolutionVariant = profileNeedsResolutionVariant(profile, normalized);
    return decode(
      needsResolutionVariant
        ? reusableResolutionVariantKey(
            owner,
            imagePath,
            normalized,
            profile.inspection.dimensions,
          )
          ?? cachedBitmapVariantKey(imagePath, normalized)
        : cachedBitmapVariantKey(imagePath),
      profile,
      needsResolutionVariant,
    );
  });
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
  // A compatibility teardown adapter can reach this path without an initialized
  // image owner. Preserve the historical no-op behavior of WeakMap.get/delete
  // for that case before advancing the new epoch map (WeakMap.set would throw).
  if (!isDecodedBitmapCacheOwner(owner)) return;
  cacheGenerationByOwner.set(owner, cacheGeneration(owner) + 1);
  derivedNamespaceGenerationByOwner.delete(owner);
  rasterProfilesByOwner.delete(owner);
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
