import {
  defaultDpr,
  isHTMLCanvas,
  clampCanvasSize,
  getCachedSvgImageByPath,
  getCachedDuotoneBitmapByPath,
  acquireBitmapCacheLease,
  preferVectorBlip,
  metafileRasterSize,
  isOoxmlDecodedImageLimitError,
  EMU_PER_PT,
  EMU_PER_PX,
  type MathRenderer,
  type ChartThreeDRenderer,
  type ChartRegionMapRenderer,
  type ChartExRenderer,
  type SrcRect,
  type Duotone,
  type OffscreenFactory,
  chartImageFillKey,
  collectChartMarkerImageFills,
  collectChartMarkerImageFillsForCharts,
} from '@silurus/ooxml-core';
import type { ParsedWorkbook, Worksheet, ViewportRange, RenderViewportOptions } from './types.js';
import {
  renderViewport,
  prepareWorksheetMath,
  worksheetHasUncachedMath,
  imageCacheKey,
  getGridGeometryForWorksheet,
  applyAutoRowHeights,
  hasPreparedAutoRowHeights,
  inheritSheetRenderCache,
  HEADER_W,
  HEADER_H,
} from './renderer.js';
import { GridGeometry, type GridAxisGeometry } from './internal/grid-geometry.js';
import { usesNativeOneCellExtent } from './internal/cell-anchor-geometry.js';

/** Internal viewer-to-renderer commit latch. It is intentionally not re-exported
 * from the package API: standalone renderer callers do not own viewer lifecycle. */
export const XLSX_RENDER_COMMIT_GUARD: unique symbol = Symbol('xlsx-render-commit-guard');

type GuardedRenderViewportOptions = RenderViewportOptions & {
  readonly [XLSX_RENDER_COMMIT_GUARD]?: () => boolean;
};

/** Attach the internal lifecycle latch without widening the public renderer
 * option type. Kept format-local because only the XLSX viewer has this frame
 * preparation/commit split. */
export function withXlsxRenderCommitGuard(
  options: RenderViewportOptions,
  guard: () => boolean,
): RenderViewportOptions {
  return { ...options, [XLSX_RENDER_COMMIT_GUARD]: guard } as RenderViewportOptions;
}

/** What `prefetchImages` needs to decode one picture: the raster `imagePath`
 *  (also the cache key), its `mimeType`, the optional svgBlip vector path, and
 *  the picture's intended draw size in points (sizes a metafile raster; 0 ⇒
 *  decoder fallback). */
interface ImageRef {
  imagePath: string;
  mimeType: string;
  svgImagePath?: string;
  widthPt?: number;
  heightPt?: number;
  /** The picture's `<a:srcRect>` crop (§20.1.8.55), when present. Forces the
   *  raster decode (the crop math needs native bitmap pixels) and, for a
   *  metafile, scales the raster up to the full picture frame so the fractional
   *  crop lands correctly (see `metafileRasterSize`). */
  srcRect?: SrcRect | null;
  /** The picture's `<a:duotone>` recolour (§20.1.8.23), when present. The base
   *  bitmap is decoded, then recoloured along the `clr1`→`clr2` ramp; the result
   *  is cached under {@link imageCacheKey}(imagePath, duotone). */
  duotone?: Duotone | null;
  /** Chart marker effects are authored paint. If the pixel pipeline cannot
   * apply them, omit the marker instead of substituting the original image. */
  failClosedOnDuotoneFailure?: boolean;
}

interface CellAnchorRange {
  fromCol: number;
  fromColOff: number;
  fromRow: number;
  fromRowOff: number;
  toCol: number;
  toColOff: number;
  toRow: number;
  toRowOff: number;
  editAs?: string;
  nativeExtCx?: number;
  nativeExtCy?: number;
}

function anchorMayIntersectViewport(
  anchor: CellAnchorRange,
  ws: Worksheet,
  viewport: ViewportRange | undefined,
  geometry?: GridGeometry,
  frame?: {
    readonly width: number;
    readonly height: number;
    readonly scale: number;
    readonly freezeRows: number;
    readonly freezeCols: number;
  },
): boolean {
  if (!viewport) return true;
  const axes = geometry ?? getGridGeometryForWorksheet(ws);
  const scale = frame?.scale ?? 1;
  const { col, row } = axes.axesAtScale(scale);
  const marker = (axis: GridAxisGeometry, index: number, offset: number) =>
    axis.offsetOf(index + 1) + (offset * scale) / EMU_PER_PX;
  const fromX = marker(col, anchor.fromCol, anchor.fromColOff);
  const fromY = marker(row, anchor.fromRow, anchor.fromRowOff);
  const useNativeExtent = usesNativeOneCellExtent(anchor);
  const toX = useNativeExtent
    ? fromX + ((anchor.nativeExtCx as number) * scale) / EMU_PER_PX
    : marker(col, anchor.toCol, anchor.toColOff);
  const toY = useNativeExtent
    ? fromY + ((anchor.nativeExtCy as number) * scale) / EMU_PER_PX
    : marker(row, anchor.toRow, anchor.toRowOff);
  if (toX <= fromX || toY <= fromY) return false;
  const effectiveFreeze = frame
    ? axes.effectiveFrozenBands({
        scale,
        width: frame.width,
        height: frame.height,
        headerWidth: HEADER_W,
        headerHeight: HEADER_H,
        rows: frame.freezeRows,
        cols: frame.freezeCols,
      })
    : { rows: ws.freezeRows ?? 0, cols: ws.freezeCols ?? 0 };
  const intersects = (
    start: number,
    end: number,
    frozenEnd: number,
    scrollStart: number,
    scrollEnd: number,
  ) => {
    const low = Math.min(start, end);
    const high = Math.max(start, end);
    return (low < frozenEnd && high > 0)
      || (low < scrollEnd && high > scrollStart);
  };
  return intersects(
    fromX,
    toX,
    col.offsetOf(effectiveFreeze.cols + 1),
    col.offsetOf(viewport.col),
    col.offsetOf(viewport.col + viewport.cols),
  ) && intersects(
    fromY,
    toY,
    row.offsetOf(effectiveFreeze.rows + 1),
    row.offsetOf(viewport.row),
    row.offsetOf(viewport.row + viewport.rows),
  );
}

/** Fetch one image's bytes by zip path and resolve them to a drawable
 *  `CanvasImageSource`, preferring the Microsoft svgBlip vector original
 *  (MS-ODRAWXML). Unified across the top-level twoCellAnchor picture
 *  (`ImageAnchor`) and the `<xdr:grpSp>` leaf (`ShapeGeom` image) — both carry a
 *  raster `imagePath` fallback plus an optional `svgImagePath`. The svgBlip
 *  vector branch applies only when the picture is NOT cropped (shared
 *  `preferVectorBlip` gate): with an `<a:srcRect>` crop we force the raster,
 *  because the renderer's crop math needs the decoded bitmap's native pixel grid
 *  (an SVG element has none).
 *
 *  All three decode paths go through the SAME per-`fetchImage` core caches that
 *  docx and pptx use (issue #781), so xlsx no longer keeps its own owned bitmap
 *  map:
 *   - raster/metafile (+ any `<a:duotone>` recolour, §20.1.8.23) →
 *     {@link getCachedDuotoneBitmapByPath}, a thin two-layer wrapper over the
 *     path-keyed `getCachedBitmapByPath` (content-sniffs the bytes: a WMF, which
 *     `createImageBitmap` can't decode, is rasterized by the shared minimal
 *     player at a size derived from `widthPt`/`heightPt`; a true EMF — or a WMF
 *     with no geometry — resolves to `null`, so the picture is skipped rather
 *     than crashing). With no duotone this is exactly the base-bitmap decode.
 *   - SVG vector original → `getCachedSvgImageByPath` (decodes to an
 *     `HTMLImageElement`, because `createImageBitmap` cannot rasterize SVG in
 *     every browser).
 *  Bytes are fetched lazily by zip path through `fetchImage` (twin of
 *  pptx/docx's `fetchImage`) instead of being inlined as base64; the decoded
 *  bitmaps are owned by those shared caches (LRU-bounded, closed on eviction and
 *  on the per-document `drop*` at destroy / re-parse) rather than by the caller's
 *  lookup map.
 *
 *  Returns `null` for an unsupported metafile so the renderer skips a missing
 *  source. */
export async function decodeImageSource(
  imagePath: string,
  mimeType: string,
  svgImagePath: string | undefined,
  fetchImage: (path: string, mime: string) => Promise<Blob>,
  widthPt = 0,
  heightPt = 0,
  srcRect: SrcRect | null = null,
  duotone: Duotone | null = null,
  offscreenFactory?: OffscreenFactory,
  failClosedOnDuotoneFailure = false,
): Promise<CanvasImageSource | null> {
  const dataIsSvg = mimeType === 'image/svg+xml';
  // SVG pixels are not exposed to the shared bitmap effect pipeline. Without
  // a raster twin, drawing the original SVG would silently discard duotone.
  if (dataIsSvg && duotone) return null;
  // A cropped metafile must rasterize at its FULL picture frame, not the visible
  // sub-rect, so the fractional crop lands correctly; raster blips and uncropped
  // metafiles pass the box through unchanged. The shared base cache is path-keyed
  // ("first size wins"), matching pptx/docx.
  const sized = metafileRasterSize(mimeType, srcRect, widthPt, heightPt);
  if (!sized) return null;
  const decodeRaster = (): Promise<ImageBitmap | null> =>
    getCachedDuotoneBitmapByPath(imagePath, mimeType, duotone, fetchImage, {
      widthPt: sized.widthPt,
      heightPt: sized.heightPt,
      offscreenFactory,
      failClosedOnDuotoneFailure,
    });
  // Shared vector-vs-raster gate (see core preferVectorBlip). When it returns
  // true, `blip.svgImagePath` is narrowed to string.
  const blip = { svgImagePath, srcRect };
  if (!duotone && preferVectorBlip(blip)) {
    // No crop: prefer the vector original; fall back to the raster on decode
    // failure (or, when `imagePath` is itself the SVG, the SVG decoder again).
    // A cropped picture skips this branch so the crop math (below, in the
    // renderer) runs on the raster bitmap's native pixel dimensions. §20.1.8.23
    // duotone applies only to the raster fallback — an SVG vector original has no
    // readable pixel grid (matches docx/pptx).
    try {
      return await getCachedSvgImageByPath(blip.svgImagePath, fetchImage);
    } catch {
      return dataIsSvg ? getCachedSvgImageByPath(imagePath, fetchImage) : decodeRaster();
    }
  }
  if (dataIsSvg) {
    // svg-only picture with no separate `svgImagePath` field (defensive): the
    // raster decoder (createImageBitmap) can't rasterize SVG.
    return getCachedSvgImageByPath(imagePath, fetchImage);
  }
  return decodeRaster();
}

/** Collect every embedded image referenced by a worksheet, resolve each against
 *  the shared per-`fetchImage` core caches, and record the drawable in
 *  `imageCache` under {@link imageCacheKey}(path, duotone) — the renderer's
 *  synchronous lookup key. Images appear either as a top-level twoCellAnchor
 *  `<xdr:pic>` (in `ws.images`) or as a leaf inside an `<xdr:grpSp>` (a
 *  `ShapeGeom` with `type: 'image'`); BOTH are collected so the renderer never
 *  hits a missing source during the synchronous draw. De-duped by lookup key so a
 *  path shared across anchors is resolved once per pass.
 *
 *  `imageCache` is a pure synchronous-lookup layer, NOT the owner of the decoded
 *  bitmaps: every image is re-resolved through `decodeImageSource` on each pass
 *  (the way docx/pptx do), so a still-referenced blip whose bitmap was
 *  LRU-evicted (and closed) by the shared cache is transparently re-decoded
 *  rather than served stale/closed — a resolved bitmap always comes from a live
 *  shared-cache entry. A shared-cache hit re-fetches no bytes and re-runs no
 *  decode, so a steady-state pass only awaits already-settled promises. Storing
 *  `null` for an unsupported metafile (true EMF / geometry-less WMF) lets the
 *  renderer skip a falsy source without a re-fetch.
 *
 *  A no-op when `fetchImage` is absent (no byte source). Per-image failures are
 *  swallowed so one broken picture doesn't sink the grid. */
export async function prefetchImages(
  ws: Worksheet,
  imageCache: Map<string, CanvasImageSource | null>,
  fetchImage: ((path: string, mime: string) => Promise<Blob>) | undefined,
  // Optional offscreen-surface factory for the `<a:duotone>` pixel transform,
  // injected in environments without a global `OffscreenCanvas` (or by tests).
  // Defaults to the real `OffscreenCanvas` when the runtime provides one.
  opts?: {
    offscreenFactory?: OffscreenFactory;
    viewport?: ViewportRange;
    width?: number;
    height?: number;
    cellScale?: number;
    freezeRows?: number;
    freezeCols?: number;
  },
): Promise<void> {
  // This map is only the synchronous lookup for the current frame. Never keep
  // entries from another worksheet/frame: their shared-cache owners may have
  // evicted and closed the underlying drawable in the meantime.
  imageCache.clear();
  if (!fetchImage) return;
  const fetch = fetchImage;
  const refs = new Map<string, ImageRef>();
  const geometry = opts?.viewport ? getGridGeometryForWorksheet(ws) : undefined;
  const frame = opts?.viewport && opts.width !== undefined && opts.height !== undefined
    ? {
        width: opts.width,
        height: opts.height,
        scale: opts.cellScale ?? 1,
        freezeRows: opts.freezeRows ?? ws.freezeRows ?? 0,
        freezeCols: opts.freezeCols ?? ws.freezeCols ?? 0,
      }
    : undefined;
  if (ws.images) {
    for (const img of ws.images) {
      if (!anchorMayIntersectViewport(
        img,
        ws,
        opts?.viewport,
        geometry,
        frame,
      )) continue;
      // Key by (path + duotone colours) so a recoloured picture is looked up
      // separately from the raw blip (§20.1.8.23).
      refs.set(imageCacheKey(img.imagePath, img.duotone), {
        imagePath: img.imagePath,
        mimeType: img.mimeType,
        svgImagePath: img.svgImagePath,
        // Saved EMU extent → pt sizes a metafile raster (0 ⇒ decoder fallback).
        widthPt: img.nativeExtCx > 0 ? img.nativeExtCx / EMU_PER_PT : 0,
        heightPt: img.nativeExtCy > 0 ? img.nativeExtCy / EMU_PER_PT : 0,
        // An `<a:srcRect>` crop forces the raster decode (native pixel grid)
        // and, for a metafile, the full-frame raster size.
        srcRect: img.srcRect ?? null,
        duotone: img.duotone ?? null,
      });
    }
  }
  if (ws.shapeGroups) {
    for (const grp of ws.shapeGroups) {
      if (!anchorMayIntersectViewport(
        grp,
        ws,
        opts?.viewport,
        geometry,
        frame,
      )) continue;
      for (const shape of grp.shapes) {
        if (shape.geom.type === 'image') {
          refs.set(imageCacheKey(shape.geom.imagePath, shape.geom.duotone), {
            imagePath: shape.geom.imagePath,
            mimeType: shape.geom.mimeType,
            svgImagePath: shape.geom.svgImagePath,
            // Group's saved EMU extent scaled by the leaf's normalized w/h → pt.
            widthPt: grp.nativeExtCx > 0 ? (grp.nativeExtCx * shape.w) / EMU_PER_PT : 0,
            heightPt: grp.nativeExtCy > 0 ? (grp.nativeExtCy * shape.h) / EMU_PER_PT : 0,
            // A crop forces the raster decode (native pixel grid for the crop)
            // and, for a metafile, the full-frame raster size.
            srcRect: shape.geom.srcRect ?? null,
            duotone: shape.geom.duotone ?? null,
          });
        }
      }
    }
  }
  const visibleChartModels = (ws.charts ?? [])
    .filter(chart => anchorMayIntersectViewport(chart, ws, opts?.viewport, geometry, frame))
    .map(chart => chart.chart);
  for (const fill of collectChartMarkerImageFillsForCharts(visibleChartModels)) {
      refs.set(chartImageFillKey(fill), {
        imagePath: fill.imagePath,
        mimeType: fill.mimeType,
        svgImagePath: fill.svgImagePath,
        widthPt: 72,
        heightPt: 72,
        srcRect: fill.srcRect ?? null,
        duotone: fill.duotone ?? null,
        failClosedOnDuotoneFailure: true,
      });
  }
  if (refs.size === 0) return;
  await Promise.all(
    [...refs.entries()].map(async ([key, ref]) => {
      try {
        // The §20.1.8.23 duotone recolour is applied inside the shared decode
        // (getCachedDuotoneBitmapByPath) and cached under a colour-suffixed key,
        // so the per-frame draw stays synchronous. Only raster/bitmap sources are
        // recoloured — an SVG element (vector blip) has no readable pixel grid.
        const src = await decodeImageSource(
          ref.imagePath,
          ref.mimeType,
          ref.svgImagePath,
          fetch,
          ref.widthPt,
          ref.heightPt,
          ref.srcRect,
          ref.duotone,
          opts?.offscreenFactory,
          ref.failClosedOnDuotoneFailure ?? false,
        );
        // Record the resolved drawable (INCLUDING a null for an unsupported
        // metafile, so the renderer skips a falsy source without a re-fetch).
        imageCache.set(key, src);
      } catch (error) {
        if (isOoxmlDecodedImageLimitError(error)) throw error;
        // Transient failure: DELETE any prior lookup entry rather than leaving
        // it. A prior entry is re-resolved precisely because its shared-cache
        // backing may be gone (LRU-evicted and GPU-closed); when the re-resolve
        // fails we cannot vouch for that bitmap's liveness, and the renderer
        // skips only a missing/falsy source — it would draw a closed one.
        imageCache.delete(key);
      }
    }),
  );
}

export interface RenderDeps {
  ws: Worksheet;
  styles: ParsedWorkbook['styles'];
  math?: MathRenderer;
  threeD?: ChartThreeDRenderer;
  regionMap?: ChartRegionMapRenderer;
  chartEx?: ChartExRenderer;
}

const autoHeightProjectionCache = new WeakMap<Worksheet, Worksheet>();

export function worksheetWithAutoRowHeights(
  ctx: CanvasRenderingContext2D,
  source: Worksheet,
  styles: ParsedWorkbook['styles'],
): Worksheet {
  if (hasPreparedAutoRowHeights(source)) return source;
  const cached = autoHeightProjectionCache.get(source);
  if (cached) return cached;
  const projection: Worksheet = {
    ...source,
    rowHeights: { ...source.rowHeights },
  };
  inheritSheetRenderCache(source, projection);
  // The viewer/main realm supplies the authoritative Normal-font MDW used by
  // hit-testing and spacer geometry. Preserve it across the render-local clone
  // so worker/direct auto-fit wraps at the exact same column pixels.
  const mdw = getGridGeometryForWorksheet(source).maximumDigitWidth;
  GridGeometry.forWorksheet(projection, mdw);
  applyAutoRowHeights(ctx, projection, styles);
  // applyAutoRowHeights invalidates geometry after deriving row sizes; seed the
  // rebuilt row axis with the same authoritative MDW rather than remeasuring in
  // another Canvas realm.
  GridGeometry.forWorksheet(projection, mdw);
  autoHeightProjectionCache.set(source, projection);
  return projection;
}

/** The full per-frame orchestration: preload uncached images, pre-rasterize
 *  equations, size the target, draw. Shared verbatim by the main-thread
 *  XlsxWorkbook and the render worker.
 *
 *  The whole pass (prefetch → synchronous draw) runs under a core render-pass
 *  lease ({@link acquireBitmapCacheLease}): the shared bitmap cache is
 *  LRU-bounded, so a pass resolving more images than the cap — or a concurrent
 *  pass on the same workbook — would otherwise evict AND GPU-close bitmaps this
 *  pass's lookup map still references before the draw runs. Under the lease the
 *  eviction still removes the cache entry (size stays bounded; the next pass
 *  re-decodes), but the close is deferred until the lease is released after the
 *  draw, so drawImage never receives a closed bitmap. */
export async function renderWorksheetViewport(
  deps: RenderDeps,
  target: HTMLCanvasElement | OffscreenCanvas,
  viewport: ViewportRange,
  opts: RenderViewportOptions = {},
): Promise<void> {
  const releaseLease = opts.fetchImage ? acquireBitmapCacheLease(opts.fetchImage) : undefined;
  try {
    await renderWorksheetViewportLeased(deps, target, viewport, opts);
  } finally {
    releaseLease?.();
  }
}

/** {@link renderWorksheetViewport}'s body, verbatim; runs under the caller's
 *  render-pass lease. */
async function renderWorksheetViewportLeased(
  deps: RenderDeps,
  target: HTMLCanvasElement | OffscreenCanvas,
  viewport: ViewportRange,
  opts: RenderViewportOptions = {},
): Promise<void> {
  const styles = deps.styles;
  const measurementCtx = target.getContext('2d') as CanvasRenderingContext2D | null;
  if (!measurementCtx) throw new Error('XLSX render target does not provide a 2-D canvas context');
  const ws = worksheetWithAutoRowHeights(measurementCtx, deps.ws, styles);
  const rawW = isHTMLCanvas(target) ? (target.clientWidth || 800) : target.width;
  const rawH = isHTMLCanvas(target) ? (target.clientHeight || 600) : target.height;
  const width = opts.width ?? rawW;
  const height = opts.height ?? rawH;
  // Frame-local synchronous lookup only. Core owns decoded reuse/eviction;
  // retaining this map across frames would accumulate stale closed references.
  const imageCache = new Map<string, CanvasImageSource | null>();

  // ── Step 1: Preload any uncached image sources BEFORE touching the canvas.
  //
  // Images can appear either as top-level twoCellAnchor `<xdr:pic>` (captured
  // in `ws.images`) or as a leaf inside an `<xdr:grpSp>` (captured as a
  // ShapeGeom with `type: 'image'`); `prefetchImages` collects both, keyed by
  // zip `imagePath`, fetching bytes lazily via `opts.fetchImage`.
  //
  // Doing this *before* the canvas resize is critical for scroll smoothness:
  // setting `canvas.width` wipes the canvas, and an `await` after that wipe
  // yields to the browser's paint cycle, causing a visible white flash on
  // every scroll frame. By awaiting first (and only when there's something
  // uncached), the whole resize+draw runs synchronously in a single tick and
  // the old frame stays visible until the new one is ready.
  await prefetchImages(ws, imageCache, opts.fetchImage, {
    viewport,
    width,
    height,
    cellScale: opts.cellScale,
    freezeRows: opts.freezeRows,
    freezeCols: opts.freezeCols,
  });

  // ── Step 1b: Pre-rasterize equations in shapes BEFORE the canvas resize,
  // for the same no-white-flash reason as the image preload. Gated on
  // `worksheetHasUncachedMath` so steady-state scroll/zoom frames take NO
  // await and stay fully synchronous — only the first frame that reveals new
  // equations pays the (idempotently cached) MathJax cost. Opt-in: skipped
  // entirely unless the caller supplies a `math` engine.
  if (deps.math && worksheetHasUncachedMath(ws)) {
    await prepareWorksheetMath(ws, deps.math);
  }

  // Resource preparation above may yield. A viewer can be destroyed or a newer
  // frame can supersede this one while it waits; never mutate the caller-owned
  // canvas after that lifecycle generation is stale.
  if ((opts as GuardedRenderViewportOptions)[XLSX_RENDER_COMMIT_GUARD]?.() === false) return;

  // ── Step 2: Resize + draw, all synchronous from here.
  const dpr = opts.dpr ?? defaultDpr();
  // Resize only when the backing store dimensions actually change. Assigning
  // canvas.width/height re-allocates (and clears) the GPU backing store, so on a
  // steady-state scroll/zoom stream — where width/height/dpr are unchanged frame
  // to frame — re-assigning the same value wastes an allocation every frame
  // (improvement plan C4). The inner renderViewport starts with an explicit
  // clearRect + white fill, so nothing depends on the width-assignment's implicit
  // clear; skipping the same-size resize is safe.
  // Clamp the backing store to browser canvas limits (RB5). A very large viewport
  // (or high dpr × large viewport, e.g. an extreme zoom) can exceed the per-axis
  // or total-area cap, at which point the browser silently allocates a smaller-
  // or-empty buffer and the sheet renders blank. `clampCanvasSize` scales BOTH
  // axes by one factor (≤ 1) so the aspect ratio is kept; we fold that factor
  // into the effective dpr, keep the CSS box at the requested size, and the
  // browser stretches the (slightly lower-res) backing store to fill it.
  const clamped = clampCanvasSize(width * dpr, height * dpr);
  const effectiveDpr = clamped.clamped ? dpr * clamped.scale : dpr;
  const bw = clamped.width;
  const bh = clamped.height;
  if (target.width !== bw) target.width = bw;
  if (target.height !== bh) target.height = bh;
  // Set CSS display size so the browser renders at 1:1 device pixels (no browser-level scaling).
  // Without this, canvas.width=2400 on a DPR=2 display causes the canvas to be laid out at
  // 2400 CSS px, making all content appear blurry when viewed in a 1200 CSS px container.
  if (isHTMLCanvas(target)) {
    const cssW = `${width}px`;
    const cssH = `${height}px`;
    if (target.style.width !== cssW) target.style.width = cssW;
    if (target.style.height !== cssH) target.style.height = cssH;
  }

  const ctx = (target as HTMLCanvasElement).getContext('2d') as CanvasRenderingContext2D;
  // Set the DPR transform absolutely rather than ctx.scale(dpr, dpr): when the
  // resize above is skipped the backing store is NOT re-created, so its transform
  // is not reset to identity, and a relative scale() would compound the dpr every
  // frame (progressive zoom). setTransform is idempotent whether or not the store
  // was reallocated. Use the effective dpr (folded with any clamp factor) so
  // drawing fills the clamped backing store; renderViewport gets the same value
  // so its own dpr-dependent math stays aligned.
  ctx.setTransform(effectiveDpr, 0, 0, effectiveDpr, 0, 0);

  // RB7 partial degradation: a sheet whose part failed to parse (see the Rust
  // `Worksheet::placeholder`) carries `parseError` and no rows. Paint a visible
  // error overlay in place of the grid so the workbook's OTHER sheets stay usable
  // and this tab clearly reads as "broken". Healthy sheets never take this path.
  if (ws.parseError) {
    drawSheetParseErrorOverlay(ctx, width, height, ws.name, ws.parseError);
    return;
  }

  renderViewport(ctx, ws, styles, viewport, {
    ...opts,
    dpr: effectiveDpr,
    loadedImages: imageCache,
    threeD: deps.threeD,
    regionMap: deps.regionMap,
    chartEx: deps.chartEx,
  });
}

/**
 * RB7: paint a placeholder overlay for a worksheet whose part failed to parse.
 * A neutral fill, a warning glyph, a heading naming the sheet, and the
 * part-tagged error wrapped to a few lines. Coordinates are in CSS px (the ctx
 * is already dpr-scaled by the caller). Only ever called for a sheet carrying
 * `parseError`.
 */
function drawSheetParseErrorOverlay(
  ctx: CanvasRenderingContext2D,
  widthPx: number,
  heightPx: number,
  sheetName: string,
  message: string,
): void {
  ctx.save();
  ctx.fillStyle = '#f7f7f8';
  ctx.fillRect(0, 0, widthPx, heightPx);
  const cx = widthPx / 2;
  const base = Math.min(widthPx, heightPx);

  const glyph = Math.max(20, base * 0.1);
  ctx.fillStyle = '#b23b3b';
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.font = `${glyph}px sans-serif`;
  ctx.fillText('⚠', cx, heightPx * 0.32);

  const headSize = Math.max(13, base * 0.035);
  ctx.fillStyle = '#333333';
  ctx.font = `600 ${headSize}px sans-serif`;
  ctx.fillText(`Sheet "${sheetName}" could not be displayed`, cx, heightPx * 0.46);

  const detailSize = Math.max(10, base * 0.022);
  ctx.fillStyle = '#666666';
  ctx.font = `${detailSize}px sans-serif`;
  const maxLineWidth = Math.min(widthPx * 0.8, 640);
  const words = message.split(/\s+/);
  const lines: string[] = [];
  let line = '';
  for (const word of words) {
    const candidate = line ? `${line} ${word}` : word;
    if (ctx.measureText(candidate).width > maxLineWidth && line) {
      lines.push(line);
      line = word;
    } else {
      line = candidate;
    }
    if (lines.length >= 4) break;
  }
  if (line && lines.length < 4) lines.push(line);
  const lineHeight = detailSize * 1.4;
  let y = heightPx * 0.52 + lineHeight;
  for (const l of lines.slice(0, 4)) {
    ctx.fillText(l, cx, y);
    y += lineHeight;
  }
  ctx.restore();
}
