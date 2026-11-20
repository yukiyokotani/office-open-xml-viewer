import { withChartFontSubstitutions } from '@silurus/ooxml-core/internal/chart-font-sources';
import { acquireDrawingMLTileBrush, acquireVmlImageBrush } from './browser-vml-image-brush.js';
import {
  withBitmapCacheLease,
  OoxmlError,
  clampCanvasSize,
  defaultDpr,
  isHTMLCanvas,
  isOptionalImageCodecUnavailableError,
  PT_TO_PX,
  chartImageFillKey,
  chartImageFillNeedsRenderedSize,
  chartImageFillPixelTransform,
  chartFillRenderedSizeKey,
  chartImageFillStoreLookup,
  collectChartFillRenderedSizes,
  assertChartFillRenderedSizeCount,
  chartImageFillUsageSize,
  collectChartImageFillUsages,
  collectChartImageFillUsagesForCharts,
  getCachedSvgImageByPath,
  preferVectorBlip,
} from '@silurus/ooxml-core';
import type { Duotone, ImageResourceOptions } from '@silurus/ooxml-core';
import type { ChartThreeDRenderer, ChartRegionMapRenderer, ChartExRenderer, TiffRenderer } from '@silurus/ooxml-core';
import type {
  ChartPaintResourceDescriptor,
  DeepReadonly,
  DocumentLayout,
  LayoutPage,
  PaintResourceRegistry,
  RasterPaintOccurrence,
} from '../layout/types.js';
import {
  decodeRaster,
  isTerminalImageAcquisitionError,
  preloadPaintImages,
  imageKey,
  drawingMLTileCellKey,
  type DocxFetchImage,
  type DecodedPaintImage,
} from './browser-images.js';
import {
  createCanvasPaintResourcePainter,
  paintLayoutPageContent,
  paintLayoutPage,
} from './canvas-page.js';
import {
  canonicalCanvasPaintResourceHandlers,
  createCanonicalCanvasPaintResourceHandlers,
} from './canonical-resource-handlers.js';
import {
  createProductionPaintResourceSession,
  unavailablePaintResourceHandle,
} from './resource-session.js';
import type { PaintCanvas2D } from './types.js';

interface PrivatePaintResourceLookup {
  readonly keys: readonly string[];
  resolve(resourceKey: string): CanvasImageSource;
}

export interface CanvasDocumentPaintOptions<TTextRun> {
  readonly onDiagnostic?: (diagnostic: import('../layout/types.js').LayoutDiagnostic) => void;
  /** Internal provenance: bitmap/worker surfaces belong to the document, not a caller target. */
  readonly callerOwnedCanvasTarget?: boolean;
  /** The renderer validates this context before layout and paint retain the same instance. */
  readonly validatedTargetContext?: PaintCanvas2D;
  readonly width?: number;
  readonly dpr?: number;
  readonly defaultTextColor?: string;
  readonly fetchImage?: DocxFetchImage;
  readonly svgDecoder?: import('@silurus/ooxml-core').SvgBlobDecoder;
  readonly parseError: boolean;
  readonly registry: PaintResourceRegistry;
  /** Layout measurement's reference-advance decisions for retained text. */
  readonly textAdvances?: import('@silurus/ooxml-core').TextAdvanceLookup;
  /** Final retained raster/chart frames for this selected page. */
  readonly rasterPaintOccurrences: readonly DeepReadonly<RasterPaintOccurrence>[];
  readonly privateResources?: PrivatePaintResourceLookup;
  readonly textRuns: readonly TTextRun[];
  readonly onTextRun?: (run: TTextRun) => void;
  readonly threeD?: ChartThreeDRenderer;
  readonly regionMap?: ChartRegionMapRenderer;
  readonly chartEx?: ChartExRenderer;
  readonly tiff?: TiffRenderer;
  readonly imageResources?: ImageResourceOptions;
  /** The unscoped substitutions this document's font sources loaded
   * (`ChartModel.fontSubstitutions`); chart text paints them after the
   * authored face. Absent without sources. */
  readonly chartFontSubstitutions?: Readonly<Record<string, string>>;
}

/** Per-canvas cancellation token: only the newest asynchronous image preload
 * may paint after rapid navigation reuses the same canvas. */
const renderTokens = new WeakMap<HTMLCanvasElement | OffscreenCanvas, number>();

/** Invalidate an in-flight main-thread render before restoring or reusing its
 * caller-owned target. The renderer observes the same token after every await. */
export function invalidateDocxRenderTarget(
  target: HTMLCanvasElement | OffscreenCanvas,
): void {
  renderTokens.set(target, (renderTokens.get(target) ?? 0) + 1);
}

export function canvasPageScale(page: LayoutPage, width?: number): number {
  return (width ?? page.geometry.widthPt * PT_TO_PX) / page.geometry.widthPt;
}

function htmlCanvasOwnerDocument(
  target: HTMLCanvasElement | OffscreenCanvas,
): Document | null {
  if (isHTMLCanvas(target)) {
    return target.ownerDocument ?? (typeof document === 'undefined' ? null : document);
  }
  const ownerDocument = (target as unknown as HTMLCanvasElement).ownerDocument;
  const ownerConstructor = ownerDocument?.defaultView?.HTMLCanvasElement;
  return ownerConstructor && target instanceof ownerConstructor ? ownerDocument : null;
}

function isElementBackedCanvas(
  target: HTMLCanvasElement | OffscreenCanvas,
): target is HTMLCanvasElement {
  return htmlCanvasOwnerDocument(target) !== null;
}

function acquireElementBackedVerticalPaintSurface(
  target: HTMLCanvasElement | OffscreenCanvas,
  required: boolean,
): Readonly<{
  canvas: HTMLCanvasElement | OffscreenCanvas;
  release?: () => void;
}> {
  const targetDocument = htmlCanvasOwnerDocument(target);
  if (!required || (targetDocument && (target as HTMLCanvasElement).isConnected)) {
    return { canvas: target };
  }
  const paintDocument = targetDocument ?? (
    typeof document === 'undefined' ? undefined : document
  );
  if (!paintDocument) {
    throw new Error('OpenType vertical glyph paint requires an element-backed document surface');
  }
  const parent = paintDocument.body ?? paintDocument.documentElement;
  if (!parent) {
    throw new Error('OpenType vertical glyph paint requires an attached document surface');
  }
  const canvas = paintDocument.createElement('canvas');
  canvas.setAttribute('aria-hidden', 'true');
  Object.assign(canvas.style, {
    position: 'fixed',
    left: '-99999px',
    top: '0',
    opacity: '0',
    pointerEvents: 'none',
  });
  parent.appendChild(canvas);
  return {
    canvas,
    release: () => canvas.remove(),
  };
}

export async function renderSelectedDocumentPage<TTextRun>(
  layout: DocumentLayout,
  page: LayoutPage,
  canvas: HTMLCanvasElement | OffscreenCanvas,
  options: CanvasDocumentPaintOptions<TTextRun>,
): Promise<void> {
  const token = (renderTokens.get(canvas) ?? 0) + 1;
  renderTokens.set(canvas, token);
  const superseded = (): boolean => renderTokens.get(canvas) !== token;
  // The §17.2.1 page background owns retained picture brushes outside the
  // page story graph; prepare them with the page's own resources.
  const backgroundResourceKeys = (page.pageBackground?.commands ?? []).flatMap((command) => (
    command.kind === 'drawingml-image-fill' ? [command.resourceKey] : []
  ));
  const pageResourceKeys = page.layers.capabilities.resourceKeys
    && [...backgroundResourceKeys, ...page.layers.capabilities.resourceKeys];
  const descriptorByKey = new Map(
    options.registry.descriptors.map((descriptor) => [descriptor.resourceKey, descriptor]),
  );
  const descriptors = pageResourceKeys
    ? pageResourceKeys.map((key) => {
        const descriptor = descriptorByKey.get(key);
        if (!descriptor) throw new Error(`Missing retained paint resource descriptor: ${key}`);
        return descriptor;
      })
    : options.registry.descriptors;
  const hasDecodedImages = descriptors.some(
    descriptor => descriptor.kind === 'image'
      || descriptor.kind === 'picture-bullet'
      || (descriptor.kind === 'chart'
        && collectChartImageFillUsages(
          descriptor.model as import('@silurus/ooxml-core').ChartModel,
        ).length > 0),
  );
  const paint = () => superseded()
    ? Promise.resolve()
    : renderSelectedDocumentPageLeased(layout, page, canvas, options, descriptors, superseded);
  return options.fetchImage && hasDecodedImages
    ? withBitmapCacheLease(options.fetchImage, options.imageResources, paint)
    : paint();
}

type PreparedPageImages = Readonly<{
  images: Map<string, DecodedPaintImage>;
  chartImages: Map<string, CanvasImageSource | null>;
}>;

/**
 * Resolve every asynchronous paint input of one page before its target is
 * cleared. The caller then paints the whole page in one synchronous section.
 *
 * Chromium records Canvas 2D calls and rasterizes the pending recording when
 * the canvas flushes at the end of a task. A paint that awaited decodes after
 * clearing was split into chunks at timing-dependent draw boundaries, and the
 * software rasterizer's antialiasing of concave paths and clips depends on the
 * other work in the same unflushed chunk, so a cold first paint could differ
 * from a warm repaint by a few edge pixels. Preparing first makes the chunking
 * independent of decode timing; it also means a page is never presented
 * half-painted while its decodes are pending.
 *
 * Failures are returned rather than thrown so the caller can raise them at
 * their former position in paint order, after the page background is cleared.
 */
async function preparePageImages<TTextRun>(
  options: CanvasDocumentPaintOptions<TTextRun>,
  descriptors: PaintResourceRegistry['descriptors'],
  scale: number,
  effectiveDpr: number,
  superseded: () => boolean,
): Promise<
  | { readonly kind: 'ready'; readonly images: PreparedPageImages }
  | { readonly kind: 'superseded' }
  | { readonly kind: 'failed'; readonly error: unknown; readonly whenSuperseded: 'drop' | 'throw' }
> {
  let images: Map<string, DecodedPaintImage>;
  try {
    images = await preloadPaintImages(
      descriptors,
      options.rasterPaintOccurrences,
      options.fetchImage,
      options.tiff,
      scale * effectiveDpr,
      options.svgDecoder,
      options.imageResources,
      options.onDiagnostic,
    );
  } catch (error) {
    // A superseded render never reported an image preload failure.
    return { kind: 'failed', error, whenSuperseded: 'drop' };
  }
  if (superseded()) return { kind: 'superseded' };
  try {
    const chartImages = await resolveChartImages(options, descriptors, images, scale, effectiveDpr);
    return { kind: 'ready', images: { images, chartImages } };
  } catch (error) {
    // Chart picture-fill budget and TIFF decode failures always propagated,
    // even from a superseded render.
    return { kind: 'failed', error, whenSuperseded: 'throw' };
  }
}

async function resolveChartImages<TTextRun>(
  options: CanvasDocumentPaintOptions<TTextRun>,
  descriptors: PaintResourceRegistry['descriptors'],
  images: ReadonlyMap<string, DecodedPaintImage>,
  scale: number,
  effectiveDpr: number,
): Promise<Map<string, CanvasImageSource | null>> {
  const chartImages = new Map<string, CanvasImageSource | null>();
  if (options.fetchImage) {
    const fetchImage = options.fetchImage;
    const chartOccurrencesByResource = new Map<string, DeepReadonly<RasterPaintOccurrence>[]>();
    for (const occurrence of options.rasterPaintOccurrences) {
      if (occurrence.resourceKind !== 'chart') continue;
      const prior = chartOccurrencesByResource.get(occurrence.resourceKey) ?? [];
      if (!chartOccurrencesByResource.has(occurrence.resourceKey)) {
        chartOccurrencesByResource.set(occurrence.resourceKey, prior);
      }
      prior.push(occurrence);
    }
    // A retained chart occurrence whose frame or derived decode size is
    // non-positive or non-finite cannot paint an image safely. Keep every
    // valid occurrence/frame pairing through source gating; different uses of
    // one chart can have different final aspect ratios before their decoded
    // picture sources are deduplicated.
    const chartOccurrences: Array<{
      descriptor: DeepReadonly<ChartPaintResourceDescriptor>;
      frame: Parameters<typeof chartImageFillUsageSize>[1];
      usages: Array<{
        usage: ReturnType<typeof collectChartImageFillUsages>[number];
        size: NonNullable<ReturnType<typeof chartImageFillUsageSize>>;
      }>;
    }> = [];
    for (const descriptor of descriptors) {
      if (descriptor.kind !== 'chart') continue;
      for (const occurrence of chartOccurrencesByResource.get(descriptor.resourceKey) ?? []) {
        if (!Number.isFinite(occurrence.widthPt) || !Number.isFinite(occurrence.heightPt) || occurrence.widthPt < 0 || occurrence.heightPt < 0) {
          throw new OoxmlError('unsupported-image-effect', 'Invalid chart image occurrence geometry');
        }
        if (occurrence.widthPt <= 0 || occurrence.heightPt <= 0) continue;
        const frame = {
          widthPt: occurrence.widthPt,
          heightPt: occurrence.heightPt,
          targetWidthPx: occurrence.widthPt * scale * effectiveDpr,
          targetHeightPx: occurrence.heightPt * scale * effectiveDpr,
        };
        const usages = [] as typeof chartOccurrences[number]['usages'];
        let valid = true;
        for (const usage of collectChartImageFillUsages(
          descriptor.model as import('@silurus/ooxml-core').ChartModel,
        )) {
          const size = chartImageFillUsageSize(usage, frame);
          if (!size) {
            valid = false;
            break;
          }
          usages.push({ usage, size });
        }
        if (valid) chartOccurrences.push({ descriptor, frame, usages });
      }
    }
    type ChartEntry = {
      fill: ReturnType<typeof collectChartImageFillUsages>[number]['fill'];
      widthPt: number;
      heightPt: number;
      targetWidthPx?: number;
      targetHeightPx?: number;
      preserveNaturalSize: boolean;
      hasSourceCrop: boolean;
      renderedSize?: import('@silurus/ooxml-core').ChartFillRenderedSize;
    };
    const chartEntries = new Map<string, ChartEntry>();
    for (const usage of collectChartImageFillUsagesForCharts(
      chartOccurrences.map(
        ({ descriptor }) => descriptor.model as import('@silurus/ooxml-core').ChartModel,
      ),
      (usage, chartIndex) => chartImageFillUsageSize(
        usage,
        chartOccurrences[chartIndex]!.frame,
      ) != null,
    )) {
      const { fill } = usage;
      // Size-dependent chains are prepared per rendered size below.
      if (chartImageFillNeedsRenderedSize(fill)) continue;
      const key = chartFillRenderedSizeKey(fill);
      if (!chartEntries.has(key)) chartEntries.set(key, {
        fill,
        widthPt: 0,
        heightPt: 0,
        preserveNaturalSize: usage.preserveNaturalSize,
        hasSourceCrop: usage.hasSourceCrop,
      });
    }
    const admittedChartFills = new Set(collectChartImageFillUsagesForCharts(
      chartOccurrences.map(({ descriptor }) => descriptor.model as import('@silurus/ooxml-core').ChartModel),
      (usage, chartIndex) => chartImageFillUsageSize(usage, chartOccurrences[chartIndex]!.frame) != null,
    ).map(({ fill }) => chartImageFillKey(fill)));
    // A stretched fill whose chain is measured in physical units is
    // evaluated per rendered size: a preparation pass of the same chart paint
    // (canonical handler: point units, ptToPx 1) records each size. Returns
    // the keys this round added.
    const prepareRenderedSizes = (resolved?: import('@silurus/ooxml-core').ChartImageLookup): string[] => {
      const added: string[] = [];
      for (const { descriptor, frame } of chartOccurrences) {
        for (const { fill, occurrence } of collectChartFillRenderedSizes(
          // Laid out exactly as the paint draws it (with its substitutions).
          withChartFontSubstitutions(descriptor.model as import('@silurus/ooxml-core').ChartModel,
            options.chartFontSubstitutions),
          { x: 0, y: 0, w: frame.widthPt, h: frame.heightPt },
          1, 0, options.threeD, options.regionMap, options.chartEx, resolved,
        )) {
          if (!admittedChartFills.has(chartImageFillKey(fill))) continue;
          const key = chartFillRenderedSizeKey(fill, undefined, occurrence);
          if (chartEntries.has(key)) continue;
          chartEntries.set(key, {
            fill, widthPt: occurrence.widthPt, heightPt: occurrence.heightPt,
            preserveNaturalSize: true, hasSourceCrop: fill.srcRect != null, renderedSize: occurrence,
          });
          added.push(key);
        }
      }
      // The per-render ceiling covers every prepared chart source.
      assertChartFillRenderedSizeCount(chartEntries.size);
      return added;
    };
    prepareRenderedSizes();
    for (const { usages } of chartOccurrences) {
      for (const { usage, size } of usages) {
        const { fill } = usage;
        if (chartImageFillNeedsRenderedSize(fill)) continue;
        const key = chartFillRenderedSizeKey(fill);
        const prior = chartEntries.get(key);
        if (!prior) continue;
        const preserveNaturalSize = prior.preserveNaturalSize || usage.preserveNaturalSize;
        // A picture fill may cover a marker, plot area, wall, or floor. The
        // chart frame bounds every consumer; core usage factors retain every
        // same-chart crop and stretch fillRect before source deduplication.
        chartEntries.set(key, {
          ...prior,
          widthPt: Math.max(prior.widthPt, size.widthPt),
          heightPt: Math.max(prior.heightPt, size.heightPt),
          targetWidthPx: preserveNaturalSize
            ? undefined
            : Math.max(prior.targetWidthPx ?? 0, size.targetWidthPx ?? 0) || undefined,
          targetHeightPx: preserveNaturalSize
            ? undefined
            : Math.max(prior.targetHeightPx ?? 0, size.targetHeightPx ?? 0) || undefined,
          preserveNaturalSize,
          hasSourceCrop: prior.hasSourceCrop || usage.hasSourceCrop,
        });
      }
    }
    const decodeChartEntry = async ([key, entry]: [string, ChartEntry]) => {
      // The source preload keys plain chart fills by chartImageFillKey.
      const preloadKey = entry.renderedSize ? undefined : chartImageFillKey(entry.fill);
      if (preloadKey !== undefined && images.has(preloadKey)) {
        const image = images.get(preloadKey);
        chartImages.set(
          key,
          isOptionalImageCodecUnavailableError(image, 'tiff') ? null : image ?? null,
        );
        return;
      }
      const {
        fill, widthPt, heightPt, targetWidthPx, targetHeightPx, hasSourceCrop, renderedSize,
      } = entry;
      // A per-size chain is evaluated on its rendered frame; other chains
      // sample the source grid at the fill DPI (chartImageFillPixelTransform).
      const transform = chartImageFillPixelTransform(fill, undefined, renderedSize);
      const effectOrientation = transform && 'effects' in transform && !renderedSize
        ? { emuPerPixelX: transform.emuPerPixelX, emuPerPixelY: transform.emuPerPixelY }
        : undefined;
      const effectWidthPt = fill.blipEffects?.length ? renderedSize?.widthPt ?? 0 : widthPt;
      const effectHeightPt = fill.blipEffects?.length ? renderedSize?.heightPt ?? 0 : heightPt;
      const target = targetWidthPx && targetHeightPx
        ? { targetWidthPx, targetHeightPx }
        : undefined;
      try {
        const decodeSvg = (path: string) => options.svgDecoder
          ? getCachedSvgImageByPath(path, fetchImage, {
              ...(target ?? {}),
              workerDecoder: options.svgDecoder,
            })
          : getCachedSvgImageByPath(path, fetchImage);
        const decodePath = (path: string, mime: string) => mime === 'image/svg+xml' && !fill.duotone
          && !fill.blipEffects?.length
          ? decodeSvg(path)
          : decodeRaster(
              path, mime, undefined, fetchImage as DocxFetchImage,
              effectWidthPt, effectHeightPt, fill.duotone, options.tiff,
              fill.blipEffects?.length ? undefined : target,
              undefined, options.svgDecoder, undefined,
              fill.blipEffects?.length ? fill.blipEffects : undefined, effectOrientation,
            );
        let image: CanvasImageSource | null;
        const blip = { svgImagePath: fill.svgImagePath, srcRect: hasSourceCrop ? true : null };
        if (!fill.blipEffects?.length && preferVectorBlip(blip)) {
          try {
            image = await decodePath(blip.svgImagePath, 'image/svg+xml');
          } catch (error) {
            if (isTerminalImageAcquisitionError(error)) throw error;
            image = await decodePath(fill.imagePath, fill.mimeType);
          }
        } else {
          image = await decodePath(fill.imagePath, fill.mimeType);
        }
        chartImages.set(key, image);
        if (!image) options.onDiagnostic?.({
          code: 'MISSING_RESOURCE', severity: 'warning',
          message: 'A chart picture fill produced no drawable output; the remaining chart content is rendered',
        });
      } catch (error) {
        if (isOptionalImageCodecUnavailableError(error, 'tiff')) {
          chartImages.set(key, null);
          return;
        }
        if (isTerminalImageAcquisitionError(error)) throw error;
        options.onDiagnostic?.({
          code: 'MISSING_RESOURCE', severity: 'warning',
          message: 'A chart picture fill could not be acquired; the remaining chart content is rendered',
        });
        chartImages.set(key, null);
      }
    };
    await Promise.all([...chartEntries].map(decodeChartEntry));
    // Occurrences derived from a decoded source's own frame (3-D plain-stack
    // repetitions) are recorded once that source is known; each round
    // repeats the paint with every source decoded so far.
    for (;;) {
      const added = prepareRenderedSizes(chartImageFillStoreLookup(chartImages, undefined, false));
      if (added.length === 0) break;
      await Promise.all(added.map(key => decodeChartEntry([key, chartEntries.get(key)!])));
    }
  }
  return chartImages;
}

/** Validate only the target operation, before dimensions/ink change. Null caller
 * contexts and a browser InvalidStateError on a caller-owned (e.g. transferred)
 * surface are input errors. OOM/resource failures and all internal surfaces remain
 * terminal acquisition errors. No paint-to-layout runtime dependency is needed. */
export function acquireDocumentCanvasTargetContext(
  canvas: HTMLCanvasElement | OffscreenCanvas,
  callerOwned: boolean = true,
): PaintCanvas2D {
  let context: PaintCanvas2D | null;
  try { context = canvas.getContext('2d') as PaintCanvas2D | null; }
  catch (error) {
    let callerInvalidState = false;
    if (callerOwned) {
      try {
        let provenDOMException = typeof DOMException !== 'undefined' && error instanceof DOMException;
        if (!provenDOMException) {
          const ownerView = htmlCanvasOwnerDocument(canvas)?.defaultView;
          const ownerDOMException = ownerView && 'DOMException' in ownerView ? ownerView.DOMException : undefined;
          provenDOMException = typeof ownerDOMException === 'function' && error instanceof ownerDOMException;
        }
        callerInvalidState = provenDOMException && typeof error === 'object' && error !== null
          && 'name' in error && error.name === 'InvalidStateError';
      } catch {
        // Failed owner, prototype or name inspection cannot replace the exact
        // original target failure or establish a caller exemption.
        throw error;
      }
    }
    if (callerInvalidState) {
      throw Object.assign(new RangeError('Invalid state of caller DOCX canvas target'), {
        code: 'docx-caller-input', cause: error,
      });
    }
    throw error;
  }
  if (!context) {
    if (callerOwned) throw Object.assign(new RangeError('2D canvas is unavailable for DOCX paint projection'), { code: 'docx-caller-input' });
    throw new Error('2D canvas is unavailable for internally acquired DOCX paint surface');
  }
  return context;
}

async function renderSelectedDocumentPageLeased<TTextRun>(
  layout: DocumentLayout,
  page: LayoutPage,
  canvas: HTMLCanvasElement | OffscreenCanvas,
  options: CanvasDocumentPaintOptions<TTextRun>,
  descriptors: PaintResourceRegistry['descriptors'],
  superseded: () => boolean,
): Promise<void> {
  let releasePaintSurface: (() => void) | undefined;
  try {
    const destination = options.validatedTargetContext ?? acquireDocumentCanvasTargetContext(canvas, options.callerOwnedCanvasTarget !== false);
    const dpr = options.dpr ?? defaultDpr();
    const paintSurface = acquireElementBackedVerticalPaintSurface(
      canvas,
      !options.parseError && page.layers.capabilities.requiresElementBackedVerticalGlyphPaint,
    );
    const paintCanvas = paintSurface.canvas;
    releasePaintSurface = paintSurface.release;
    const context = paintCanvas === canvas ? destination : acquireDocumentCanvasTargetContext(paintCanvas, false);
    const scale = canvasPageScale(page, options.width);
    const cssWidth = page.geometry.widthPt * scale;
    const cssHeight = page.geometry.heightPt * scale;
    const clamped = clampCanvasSize(cssWidth * dpr, cssHeight * dpr);
    const effectiveDpr = clamped.clamped ? dpr * clamped.scale : dpr;

    // Clearing the target (by resizing it) starts the page paint.
    const clearPage = (): void => {
      canvas.width = clamped.width;
      canvas.height = clamped.height;
      if (paintCanvas !== canvas) {
        paintCanvas.width = clamped.width;
        paintCanvas.height = clamped.height;
      }
      if (isElementBackedCanvas(canvas)) {
        canvas.style.width = `${cssWidth}px`;
        canvas.style.height = `${cssHeight}px`;
        if (!canvas.style.display) canvas.style.display = 'block';
      }
      if (isElementBackedCanvas(paintCanvas) && paintCanvas !== canvas) {
        paintCanvas.style.width = `${cssWidth}px`;
        paintCanvas.style.height = `${cssHeight}px`;
      }
      context.scale(effectiveDpr, effectiveDpr);
      // ECMA-376 §17.2.1: a solid document background replaces the white
      // page surface (one fill, same geometry).
      context.fillStyle = page.pageBackground?.color ?? '#ffffff';
      context.fillRect(0, 0, cssWidth, cssHeight);
    };

    if (options.parseError) {
      clearPage();
      await paintLayoutPage(layout, 0, canvas, { scale, dpr: effectiveDpr });
      return;
    }

    // Prepare: nothing touches the target until every input settles, so a
    // superseded render returns before clearing the newer render's page.
    const outcome = await preparePageImages(
      options, descriptors, scale, effectiveDpr, superseded,
    );
    if (outcome.kind === 'superseded') return;
    if (superseded()) {
      if (outcome.kind === 'failed' && outcome.whenSuperseded === 'throw') throw outcome.error;
      return;
    }

    // Decode and budget failures surface after the page background, before
    // any content.
    if (outcome.kind === 'failed') { clearPage(); throw outcome.error; }
    const { images, chartImages } = outcome.images;

    const session = createProductionPaintResourceSession(options.registry, (descriptor) => {
      if (descriptor.kind === 'math') {
        return options.privateResources?.keys.includes(descriptor.resourceKey)
          ? options.privateResources.resolve(descriptor.resourceKey)
          : unavailablePaintResourceHandle('optional math renderer unavailable');
      }
      if (descriptor.kind === 'image' || descriptor.kind === 'picture-bullet') {
        const image = images.get(descriptor.blipEffects?.length ? descriptor.resourceKey : imageKey(
          descriptor.partPath,
          descriptor.colorReplaceFrom,
          descriptor.duotone as Duotone | undefined,
          descriptor.vmlStrokePattern as Duotone | undefined,
          descriptor.blipEffects,
        ));
        if (isOptionalImageCodecUnavailableError(image, 'tiff')) {
          return unavailablePaintResourceHandle(
            'optional TIFF codec unavailable',
            { placeholder: 'tiff' },
          );
        }
        if (image && descriptor.vmlFill) return acquireVmlImageBrush(descriptor, image);
        if (image && descriptor.tile) {
          const cell = images.get(drawingMLTileCellKey(descriptor));
          // A §20.1.8.55 window with no positive extent selects nothing.
          return acquireDrawingMLTileBrush(descriptor, image, cell as CanvasImageSource | null | undefined)
            ?? unavailablePaintResourceHandle('picture tile window selects no content');
        }
        return image ?? unavailablePaintResourceHandle(
          options.fetchImage
            ? 'unsupported image format produced no drawable output'
            : 'image byte source unavailable',
        );
      }
      return undefined;
    });
    // From this single clear through completed painting there is no await.
    clearPage();

    const resources = createCanvasPaintResourcePainter(
      session,
      options.threeD || options.regionMap || options.chartEx || chartImages.size > 0
        || options.chartFontSubstitutions
        ? createCanonicalCanvasPaintResourceHandlers(
            options.threeD,
            options.regionMap,
            chartImageFillStoreLookup(chartImages),
            options.chartEx,
            options.chartFontSubstitutions,
          )
        : canonicalCanvasPaintResourceHandlers,
    );
    context.save();
    try {
      context.scale(scale, scale);
      paintLayoutPageContent(page, {
        ctx: context,
        scale,
        dpr: effectiveDpr,
        resources,
        documentDefaultTextColor: options.defaultTextColor ?? '#000000',
        defaultTextColor: options.defaultTextColor ?? '#000000',
        ...(options.textAdvances ? { textAdvances: options.textAdvances } : {}),
      });
    } finally {
      context.restore();
    }
    if (paintCanvas !== canvas) {
      if (superseded()) return;
      destination.drawImage(paintCanvas, 0, 0);
    }
    if (options.onTextRun) {
      for (const run of options.textRuns) options.onTextRun(run);
    }
  } finally {
    releasePaintSurface?.();
  }
}
