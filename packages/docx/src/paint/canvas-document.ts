import {
  withBitmapCacheLease,
  clampCanvasSize,
  defaultDpr,
  isHTMLCanvas,
  isOoxmlDecodedImageLimitError,
  isTiffDecodeError,
  isOptionalImageCodecUnavailableError,
  PT_TO_PX,
  chartImageFillKey,
  chartImageFillUsageSize,
  collectChartImageFillUsages,
  collectChartImageFillUsagesForCharts,
  getCachedSvgImageByPath,
  preferVectorBlip,
} from '@silurus/ooxml-core';
import type { Duotone, FontFamilyRoutes, ImageResourceOptions } from '@silurus/ooxml-core';
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
  preloadPaintImages,
  imageKey,
  type DocxFetchImage,
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
  readonly width?: number;
  readonly dpr?: number;
  readonly defaultTextColor?: string;
  readonly fetchImage?: DocxFetchImage;
  readonly svgDecoder?: import('@silurus/ooxml-core').SvgBlobDecoder;
  readonly parseError: boolean;
  readonly registry: PaintResourceRegistry;
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
  readonly providerFontRoutes?: FontFamilyRoutes;
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
  const pageResourceKeys = page.layers.capabilities.resourceKeys;
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
    const dpr = options.dpr ?? defaultDpr();
    const paintSurface = acquireElementBackedVerticalPaintSurface(
      canvas,
      !options.parseError && page.layers.capabilities.requiresElementBackedVerticalGlyphPaint,
    );
    const paintCanvas = paintSurface.canvas;
    releasePaintSurface = paintSurface.release;
    const context = paintCanvas.getContext('2d') as PaintCanvas2D | null;
    if (!context) throw new Error('2D canvas is unavailable for DOCX paint');
    const scale = canvasPageScale(page, options.width);
    const cssWidth = page.geometry.widthPt * scale;
    const cssHeight = page.geometry.heightPt * scale;
    const clamped = clampCanvasSize(cssWidth * dpr, cssHeight * dpr);
    const effectiveDpr = clamped.clamped ? dpr * clamped.scale : dpr;
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
    context.fillStyle = '#ffffff';
    context.fillRect(0, 0, cssWidth, cssHeight);

    if (options.parseError) {
      await paintLayoutPage(layout, 0, canvas, { scale, dpr: effectiveDpr });
      return;
    }

    let images;
    try {
      images = await preloadPaintImages(
        descriptors,
        options.rasterPaintOccurrences,
        options.fetchImage,
        options.tiff,
        scale * effectiveDpr,
        options.svgDecoder,
        options.imageResources,
      );
    } catch (error) {
      if (superseded()) return;
      throw error;
    }
    if (superseded()) return;

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
          if (!Number.isFinite(occurrence.widthPt)
            || occurrence.widthPt <= 0
            || !Number.isFinite(occurrence.heightPt)
            || occurrence.heightPt <= 0) continue;
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
      const chartEntries = new Map<string, {
        fill: ReturnType<typeof collectChartImageFillUsages>[number]['fill'];
        widthPt: number;
        heightPt: number;
        targetWidthPx?: number;
        targetHeightPx?: number;
        preserveNaturalSize: boolean;
        hasSourceCrop: boolean;
      }>();
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
        const key = chartImageFillKey(fill);
        if (!chartEntries.has(key)) chartEntries.set(key, {
          fill,
          widthPt: 0,
          heightPt: 0,
          preserveNaturalSize: usage.preserveNaturalSize,
          hasSourceCrop: usage.hasSourceCrop,
        });
      }
      for (const { usages } of chartOccurrences) {
        for (const { usage, size } of usages) {
          const { fill } = usage;
          const key = chartImageFillKey(fill);
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
      await Promise.all([...chartEntries].map(async ([key, entry]) => {
        if (images.has(key)) {
          const image = images.get(key);
          chartImages.set(
            key,
            isOptionalImageCodecUnavailableError(image, 'tiff') ? null : image ?? null,
          );
          return;
        }
        const {
          fill, widthPt, heightPt, targetWidthPx, targetHeightPx, hasSourceCrop,
        } = entry;
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
          const decodeFallback = () => fill.mimeType === 'image/svg+xml'
            ? fill.duotone ? Promise.resolve(null) : decodeSvg(fill.imagePath)
            : decodeRaster(
                fill.imagePath, fill.mimeType, undefined, fetchImage as DocxFetchImage,
                widthPt, heightPt, fill.duotone, true, options.tiff, target,
              );
          let image: CanvasImageSource | null;
          const blip = {
            svgImagePath: fill.svgImagePath,
            srcRect: hasSourceCrop ? true : null,
          };
          if (!fill.duotone && preferVectorBlip(blip)) {
            try {
              image = await decodeSvg(blip.svgImagePath);
            } catch {
              image = await decodeFallback();
            }
          } else {
            image = await decodeFallback();
          }
          chartImages.set(key, image);
        } catch (error) {
          if (isOptionalImageCodecUnavailableError(error, 'tiff')) {
            chartImages.set(key, null);
            return;
          }
          if (isOoxmlDecodedImageLimitError(error) || isTiffDecodeError(error)) throw error;
          chartImages.set(key, null);
        }
      }));
    }
    if (superseded()) return;

    const session = createProductionPaintResourceSession(options.registry, (descriptor) => {
      if (descriptor.kind === 'math') {
        return options.privateResources?.keys.includes(descriptor.resourceKey)
          ? options.privateResources.resolve(descriptor.resourceKey)
          : unavailablePaintResourceHandle('optional math renderer unavailable');
      }
      if (descriptor.kind === 'image' || descriptor.kind === 'picture-bullet') {
        const image = images.get(imageKey(
          descriptor.partPath,
          descriptor.colorReplaceFrom,
          descriptor.duotone as Duotone | undefined,
        ));
        if (isOptionalImageCodecUnavailableError(image, 'tiff')) {
          return unavailablePaintResourceHandle(
            'optional TIFF codec unavailable',
            { placeholder: 'tiff' },
          );
        }
        return image ?? unavailablePaintResourceHandle(
          options.fetchImage
            ? 'unsupported image format produced no drawable output'
            : 'image byte source unavailable',
        );
      }
      return undefined;
    });
    const resources = createCanvasPaintResourcePainter(
      session,
      options.threeD || options.regionMap || options.chartEx || chartImages.size > 0 || options.providerFontRoutes
        ? createCanonicalCanvasPaintResourceHandlers(
            options.threeD,
            options.regionMap,
            fill => chartImages.get(chartImageFillKey(fill)),
            options.chartEx,
            options.providerFontRoutes,
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
      });
    } finally {
      context.restore();
    }
    if (paintCanvas !== canvas) {
      if (superseded()) return;
      const destination = canvas.getContext('2d') as PaintCanvas2D | null;
      if (!destination) throw new Error('2D canvas is unavailable for DOCX paint projection');
      destination.drawImage(paintCanvas, 0, 0);
    }
    if (options.onTextRun) {
      for (const run of options.textRuns) options.onTextRun(run);
    }
  } finally {
    releasePaintSurface?.();
  }
}
