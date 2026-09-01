import {
  acquireBitmapCacheLease,
  clampCanvasSize,
  defaultDpr,
  isHTMLCanvas,
  isOoxmlDecodedImageLimitError,
  metafileRasterSize,
  PT_TO_PX,
  chartImageFillKey,
  collectChartMarkerImageFills,
  collectChartMarkerImageFillsForCharts,
  getCachedSvgImageByPath,
  preferVectorBlip,
} from '@silurus/ooxml-core';
import type { Duotone } from '@silurus/ooxml-core';
import type { ChartThreeDRenderer, ChartRegionMapRenderer, ChartExRenderer, TiffRenderer } from '@silurus/ooxml-core';
import type {
  DocumentLayout,
  LayoutPage,
  PaintResourceRegistry,
} from '../layout/types.js';
import { decodeRaster, preloadPaintImages, imageKey, type DocxFetchImage } from './browser-images.js';
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
  readonly parseError: boolean;
  readonly registry: PaintResourceRegistry;
  readonly privateResources?: PrivatePaintResourceLookup;
  readonly textRuns: readonly TTextRun[];
  readonly onTextRun?: (run: TTextRun) => void;
  readonly threeD?: ChartThreeDRenderer;
  readonly regionMap?: ChartRegionMapRenderer;
  readonly chartEx?: ChartExRenderer;
  readonly tiff?: TiffRenderer;
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
  const releaseLease = options.fetchImage
    ? acquireBitmapCacheLease(options.fetchImage)
    : undefined;
  let releasePaintSurface: (() => void) | undefined;
  try {
    const token = (renderTokens.get(canvas) ?? 0) + 1;
    renderTokens.set(canvas, token);
    const superseded = (): boolean => renderTokens.get(canvas) !== token;
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
      images = await preloadPaintImages(options.registry.descriptors, options.fetchImage, options.tiff);
    } catch (error) {
      if (superseded()) return;
      throw error;
    }
    if (superseded()) return;

    const chartImages = new Map<string, CanvasImageSource | null>();
    if (options.fetchImage) {
      const fetchImage = options.fetchImage;
      const uniqueFills = new Map<string, ReturnType<typeof collectChartMarkerImageFills>[number]>();
      for (const fill of collectChartMarkerImageFillsForCharts(
        options.registry.descriptors
          .filter(descriptor => descriptor.kind === 'chart')
          .map(descriptor => descriptor.model as import('@silurus/ooxml-core').ChartModel),
      )) {
        const key = chartImageFillKey(fill);
        if (!uniqueFills.has(key)) uniqueFills.set(key, fill);
      }
      await Promise.all([...uniqueFills].map(async ([key, fill]) => {
        const raster = metafileRasterSize(fill.mimeType, fill.srcRect, 72, 72);
        if (!raster) {
          chartImages.set(key, null);
          return;
        }
        try {
          const decodeFallback = () => fill.mimeType === 'image/svg+xml'
            ? fill.duotone ? Promise.resolve(null) : getCachedSvgImageByPath(fill.imagePath, fetchImage)
            : decodeRaster(
                fill.imagePath, fill.mimeType, undefined, fetchImage as DocxFetchImage,
                raster.widthPt, raster.heightPt, fill.duotone, true, options.tiff,
              );
          let image: CanvasImageSource | null;
          if (!fill.duotone && preferVectorBlip(fill)) {
            try {
              image = await getCachedSvgImageByPath(fill.svgImagePath, fetchImage);
            } catch {
              image = await decodeFallback();
            }
          } else {
            image = await decodeFallback();
          }
          chartImages.set(key, image);
        } catch (error) {
          if (isOoxmlDecodedImageLimitError(error)) throw error;
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
        return images.get(imageKey(
          descriptor.partPath,
          descriptor.colorReplaceFrom,
          descriptor.duotone as Duotone | undefined,
        )) ?? unavailablePaintResourceHandle(
          options.fetchImage
            ? 'unsupported image format produced no drawable output'
            : 'image byte source unavailable',
        );
      }
      return undefined;
    });
    const resources = createCanvasPaintResourcePainter(
      session,
      options.threeD || options.regionMap || options.chartEx || chartImages.size > 0
        ? createCanonicalCanvasPaintResourceHandlers(
            options.threeD,
            options.regionMap,
            fill => chartImages.get(chartImageFillKey(fill)),
            options.chartEx,
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
    releaseLease?.();
  }
}
