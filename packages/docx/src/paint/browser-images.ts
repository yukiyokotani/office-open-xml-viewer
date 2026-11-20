import {
  applyImageDataTransform,
  getCachedDuotoneBitmapByPath,
  blipPixelEffectsKey,
  OoxmlError,
  decodeSvgBlobOnMainThread,
  OoxmlDecodedImageLimitError,
  OoxmlResourceLimitError,
  isOoxmlDecodedImageLimitError,
  isTiffDecodeError,
  captureDecodedBitmapCacheEpoch,
  decodedBitmapTargetResizeOptions,
  duotoneImageData,
  MAX_RASTER_DIMENSION,
  MAX_RASTER_PIXELS,
  createAuxCanvas,
  imageNaturalSize,
  dropCachedDerivedBitmapNamespace,
  getCachedBitmapByPath,
  getCachedDerivedBitmap,
  getCachedSvgImageByPath,
  inspectCachedRasterSource,
  isBrowserResizableRasterMimeType,
  isDecodeTargetResizableRasterFormat,
  metafileRasterSize,
  preferVectorBlip,
  normalizeImageResourceOptions,
  planDecodedImageTargets,
  resolvedCachedBitmapVariantKey,
  sourceRasterTargetSize,
  chartImageFillKey,
  stretchedSourceFrame,
  chartImageFillPixelTransform,
  chartImageFillNeedsRenderedSize,
  chartImageFillUsageSize,
  collectChartImageFillUsages,
  isOptionalImageCodecUnavailableError,
} from '@silurus/ooxml-core';
import type {
  Duotone,
  BlipEffect,
  ImageResourceOptions,
  OptionalImageCodecUnavailableError,
  TiffRenderer,
} from '@silurus/ooxml-core';
import type {
  DeepReadonly,
  ImagePaintResourceDescriptor,
  PaintResourceDescriptor,
  RasterPaintOccurrence,
} from '../layout/types.js';

/** Resource ceilings and recognized TIFF decode failures stay terminal even
 * when an SVG alternative or a chart's remaining content could be displayed.
 * Invalid authored effect chains also fail closed rather than selecting raw pixels. */
export function isTerminalImageAcquisitionError(error: unknown): boolean {
  return error instanceof OoxmlResourceLimitError || isOoxmlDecodedImageLimitError(error)
    || isTiffDecodeError(error)
    || (error instanceof OoxmlError && error.code === 'unsupported-image-effect');
}

export type DecodedImage = ImageBitmap | HTMLImageElement;
export type DecodedPaintImage = DecodedImage | OptionalImageCodecUnavailableError | null;
export type DocxFetchImage = (path: string, mime: string) => Promise<Blob>;

interface ImageDecodeRequest {
  cacheKey?: string;
  imagePath: string;
  mimeType: string;
  svgImagePath?: string;
  colorReplaceFrom?: string;
  duotone?: Duotone;
  vmlStrokePattern?: Duotone;
  blipEffects?: readonly BlipEffect[];
  /** Host orientation and, for chart fills, fixed physical sampling that the
   * shared CT_Blip interpreter evaluates the chain against. */
  orientation?: {
    readonly hostRotation?: number; readonly hostFlipH?: boolean; readonly hostFlipV?: boolean;
    readonly emuPerPixelX?: number; readonly emuPerPixelY?: number;
  };
  widthPt: number;
  heightPt: number;
  hasCrop: boolean;
  targetWidthPx?: number;
  targetHeightPx?: number;
  /** Set only for browser-resizable rasters admitted by the shared plan. */
  plannedPixelLimit?: number;
  chartPictureFill?: boolean;
}

export function imageKey(
  imagePath: string,
  colorReplaceFrom?: string,
  duotone?: Duotone,
  vmlStrokePattern?: Duotone,
  blipEffects?: readonly BlipEffect[],
): string {
  const clr = colorReplaceFrom ? `|clr:${colorReplaceFrom}` : '';
  const duo = duotone ? `|duo:${duotone.clr1}:${duotone.clr2}` : '';
  const effects = blipEffects?.length
    ? `|effects:${blipPixelEffectsKey({ effects: blipEffects, duotone })}` : '';
  return `${imagePath}${clr}${duo}${effects}${vmlStrokePattern ? `|vml-pen:${vmlStrokePattern.clr1}:${vmlStrokePattern.clr2}` : ''}`;
}

function requestKey(request: ImageDecodeRequest): string {
  return request.cacheKey
    ?? imageKey(request.imagePath, request.colorReplaceFrom, request.duotone, request.vmlStrokePattern, request.blipEffects);
}

const DOCX_COLOR_EFFECT_CACHE_NAMESPACE = 'docx-color-effects';
const DOCX_TILE_CELL_CACHE_NAMESPACE = 'docx-tile-cells';

export function dropBrowserImageCache(fetchImage: DocxFetchImage): void {
  dropCachedDerivedBitmapNamespace(fetchImage, DOCX_COLOR_EFFECT_CACHE_NAMESPACE);
  dropCachedDerivedBitmapNamespace(fetchImage, DOCX_TILE_CELL_CACHE_NAMESPACE);
}

/** Decoded-image map key of an image descriptor's (effect-applied) source. */
export function paintImageDecodeKey(descriptor: DeepReadonly<ImagePaintResourceDescriptor>): string {
  return descriptor.blipEffects?.length ? descriptor.resourceKey : imageKey(
    descriptor.partPath,
    descriptor.colorReplaceFrom,
    descriptor.duotone as Duotone | undefined,
    descriptor.vmlStrokePattern as Duotone | undefined,
    descriptor.blipEffects,
  );
}

/** Whether a §20.1.8.58 tile repeats a composed cell rather than the decoded
 * source itself: a §20.1.8.55 crop or a `flip` mirror needs one. */
export function drawingMLTileNeedsCell(descriptor: DeepReadonly<ImagePaintResourceDescriptor>): boolean {
  const flip = descriptor.tile?.flip;
  return descriptor.tile !== undefined
    && (descriptor.srcRect != null || flip === 'x' || flip === 'y' || flip === 'xy');
}

/** Decoded-image map key of a descriptor's composed tile cell. */
export function drawingMLTileCellKey(descriptor: DeepReadonly<ImagePaintResourceDescriptor>): string {
  const crop = descriptor.srcRect ?? { l: 0, t: 0, r: 0, b: 0 };
  return `${paintImageDecodeKey(descriptor)}|tile-cell:${descriptor.tile?.flip ?? 'none'}:${crop.l}:${crop.t}:${crop.r}:${crop.b}`;
}

/** Native source-pixel geometry of one tile cell: the §20.1.8.55 crop
 * window of the decoded source (whose grid stays native for tiles), and the
 * cell that mirrors it per `flip`. */
export function drawingMLTileCellGeometry(
  descriptor: DeepReadonly<ImagePaintResourceDescriptor>,
  source: CanvasImageSource,
): Readonly<{
  sourceX: number; sourceY: number; sourceWidth: number; sourceHeight: number;
  tileWidthPx: number; tileHeightPx: number; cellWidthPx: number; cellHeightPx: number;
  flipX: boolean; flipY: boolean;
}> {
  const natural = imageNaturalSize(source);
  const crop = descriptor.srcRect ?? { l: 0, t: 0, r: 0, b: 0 };
  const sourceWidth = natural.w * (1 - crop.l - crop.r);
  const sourceHeight = natural.h * (1 - crop.t - crop.b);
  const flipX = descriptor.tile?.flip === 'x' || descriptor.tile?.flip === 'xy';
  const flipY = descriptor.tile?.flip === 'y' || descriptor.tile?.flip === 'xy';
  const tileWidthPx = Math.round(sourceWidth);
  const tileHeightPx = Math.round(sourceHeight);
  if (!(Number.isFinite(sourceWidth) && Number.isFinite(sourceHeight) && tileWidthPx >= 1 && tileHeightPx >= 1)) {
    throw new OoxmlError('unsupported-image-effect', 'Unrepresentable picture tile size');
  }
  return {
    sourceX: natural.w * crop.l, sourceY: natural.h * crop.t, sourceWidth, sourceHeight,
    tileWidthPx, tileHeightPx,
    cellWidthPx: tileWidthPx * (flipX ? 2 : 1), cellHeightPx: tileHeightPx * (flipY ? 2 : 1),
    flipX, flipY,
  };
}

/** Compose a tile cell as a document-owned derived bitmap. The cell is
 * admitted against the shared per-surface axis and pixel ceilings before any
 * allocation (a negative crop inset or a mirror can otherwise amplify a
 * source), and is then retained under the same weighted LRU, decode slots,
 * render-pass lease and aggregate decoded-byte ceiling as its source. The
 * scratch surface is released once the bitmap is baked. */
async function drawingMLTileCell(
  descriptor: DeepReadonly<ImagePaintResourceDescriptor>,
  source: DecodedImage,
  fetchImage: DocxFetchImage,
): Promise<ImageBitmap | null> {
  const geometry = drawingMLTileCellGeometry(descriptor, source);
  const pixels = geometry.cellWidthPx * geometry.cellHeightPx;
  if (geometry.cellWidthPx > MAX_RASTER_DIMENSION || geometry.cellHeightPx > MAX_RASTER_DIMENSION
    || pixels > MAX_RASTER_PIXELS) {
    throw new OoxmlDecodedImageLimitError('image-pixels', MAX_RASTER_PIXELS, pixels);
  }
  const epoch = captureDecodedBitmapCacheEpoch(fetchImage, DOCX_TILE_CELL_CACHE_NAMESPACE);
  return getCachedDerivedBitmap(
    DOCX_TILE_CELL_CACHE_NAMESPACE,
    drawingMLTileCellKey(descriptor),
    fetchImage,
    async () => {
      const canvas = createAuxCanvas(geometry.cellWidthPx, geometry.cellHeightPx);
      const context = canvas?.getContext('2d') as CanvasRenderingContext2D | null | undefined;
      if (!canvas || !context) throw new Error('2D canvas is unavailable for a picture tile cell');
      try {
        for (const [mirrorX, mirrorY] of [[false, false], [true, false], [false, true], [true, true]] as const) {
          if ((mirrorX && !geometry.flipX) || (mirrorY && !geometry.flipY)) continue;
          context.save();
          context.translate(geometry.tileWidthPx * (mirrorX ? 2 : 0), geometry.tileHeightPx * (mirrorY ? 2 : 0));
          context.scale(mirrorX ? -1 : 1, mirrorY ? -1 : 1);
          context.drawImage(source, geometry.sourceX, geometry.sourceY,
            geometry.sourceWidth, geometry.sourceHeight, 0, 0, geometry.tileWidthPx, geometry.tileHeightPx);
          context.restore();
        }
        return { bitmap: await createImageBitmap(canvas as unknown as ImageBitmapSource), owned: true };
      } finally {
        canvas.width = 1;
        canvas.height = 1;
      }
    },
    epoch,
  );
}

function applyExactColorReplacement(data: ImageData, colorHex: string): void {
  const red = parseInt(colorHex.slice(0, 2), 16);
  const green = parseInt(colorHex.slice(2, 4), 16);
  const blue = parseInt(colorHex.slice(4, 6), 16);
  for (let index = 0; index < data.data.length; index += 4) {
    if (data.data[index] === red
      && data.data[index + 1] === green
      && data.data[index + 2] === blue) {
      data.data[index + 3] = 0;
    }
  }
}

async function applyColorEffects(
  bitmap: ImageBitmap,
  colorReplaceFrom: string | undefined,
  duotone: Duotone | undefined,
  target?: Readonly<{ targetWidthPx: number; targetHeightPx: number }>,
  vmlStrokePattern?: Duotone,
): Promise<ImageBitmap | null> {
  // Reuse the shared source-grid transform used by PPTX/XLSX. DOCX combines
  // the legacy exact-colour shorthand and a lone §20.1.8.23 duotone in one
  // buffer; ordered CT_Blip chains go through the shared core interpreter
  // (see decodeRaster). A DOM surface supports browsers without
  // OffscreenCanvas. Returning the source would drop an authored effect, so
  // unavailable readback is explicit.
  const transformed = await applyImageDataTransform(bitmap, data => {
    if (colorReplaceFrom) applyExactColorReplacement(data, colorReplaceFrom);
    if (duotone) duotoneImageData(data, duotone.clr1, duotone.clr2);
    if (vmlStrokePattern) {
      // Part 4 §19.1.2.21 color2 is a black/coloured predicate, not a
      // luminance transfer. Preserve source alpha and the exact native grid.
      const rgb = (hex: string) => [0,2,4].map(i => parseInt(hex.slice(i,i+2),16));
      const black=rgb(vmlStrokePattern.clr1), colored=rgb(vmlStrokePattern.clr2);
      for(let i=0;i<data.data.length;i+=4) {
        const value=data.data[i]===0 && data.data[i+1]===0 && data.data[i+2]===0 ? black : colored;
        data.data[i]=value[0]!; data.data[i+1]=value[1]!; data.data[i+2]=value[2]!;
      }
    }
  }, {
    width: bitmap.width,
    height: bitmap.height,
    ...target,
    offscreenFactory: (width, height) => {
      if (typeof OffscreenCanvas !== 'undefined') return new OffscreenCanvas(width, height);
      if (typeof document === 'undefined') return null;
      const canvas = document.createElement('canvas');
      canvas.width = width;
      canvas.height = height;
      return canvas;
    },
  });
  if (transformed === bitmap) throw new Error('2D canvas is unavailable for image color effects');
  return transformed as ImageBitmap;
}

export async function decodeRaster(
  imagePath: string,
  mimeType: string,
  colorReplaceFrom: string | undefined,
  fetchImage: DocxFetchImage,
  widthPt = 0,
  heightPt = 0,
  duotone?: Duotone,
  tiff?: TiffRenderer,
  target?: Readonly<{ targetWidthPx: number; targetHeightPx: number }>,
  plannedPixelLimit?: number,
  svgDecoder?: import('@silurus/ooxml-core').SvgBlobDecoder,
  vmlStrokePattern?: Duotone,
  blipEffects?: readonly BlipEffect[],
  orientation?: ImageDecodeRequest['orientation'],
): Promise<ImageBitmap | null> {
  if (blipEffects?.length) {
    // One shared CT_Blip interpreter/cache/resource policy for DOCX and PPTX.
    // Synthesize the legacy exact-color shorthand only when the complete
    // parsed clrChange entry was not supplied by the model caller.
    const effects: readonly BlipEffect[] = colorReplaceFrom && !blipEffects.some(e => e.type === 'colorChange')
      ? [{ type: 'colorChange', from: colorReplaceFrom, fromAlpha: 1,
          to: colorReplaceFrom, toAlpha: 0, useAlpha: false }, ...blipEffects]
      : blipEffects;
    return getCachedDuotoneBitmapByPath(imagePath, mimeType, { effects, duotone, ...orientation }, fetchImage, {
      widthPt, heightPt, tiff, ...target,
      ...(plannedPixelLimit ? { maxRetainedPixels: plannedPixelLimit } : {}),
      svgDecoder: svgDecoder ?? (typeof Image !== 'undefined' ? decodeSvgBlobOnMainThread : undefined),
      failClosedOnDuotoneFailure: true,
    });
  }
  // Pixel effects temporarily retain more than their cached input/output:
  // source + offscreen backing + ImageData + result = four surfaces. clrChange
  // and duotone mutate the same ImageData before the one final bitmap bake, so
  // chaining them does not add another full-size intermediate.
  const effectSurfaceCount = colorReplaceFrom || duotone || vmlStrokePattern ? 4 : 1;
  const effectPixelLimit = Math.floor(MAX_RASTER_PIXELS / effectSurfaceCount);
  const maxRetainedPixels = Number.isSafeInteger(plannedPixelLimit) && (plannedPixelLimit ?? 0) > 0
    ? Math.min(effectPixelLimit, plannedPixelLimit as number)
    : effectPixelLimit;
  const epoch = colorReplaceFrom || duotone || vmlStrokePattern
    ? captureDecodedBitmapCacheEpoch(fetchImage, DOCX_COLOR_EFFECT_CACHE_NAMESPACE)
    : undefined;
  const sourceBitmapOptions = {
    widthPt,
    heightPt,
    svgDecoder: svgDecoder ?? (mimeType === 'image/svg+xml' && typeof Image !== 'undefined'
      ? decodeSvgBlobOnMainThread : undefined),
    tiff,
    maxRetainedPixels,
    // The shared decoder preserves the native grid while it fits the explicit
    // effect working-set ceiling. Supplying the display target as a fallback
    // lets adaptive mode downsample an oversized effect source before the four
    // source/offscreen/ImageData/result surfaces would cross that ceiling.
    ...(target ?? {}),
  };
  const base = await getCachedBitmapByPath(imagePath, mimeType, fetchImage, {
    ...sourceBitmapOptions,
  });
  if (!base) return null;
  if (!colorReplaceFrom && !duotone && !vmlStrokePattern) return base;
  // Injected SVG decoders share the decoded owner but can return a larger
  // vector raster than the effect's four-surface working-set allowance.
  const pixels = base.width * base.height;
  if (pixels > maxRetainedPixels) {
    throw new OoxmlDecodedImageLimitError('image-pixels', maxRetainedPixels, pixels);
  }
  const resolvedBaseKey = await resolvedCachedBitmapVariantKey(
    imagePath,
    mimeType,
    fetchImage,
    sourceBitmapOptions,
    epoch,
    base,
  );
  const resizeOptions = decodedBitmapTargetResizeOptions(
    base.width,
    base.height,
    target?.targetWidthPx,
    target?.targetHeightPx,
  );
  const resizeKey = resizeOptions
    ? `|resize:${resizeOptions.resizeWidth}x${resizeOptions.resizeHeight}`
    : '';
  const key = `${imageKey(resolvedBaseKey, colorReplaceFrom, duotone, vmlStrokePattern)}${resizeKey}`;
  return getCachedDerivedBitmap(
    DOCX_COLOR_EFFECT_CACHE_NAMESPACE,
    key,
    fetchImage,
    async () => {
      const bitmap = await applyColorEffects(
        base,
        colorReplaceFrom,
        duotone,
        target,
        vmlStrokePattern,
      );
      return { bitmap, owned: bitmap !== null && bitmap !== base };
    },
    epoch,
  );
}

function imageDecodeRequests(
  descriptors: readonly DeepReadonly<PaintResourceDescriptor>[],
  rasterPaintOccurrences: readonly DeepReadonly<RasterPaintOccurrence>[],
  devicePixelsPerPoint?: number,
): ImageDecodeRequest[] {
  const requests = new Map<string, ImageDecodeRequest>();
  const naturalSizeChartKeys = new Set<string>();
  const demandByResource = new Map<string, { widthPt: number; heightPt: number }>();
  for (const occurrence of rasterPaintOccurrences) {
    if (occurrence.resourceKind !== 'image' && occurrence.resourceKind !== 'picture-bullet') {
      continue;
    }
    if (!Number.isFinite(occurrence.widthPt) || !Number.isFinite(occurrence.heightPt) || occurrence.widthPt < 0 || occurrence.heightPt < 0) {
      throw new OoxmlError('unsupported-image-effect', 'Invalid image occurrence geometry');
    }
    if (occurrence.widthPt <= 0 || occurrence.heightPt <= 0) continue;
    const key = `${occurrence.resourceKind}:${occurrence.resourceKey}`;
    const prior = demandByResource.get(key);
    demandByResource.set(key, {
      widthPt: Math.max(prior?.widthPt ?? 0, occurrence.widthPt),
      heightPt: Math.max(prior?.heightPt ?? 0, occurrence.heightPt),
    });
  }
  const images = descriptors
    .filter((descriptor): descriptor is DeepReadonly<ImagePaintResourceDescriptor> => (
      descriptor.kind === 'image' || descriptor.kind === 'picture-bullet'
    ))
    .sort((left, right) => (
      (left.documentOrder ?? Number.MAX_SAFE_INTEGER)
      - (right.documentOrder ?? Number.MAX_SAFE_INTEGER)
    ));
  for (const image of images) {
    const demand = demandByResource.get(`${image.kind}:${image.resourceKey}`);
    if (!demand) continue;
    // A chain measured in EMUs (§20.1.8.15) is evaluated on the full source
    // frame: the srcRect window stretched over the (fillRect-inset) extent.
    // A tiled fill instead samples its source at the fill DPI (§20.1.8.58).
    const tileDpi = image.blipEffects?.length ? image.effectHost?.tileDpi : undefined;
    const raster = tileDpi !== undefined
      ? { widthPt: 0, heightPt: 0 }
      : image.blipEffects?.length
        ? (frame => ({ widthPt: frame.width, heightPt: frame.height }))(stretchedSourceFrame(
          image.intrinsicSize.widthPt, image.intrinsicSize.heightPt, image.srcRect, image.effectHost?.fillRect))
        : metafileRasterSize(image.mimeType, image.srcRect, demand.widthPt, demand.heightPt);
    if (!raster) continue;
    const request: ImageDecodeRequest = {
      // §20.1.8.15: rad is physical. Keep each authored occurrence's
      // dimensions through preload and synchronous paint lookup; merging by
      // PNG/effect recipe changes the blur when a larger copy is added.
      ...(image.blipEffects?.length ? { cacheKey: image.resourceKey } : {}),
      imagePath: image.partPath,
      mimeType: image.mimeType,
      ...(image.svgImagePath === undefined ? {} : { svgImagePath: image.svgImagePath }),
      ...(image.colorReplaceFrom === undefined ? {} : { colorReplaceFrom: image.colorReplaceFrom }),
      ...(image.duotone === undefined ? {} : { duotone: image.duotone as Duotone }),
      ...(image.blipEffects?.length ? { blipEffects: image.blipEffects,
        orientation: {
          ...(image.effectHost
            ? { hostRotation: image.effectHost.rotation, hostFlipH: image.effectHost.flipH, hostFlipV: image.effectHost.flipV }
            : { hostRotation: image.rotation, hostFlipH: image.flipH, hostFlipV: image.flipV }),
          ...(tileDpi === undefined ? {} : { emuPerPixelX: 914400 / tileDpi, emuPerPixelY: 914400 / tileDpi }),
        } } : {}),
      ...(image.vmlStrokePattern === undefined ? {} : { vmlStrokePattern: image.vmlStrokePattern as Duotone }),
      widthPt: raster.widthPt,
      heightPt: raster.heightPt,
      hasCrop: image.srcRect != null,
    };
    const target = devicePixelsPerPoint === undefined
      ? null
      : sourceRasterTargetSize(
          demand.widthPt * devicePixelsPerPoint,
          demand.heightPt * devicePixelsPerPoint,
          image.srcRect,
        );
    if (target) {
      request.targetWidthPx = target.width;
      request.targetHeightPx = target.height;
    }
    const key = requestKey(request);
    // VML image brushes and §20.1.8.58 DrawingML tiles take their size from
    // the native pixel dimensions. Preserve the source grid even when another
    // occurrence shares this decoded image.
    if (image.vmlFill || image.tile) naturalSizeChartKeys.add(key);
    const existing = requests.get(key);
    if (!existing) {
      requests.set(key, request);
    } else {
      existing.widthPt = Math.max(existing.widthPt, request.widthPt);
      existing.heightPt = Math.max(existing.heightPt, request.heightPt);
      existing.hasCrop ||= request.hasCrop;
      existing.targetWidthPx = Math.max(existing.targetWidthPx ?? 0, request.targetWidthPx ?? 0) || undefined;
      existing.targetHeightPx = Math.max(existing.targetHeightPx ?? 0, request.targetHeightPx ?? 0) || undefined;
    }
  }
  const chartOccurrencesByResource = new Map<string, DeepReadonly<RasterPaintOccurrence>[]>();
  for (const occurrence of rasterPaintOccurrences) {
    if (occurrence.resourceKind !== 'chart') continue;
    const prior = chartOccurrencesByResource.get(occurrence.resourceKey) ?? [];
    if (!chartOccurrencesByResource.has(occurrence.resourceKey)) {
      chartOccurrencesByResource.set(occurrence.resourceKey, prior);
    }
    prior.push(occurrence);
  }
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
        targetWidthPx: devicePixelsPerPoint === undefined
          ? undefined
          : occurrence.widthPt * devicePixelsPerPoint,
        targetHeightPx: devicePixelsPerPoint === undefined
          ? undefined
          : occurrence.heightPt * devicePixelsPerPoint,
      };
      const usages = collectChartImageFillUsages(
        descriptor.model as import('@silurus/ooxml-core').ChartModel,
      ).map(usage => ({ usage, size: chartImageFillUsageSize(usage, frame) }));
      if (usages.some(({ size }) => size === null)) continue;
      for (const { usage, size } of usages) {
        if (!size) continue;
        const fill = usage.fill;
        // Size-dependent chains are prepared per rendered size by the chart
        // paint preparation (resolveChartImages), not by this source preload.
        if (chartImageFillNeedsRenderedSize(fill)) continue;
        const raster = metafileRasterSize(
          fill.mimeType,
          fill.srcRect,
          size.widthPt,
          size.heightPt,
        );
        if (!raster) continue;
        const request: ImageDecodeRequest = {
          cacheKey: chartImageFillKey(fill),
          imagePath: fill.imagePath,
          mimeType: fill.mimeType,
          ...(fill.svgImagePath === undefined ? {} : { svgImagePath: fill.svgImagePath }),
          ...(fill.duotone === undefined ? {} : { duotone: fill.duotone }),
          // An ordered chain samples the source grid at the fill DPI (see core
          // chartImageFillPixelTransform), independent of any usage frame.
          ...(fill.blipEffects?.length ? {
            blipEffects: fill.blipEffects,
            orientation: (transform => transform && 'effects' in transform
              ? { emuPerPixelX: transform.emuPerPixelX, emuPerPixelY: transform.emuPerPixelY }
              : {})(chartImageFillPixelTransform(fill)),
          } : {}),
          widthPt: fill.blipEffects?.length ? 0 : raster.widthPt,
          heightPt: fill.blipEffects?.length ? 0 : raster.heightPt,
          hasCrop: fill.srcRect != null,
          chartPictureFill: true,
          ...(!usage.preserveNaturalSize && !fill.blipEffects?.length && size.targetWidthPx && size.targetHeightPx
            ? { targetWidthPx: size.targetWidthPx, targetHeightPx: size.targetHeightPx }
            : {}),
        };
        const key = requestKey(request);
        const existing = requests.get(key);
        if (!existing) {
          requests.set(key, request);
        } else {
          existing.widthPt = Math.max(existing.widthPt, request.widthPt);
          existing.heightPt = Math.max(existing.heightPt, request.heightPt);
          existing.hasCrop ||= request.hasCrop;
        }
        if (usage.preserveNaturalSize) naturalSizeChartKeys.add(key);
        const merged = requests.get(key) as ImageDecodeRequest;
        if (naturalSizeChartKeys.has(key)) {
          merged.targetWidthPx = undefined;
          merged.targetHeightPx = undefined;
        } else {
          merged.targetWidthPx = Math.max(
            merged.targetWidthPx ?? 0,
            request.targetWidthPx ?? 0,
          ) || undefined;
          merged.targetHeightPx = Math.max(
            merged.targetHeightPx ?? 0,
            request.targetHeightPx ?? 0,
          ) || undefined;
        }
      }
    }
  }
  for (const key of naturalSizeChartKeys) {
    const request = requests.get(key);
    if (request) { request.targetWidthPx = undefined; request.targetHeightPx = undefined; }
  }
  return [...requests.values()];
}

export async function preloadPaintImages(
  descriptors: readonly DeepReadonly<PaintResourceDescriptor>[],
  rasterPaintOccurrences: readonly DeepReadonly<RasterPaintOccurrence>[],
  fetchImage: DocxFetchImage | undefined,
  tiff?: TiffRenderer,
  devicePixelsPerPoint?: number,
  svgDecoder?: import('@silurus/ooxml-core').SvgBlobDecoder,
  imageResources?: ImageResourceOptions,
  onDiagnostic?: (diagnostic: import('../layout/types.js').LayoutDiagnostic) => void,
): Promise<Map<string, DecodedPaintImage>> {
  if (!fetchImage) return new Map();
  const policy = normalizeImageResourceOptions(imageResources);
  // SVG targets are coverage minimums: the rasterizer keeps the SVG's aspect,
  // so a square SVG in a 96x87 box retains 96x96. Core admits exactly that
  // coverage grid (plus the per-image hard quota); a `width x height` ceiling
  // here would reject every aspect-mismatched picture in worker mode only,
  // while main mode draws the same SVG element without a raster surface.
  const decodeSvg = (path: string, request: ImageDecodeRequest) => svgDecoder
    ? getCachedSvgImageByPath(path, fetchImage, {
        targetWidthPx: request.targetWidthPx,
        targetHeightPx: request.targetHeightPx,
        workerDecoder: svgDecoder,
      })
    : getCachedSvgImageByPath(path, fetchImage);
  const requests = imageDecodeRequests(
    descriptors,
    rasterPaintOccurrences,
    devicePixelsPerPoint,
  );
  const demands = (await Promise.all(requests.map(async (request) => {
    if (!request.targetWidthPx || !request.targetHeightPx) return null;
    // DrawingML pixel effects consume the authored source grid. They may
    // resample only after the exact transform, so an adaptive pre-decode target
    // would change semantics instead of merely reducing display quality.
    if (request.colorReplaceFrom || request.duotone || request.vmlStrokePattern || request.blipEffects?.length) return null;
    const dataIsSvg = request.mimeType === 'image/svg+xml';
    const blip = { svgImagePath: request.svgImagePath, srcRect: request.hasCrop || null };
    if (dataIsSvg || preferVectorBlip(blip)) return null;
    if (isBrowserResizableRasterMimeType(request.mimeType)
      && (policy.resolution === 'display' || policy.strategy === 'adaptive')) {
      return {
        key: requestKey(request),
        targetWidthPx: request.targetWidthPx,
        targetHeightPx: request.targetHeightPx,
        retainedSurfaceCount: 1,
      };
    }
    const inspection = await inspectCachedRasterSource(
      request.imagePath,
      request.mimeType,
      fetchImage,
    ).catch((error: unknown) => {
      if (error instanceof OoxmlResourceLimitError || isOoxmlDecodedImageLimitError(error)) throw error;
      return null;
    });
    if (!inspection?.dimensions
      || !isDecodeTargetResizableRasterFormat(inspection.format, tiff !== undefined)) return null;
    return {
      key: requestKey(request),
      targetWidthPx: request.targetWidthPx,
      targetHeightPx: request.targetHeightPx,
      sourceWidthPx: inspection.dimensions.width,
      sourceHeightPx: inspection.dimensions.height,
      retainedSurfaceCount: 1,
    };
  }))).filter((demand): demand is NonNullable<typeof demand> => demand !== null);
  const plan = planDecodedImageTargets(demands, policy);
  for (const request of requests) {
    const dataIsSvg = request.mimeType === 'image/svg+xml';
    const usesVector = dataIsSvg || (!request.colorReplaceFrom && !request.duotone && !request.vmlStrokePattern && !request.blipEffects?.length
      && preferVectorBlip({
        svgImagePath: request.svgImagePath,
        srcRect: request.hasCrop || null,
      }));
    // Vector rasterization is inherently target-sized. Raster/metafile effects,
    // natural-size consumers, and formats not admitted by the shared planner
    // must retain their authored source grid instead of inheriting the raw
    // display target collected above.
    if (usesVector) continue;
    const target = plan.targets.get(requestKey(request));
    request.targetWidthPx = target?.width;
    request.targetHeightPx = target?.height;
    request.plannedPixelLimit = target?.maxRetainedPixels;
  }
  const entries = await Promise.all(requests.map(async (request) => {
    try {
      const dataIsSvg = request.mimeType === 'image/svg+xml';
      const blip = { svgImagePath: request.svgImagePath, srcRect: request.hasCrop || null };
      let image: DecodedImage | null;
      if (preferVectorBlip(blip) && !request.colorReplaceFrom && !request.duotone && !request.vmlStrokePattern && !request.blipEffects?.length) {
        try {
          image = await decodeSvg(blip.svgImagePath, request);
        } catch (vectorError) {
          if (isTerminalImageAcquisitionError(vectorError)) throw vectorError;
          const fallback = dataIsSvg
            ? await decodeSvg(request.imagePath, request)
            : await decodeRaster(
                request.imagePath,
                request.mimeType,
                request.colorReplaceFrom,
                fetchImage,
                request.widthPt,
                request.heightPt,
                request.duotone,
                tiff,
                request.targetWidthPx && request.targetHeightPx
                  ? { targetWidthPx: request.targetWidthPx, targetHeightPx: request.targetHeightPx }
                  : undefined,
                request.plannedPixelLimit,
                svgDecoder, request.vmlStrokePattern, request.blipEffects, request.orientation,
              );
          if (!fallback) throw vectorError;
          image = fallback;
        }
      } else if (dataIsSvg && !request.colorReplaceFrom && !request.duotone && !request.vmlStrokePattern && !request.blipEffects?.length) {
        image = await decodeSvg(request.imagePath, request);
      } else {
        const decodeEffect = (path: string, mime: string) => decodeRaster(
          path, mime, request.colorReplaceFrom, fetchImage,
          request.widthPt, request.heightPt, request.duotone, tiff,
          request.targetWidthPx && request.targetHeightPx
            ? { targetWidthPx: request.targetWidthPx, targetHeightPx: request.targetHeightPx }
            : undefined,
          request.plannedPixelLimit, svgDecoder, request.vmlStrokePattern, request.blipEffects, request.orientation,
        );
        if (preferVectorBlip(blip)) {
          try { image = await decodeEffect(blip.svgImagePath, 'image/svg+xml'); }
          catch (error) {
            if (isTerminalImageAcquisitionError(error)) throw error;
            image = await decodeEffect(request.imagePath, request.mimeType);
          }
        } else {
          image = await decodeEffect(request.imagePath, request.mimeType);
        }
      }
      if (image == null && request.chartPictureFill) {
        onDiagnostic?.({ code: 'MISSING_RESOURCE', severity: 'warning',
          message: 'A chart picture fill produced no drawable output; the remaining chart content is rendered' });
        return [requestKey(request), null] as const;
      }
      return image == null ? null : [requestKey(request), image] as const;
    } catch (error) {
      if (isOptionalImageCodecUnavailableError(error, 'tiff')) {
        return [requestKey(request), error] as const;
      }
      if (isTerminalImageAcquisitionError(error)) throw error;
      if (request.chartPictureFill) {
        onDiagnostic?.({ code: 'MISSING_RESOURCE', severity: 'warning',
          message: 'A chart picture fill could not be acquired; the remaining chart content is rendered' });
        return [requestKey(request), null] as const;
      }
      throw error;
    }
  }));
  const decoded = new Map<string, DecodedPaintImage>();
  for (const entry of entries) {
    if (entry) decoded.set(entry[0], entry[1]);
  }
  // §20.1.8.58 cells derive from the decoded sources above.
  for (const descriptor of descriptors) {
    if ((descriptor.kind !== 'image' && descriptor.kind !== 'picture-bullet')
      || !drawingMLTileNeedsCell(descriptor)) continue;
    const key = drawingMLTileCellKey(descriptor);
    if (decoded.has(key)) continue;
    const source = decoded.get(paintImageDecodeKey(descriptor));
    if (!source || isOptionalImageCodecUnavailableError(source, 'tiff')) continue;
    decoded.set(key, await drawingMLTileCell(descriptor, source as DecodedImage, fetchImage));
  }
  return decoded;
}
