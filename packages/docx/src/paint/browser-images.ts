import {
  captureDecodedBitmapCacheEpoch,
  decodedBitmapTargetResizeOptions,
  duotoneImageData,
  MAX_RASTER_PIXELS,
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
  chartImageFillUsageSize,
  collectChartImageFillUsages,
  isOptionalImageCodecUnavailableError,
} from '@silurus/ooxml-core';
import type {
  Duotone,
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

export type DecodedImage = ImageBitmap | HTMLImageElement;
export type DecodedPaintImage = DecodedImage | OptionalImageCodecUnavailableError;
export type DocxFetchImage = (path: string, mime: string) => Promise<Blob>;

interface ImageDecodeRequest {
  cacheKey?: string;
  imagePath: string;
  mimeType: string;
  svgImagePath?: string;
  colorReplaceFrom?: string;
  duotone?: Duotone;
  widthPt: number;
  heightPt: number;
  hasCrop: boolean;
  targetWidthPx?: number;
  targetHeightPx?: number;
  /** Set only for browser-resizable rasters admitted by the shared plan. */
  plannedPixelLimit?: number;
  failClosedOnDuotoneFailure?: boolean;
}

export function imageKey(
  imagePath: string,
  colorReplaceFrom?: string,
  duotone?: Duotone,
): string {
  const clr = colorReplaceFrom ? `|clr:${colorReplaceFrom}` : '';
  const duo = duotone ? `|duo:${duotone.clr1}:${duotone.clr2}` : '';
  return `${imagePath}${clr}${duo}`;
}

function requestKey(request: ImageDecodeRequest): string {
  return request.cacheKey
    ?? imageKey(request.imagePath, request.colorReplaceFrom, request.duotone);
}

const DOCX_COLOR_EFFECT_CACHE_NAMESPACE = 'docx-color-effects';

export function dropBrowserImageCache(fetchImage: DocxFetchImage): void {
  dropCachedDerivedBitmapNamespace(fetchImage, DOCX_COLOR_EFFECT_CACHE_NAMESPACE);
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
  failClosedOnDuotoneFailure: boolean,
  target?: Readonly<{ targetWidthPx: number; targetHeightPx: number }>,
): Promise<ImageBitmap | null> {
  const unavailable = async (error?: unknown): Promise<ImageBitmap | null> => {
    // clrChange has no compatibility pass-through: silently dropping its exact
    // alpha transform would change authored content. Duotone-only consumers
    // retain their established compatibility behavior, but at the requested
    // display resolution; strict chart consumers still fail closed.
    if (colorReplaceFrom) {
      throw error instanceof Error
        ? error
        : new Error('2D canvas is unavailable for image color effects');
    }
    if (failClosedOnDuotoneFailure) return null;
    const resizeOptions = decodedBitmapTargetResizeOptions(
      bitmap.width,
      bitmap.height,
      target?.targetWidthPx,
      target?.targetHeightPx,
    );
    if (!resizeOptions) return bitmap;
    if (typeof createImageBitmap === 'undefined') {
      throw new Error('createImageBitmap is unavailable for duotone fallback resampling');
    }
    return createImageBitmap(bitmap, resizeOptions);
  };
  if (typeof OffscreenCanvas === 'undefined') return unavailable();
  let offscreen: OffscreenCanvas;
  let context: OffscreenCanvasRenderingContext2D | null;
  try {
    offscreen = new OffscreenCanvas(bitmap.width, bitmap.height);
    context = offscreen.getContext('2d');
  } catch (error) {
    return unavailable(error);
  }
  if (!context) return unavailable();
  context.drawImage(bitmap, 0, 0);
  let imageData: ImageData;
  try {
    imageData = context.getImageData(0, 0, bitmap.width, bitmap.height);
  } catch (error) {
    return unavailable(error);
  }
  if (colorReplaceFrom) applyExactColorReplacement(imageData, colorReplaceFrom);
  if (duotone) {
    try {
      duotoneImageData(imageData, duotone.clr1, duotone.clr2);
    } catch {
      // The exact clrChange mutation already lives in this one source-grid
      // buffer. Compatibility mode preserves and resamples that current result;
      // strict chart consumers must not draw it without the authored duotone.
      if (failClosedOnDuotoneFailure) return null;
    }
  }
  context.putImageData(imageData, 0, 0);
  const resizeOptions = decodedBitmapTargetResizeOptions(
    bitmap.width,
    bitmap.height,
    target?.targetWidthPx,
    target?.targetHeightPx,
  );
  return resizeOptions
    ? createImageBitmap(offscreen, resizeOptions)
    : createImageBitmap(offscreen);
}

export async function decodeRaster(
  imagePath: string,
  mimeType: string,
  colorReplaceFrom: string | undefined,
  fetchImage: DocxFetchImage,
  widthPt = 0,
  heightPt = 0,
  duotone?: Duotone,
  failClosedOnDuotoneFailure = false,
  tiff?: TiffRenderer,
  target?: Readonly<{ targetWidthPx: number; targetHeightPx: number }>,
  plannedPixelLimit?: number,
): Promise<ImageBitmap | null> {
  // Pixel effects temporarily retain more than their cached input/output:
  // source + offscreen backing + ImageData + result = four surfaces. clrChange
  // and duotone mutate the same ImageData before the one final bitmap bake, so
  // chaining them does not add another full-size intermediate.
  const effectSurfaceCount = colorReplaceFrom || duotone ? 4 : 1;
  const effectPixelLimit = Math.floor(MAX_RASTER_PIXELS / effectSurfaceCount);
  const maxRetainedPixels = Number.isSafeInteger(plannedPixelLimit) && (plannedPixelLimit ?? 0) > 0
    ? Math.min(effectPixelLimit, plannedPixelLimit as number)
    : effectPixelLimit;
  const epoch = colorReplaceFrom || duotone
    ? captureDecodedBitmapCacheEpoch(fetchImage, DOCX_COLOR_EFFECT_CACHE_NAMESPACE)
    : undefined;
  const sourceBitmapOptions = {
    widthPt,
    heightPt,
    suppressBoundaryFrame: true,
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
  if (!colorReplaceFrom && !duotone) return base;
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
  const key = `${imageKey(resolvedBaseKey, colorReplaceFrom, duotone)}${resizeKey}${failClosedOnDuotoneFailure ? '|strict' : ''}`;
  return getCachedDerivedBitmap(
    DOCX_COLOR_EFFECT_CACHE_NAMESPACE,
    key,
    fetchImage,
    async () => {
      const bitmap = await applyColorEffects(
        base,
        colorReplaceFrom,
        duotone,
        failClosedOnDuotoneFailure,
        target,
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
  const demandByResource = new Map<string, { widthPt: number; heightPt: number }>();
  for (const occurrence of rasterPaintOccurrences) {
    if (occurrence.resourceKind !== 'image' && occurrence.resourceKind !== 'picture-bullet') {
      continue;
    }
    if (!Number.isFinite(occurrence.widthPt) || occurrence.widthPt <= 0
      || !Number.isFinite(occurrence.heightPt) || occurrence.heightPt <= 0) continue;
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
    const raster = metafileRasterSize(
      image.mimeType,
      image.srcRect,
      demand.widthPt,
      demand.heightPt,
    );
    if (!raster) continue;
    const request: ImageDecodeRequest = {
      imagePath: image.partPath,
      mimeType: image.mimeType,
      ...(image.svgImagePath === undefined ? {} : { svgImagePath: image.svgImagePath }),
      ...(image.colorReplaceFrom === undefined ? {} : { colorReplaceFrom: image.colorReplaceFrom }),
      ...(image.duotone === undefined ? {} : { duotone: image.duotone as Duotone }),
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
  const naturalSizeChartKeys = new Set<string>();
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
      if (!Number.isFinite(occurrence.widthPt) || occurrence.widthPt <= 0
        || !Number.isFinite(occurrence.heightPt) || occurrence.heightPt <= 0) continue;
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
          widthPt: raster.widthPt,
          heightPt: raster.heightPt,
          hasCrop: fill.srcRect != null,
          failClosedOnDuotoneFailure: true,
          ...(!usage.preserveNaturalSize && size.targetWidthPx && size.targetHeightPx
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
): Promise<Map<string, DecodedPaintImage>> {
  if (!fetchImage) return new Map();
  const policy = normalizeImageResourceOptions(imageResources);
  const decodeSvg = (path: string, request: ImageDecodeRequest) => svgDecoder
    ? getCachedSvgImageByPath(path, fetchImage, {
        targetWidthPx: request.targetWidthPx,
        targetHeightPx: request.targetHeightPx,
        maxRetainedPixels: request.targetWidthPx && request.targetHeightPx
          ? request.targetWidthPx * request.targetHeightPx
          : undefined,
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
    if (request.colorReplaceFrom || request.duotone) return null;
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
    ).catch(() => null);
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
    const usesVector = dataIsSvg || (!request.colorReplaceFrom && !request.duotone
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
      if (preferVectorBlip(blip)) {
        try {
          image = await decodeSvg(blip.svgImagePath, request);
        } catch (vectorError) {
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
                request.failClosedOnDuotoneFailure ?? false,
                tiff,
                request.targetWidthPx && request.targetHeightPx
                  ? { targetWidthPx: request.targetWidthPx, targetHeightPx: request.targetHeightPx }
                  : undefined,
                request.plannedPixelLimit,
              );
          if (!fallback) throw vectorError;
          image = fallback;
        }
      } else if (dataIsSvg) {
        image = await decodeSvg(request.imagePath, request);
      } else {
        image = await decodeRaster(
          request.imagePath,
          request.mimeType,
          request.colorReplaceFrom,
          fetchImage,
          request.widthPt,
          request.heightPt,
          request.duotone,
          request.failClosedOnDuotoneFailure ?? false,
          tiff,
          request.targetWidthPx && request.targetHeightPx
            ? { targetWidthPx: request.targetWidthPx, targetHeightPx: request.targetHeightPx }
            : undefined,
          request.plannedPixelLimit,
        );
      }
      return image == null
        ? null
        : [requestKey(request), image] as const;
    } catch (error) {
      if (isOptionalImageCodecUnavailableError(error, 'tiff')) {
        return [requestKey(request), error] as const;
      }
      throw error;
    }
  }));
  const decoded = new Map<string, DecodedPaintImage>();
  for (const entry of entries) {
    if (entry) decoded.set(entry[0], entry[1]);
  }
  return decoded;
}
