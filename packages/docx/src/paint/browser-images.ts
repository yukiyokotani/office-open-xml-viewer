import {
  applyDuotone,
  dropCachedDerivedBitmapNamespace,
  getCachedBitmapByPath,
  getCachedDerivedBitmap,
  getCachedSvgImageByPath,
  imageNaturalSize,
  metafileRasterSize,
  preferVectorBlip,
  releaseOwnedBitmap,
} from '@silurus/ooxml-core';
import type { Duotone, TiffRenderer } from '@silurus/ooxml-core';
import type {
  DeepReadonly,
  ImagePaintResourceDescriptor,
  PaintResourceDescriptor,
} from '../layout/types.js';

export type DecodedImage = ImageBitmap | HTMLImageElement;
export type DocxFetchImage = (path: string, mime: string) => Promise<Blob>;

interface ImageDecodeRequest {
  imagePath: string;
  mimeType: string;
  svgImagePath?: string;
  colorReplaceFrom?: string;
  duotone?: Duotone;
  widthPt: number;
  heightPt: number;
  hasCrop: boolean;
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

const DOCX_COLOR_EFFECT_CACHE_NAMESPACE = 'docx-color-effects';

export function dropBrowserImageCache(fetchImage: DocxFetchImage): void {
  dropCachedDerivedBitmapNamespace(fetchImage, DOCX_COLOR_EFFECT_CACHE_NAMESPACE);
}

async function applyColorReplacement(
  bitmap: ImageBitmap,
  colorHex: string,
): Promise<ImageBitmap> {
  const red = parseInt(colorHex.slice(0, 2), 16);
  const green = parseInt(colorHex.slice(2, 4), 16);
  const blue = parseInt(colorHex.slice(4, 6), 16);
  const offscreen = new OffscreenCanvas(bitmap.width, bitmap.height);
  const context = offscreen.getContext('2d');
  if (!context) throw new Error('2D canvas is unavailable for image color replacement');
  context.drawImage(bitmap, 0, 0);
  const imageData = context.getImageData(0, 0, bitmap.width, bitmap.height);
  for (let index = 0; index < imageData.data.length; index += 4) {
    if (imageData.data[index] === red
      && imageData.data[index + 1] === green
      && imageData.data[index + 2] === blue) {
      imageData.data[index + 3] = 0;
    }
  }
  context.putImageData(imageData, 0, 0);
  return createImageBitmap(offscreen);
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
): Promise<ImageBitmap | null> {
  const base = await getCachedBitmapByPath(imagePath, mimeType, fetchImage, {
    widthPt,
    heightPt,
    suppressBoundaryFrame: true,
    tiff,
  });
  if (!base) return null;
  if (!colorReplaceFrom && !duotone) return base;
  const key = `${imageKey(imagePath, colorReplaceFrom, duotone)}${failClosedOnDuotoneFailure ? '|strict' : ''}`;
  return getCachedDerivedBitmap(
    DOCX_COLOR_EFFECT_CACHE_NAMESPACE,
    key,
    fetchImage,
    async () => {
      let bitmap: ImageBitmap = base;
      try {
        if (colorReplaceFrom) bitmap = await applyColorReplacement(bitmap, colorReplaceFrom);
        if (duotone) {
          const { w, h } = imageNaturalSize(bitmap);
          if (w > 0 && h > 0) {
            const source = bitmap;
            const recoloured = await applyDuotone(bitmap, duotone, { width: w, height: h });
            if (failClosedOnDuotoneFailure && recoloured === source) {
              if (source !== base) releaseOwnedBitmap(source);
              return { bitmap: null, owned: false };
            }
            bitmap = recoloured as ImageBitmap;
            // A clrChange surface is only an intermediate when duotone creates a
            // second surface. It is never cached or drawn, so release it now.
            if (source !== base && bitmap !== source) releaseOwnedBitmap(source);
          }
        }
        return { bitmap, owned: bitmap !== base };
      } catch (error) {
        if (bitmap !== base) releaseOwnedBitmap(bitmap);
        throw error;
      }
    },
  );
}

function imageDecodeRequests(
  descriptors: readonly DeepReadonly<PaintResourceDescriptor>[],
): ImageDecodeRequest[] {
  const requests = new Map<string, ImageDecodeRequest>();
  const images = descriptors
    .filter((descriptor): descriptor is DeepReadonly<ImagePaintResourceDescriptor> => (
      descriptor.kind === 'image' || descriptor.kind === 'picture-bullet'
    ))
    .sort((left, right) => (
      (left.documentOrder ?? Number.MAX_SAFE_INTEGER)
      - (right.documentOrder ?? Number.MAX_SAFE_INTEGER)
    ));
  for (const image of images) {
    const raster = metafileRasterSize(
      image.mimeType,
      image.srcRect,
      image.intrinsicSize.widthPt,
      image.intrinsicSize.heightPt,
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
    const key = imageKey(request.imagePath, request.colorReplaceFrom, request.duotone);
    const existing = requests.get(key);
    if (!existing) {
      requests.set(key, request);
    } else {
      existing.widthPt = Math.max(existing.widthPt, request.widthPt);
      existing.heightPt = Math.max(existing.heightPt, request.heightPt);
      existing.hasCrop ||= request.hasCrop;
    }
  }
  return [...requests.values()];
}

export async function preloadPaintImages(
  descriptors: readonly DeepReadonly<PaintResourceDescriptor>[],
  fetchImage: DocxFetchImage | undefined,
  tiff?: TiffRenderer,
): Promise<Map<string, DecodedImage>> {
  if (!fetchImage) return new Map();
  const entries = await Promise.all(imageDecodeRequests(descriptors).map(async (request) => {
    const dataIsSvg = request.mimeType === 'image/svg+xml';
    const blip = { svgImagePath: request.svgImagePath, srcRect: request.hasCrop || null };
    let image: DecodedImage | null;
    if (preferVectorBlip(blip)) {
      try {
        image = await getCachedSvgImageByPath(blip.svgImagePath, fetchImage);
      } catch (vectorError) {
        const fallback = dataIsSvg
          ? await getCachedSvgImageByPath(request.imagePath, fetchImage)
          : await decodeRaster(
              request.imagePath,
              request.mimeType,
              request.colorReplaceFrom,
              fetchImage,
              request.widthPt,
              request.heightPt,
              request.duotone,
              false,
              tiff,
            );
        if (!fallback) throw vectorError;
        image = fallback;
      }
    } else if (dataIsSvg) {
      image = await getCachedSvgImageByPath(request.imagePath, fetchImage);
    } else {
      image = await decodeRaster(
        request.imagePath,
        request.mimeType,
        request.colorReplaceFrom,
        fetchImage,
        request.widthPt,
        request.heightPt,
        request.duotone,
        false,
        tiff,
      );
    }
    return image == null
      ? null
      : [imageKey(request.imagePath, request.colorReplaceFrom, request.duotone), image] as const;
  }));
  return new Map(entries.filter((entry): entry is readonly [string, DecodedImage] => entry !== null));
}
