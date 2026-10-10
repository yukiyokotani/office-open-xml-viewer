// Shared `<a:srcRect>` source-rectangle crop (ECMA-376 §20.1.8.55) for the docx,
// pptx and xlsx renderers. The crop is a fraction of the image's NATIVE pixel
// grid. Browser rasters may use a display-sized grid because the same fractional
// mapping is preserved under axis-wise resampling; metafiles must rasterize the
// whole picture FRAME (see `metafileRasterSize`) because their player maps that
// authored frame onto the output before the crop. Centralised here so all three
// renderers crop identically (previously triplicated, and metafiles diverged).

import { isMetafileMime } from './wmf';

type AnyCtx = CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;

/** A `<a:srcRect>` crop: signed fractional insets measured from each edge.
 *  The visible region is `[l, t, 1−r, 1−b]` of the source. `ST_Percentage` is
 *  not range-limited, so negative and greater-than-one fractions are retained. */
export interface SrcRect {
  l: number;
  t: number;
  r: number;
  b: number;
}

/** Whether the logical source rectangle intersects the bitmap at positive
 * area. Shared by decode preflight and paint so a fully cropped blip never
 * allocates an oversized fallback raster. */
export function srcRectHasVisibleArea(srcRect: SrcRect | null | undefined): boolean {
  if (!srcRect) return true;
  const values = [srcRect.l, srcRect.t, srcRect.r, srcRect.b];
  if (!values.every(Number.isFinite)) return false;
  const x0 = srcRect.l;
  const y0 = srcRect.t;
  const x1 = 1 - srcRect.r;
  const y1 = 1 - srcRect.b;
  return x1 > x0 && y1 > y0
    && Math.min(1, x1) > Math.max(0, x0)
    && Math.min(1, y1) > Math.max(0, y0);
}

/** Native pixel size of a decoded image (ImageBitmap exposes `width`/`height`;
 *  an `<img>` element exposes `naturalWidth`/`naturalHeight`). */
export function imageNaturalSize(img: CanvasImageSource): { w: number; h: number } {
  const el = img as {
    naturalWidth?: number;
    naturalHeight?: number;
    width?: number;
    height?: number;
  };
  const w = el.naturalWidth || (typeof el.width === 'number' ? el.width : 0) || 0;
  const h = el.naturalHeight || (typeof el.height === 'number' ? el.height : 0) || 0;
  return { w, h };
}

/** Immutable normalized source intersection and destination projection. */
export interface ImageCropProjection {
  readonly sourceX: number; readonly sourceY: number;
  readonly sourceWidth: number; readonly sourceHeight: number;
  readonly dxFraction: number; readonly dyFraction: number;
  readonly dwFraction: number; readonly dhFraction: number;
}

interface CropMapping {
  sx: number; sy: number; sw: number; sh: number;
  dxFraction: number; dyFraction: number; dwFraction: number; dhFraction: number;
}

/** One §20.1.8.55 calculation for ordinary and retained image painting.
 * Strict acquisition distinguishes numerical overflow from legitimate empty
 * crop; ordinary callers retain their historical invalid/empty behavior. */
export function imageCropProjection(
  srcRect: SrcRect | null | undefined, strict = false,
): Readonly<ImageCropProjection> | null {
  if (!srcRect) return null;
  const values = [srcRect.l, srcRect.t, srcRect.r, srcRect.b];
  if (!values.every(Number.isFinite)) {
    if (strict) throw new RangeError('Image crop must contain finite fractions');
    return null;
  }
  if (!(srcRect.l || srcRect.t || srcRect.r || srcRect.b)) return null;
  const logicalX0 = srcRect.l, logicalY0 = srcRect.t;
  const logicalX1 = 1 - srcRect.r, logicalY1 = 1 - srcRect.b;
  const logicalW = logicalX1 - logicalX0, logicalH = logicalY1 - logicalY0;
  // Checking output fractions alone would accept l=r=-1e308: an infinite
  // logical width can hide behind finite zero fractions.
  if (strict && ![logicalX1, logicalY1, logicalW, logicalH].every(Number.isFinite)) {
    throw new RangeError('Image crop intermediate overflow');
  }
  if (!(logicalW > 0) || !(logicalH > 0)) return Object.freeze({
    sourceX: 0, sourceY: 0, sourceWidth: 0, sourceHeight: 0,
    dxFraction: 0, dyFraction: 0, dwFraction: 0, dhFraction: 0,
  });
  const sourceX0 = Math.max(0, logicalX0), sourceY0 = Math.max(0, logicalY0);
  const sourceX1 = Math.min(1, logicalX1), sourceY1 = Math.min(1, logicalY1);
  const sourceW = Math.max(0, sourceX1 - sourceX0), sourceH = Math.max(0, sourceY1 - sourceY0);
  const plan = Object.freeze({
    sourceX: sourceX0, sourceY: sourceY0, sourceWidth: sourceW, sourceHeight: sourceH,
    dxFraction: (sourceX0 - logicalX0) / logicalW,
    dyFraction: (sourceY0 - logicalY0) / logicalH,
    dwFraction: sourceW / logicalW, dhFraction: sourceH / logicalH,
  });
  if (strict && !Object.values(plan).every(Number.isFinite)) throw new RangeError('Image crop projection overflow');
  return plan;
}

/** The decoded grid supplies pixels only, never retained layout geometry. */
export function cropSourceMapping(
  img: CanvasImageSource, srcRect: SrcRect | null | undefined,
  projection?: Readonly<ImageCropProjection> | null,
): CropMapping | null {
  const plan = projection === undefined ? imageCropProjection(srcRect) : projection;
  if (!plan) return null;
  const { w, h } = imageNaturalSize(img);
  if (w <= 0 || h <= 0) return null;
  return {
    sx: plan.sourceX * w, sy: plan.sourceY * h,
    sw: plan.sourceWidth * w, sh: plan.sourceHeight * h,
    dxFraction: plan.dxFraction, dyFraction: plan.dyFraction,
    dwFraction: plan.dwFraction, dhFraction: plan.dhFraction,
  };
}

export interface ImagePaintTransform {
  readonly rotation?: number; readonly flipH?: boolean; readonly flipV?: boolean; readonly alpha?: number;
}
export interface ImagePaintProjection {
  readonly width: number; readonly height: number;
  readonly transformed: boolean;
  readonly centerX: number; readonly centerY: number; readonly rotationRad: number;
  readonly scaleX: 1 | -1; readonly scaleY: 1 | -1; readonly alpha: number;
  readonly destination: Readonly<{ x: number; y: number; width: number; height: number }>;
  readonly crop: Readonly<ImageCropProjection> | null;
}
export interface ImagePaintProjectionOptions {
  readonly transform?: Readonly<ImagePaintTransform>;
  readonly projection?: Readonly<ImagePaintProjection>;
}

/** Canonical local-coordinate crop/transform projection. This is a pure
 * drawing primitive, not a layout or contour algorithm. Layout retains its
 * plan and proves allocation; ordinary painting uses the same constructor. */
export function imagePaintProjection(
  width: number, height: number, srcRect: SrcRect | null | undefined,
  transform: Readonly<ImagePaintTransform> = {}, strict = true,
): Readonly<ImagePaintProjection> {
  const rotation = transform.rotation ?? 0;
  const rotationProduct = rotation * Math.PI;
  const rotationRad = rotationProduct / 180;
  const transformed = rotation !== 0 || Boolean(transform.flipH) || Boolean(transform.flipV);
  const centerX = width / 2, centerY = height / 2;
  if (strict && (!Number.isFinite(width) || !Number.isFinite(height) || width < 0 || height < 0
    || ![rotation, rotationProduct, rotationRad, centerX, centerY].every(Number.isFinite)
    || (transform.flipH !== undefined && typeof transform.flipH !== 'boolean')
    || (transform.flipV !== undefined && typeof transform.flipV !== 'boolean')
    || (transform.alpha !== undefined && (!Number.isFinite(transform.alpha) || transform.alpha < 0 || transform.alpha > 1)))) {
    throw new RangeError('Image projection has invalid or overflowed transform');
  }
  const crop = imageCropProjection(srcRect, strict);
  const destination = Object.freeze({ x: transformed ? -centerX : 0, y: transformed ? -centerY : 0, width, height });
  // No clamping of authored fractions. With a nonempty crop, the canonical
  // source intersection is mathematically inside the full destination frame;
  // the bound admits only rounding of its subtraction/division/addition.
  if (strict && crop && crop.sourceWidth > 0 && crop.sourceHeight > 0) {
    const precision = 4 * Number.EPSILON;
    if (crop.dxFraction < 0 || crop.dyFraction < 0 || crop.dwFraction < 0 || crop.dhFraction < 0
      || crop.dxFraction + crop.dwFraction > 1 + precision || crop.dyFraction + crop.dhFraction > 1 + precision
      || ![crop.dxFraction * width, crop.dyFraction * height, crop.dwFraction * width, crop.dhFraction * height].every(Number.isFinite)) {
      throw new RangeError('Image crop exceeds its full destination frame');
    }
  }
  return Object.freeze({ width, height, transformed, centerX, centerY, rotationRad,
    scaleX: transform.flipH ? -1 : 1, scaleY: transform.flipV ? -1 : 1,
    alpha: transform.alpha ?? 1, destination, crop });
}

/** Apply retained drawing transforms only. No recorder, bounds or measurement
 * result is produced. Origins may move while the immutable local plan stays. */
export function withImagePaintProjection(
  ctx: AnyCtx, x: number, y: number, plan: Readonly<ImagePaintProjection>,
  draw: (x: number, y: number, width: number, height: number) => void,
): void {
  const hasAlpha = plan.alpha < 1;
  if (hasAlpha) { ctx.save(); ctx.globalAlpha *= plan.alpha; }
  try {
    if (!plan.transformed) {
      draw(x + plan.destination.x, y + plan.destination.y, plan.width, plan.height);
    } else {
      ctx.save();
      try {
        ctx.translate(x + plan.centerX, y + plan.centerY);
        ctx.rotate(plan.rotationRad); ctx.scale(plan.scaleX, plan.scaleY);
        draw(plan.destination.x, plan.destination.y, plan.width, plan.height);
      } finally { ctx.restore(); }
    }
  } finally { if (hasAlpha) ctx.restore(); }
}

/** The 9-arg `drawImage` source rectangle for an `<a:srcRect>` crop, or `null`
 *  when there is no (non-empty) crop or the image reports no native size.
 *
 *  The returned rectangle is the intersection of the logical source rectangle
 *  with the actual bitmap. Negative edges therefore return the full affected
 *  source range; {@link drawImageCropped} additionally preserves their outset
 *  as transparent destination space. Callers
 *  that need the rect for auxiliary paints (e.g. pptx effect passes) call this
 *  directly; the common path uses {@link drawImageCropped}. */
export function cropSourceRect(
  img: CanvasImageSource,
  srcRect: SrcRect | null | undefined,
): { sx: number; sy: number; sw: number; sh: number } | null {
  const mapping = cropSourceMapping(img, srcRect);
  if (!mapping) return null;
  return { sx: mapping.sx, sy: mapping.sy, sw: mapping.sw, sh: mapping.sh };
}

/** Draw `img` into the destination box `[dx, dy, dw, dh]`, honoring an optional
 *  `<a:srcRect>` crop. The destination box is unchanged — the visible slice is
 *  stretched to fill it (the 9-arg `drawImage` behavior). Crop applies to raster
 *  blips AND metafiles alike: a cropped metafile must have been rasterized at its
 *  full frame via {@link metafileRasterSize}, so its bitmap is the full source. */
export function drawImageCropped(
  ctx: AnyCtx, img: CanvasImageSource, srcRect: SrcRect | null | undefined,
  dx: number, dy: number, dw: number, dh: number,
  options?: Readonly<ImagePaintProjectionOptions>,
): void {
  const paint = (x: number, y: number, width: number, height: number, projection?: Readonly<ImageCropProjection> | null) => {
    const c = cropSourceMapping(img, srcRect, projection);
    if (!c) ctx.drawImage(img, x, y, width, height);
    else if (c.sw > 0 && c.sh > 0 && c.dwFraction > 0 && c.dhFraction > 0) {
      ctx.drawImage(img, c.sx, c.sy, c.sw, c.sh,
        x + c.dxFraction * width, y + c.dyFraction * height, c.dwFraction * width, c.dhFraction * height);
    }
  };
  // Keep all existing seven-argument DOCX/XLSX/PPTX calls exactly on their
  // crop-only path; a retained projection is an explicit eighth argument.
  if (!options) { paint(dx, dy, dw, dh); return; }
  const plan = options.projection ?? imagePaintProjection(dw, dh, srcRect, options.transform, false);
  if (plan.width !== dw || plan.height !== dh) throw new RangeError('Image projection destination mismatch');
  withImagePaintProjection(ctx, dx, dy, plan, (x, y, width, height) => paint(x, y, width, height, plan.crop));
}

/**
 * Required full-source raster resolution for a destination measured in device
 * pixels. A positive `<a:srcRect>` crop magnifies a source slice to fill the
 * destination, so the full decoded source needs proportionally more pixels;
 * negative insets create transparent outsets and therefore need fewer.
 *
 * The result is a decode request, not an allocation promise. The decoder maps
 * the full source grid to both requested axes (the same resampling the final
 * destination draw would otherwise perform) and applies the shared ceiling.
 */
export function sourceRasterTargetSize(
  destinationWidthPx: number,
  destinationHeightPx: number,
  srcRect?: SrcRect | null,
): { width: number; height: number } | null {
  if (!Number.isFinite(destinationWidthPx) || !Number.isFinite(destinationHeightPx)) return null;
  if (!(destinationWidthPx > 0) || !(destinationHeightPx > 0)) return null;
  const logicalWidth = srcRect ? 1 - srcRect.l - srcRect.r : 1;
  const logicalHeight = srcRect ? 1 - srcRect.t - srcRect.b : 1;
  if (!Number.isFinite(logicalWidth) || !Number.isFinite(logicalHeight)) return null;
  if (!(logicalWidth > 0) || !(logicalHeight > 0) || !srcRectHasVisibleArea(srcRect)) return null;
  const width = Math.ceil(destinationWidthPx / logicalWidth);
  const height = Math.ceil(destinationHeightPx / logicalHeight);
  return Number.isFinite(width) && Number.isFinite(height)
    ? { width, height }
    : null;
}

/** Raster target size (pt) for decoding an embedded image. A browser raster's
 *  retained pixel target is handled separately, so its point box passes through. A metafile
 *  (WMF/EMF) with an `<a:srcRect>` crop must be rasterized at its FULL picture
 *  frame, not the visible sub-rectangle — the player maps the frame onto the
 *  raster (see `playEmf`), and the crop is relative to that frame. Scale the box
 *  up by `1/(1−l−r)` and `1/(1−t−b)` so the rasterised frame and the fractional
 *  crop align (e.g. one composite EMF cropped into subfigures). Uncropped
 *  metafiles and all raster blips pass the box through unchanged. */
export function metafileRasterSize(
  mimeType: string,
  srcRect: SrcRect | null | undefined,
  widthPt: number,
  heightPt: number,
): { widthPt: number; heightPt: number } | null {
  if (!srcRect || !isMetafileMime(mimeType)) return { widthPt, heightPt };
  if (!srcRectHasVisibleArea(srcRect)) return null;
  const fracW = 1 - srcRect.l - srcRect.r;
  const fracH = 1 - srcRect.t - srcRect.b;
  return { widthPt: widthPt / fracW, heightPt: heightPt / fracH };
}
