// Shared `<a:srcRect>` source-rectangle crop (ECMA-376 §20.1.8.55) for the docx,
// pptx and xlsx renderers. The crop is a fraction of the image's NATIVE pixel
// grid, so the only per-renderer concern is decoding the image at full source
// size first — for a metafile that means rasterizing the whole picture FRAME
// (see `metafileRasterSize`), since the player maps the EMF/WMF frame onto the
// raster and the crop is relative to that frame. Centralised here so all three
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

interface CropMapping {
  sx: number;
  sy: number;
  sw: number;
  sh: number;
  dxFraction: number;
  dyFraction: number;
  dwFraction: number;
  dhFraction: number;
}

/** Resolve both the native source intersection and its destination placement.
 * Negative insets are normative outsets (§20.1.8.55): the logical source rect
 * extends beyond the bitmap and the unavailable part stays transparent in the
 * destination instead of being clamped and stretching the bitmap. */
export function cropSourceMapping(
  img: CanvasImageSource,
  srcRect: SrcRect | null | undefined,
): CropMapping | null {
  if (!srcRect || !(srcRect.l || srcRect.t || srcRect.r || srcRect.b)) return null;
  const values = [srcRect.l, srcRect.t, srcRect.r, srcRect.b];
  if (!values.every(Number.isFinite)) return null;
  const { w, h } = imageNaturalSize(img);
  if (w <= 0 || h <= 0) return null;
  const logicalX0 = srcRect.l;
  const logicalY0 = srcRect.t;
  const logicalX1 = 1 - srcRect.r;
  const logicalY1 = 1 - srcRect.b;
  const logicalW = logicalX1 - logicalX0;
  const logicalH = logicalY1 - logicalY0;
  if (!(logicalW > 0) || !(logicalH > 0)) {
    return {
      sx: 0, sy: 0, sw: 0, sh: 0,
      dxFraction: 0, dyFraction: 0, dwFraction: 0, dhFraction: 0,
    };
  }
  const sourceX0 = Math.max(0, logicalX0);
  const sourceY0 = Math.max(0, logicalY0);
  const sourceX1 = Math.min(1, logicalX1);
  const sourceY1 = Math.min(1, logicalY1);
  const sourceW = Math.max(0, sourceX1 - sourceX0);
  const sourceH = Math.max(0, sourceY1 - sourceY0);
  return {
    sx: sourceX0 * w,
    sy: sourceY0 * h,
    sw: sourceW * w,
    sh: sourceH * h,
    dxFraction: (sourceX0 - logicalX0) / logicalW,
    dyFraction: (sourceY0 - logicalY0) / logicalH,
    dwFraction: sourceW / logicalW,
    dhFraction: sourceH / logicalH,
  };
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
  ctx: AnyCtx,
  img: CanvasImageSource,
  srcRect: SrcRect | null | undefined,
  dx: number,
  dy: number,
  dw: number,
  dh: number,
): void {
  const c = cropSourceMapping(img, srcRect);
  if (!c) {
    ctx.drawImage(img, dx, dy, dw, dh);
  } else if (c.sw > 0 && c.sh > 0 && c.dwFraction > 0 && c.dhFraction > 0) {
    ctx.drawImage(
      img, c.sx, c.sy, c.sw, c.sh,
      dx + c.dxFraction * dw,
      dy + c.dyFraction * dh,
      c.dwFraction * dw,
      c.dhFraction * dh,
    );
  }
}

/**
 * Required full-source raster resolution for a destination measured in device
 * pixels. A positive `<a:srcRect>` crop magnifies a source slice to fill the
 * destination, so the full decoded source needs proportionally more pixels;
 * negative insets create transparent outsets and therefore need fewer.
 *
 * The result is a decode request, not an allocation promise. The decoder keeps
 * the source aspect ratio and applies the shared decoded-surface ceiling.
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
  return {
    width: Math.ceil(destinationWidthPx / logicalWidth),
    height: Math.ceil(destinationHeightPx / logicalHeight),
  };
}

/** Raster target size (pt) for decoding an embedded image. A raster blip decodes
 *  at its native pixel grid, so its display box passes through. A metafile
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
