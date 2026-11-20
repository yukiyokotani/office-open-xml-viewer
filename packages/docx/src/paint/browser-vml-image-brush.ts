import { createAuxCanvas, imageNaturalSize } from '@silurus/ooxml-core';
import type { DeepReadonly, ImagePaintResourceDescriptor } from '../layout/types.js';
import type { DrawingMLTileBrush, VmlImageBrush } from './vml-image-brush.js';

/** Resource acquisition only: resolve native pixel size once after lazy decode,
 * before paint. Part 4 §19.1.2.5 defines size, origin, position and aspect;
 * native image pixels use the shared 96dpi image-coordinate convention. */
export function acquireVmlImageBrush(
  descriptor: DeepReadonly<ImagePaintResourceDescriptor>, image: CanvasImageSource,
): Readonly<VmlImageBrush> {
  const fill = descriptor.vmlFill;
  if (!fill) throw new Error('Missing VML image brush acquisition facts');
  const natural = imageNaturalSize(image);
  const box = descriptor.intrinsicSize;
  if (!(natural.w > 0 && natural.h > 0 && box.widthPt > 0 && box.heightPt > 0)) {
    throw new RangeError('VML image brush requires positive image and host dimensions');
  }
  let width = fill.size
    ? fill.size[0].value * (fill.size[0].fraction ? box.widthPt : 1)
    : fill.repeat ? natural.w * 72 / 96 : box.widthPt;
  let height = fill.size
    ? fill.size[1].value * (fill.size[1].fraction ? box.heightPt : 1)
    : fill.repeat ? natural.h * 72 / 96 : box.heightPt;
  if (fill.aspect === 'atleast' || fill.aspect === 'atmost') {
    const ratio = (fill.aspect === 'atleast' ? Math.max : Math.min)(width / natural.w, height / natural.h);
    width = natural.w * ratio; height = natural.h * ratio;
  }
  if (![width, height].every(n => Number.isFinite(n) && n > 0)) {
    throw new RangeError('VML image brush size must be finite and positive');
  }
  return Object.freeze({ kind: 'vml-image-brush', image, repeat: fill.repeat,
    width: width / box.widthPt, height: height / box.heightPt,
    x: fill.position[0] - fill.origin[0] * width / box.widthPt,
    y: fill.position[1] - fill.origin[1] * height / box.heightPt,
    pixelWidth: natural.w, pixelHeight: natural.h,
  });
}

/** ECMA-376 §20.1.8.58 DrawingML tile, resolved once after lazy decode. The
 * tile repeats the §20.1.8.55 cropped source (Annex L.4.8.4.3) at its native
 * 96-DPI size scaled by `sx`/`sy`; `algn` registers the first tile against
 * the painted box after scaling, `tx`/`ty` then offset it, and `flip`
 * mirrors alternate tiles (the first is unmirrored), composed here into one
 * repeating cell.
 *
 * CT_TileInfoProperties gives `tx`, `ty`, `sx`, `sy` and `algn` no schema
 * default (`flip` defaults to none). Library policy reads an absent value as
 * the identity: no offset, 100% scale, top-left registration. */
export function acquireDrawingMLTileBrush(
  descriptor: DeepReadonly<ImagePaintResourceDescriptor>, image: CanvasImageSource,
): Readonly<DrawingMLTileBrush> {
  const tile = descriptor.tile;
  if (!tile) throw new Error('Missing DrawingML tile acquisition facts');
  const natural = imageNaturalSize(image);
  const crop = descriptor.srcRect ?? { l: 0, t: 0, r: 0, b: 0 };
  const sourceX = natural.w * crop.l;
  const sourceY = natural.h * crop.t;
  const sourceW = natural.w * (1 - crop.l - crop.r);
  const sourceH = natural.h * (1 - crop.t - crop.b);
  const tileWidthPt = sourceW * 72 / 96 * (tile.sx ?? 1);
  const tileHeightPt = sourceH * 72 / 96 * (tile.sy ?? 1);
  if (![sourceW, sourceH, tileWidthPt, tileHeightPt].every((value) => Number.isFinite(value) && value > 0)) {
    throw new RangeError('DrawingML tile requires positive image and tile dimensions');
  }
  const offsetXPt = (tile.tx ?? 0) / 12700;
  const offsetYPt = (tile.ty ?? 0) / 12700;
  if (![offsetXPt, offsetYPt].every(Number.isFinite)) throw new RangeError('DrawingML tile offset must be finite');
  const flipX = tile.flip === 'x' || tile.flip === 'xy';
  const flipY = tile.flip === 'y' || tile.flip === 'xy';
  const pixelWidth = Math.max(1, Math.round(sourceW));
  const pixelHeight = Math.max(1, Math.round(sourceH));
  let cell: CanvasImageSource = image;
  if (flipX || flipY || descriptor.srcRect) {
    const canvas = createAuxCanvas(pixelWidth * (flipX ? 2 : 1), pixelHeight * (flipY ? 2 : 1));
    const context = canvas?.getContext('2d') as CanvasRenderingContext2D | null | undefined;
    if (!canvas || !context) throw new Error('DrawingML tile brush requires a scratch canvas');
    for (const [mirrorX, mirrorY] of [[false, false], [true, false], [false, true], [true, true]] as const) {
      if ((mirrorX && !flipX) || (mirrorY && !flipY)) continue;
      context.save();
      context.translate(pixelWidth * (mirrorX ? 2 : 0), pixelHeight * (mirrorY ? 2 : 0));
      context.scale(mirrorX ? -1 : 1, mirrorY ? -1 : 1);
      context.drawImage(image, sourceX, sourceY, sourceW, sourceH, 0, 0, pixelWidth, pixelHeight);
      context.restore();
    }
    cell = canvas as CanvasImageSource;
  }
  return Object.freeze({
    kind: 'drawingml-tile-brush', image: cell,
    tileWidthPt, tileHeightPt, alignment: tile.algn ?? 'tl', offsetXPt, offsetYPt,
    pixelWidth, pixelHeight,
  });
}
