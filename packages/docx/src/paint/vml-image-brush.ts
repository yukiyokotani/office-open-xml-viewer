export interface VmlImageBrush {
  readonly kind: 'vml-image-brush';
  readonly image: CanvasImageSource;
  readonly repeat: boolean;
  readonly x: number; readonly y: number;
  readonly width: number; readonly height: number;
  readonly pixelWidth: number; readonly pixelHeight: number;
}
export function isVmlImageBrush(handle: unknown): handle is Readonly<VmlImageBrush> {
  return typeof handle === 'object' && handle !== null && 'kind' in handle && handle.kind === 'vml-image-brush';
}

/** ECMA-376 §20.1.8.58 DrawingML tile brush. Unlike a VML brush, every
 * placement fact is absolute (points) or an ST_RectAlignment resolved
 * against the painted bounds, so one acquired resource serves every extent
 * its host takes (for instance a page-relative anchor size). `image` is the
 * repeating cell: one tile, or the mirrored 2x2 / 2x1 block of `flip`. */
export interface DrawingMLTileBrush {
  readonly kind: 'drawingml-tile-brush';
  readonly image: CanvasImageSource;
  readonly tileWidthPt: number; readonly tileHeightPt: number;
  readonly alignment: string;
  readonly offsetXPt: number; readonly offsetYPt: number;
  /** Pixel extent of one tile inside `image`. */
  readonly pixelWidth: number; readonly pixelHeight: number;
}
export function isDrawingMLTileBrush(handle: unknown): handle is Readonly<DrawingMLTileBrush> {
  return typeof handle === 'object' && handle !== null && 'kind' in handle && handle.kind === 'drawingml-tile-brush';
}

/** Top-left of the first tile within `bounds` (§20.1.10.53 alignment after
 * scaling, then the tile offset). */
export function drawingMLTileOrigin(
  brush: Readonly<DrawingMLTileBrush>,
  bounds: Readonly<{ xPt: number; yPt: number; widthPt: number; heightPt: number }>,
): Readonly<{ xPt: number; yPt: number }> {
  const a = brush.alignment;
  const x = a === 't' || a === 'ctr' || a === 'b'
    ? (bounds.widthPt - brush.tileWidthPt) / 2
    : a === 'tr' || a === 'r' || a === 'br' ? bounds.widthPt - brush.tileWidthPt : 0;
  const y = a === 'l' || a === 'ctr' || a === 'r'
    ? (bounds.heightPt - brush.tileHeightPt) / 2
    : a === 'bl' || a === 'b' || a === 'br' ? bounds.heightPt - brush.tileHeightPt : 0;
  return { xPt: bounds.xPt + x + brush.offsetXPt, yPt: bounds.yPt + y + brush.offsetYPt };
}
