/** Physical resolution stored in a raster's own header, in dots per inch.
 * ECMA-376 §20.1.8.14 `blipFill@dpi`: an absent or zero value means "the DPI
 * in the blip is used", which is this metadata. */
export interface RasterResolution {
  readonly dpiX: number;
  readonly dpiY: number;
}

const METRES_PER_INCH = 0.0254;
const CM_PER_INCH = 2.54;

function resolution(dpiX: number, dpiY: number): RasterResolution | null {
  return Number.isFinite(dpiX) && Number.isFinite(dpiY) && dpiX > 0 && dpiY > 0
    ? { dpiX, dpiY }
    : null;
}

/** PNG `pHYs` (ISO/IEC 15948 §11.3.5.3): pixels per unit; unit 1 is the metre.
 * Unit 0 states only an aspect ratio and carries no physical size. */
function pngResolution(b: Uint8Array): RasterResolution | null {
  const view = new DataView(b.buffer, b.byteOffset, b.byteLength);
  let offset = 8;
  while (offset + 8 <= b.length) {
    const length = view.getUint32(offset);
    const type = String.fromCharCode(b[offset + 4]!, b[offset + 5]!, b[offset + 6]!, b[offset + 7]!);
    if (type === 'IDAT' || type === 'IEND') return null;
    if (type === 'pHYs') {
      if (length < 9 || offset + 8 + 9 > b.length || b[offset + 16] !== 1) return null;
      return resolution(view.getUint32(offset + 8) * METRES_PER_INCH, view.getUint32(offset + 12) * METRES_PER_INCH);
    }
    offset += 12 + length;
  }
  return null;
}

/** TIFF IFD0 XResolution / YResolution (RATIONAL) with ResolutionUnit
 * (2 = inch, the default; 3 = centimetre; 1 = none). Shared by TIFF files and
 * JPEG EXIF, whose offsets are relative to `tiffOffset`. */
function tiffResolution(b: Uint8Array, tiffOffset: number): RasterResolution | null {
  if (tiffOffset + 8 > b.length) return null;
  const little = b[tiffOffset] === 0x49 && b[tiffOffset + 1] === 0x49;
  const big = b[tiffOffset] === 0x4d && b[tiffOffset + 1] === 0x4d;
  if (!little && !big) return null;
  const view = new DataView(b.buffer, b.byteOffset, b.byteLength);
  if (view.getUint16(tiffOffset + 2, little) !== 42) return null;
  const ifd = tiffOffset + view.getUint32(tiffOffset + 4, little);
  if (ifd + 2 > b.length) return null;
  const count = Math.min(view.getUint16(ifd, little), Math.floor((b.length - ifd - 2) / 12));
  let x: number | null = null;
  let y: number | null = null;
  let unit = 2;
  const rational = (entry: number): number | null => {
    if (view.getUint16(entry + 2, little) !== 5 || view.getUint32(entry + 4, little) !== 1) return null;
    const at = tiffOffset + view.getUint32(entry + 8, little);
    if (at + 8 > b.length) return null;
    const denominator = view.getUint32(at + 4, little);
    return denominator === 0 ? null : view.getUint32(at, little) / denominator;
  };
  for (let index = 0; index < count; index++) {
    const entry = ifd + 2 + index * 12;
    const tag = view.getUint16(entry, little);
    if (tag === 0x011a) x = rational(entry);
    else if (tag === 0x011b) y = rational(entry);
    else if (tag === 0x0128 && view.getUint16(entry + 2, little) === 3) unit = view.getUint16(entry + 8, little);
  }
  if (x === null || y === null || (unit !== 2 && unit !== 3)) return null;
  return unit === 3 ? resolution(x * CM_PER_INCH, y * CM_PER_INCH) : resolution(x, y);
}

/** JPEG JFIF APP0 density (units 1 = dots per inch, 2 = dots per cm; 0 is an
 * aspect ratio only), else EXIF APP1 IFD0 resolution. */
function jpegResolution(b: Uint8Array): RasterResolution | null {
  let offset = 2;
  let exif: RasterResolution | null = null;
  while (offset + 4 <= b.length && b[offset] === 0xff) {
    const marker = b[offset + 1]!;
    if (marker === 0xda || marker === 0xd9) break;
    const length = (b[offset + 2]! << 8) | b[offset + 3]!;
    const payload = offset + 4;
    if (marker === 0xe0 && payload + 12 <= b.length
      && b[payload] === 0x4a && b[payload + 1] === 0x46 && b[payload + 2] === 0x49
      && b[payload + 3] === 0x46 && b[payload + 4] === 0x00) {
      const units = b[payload + 7];
      const dx = (b[payload + 8]! << 8) | b[payload + 9]!;
      const dy = (b[payload + 10]! << 8) | b[payload + 11]!;
      if (units === 1) return resolution(dx, dy);
      if (units === 2) return resolution(dx * CM_PER_INCH, dy * CM_PER_INCH);
    }
    if (marker === 0xe1 && exif === null && payload + 6 <= b.length
      && b[payload] === 0x45 && b[payload + 1] === 0x78 && b[payload + 2] === 0x69
      && b[payload + 3] === 0x66 && b[payload + 4] === 0 && b[payload + 5] === 0) {
      exif = tiffResolution(b.subarray(0, Math.min(b.length, payload + length - 2)), payload + 6);
    }
    if (length < 2) break;
    offset += 2 + length;
  }
  return exif;
}

/** BMP BITMAPINFOHEADER biXPelsPerMeter / biYPelsPerMeter. */
function bmpResolution(b: Uint8Array): RasterResolution | null {
  if (b.length < 46) return null;
  const view = new DataView(b.buffer, b.byteOffset, b.byteLength);
  if (view.getUint32(14, true) < 40) return null;
  return resolution(view.getInt32(38, true) * METRES_PER_INCH, view.getInt32(42, true) * METRES_PER_INCH);
}

/** The physical resolution a raster header declares, or `null` when the
 * format carries none (GIF, WebP, SVG, metafiles) or it is absent/unitless. */
export function sniffRasterResolution(head: Uint8Array): RasterResolution | null {
  if (head.length >= 8 && head[0] === 0x89 && head[1] === 0x50 && head[2] === 0x4e && head[3] === 0x47) {
    return pngResolution(head);
  }
  if (head.length >= 4 && head[0] === 0xff && head[1] === 0xd8) return jpegResolution(head);
  if (head.length >= 2 && head[0] === 0x42 && head[1] === 0x4d) return bmpResolution(head);
  if (head.length >= 4 && ((head[0] === 0x49 && head[1] === 0x49) || (head[0] === 0x4d && head[1] === 0x4d))) {
    return tiffResolution(head, 0);
  }
  return null;
}
