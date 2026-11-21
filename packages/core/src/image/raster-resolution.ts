/** Physical resolution stored in a raster's own metadata, in dots per inch,
 * on the decoder-natural (orientation-applied) axes. ECMA-376 §20.1.8.14
 * `blipFill@dpi`: an absent or zero value means "the DPI in the blip is
 * used", which is this metadata. */
export interface RasterResolution {
  readonly dpiX: number;
  readonly dpiY: number;
}

/** Bounded random access to an encoded source. A read past the end returns
 * fewer bytes; callers treat a short read as absent metadata. */
export type RasterByteReader = (offset: number, length: number) => Promise<Uint8Array>;

const METRES_PER_INCH = 0.0254;
const CM_PER_INCH = 2.54;
/** Resource policy: chunks/entries visited while seeking one metadata field
 * past the header prefix. Each visit reads a fixed, small number of bytes. */
const MAX_PNG_METADATA_CHUNKS = 1024;
const MAX_TIFF_IFD_ENTRIES = 4096;

function resolution(dpiX: number, dpiY: number): RasterResolution | null {
  return Number.isFinite(dpiX) && Number.isFinite(dpiY) && dpiX > 0 && dpiY > 0
    ? { dpiX, dpiY }
    : null;
}

/** EXIF orientations 5–8 transpose the decoded image, so its axes swap. */
export function orientRasterResolution(
  value: RasterResolution | null,
  orientation: number | null,
): RasterResolution | null {
  return value && orientation !== null && orientation >= 5 && orientation <= 8
    ? { dpiX: value.dpiY, dpiY: value.dpiX }
    : value;
}

/** JPEG JFIF APP0 payload (after the length field): units 1 = dots per inch,
 * 2 = dots per cm; 0 states only an aspect ratio. Every read stays inside the
 * declared payload. */
export function jfifResolution(payload: Uint8Array): RasterResolution | null {
  if (payload.length < 12 || payload[0] !== 0x4a || payload[1] !== 0x46 || payload[2] !== 0x49
    || payload[3] !== 0x46 || payload[4] !== 0x00) return null;
  const units = payload[7];
  const x = (payload[8]! << 8) | payload[9]!;
  const y = (payload[10]! << 8) | payload[11]!;
  if (units === 1) return resolution(x, y);
  if (units === 2) return resolution(x * CM_PER_INCH, y * CM_PER_INCH);
  return null;
}

/** TIFF IFD0 XResolution / YResolution (RATIONAL) with ResolutionUnit
 * (2 = inch, the default; 3 = centimetre; 1 = none), read through `read` at
 * offsets relative to the TIFF header at `base`. Shared by TIFF files and
 * JPEG EXIF. */
export async function tiffResolution(read: RasterByteReader, base = 0): Promise<RasterResolution | null> {
  const header = await read(base, 8);
  if (header.length < 8) return null;
  const little = header[0] === 0x49 && header[1] === 0x49;
  const big = header[0] === 0x4d && header[1] === 0x4d;
  if (!little && !big) return null;
  const headerView = new DataView(header.buffer, header.byteOffset, header.byteLength);
  if (headerView.getUint16(2, little) !== 42) return null;
  const ifd = headerView.getUint32(4, little);
  if (ifd < 8) return null;
  const countBytes = await read(base + ifd, 2);
  if (countBytes.length < 2) return null;
  const count = Math.min(
    new DataView(countBytes.buffer, countBytes.byteOffset, 2).getUint16(0, little),
    MAX_TIFF_IFD_ENTRIES,
  );
  const entries = await read(base + ifd + 2, count * 12);
  const view = new DataView(entries.buffer, entries.byteOffset, entries.byteLength);
  const readable = Math.floor(entries.length / 12);
  const rational = async (entry: number): Promise<number | null> => {
    if (view.getUint16(entry + 2, little) !== 5 || view.getUint32(entry + 4, little) !== 1) return null;
    const value = await read(base + view.getUint32(entry + 8, little), 8);
    if (value.length < 8) return null;
    const valueView = new DataView(value.buffer, value.byteOffset, 8);
    const denominator = valueView.getUint32(4, little);
    return denominator === 0 ? null : valueView.getUint32(0, little) / denominator;
  };
  let x: number | null = null;
  let y: number | null = null;
  let unit = 2;
  for (let index = 0; index < readable; index++) {
    const entry = index * 12;
    const tag = view.getUint16(entry, little);
    if (tag === 0x011a) x = await rational(entry);
    else if (tag === 0x011b) y = await rational(entry);
    else if (tag === 0x0128 && view.getUint16(entry + 2, little) === 3) unit = view.getUint16(entry + 8, little);
  }
  if (x === null || y === null || (unit !== 2 && unit !== 3)) return null;
  return unit === 3 ? resolution(x * CM_PER_INCH, y * CM_PER_INCH) : resolution(x, y);
}

/** JPEG EXIF APP1 payload (after the length field): its IFD0 resolution,
 * bounded by the payload. */
export async function exifResolution(payload: Uint8Array): Promise<RasterResolution | null> {
  if (payload.length < 14 || payload[0] !== 0x45 || payload[1] !== 0x78 || payload[2] !== 0x69
    || payload[3] !== 0x66 || payload[4] !== 0 || payload[5] !== 0) return null;
  return tiffResolution(async (offset, length) => (
    offset < 0 || offset >= payload.length ? new Uint8Array(0) : payload.subarray(offset, offset + length)
  ), 6);
}

/** PNG `pHYs` (ISO/IEC 15948 §11.3.5.3): pixels per unit, unit 1 = metre
 * (unit 0 states only an aspect ratio). It precedes IDAT; the chunk walk
 * reads only chunk headers until pHYs, IDAT or IEND, under a chunk cap. */
export async function pngResolution(read: RasterByteReader): Promise<RasterResolution | null> {
  let offset = 8;
  for (let chunk = 0; chunk < MAX_PNG_METADATA_CHUNKS; chunk++) {
    const header = await read(offset, 8);
    if (header.length < 8) return null;
    const length = new DataView(header.buffer, header.byteOffset, 8).getUint32(0);
    const type = String.fromCharCode(header[4]!, header[5]!, header[6]!, header[7]!);
    if (type === 'IDAT' || type === 'IEND') return null;
    if (type === 'pHYs') {
      if (length < 9) return null;
      const data = await read(offset + 8, 9);
      if (data.length < 9 || data[8] !== 1) return null;
      const view = new DataView(data.buffer, data.byteOffset, 9);
      return resolution(view.getUint32(0) * METRES_PER_INCH, view.getUint32(4) * METRES_PER_INCH);
    }
    offset += 12 + length;
    if (!Number.isSafeInteger(offset)) return null;
  }
  return null;
}

/** BMP BITMAPINFOHEADER biXPelsPerMeter / biYPelsPerMeter (a fixed offset
 * inside the header prefix). */
export function bmpResolution(head: Uint8Array): RasterResolution | null {
  if (head.length < 46) return null;
  const view = new DataView(head.buffer, head.byteOffset, head.byteLength);
  if (view.getUint32(14, true) < 40) return null;
  return resolution(view.getInt32(38, true) * METRES_PER_INCH, view.getInt32(42, true) * METRES_PER_INCH);
}
