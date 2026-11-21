import {
  isJpegExifPayload,
  orientJpegDimensions,
  sniffJpegExifOrientation,
  sniffRasterDimensions,
  type RasterDimensions,
} from './raster-dimensions.js';
import { isTiff } from './tiff-contract.js';
import {
  bmpResolution,
  exifResolution,
  jfifResolution,
  orientRasterResolution,
  pngResolution,
  tiffResolution,
  type RasterByteReader,
  type RasterResolution,
} from './raster-resolution.js';
import { isEmf, isWmf } from './wmf.js';

export type RasterFormat = 'png' | 'jpeg' | 'gif' | 'bmp' | 'webp' | 'tiff' | 'wmf' | 'emf';

export interface RasterBlobInspection {
  readonly format: RasterFormat | null;
  readonly dimensions: RasterDimensions | null;
  /** Physical resolution declared in the header prefix, when present. */
  readonly resolution?: RasterResolution | null;
}

/** MIME types whose browser decoder accepts a bounded two-axis resize request.
 * A renderer can plan these from authored display geometry without first
 * extracting every source merely to classify it. The decode boundary still
 * sniffs the bytes and enforces the encoded-source hard limits. */
export function isBrowserResizableRasterMimeType(mimeType: string): boolean {
  switch (mimeType.split(';', 1)[0]?.trim().toLowerCase()) {
    case 'image/png':
    case 'image/jpeg':
    case 'image/jpg':
    case 'image/gif':
    case 'image/bmp':
    case 'image/x-ms-bmp':
    case 'image/webp':
      return true;
    default:
      return false;
  }
}

/** Raster formats whose active decoder can honor a retained-grid target.
 * Browser formats use `createImageBitmap` resize; TIFF participates only when
 * the caller has installed the optional codec. WMF/EMF are deliberately
 * excluded because their authored coordinate space is not a source pixel grid. */
export function isDecodeTargetResizableRasterFormat(
  format: RasterFormat | null,
  hasTiffRenderer = false,
): boolean {
  return format !== null
    && format !== 'wmf'
    && format !== 'emf'
    && (format !== 'tiff' || hasTiffRenderer);
}

const INITIAL_SNIFF_BYTES = 64 * 1024;

function rasterFormat(head: Uint8Array): RasterFormat | null {
  if (isWmf(head)) return 'wmf';
  if (isEmf(head)) return 'emf';
  if (head.length >= 8
    && head[0] === 0x89 && head[1] === 0x50 && head[2] === 0x4e && head[3] === 0x47
    && head[4] === 0x0d && head[5] === 0x0a && head[6] === 0x1a && head[7] === 0x0a) return 'png';
  if (head.length >= 2 && head[0] === 0xff && head[1] === 0xd8) return 'jpeg';
  if (head.length >= 6 && head[0] === 0x47 && head[1] === 0x49 && head[2] === 0x46
    && head[3] === 0x38 && (head[4] === 0x37 || head[4] === 0x39) && head[5] === 0x61) return 'gif';
  if (head.length >= 2 && head[0] === 0x42 && head[1] === 0x4d) return 'bmp';
  if (head.length >= 12 && head[0] === 0x52 && head[1] === 0x49 && head[2] === 0x46
    && head[3] === 0x46 && head[8] === 0x57 && head[9] === 0x45 && head[10] === 0x42
    && head[11] === 0x50) return 'webp';
  if (isTiff(head)) return 'tiff';
  return null;
}

function isJpegSof(marker: number): boolean {
  return marker >= 0xc0 && marker <= 0xcf
    && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc;
}

interface JpegMetadata {
  readonly dimensions: RasterDimensions | null;
  /** JFIF density, else EXIF IFD0 resolution, on the oriented axes. */
  readonly resolution: RasterResolution | null;
}

/**
 * Find JPEG SOF dimensions and declared resolution without retaining the
 * compressed image in a JS ArrayBuffer. JPEG permits arbitrarily many APP/ICC
 * segments before SOF, so a fixed prefix is not a security boundary. The state
 * machine consumes the Blob stream once, skips fill bytes, standalone markers
 * and segment payloads, retains only SOF, JFIF APP0 and EXIF APP1 payloads
 * (each bounded by its 16-bit declared length), and continues through metadata
 * to SOS/EOI for a later EXIF orientation. Orientation applies to the
 * dimensions and the resolution axes together.
 */
async function streamJpegMetadata(blob: Blob): Promise<JpegMetadata> {
  const reader = blob.slice(2).stream().getReader();
  let sawMarkerPrefix = false;
  let marker = -1;
  let lengthHi = -1;
  let skip = 0;
  let sofPayload: Uint8Array | null = null;
  let sofOffset = 0;
  let sofTrailingBytes = 0;
  let dimensions: RasterDimensions | null = null;
  let app1Payload: Uint8Array | null = null;
  let app1Offset = 0;
  let orientation: number | null = null;
  let sawExif = false;
  let app0Payload: Uint8Array | null = null;
  let app0Offset = 0;
  let jfif: RasterResolution | null = null;
  let exif: RasterResolution | null = null;
  const result = (): JpegMetadata => ({
    dimensions: dimensions ? orientJpegDimensions(dimensions, orientation) : null,
    resolution: orientRasterResolution(jfif ?? exif, orientation),
  });
  try {
    while (true) {
      const { value, done } = await reader.read();
      if (done) return result();
      for (const byte of value) {
        if (sofPayload) {
          sofPayload[sofOffset++] = byte;
          if (sofOffset === sofPayload.length) {
            const height = (sofPayload[1] << 8) | sofPayload[2];
            const width = (sofPayload[3] << 8) | sofPayload[4];
            dimensions ??= { width, height };
            sofPayload = null;
            sofOffset = 0;
            if (sawExif) {
              await reader.cancel();
              return result();
            }
            skip = sofTrailingBytes;
            sofTrailingBytes = 0;
          }
          continue;
        }
        if (app0Payload) {
          app0Payload[app0Offset++] = byte;
          if (app0Offset === app0Payload.length) {
            jfif ??= jfifResolution(app0Payload);
            app0Payload = null;
            app0Offset = 0;
          }
          continue;
        }
        if (app1Payload) {
          app1Payload[app1Offset++] = byte;
          if (app1Offset === app1Payload.length) {
            if (isJpegExifPayload(app1Payload)) {
              sawExif = true;
              orientation = sniffJpegExifOrientation(app1Payload);
              exif = await exifResolution(app1Payload);
            }
            app1Payload = null;
            app1Offset = 0;
            if (sawExif && dimensions) {
              await reader.cancel();
              return result();
            }
          }
          continue;
        }
        if (skip > 0) {
          skip--;
          continue;
        }
        if (marker >= 0) {
          if (lengthHi < 0) {
            lengthHi = byte;
            continue;
          }
          const segmentLength = (lengthHi << 8) | byte;
          if (segmentLength < 2) return result();
          if (isJpegSof(marker)) {
            if (segmentLength < 7) return result();
            if (dimensions) {
              skip = segmentLength - 2;
            } else {
              sofPayload = new Uint8Array(5);
              sofOffset = 0;
              sofTrailingBytes = segmentLength - 7;
            }
          } else if (marker === 0xe1 && !sawExif && segmentLength > 2) {
            app1Payload = new Uint8Array(segmentLength - 2);
            app1Offset = 0;
          } else if (marker === 0xe0 && jfif === null && segmentLength > 2) {
            app0Payload = new Uint8Array(segmentLength - 2);
            app0Offset = 0;
          } else {
            skip = segmentLength - 2;
          }
          marker = -1;
          lengthHi = -1;
          continue;
        }
        if (!sawMarkerPrefix) {
          if (byte === 0xff) sawMarkerPrefix = true;
          continue;
        }
        if (byte === 0xff) continue;
        sawMarkerPrefix = false;
        if (byte === 0x00) continue;
        if (byte === 0xd9 || byte === 0xda) {
          await reader.cancel();
          return result();
        }
        if (byte === 0xd8 || byte === 0x01 || (byte >= 0xd0 && byte <= 0xd7)) continue;
        marker = byte;
      }
    }
  } finally {
    reader.releaseLock();
  }
}

async function deepTiffDimensions(blob: Blob, head: Uint8Array): Promise<RasterDimensions | null> {
  if (head.length < 8 || !isTiff(head)) return null;
  const little = head[0] === 0x49;
  const headView = new DataView(head.buffer, head.byteOffset, head.byteLength);
  const ifdOffset = headView.getUint32(4, little);
  if (!Number.isSafeInteger(ifdOffset) || ifdOffset < 8 || ifdOffset > blob.size - 2) return null;
  const countBytes = new Uint8Array(await blob.slice(ifdOffset, ifdOffset + 2).arrayBuffer());
  if (countBytes.length !== 2) return null;
  const count = new DataView(countBytes.buffer, countBytes.byteOffset, 2).getUint16(0, little);
  const directoryBytes = count * 12;
  if (ifdOffset + 2 + directoryBytes > blob.size) return null;
  const directory = new Uint8Array(
    await blob.slice(ifdOffset + 2, ifdOffset + 2 + directoryBytes).arrayBuffer(),
  );
  const view = new DataView(directory.buffer, directory.byteOffset, directory.byteLength);
  let width: number | undefined;
  let height: number | undefined;
  for (let index = 0; index < count; index++) {
    const offset = index * 12;
    const tag = view.getUint16(offset, little);
    if (tag !== 256 && tag !== 257) continue;
    const type = view.getUint16(offset + 2, little);
    const valueCount = view.getUint32(offset + 4, little);
    if (valueCount !== 1 || (type !== 1 && type !== 3 && type !== 4)) return null;
    const value = type === 1
      ? view.getUint8(offset + 8)
      : type === 3
        ? view.getUint16(offset + 8, little)
        : view.getUint32(offset + 8, little);
    if (tag === 256) width = value;
    else height = value;
  }
  return width === undefined || height === undefined ? null : { width, height };
}

/** Declared resolution of a non-JPEG raster. Metadata past the header prefix
 * (a PNG pHYs after large ancillary chunks, TIFF values at far offsets) is
 * sought with small bounded reads; nothing else is read. */
async function rasterResolution(
  format: RasterFormat,
  blob: Blob,
  head: Uint8Array,
): Promise<RasterResolution | null> {
  const read: RasterByteReader = async (offset, length) => {
    if (!(offset >= 0) || !(length >= 0) || offset >= blob.size) return new Uint8Array(0);
    if (offset + length <= head.length) return head.subarray(offset, offset + length);
    return new Uint8Array(await blob.slice(offset, offset + length).arrayBuffer());
  };
  if (format === 'png') return pngResolution(read);
  if (format === 'tiff') return tiffResolution(read);
  if (format === 'bmp') return bmpResolution(head);
  return null;
}

/** Inspect the decoder-natural raster grid before handing bytes to a decoder. */
export async function inspectRasterBlob(
  blob: Blob,
  initialHead?: Uint8Array,
): Promise<RasterBlobInspection> {
  const head = initialHead ?? new Uint8Array(
    await blob.slice(0, INITIAL_SNIFF_BYTES).arrayBuffer(),
  );
  const format = rasterFormat(head);
  if (!format) return { format: null, dimensions: null };
  const dimensions = sniffRasterDimensions(head);
  // A prefix SOF is insufficient: Chrome also honors EXIF APP1 metadata that
  // follows SOF. Scan every JPEG to the first scan boundary so retained-surface
  // planning uses the decoder-natural orientation.
  if (format === 'jpeg') {
    const metadata = await streamJpegMetadata(blob);
    return { format, dimensions: metadata.dimensions ?? dimensions, resolution: metadata.resolution };
  }
  const resolution = await rasterResolution(format, blob, head);
  if (dimensions) return { format, dimensions, resolution };
  if (format === 'tiff') return { format, dimensions: await deepTiffDimensions(blob, head), resolution };
  return { format, dimensions: null, resolution };
}
