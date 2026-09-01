// TIFF Revision 6.0 decoder for OOXML image parts. ECMA-376 Part 1 §15.2.14
// explicitly permits `image/tiff`, but browser ImageBitmap decoders generally
// do not. Keep the supported class tag-defined: classic TIFF, first IFD,
// uncompressed 8-bit chunky process-CMYK strips and top-left orientation. The
// codec deliberately rejects other TIFF classes instead of guessing or handing
// them to createImageBitmap and aborting the document render.

import { createAuxCanvas } from '../canvas/aux-canvas.js';
import {
  MAX_RASTER_DIMENSION,
  MAX_RASTER_PIXELS,
  OoxmlDecodedImageLimitError,
} from './pixel-budget.js';
import { isTiff } from './tiff-contract.js';

/** Stable bundle-audit marker retained only by the opt-in TIFF entry. */
export const TIFF_IMPLEMENTATION_MARKER = 'packages/core/src/image/tiff.ts';

const TYPE_BYTE = 1;
const TYPE_SHORT = 3;
const TYPE_LONG = 4;

const TAG = {
  width: 256,
  height: 257,
  bitsPerSample: 258,
  compression: 259,
  photometric: 262,
  stripOffsets: 273,
  orientation: 274,
  samplesPerPixel: 277,
  rowsPerStrip: 278,
  stripByteCounts: 279,
  planarConfiguration: 284,
  inkSet: 332,
  extraSamples: 338,
} as const;

interface Field {
  type: number;
  count: number;
  entryOffset: number;
}

class Reader {
  readonly view: DataView;

  constructor(
    readonly bytes: Uint8Array,
    readonly littleEndian: boolean,
  ) {
    this.view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  }

  contains(offset: number, length: number): boolean {
    return Number.isSafeInteger(offset)
      && Number.isSafeInteger(length)
      && offset >= 0
      && length >= 0
      && offset <= this.view.byteLength - length;
  }

  u16(offset: number): number | null {
    return this.contains(offset, 2) ? this.view.getUint16(offset, this.littleEndian) : null;
  }

  u32(offset: number): number | null {
    return this.contains(offset, 4) ? this.view.getUint32(offset, this.littleEndian) : null;
  }
}

interface Directory {
  reader: Reader;
  fields: Map<number, Field>;
}

function firstDirectory(bytes: Uint8Array): Directory | null {
  if (!isTiff(bytes)) return null;
  const reader = new Reader(bytes, bytes[0] === 0x49);
  const ifdOffset = reader.u32(4);
  if (ifdOffset == null) return null;
  const count = reader.u16(ifdOffset);
  if (count == null || !reader.contains(ifdOffset + 2, count * 12 + 4)) return null;

  const fields = new Map<number, Field>();
  for (let i = 0; i < count; i++) {
    const entryOffset = ifdOffset + 2 + i * 12;
    const tag = reader.u16(entryOffset);
    const type = reader.u16(entryOffset + 2);
    const valueCount = reader.u32(entryOffset + 4);
    if (tag == null || type == null || valueCount == null || fields.has(tag)) return null;
    fields.set(tag, { type, count: valueCount, entryOffset });
  }
  return { reader, fields };
}

function values(directory: Directory, tag: number, maxCount: number): number[] | null {
  const field = directory.fields.get(tag);
  if (!field || field.count < 1 || field.count > maxCount) return null;
  const size = field.type === TYPE_BYTE ? 1 : field.type === TYPE_SHORT ? 2 : field.type === TYPE_LONG ? 4 : 0;
  if (size === 0) return null;
  const byteLength = field.count * size;
  const offset = byteLength <= 4
    ? field.entryOffset + 8
    : directory.reader.u32(field.entryOffset + 8);
  if (offset == null || !directory.reader.contains(offset, byteLength)) return null;

  const result = new Array<number>(field.count);
  for (let i = 0; i < field.count; i++) {
    const itemOffset = offset + i * size;
    const value = field.type === TYPE_BYTE
      ? directory.reader.bytes[itemOffset]
      : field.type === TYPE_SHORT
        ? directory.reader.u16(itemOffset)
        : directory.reader.u32(itemOffset);
    if (value == null) return null;
    result[i] = value;
  }
  return result;
}

function scalar(directory: Directory, tag: number, fallback?: number): number | null {
  if (!directory.fields.has(tag)) return fallback ?? null;
  const field = directory.fields.get(tag) as Field;
  return field.count === 1 ? values(directory, tag, 1)?.[0] ?? null : null;
}

export interface DecodedTiff {
  width: number;
  height: number;
  data: Uint8ClampedArray;
}

function cmykChannel(channel: number, black: number): number {
  // TIFF 6.0 §16 defines ink coverage but no unique device-independent RGB
  // conversion. This multiplicative subtractive mapping matches Word's output
  // for the unprofiled 8-bit process-ink levels in the Office reference used to
  // verify this compatibility path; it applies uniformly to the defined class.
  return Math.round(((255 - channel) * (255 - black)) / 255);
}

/** Decode the supported TIFF class into top-down Canvas RGBA bytes. */
export function decodeTiffRgba(bytes: Uint8Array): DecodedTiff | null {
  const directory = firstDirectory(bytes);
  if (!directory) return null;
  const width = scalar(directory, TAG.width);
  const height = scalar(directory, TAG.height);
  if (width == null || height == null) return null;
  const pixels = width * height;
  if (
    width <= 0
    || height <= 0
    || width > MAX_RASTER_DIMENSION
    || height > MAX_RASTER_DIMENSION
    || pixels > MAX_RASTER_PIXELS
  ) {
    throw new OoxmlDecodedImageLimitError('image-pixels', MAX_RASTER_PIXELS, pixels);
  }

  const compression = scalar(directory, TAG.compression, 1);
  const photometric = scalar(directory, TAG.photometric);
  const samples = scalar(directory, TAG.samplesPerPixel, 1);
  const rowsPerStrip = scalar(directory, TAG.rowsPerStrip, 0xffffffff);
  if (
    compression !== 1
    || photometric == null
    || samples == null
    || rowsPerStrip == null
    || rowsPerStrip < 1
    || scalar(directory, TAG.planarConfiguration, 1) !== 1
    || scalar(directory, TAG.orientation, 1) !== 1
    || directory.fields.has(TAG.extraSamples)
  ) return null;

  if (photometric !== 5 || samples !== 4 || scalar(directory, TAG.inkSet, 1) !== 1) return null;
  const bits = directory.fields.has(TAG.bitsPerSample)
    ? values(directory, TAG.bitsPerSample, samples)
    : [1];
  if (!bits || bits.length !== samples || bits.some((bit) => bit !== 8)) return null;

  const stripCount = Math.ceil(height / rowsPerStrip);
  const offsets = values(directory, TAG.stripOffsets, stripCount);
  const byteCounts = values(directory, TAG.stripByteCounts, stripCount);
  if (!offsets || !byteCounts || offsets.length !== stripCount || byteCounts.length !== stripCount) {
    return null;
  }
  const bytesPerRow = width * samples;
  // Validate all source ranges before reserving the potentially large decoded
  // raster. Malformed metadata must fail without first consuming the pixel budget.
  for (let strip = 0; strip < stripCount; strip++) {
    const firstRow = strip * rowsPerStrip;
    const rowCount = Math.min(rowsPerStrip, height - firstRow);
    const required = rowCount * bytesPerRow;
    if (byteCounts[strip] < required || !directory.reader.contains(offsets[strip], required)) return null;
  }

  const output = new Uint8ClampedArray(pixels * 4);
  for (let strip = 0; strip < stripCount; strip++) {
    const firstRow = strip * rowsPerStrip;
    const rowCount = Math.min(rowsPerStrip, height - firstRow);
    for (let row = 0; row < rowCount; row++) {
      let source = offsets[strip] + row * bytesPerRow;
      let destination = (firstRow + row) * width * 4;
      for (let x = 0; x < width; x++, destination += 4) {
        const cyan = bytes[source++];
        const magenta = bytes[source++];
        const yellow = bytes[source++];
        const black = bytes[source++];
        output[destination] = cmykChannel(cyan, black);
        output[destination + 1] = cmykChannel(magenta, black);
        output[destination + 2] = cmykChannel(yellow, black);
        output[destination + 3] = 255;
      }
    }
  }
  return { width, height, data: output };
}

export async function renderTiffToBitmap(bytes: Uint8Array): Promise<ImageBitmap | null> {
  const decoded = decodeTiffRgba(bytes);
  if (!decoded) return null;
  const canvas = createAuxCanvas(decoded.width, decoded.height);
  if (!canvas) return null;
  const context = canvas.getContext('2d') as CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D | null;
  if (!context) return null;
  const imageData = context.createImageData(decoded.width, decoded.height);
  imageData.data.set(decoded.data);
  context.putImageData(imageData, 0, 0);
  return createImageBitmap(canvas);
}
