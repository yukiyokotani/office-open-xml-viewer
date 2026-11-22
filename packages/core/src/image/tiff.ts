// TIFF Revision 6.0 decoder for OOXML image parts. ECMA-376 Part 1
// §15.2.14 permits TIFF image parts, while browser bitmap decoders do not
// consistently support them. Keep this compatibility codec deliberately
// tag-defined: classic TIFF, first IFD, stripped top-left chunky pixels, the
// baseline bilevel/grayscale/RGB/process-CMYK classes listed below, stored
// uncompressed, PackBits-compressed or LZW-compressed (LZW optionally with the
// horizontal differencing predictor), plus CCITT Group 4 bilevel images.

import {
  MAX_RASTER_DIMENSION,
  MAX_RASTER_PIXELS,
  MAX_RASTER_SOURCE_DIMENSION,
  MAX_RASTER_SOURCE_PIXELS,
  OoxmlDecodedImageLimitError,
} from './pixel-budget.js';
import { aspectPreservingRasterTarget } from './raster-target.js';
import {
  isTiff,
  TiffDecodeError,
  type TiffRenderOptions,
} from './tiff-contract.js';

export { TiffDecodeError, isTiffDecodeError, type TiffRenderOptions } from './tiff-contract.js';

/** Stable bundle-audit marker retained only by the opt-in TIFF entry. */
export const TIFF_IMPLEMENTATION_MARKER = 'packages/core/src/image/tiff.ts';

/**
 * Keep the packed-bit decoder in this module so the optional TIFF bundle audit
 * covers both LZW and Group 4, including their shared reader.
 *
 * TIFF 6.0 §13 LZW and §11/CCITT T.6 (FillOrder 1) use MSB-first packed codes.
 * The caller validates the strip byte range. Keep byte positions as arithmetic
 * numbers: even a valid strip can exceed 2^31 bits, where bitwise indexing
 * wraps. Only the 0..7 bit position and at most 12-bit code values are shifted.
 */
export class TiffBitReader {
  private byteCursor = 0;
  private byteEnd = 0;
  private bitInByte = 0;

  constructor(
    private readonly bytes: ArrayLike<number>,
    byteOffset: number,
    byteLength: number,
    private readonly truncatedMessage: string,
  ) {
    this.beginStrip(byteOffset, byteLength);
  }

  beginStrip(byteOffset: number, byteLength: number): void {
    this.byteCursor = byteOffset;
    this.byteEnd = byteOffset + byteLength;
    this.bitInByte = 0;
  }

  /** Read one LZW code (9..12 bits), or fewer bits for another packed code. */
  readBits(count: number): number {
    if (count > (this.byteEnd - this.byteCursor) * 8 - this.bitInByte) {
      throw new TiffDecodeError(this.truncatedMessage);
    }
    let value = 0;
    // Consume up to a byte per iteration, with no per-bit division or lookup.
    while (count > 0) {
      const available = 8 - this.bitInByte;
      const take = Math.min(count, available);
      value = (value << take)
        | ((this.bytes[this.byteCursor] >> (available - take)) & ((1 << take) - 1));
      this.bitInByte += take;
      if (this.bitInByte === 8) {
        this.byteCursor++;
        this.bitInByte = 0;
      }
      count -= take;
    }
    return value;
  }

  readBit(): number {
    if (this.byteCursor >= this.byteEnd) throw new TiffDecodeError(this.truncatedMessage);
    const bit = (this.bytes[this.byteCursor] >> (7 - this.bitInByte)) & 1;
    this.bitInByte++;
    if (this.bitInByte === 8) {
      this.byteCursor++;
      this.bitInByte = 0;
    }
    return bit;
  }
}

const TYPE_BYTE = 1;
const TYPE_SHORT = 3;
const TYPE_LONG = 4;

const TAG = {
  width: 256,
  height: 257,
  bitsPerSample: 258,
  compression: 259,
  photometric: 262,
  fillOrder: 266,
  stripOffsets: 273,
  orientation: 274,
  samplesPerPixel: 277,
  rowsPerStrip: 278,
  stripByteCounts: 279,
  planarConfiguration: 284,
  t6Options: 293,
  predictor: 317,
  tileWidth: 322,
  tileLength: 323,
  tileOffsets: 324,
  tileByteCounts: 325,
  inkSet: 332,
  extraSamples: 338,
  sampleFormat: 339,
} as const;

const TILE_TAGS = [TAG.tileWidth, TAG.tileLength, TAG.tileOffsets, TAG.tileByteCounts] as const;

function fail(message: string, cause?: unknown): never {
  throw new TiffDecodeError(
    message,
    cause === undefined ? undefined : { cause },
  );
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

  u8(offset: number, description: string): number {
    if (!this.contains(offset, 1)) fail(`Malformed TIFF ${description} range`);
    return this.view.getUint8(offset);
  }

  u16(offset: number, description: string): number {
    if (!this.contains(offset, 2)) fail(`Malformed TIFF ${description} range`);
    return this.view.getUint16(offset, this.littleEndian);
  }

  u32(offset: number, description: string): number {
    if (!this.contains(offset, 4)) fail(`Malformed TIFF ${description} range`);
    return this.view.getUint32(offset, this.littleEndian);
  }
}

interface Field {
  type: number;
  count: number;
  entryOffset: number;
}

interface Directory {
  reader: Reader;
  fields: ReadonlyMap<number, Field>;
}

interface FieldValues {
  readonly count: number;
  at(index: number): number;
}

function hasTiffByteOrderMarker(bytes: Uint8Array): boolean {
  return bytes.length >= 2
    && ((bytes[0] === 0x49 && bytes[1] === 0x49)
      || (bytes[0] === 0x4d && bytes[1] === 0x4d));
}

function firstDirectory(bytes: Uint8Array): Directory | null {
  if (!isTiff(bytes)) {
    if (!hasTiffByteOrderMarker(bytes)) return null;
    if (bytes.length < 4) fail('Malformed TIFF header: missing version');
    const littleEndian = bytes[0] === 0x49;
    const version = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
      .getUint16(2, littleEndian);
    fail(`Unsupported TIFF version: ${version}`);
  }
  if (bytes.length < 8) fail('Malformed TIFF header: missing first-IFD offset');

  const reader = new Reader(bytes, bytes[0] === 0x49);
  const ifdOffset = reader.u32(4, 'first-IFD offset');
  const count = reader.u16(ifdOffset, 'first IFD');
  const directoryLength = count * 12 + 4;
  if (!reader.contains(ifdOffset + 2, directoryLength)) {
    fail('Malformed TIFF first-IFD entry range');
  }

  const fields = new Map<number, Field>();
  for (let index = 0; index < count; index++) {
    const entryOffset = ifdOffset + 2 + index * 12;
    const tag = reader.u16(entryOffset, `IFD entry ${index}`);
    const type = reader.u16(entryOffset + 2, `IFD entry ${index} type`);
    const valueCount = reader.u32(entryOffset + 4, `IFD entry ${index} count`);
    if (fields.has(tag)) fail(`Malformed TIFF: duplicate tag ${tag}`);
    fields.set(tag, { type, count: valueCount, entryOffset });
  }
  return { reader, fields };
}

function typeSize(type: number): number {
  if (type === TYPE_BYTE) return 1;
  if (type === TYPE_SHORT) return 2;
  if (type === TYPE_LONG) return 4;
  return 0;
}

function fieldValues(
  directory: Directory,
  tag: number,
  description: string,
  allowedTypes: readonly number[],
): FieldValues | undefined {
  const field = directory.fields.get(tag);
  if (!field) return undefined;
  if (!allowedTypes.includes(field.type)) {
    fail(`Unsupported TIFF ${description} field type: ${field.type}`);
  }
  if (field.count < 1) fail(`Malformed TIFF ${description}: empty value list`);
  const size = typeSize(field.type);
  const byteLength = field.count * size;
  if (!Number.isSafeInteger(byteLength)) fail(`Malformed TIFF ${description} value count`);
  const offset = byteLength <= 4
    ? field.entryOffset + 8
    : directory.reader.u32(field.entryOffset + 8, `${description} value offset`);
  if (!directory.reader.contains(offset, byteLength)) {
    fail(`Malformed TIFF ${description} value range`);
  }

  return {
    count: field.count,
    at(index: number): number {
      if (!Number.isSafeInteger(index) || index < 0 || index >= field.count) {
        fail(`Malformed TIFF ${description} value index`);
      }
      const itemOffset = offset + index * size;
      if (field.type === TYPE_BYTE) return directory.reader.u8(itemOffset, description);
      if (field.type === TYPE_SHORT) return directory.reader.u16(itemOffset, description);
      return directory.reader.u32(itemOffset, description);
    },
  };
}

function scalar(
  directory: Directory,
  tag: number,
  description: string,
  allowedTypes: readonly number[],
  fallback?: number,
): number {
  const values = fieldValues(directory, tag, description, allowedTypes);
  if (!values) {
    if (fallback !== undefined) return fallback;
    fail(`Malformed TIFF: missing ${description}`);
  }
  if (values.count !== 1) fail(`Malformed TIFF ${description}: expected one value`);
  return values.at(0);
}

function exactValues(
  directory: Directory,
  tag: number,
  description: string,
  allowedTypes: readonly number[],
  count: number,
): FieldValues {
  const values = fieldValues(directory, tag, description, allowedTypes);
  if (!values) fail(`Malformed TIFF: missing ${description}`);
  if (values.count !== count) {
    fail(`Malformed TIFF ${description}: expected ${count} value${count === 1 ? '' : 's'}`);
  }
  return values;
}

export interface DecodedTiff {
  width: number;
  height: number;
  data: Uint8ClampedArray;
}

type PixelLayout =
  | { kind: 'gray-1'; photometric: 0 | 1; bytesPerRow: number }
  | { kind: 'gray-8'; photometric: 0 | 1; bytesPerRow: number }
  | { kind: 'rgb'; bytesPerRow: number }
  | { kind: 'rgba-associated'; bytesPerRow: number }
  | { kind: 'rgba-unassociated'; bytesPerRow: number }
  | { kind: 'cmyk'; bytesPerRow: number }
  | { kind: 'group4'; photometric: 0 | 1 };

type ChunkyLayout = Exclude<PixelLayout, { kind: 'group4' }>;

const COMPRESSION_NONE = 1;
const COMPRESSION_CCITT_GROUP4 = 4;
const COMPRESSION_LZW = 5;
const COMPRESSION_PACKBITS = 32773;

interface ParsedTiff {
  directory: Directory;
  width: number;
  height: number;
  rowsPerStrip: number;
  stripCount: number;
  stripOffsets: FieldValues;
  stripByteCounts: FieldValues;
  compression: number;
  layout: PixelLayout;
  /**
   * Sample stride of the TIFF 6.0 §14 horizontal differencing predictor, or
   * `null` when rows are stored without prediction.
   */
  differencingStride: number | null;
}

function assertSourceBudget(width: number, height: number): void {
  if (width < 1 || height < 1) fail('Malformed TIFF: image dimensions must be positive');
  const largestDimension = Math.max(width, height);
  if (largestDimension > MAX_RASTER_SOURCE_DIMENSION) {
    throw new OoxmlDecodedImageLimitError(
      'image-dimension',
      MAX_RASTER_SOURCE_DIMENSION,
      largestDimension,
    );
  }
  const pixels = width * height;
  if (pixels > MAX_RASTER_SOURCE_PIXELS) {
    throw new OoxmlDecodedImageLimitError(
      'image-pixels',
      MAX_RASTER_SOURCE_PIXELS,
      Number.isSafeInteger(pixels) ? pixels : Number.MAX_SAFE_INTEGER,
    );
  }
}

function readBits(directory: Directory, samples: number): number[] {
  if (!directory.fields.has(TAG.bitsPerSample)) return [1];
  const values = exactValues(
    directory,
    TAG.bitsPerSample,
    'BitsPerSample',
    [TYPE_SHORT],
    samples,
  );
  const bits = new Array<number>(samples);
  for (let index = 0; index < samples; index++) bits[index] = values.at(index);
  return bits;
}

function assertUnsignedSamples(directory: Directory, samples: number): void {
  if (!directory.fields.has(TAG.sampleFormat)) return;
  const formats = exactValues(
    directory,
    TAG.sampleFormat,
    'SampleFormat',
    [TYPE_SHORT],
    samples,
  );
  for (let index = 0; index < formats.count; index++) {
    if (formats.at(index) !== 1) fail('Unsupported TIFF SampleFormat: only unsigned integers are supported');
  }
}

function parseLayout(directory: Directory, width: number): { compression: number; layout: PixelLayout } {
  const compression = scalar(directory, TAG.compression, 'Compression', [TYPE_SHORT], 1);
  if (
    compression !== COMPRESSION_NONE
    && compression !== COMPRESSION_CCITT_GROUP4
    && compression !== COMPRESSION_LZW
    && compression !== COMPRESSION_PACKBITS
  ) {
    fail(`Unsupported TIFF compression: ${compression}`);
  }
  const photometric = scalar(
    directory,
    TAG.photometric,
    'PhotometricInterpretation',
    [TYPE_SHORT],
  );
  const samples = scalar(directory, TAG.samplesPerPixel, 'SamplesPerPixel', [TYPE_SHORT], 1);
  if (samples !== 1 && samples !== 3 && samples !== 4) {
    fail(`Unsupported TIFF SamplesPerPixel: ${samples}`);
  }
  const bits = readBits(directory, samples);
  assertUnsignedSamples(directory, samples);

  if (compression === COMPRESSION_CCITT_GROUP4) {
    if (
      samples !== 1
      || bits.length !== 1
      || bits[0] !== 1
      || (photometric !== 0 && photometric !== 1)
    ) {
      fail('Unsupported CCITT Group 4 TIFF sample layout');
    }
    if (directory.fields.has(TAG.extraSamples)) {
      fail('Unsupported CCITT Group 4 TIFF ExtraSamples');
    }
    const t6Options = scalar(directory, TAG.t6Options, 'T6Options', [TYPE_LONG], 0);
    if (t6Options !== 0) fail(`Unsupported TIFF T6Options: ${t6Options}`);
    return { compression, layout: { kind: 'group4', photometric } };
  }

  if (photometric === 0 || photometric === 1) {
    if (samples !== 1 || directory.fields.has(TAG.extraSamples)) {
      fail('Unsupported TIFF grayscale sample layout');
    }
    if (bits[0] === 1) {
      return {
        compression,
        layout: { kind: 'gray-1', photometric, bytesPerRow: Math.ceil(width / 8) },
      };
    }
    if (bits[0] === 8) {
      return { compression, layout: { kind: 'gray-8', photometric, bytesPerRow: width } };
    }
    fail(`Unsupported TIFF grayscale BitsPerSample: ${bits[0]}`);
  }

  if (photometric === 2) {
    if (bits.some((value) => value !== 8)) {
      fail('Unsupported TIFF RGB BitsPerSample: only 8-bit samples are supported');
    }
    if (samples === 3) {
      if (directory.fields.has(TAG.extraSamples)) fail('Unsupported TIFF RGB ExtraSamples');
      return { compression, layout: { kind: 'rgb', bytesPerRow: width * 3 } };
    }
    if (samples === 4) {
      const extras = exactValues(
        directory,
        TAG.extraSamples,
        'ExtraSamples',
        [TYPE_SHORT],
        1,
      );
      const alpha = extras.at(0);
      if (alpha === 1) {
        return { compression, layout: { kind: 'rgba-associated', bytesPerRow: width * 4 } };
      }
      if (alpha === 2) {
        return { compression, layout: { kind: 'rgba-unassociated', bytesPerRow: width * 4 } };
      }
      fail(`Unsupported TIFF ExtraSamples value: ${alpha}`);
    }
    fail(`Unsupported TIFF RGB SamplesPerPixel: ${samples}`);
  }

  if (photometric === 5) {
    if (samples !== 4 || bits.length !== 4 || bits.some((value) => value !== 8)) {
      fail('Unsupported TIFF separated sample layout: expected 8-bit process CMYK');
    }
    if (directory.fields.has(TAG.extraSamples)) fail('Unsupported TIFF CMYK ExtraSamples');
    const inkSet = scalar(directory, TAG.inkSet, 'InkSet', [TYPE_SHORT], 1);
    if (inkSet !== 1) fail(`Unsupported TIFF InkSet: ${inkSet}`);
    return { compression, layout: { kind: 'cmyk', bytesPerRow: width * 4 } };
  }

  fail(`Unsupported TIFF PhotometricInterpretation: ${photometric}`);
}

/**
 * TIFF 6.0 §14 Predictor (tag 317): 1 = no prediction, 2 = horizontal
 * differencing. §14 defines the predictor for use with LZW (Compression 5), so
 * it is refused with any other compression rather than silently ignored.
 * Differencing applies to each component separately, so for chunky pixels the
 * stride is SamplesPerPixel. The arithmetic is defined per sample; this decoder
 * implements it for its 8-bit sample classes, where it is modulo 2^8, and
 * refuses sub-byte (1-bit) rows instead of guessing a bit-level variant.
 * Other values, such as the floating-point predictor 3 of TIFF Technical
 * Note 3, stay diagnostic failures.
 */
function parseDifferencingStride(
  directory: Directory,
  compression: number,
  layout: PixelLayout,
): number | null {
  const predictor = scalar(directory, TAG.predictor, 'Predictor', [TYPE_SHORT], 1);
  if (predictor === 1) return null;
  if (predictor !== 2) fail(`Unsupported TIFF Predictor: ${predictor}`);
  if (compression !== COMPRESSION_LZW) {
    fail(`Unsupported TIFF Predictor 2 with Compression ${compression}`);
  }
  switch (layout.kind) {
    case 'gray-8':
      return 1;
    case 'rgb':
      return 3;
    case 'rgba-associated':
    case 'rgba-unassociated':
    case 'cmyk':
      return 4;
    default:
      return fail('Unsupported TIFF Predictor 2 with 1-bit samples');
  }
}

function parseTiff(bytes: Uint8Array): ParsedTiff | null {
  const directory = firstDirectory(bytes);
  if (!directory) return null;

  const width = scalar(directory, TAG.width, 'ImageWidth', [TYPE_SHORT, TYPE_LONG]);
  const height = scalar(directory, TAG.height, 'ImageLength', [TYPE_SHORT, TYPE_LONG]);
  assertSourceBudget(width, height);

  if (TILE_TAGS.some((tag) => directory.fields.has(tag))) {
    fail('Unsupported tiled TIFF: only stripped images are supported');
  }
  const orientation = scalar(directory, TAG.orientation, 'Orientation', [TYPE_SHORT], 1);
  if (orientation !== 1) fail(`Unsupported TIFF orientation: ${orientation}`);
  const planar = scalar(
    directory,
    TAG.planarConfiguration,
    'PlanarConfiguration',
    [TYPE_SHORT],
    1,
  );
  if (planar !== 1) fail(`Unsupported TIFF PlanarConfiguration: ${planar}`);
  const fillOrder = scalar(directory, TAG.fillOrder, 'FillOrder', [TYPE_SHORT], 1);
  if (fillOrder !== 1) fail(`Unsupported TIFF FillOrder: ${fillOrder}`);
  const rowsPerStrip = scalar(
    directory,
    TAG.rowsPerStrip,
    'RowsPerStrip',
    [TYPE_SHORT, TYPE_LONG],
    0xffffffff,
  );
  if (rowsPerStrip < 1) fail('Malformed TIFF RowsPerStrip: expected a positive value');
  const stripCount = Math.ceil(height / rowsPerStrip);
  const stripOffsets = exactValues(
    directory,
    TAG.stripOffsets,
    'StripOffsets',
    [TYPE_SHORT, TYPE_LONG],
    stripCount,
  );
  const stripByteCounts = exactValues(
    directory,
    TAG.stripByteCounts,
    'StripByteCounts',
    [TYPE_SHORT, TYPE_LONG],
    stripCount,
  );
  const { compression, layout } = parseLayout(directory, width);
  const differencingStride = parseDifferencingStride(directory, compression, layout);

  // Validate every declared range before reserving the target RGBA surface.
  // In particular, validate the full declared count, not merely the bytes the
  // supported decoder happens to consume.
  // Classic TIFF offsets/counts can be unsigned LONGs: their sum can exceed 2^32.
  // Keep range and row-byte arithmetic as numbers, never 32-bit coercions.
  for (let strip = 0; strip < stripCount; strip++) {
    const offset = stripOffsets.at(strip);
    const byteCount = stripByteCounts.at(strip);
    if (!directory.reader.contains(offset, byteCount)) {
      fail(`Malformed TIFF strip ${strip} byte range`);
    }
    if (compression === COMPRESSION_NONE) {
      const firstRow = strip * rowsPerStrip;
      const rowCount = Math.min(rowsPerStrip, height - firstRow);
      const bytesPerRow = (layout as ChunkyLayout).bytesPerRow;
      const required = rowCount * bytesPerRow;
      if (byteCount < required) {
        fail(`Malformed TIFF strip ${strip}: ${byteCount} bytes cannot hold ${required} bytes`);
      }
    }
  }

  return {
    directory,
    width,
    height,
    rowsPerStrip,
    stripCount,
    stripOffsets,
    stripByteCounts,
    compression,
    layout,
    differencingStride,
  };
}

interface OutputPlan {
  width: number;
  height: number;
  pixels: number;
}

function positiveTarget(value: number | undefined, description: string): number | undefined {
  if (value === undefined) return undefined;
  if (!Number.isFinite(value) || value <= 0) {
    fail(`Invalid TIFF ${description}: expected a positive finite number`);
  }
  return value;
}

function outputPlan(
  sourceWidth: number,
  sourceHeight: number,
  options: Readonly<TiffRenderOptions>,
): OutputPlan {
  const targetWidth = positiveTarget(options.targetWidthPx, 'targetWidthPx');
  const targetHeight = positiveTarget(options.targetHeightPx, 'targetHeightPx');
  const target = aspectPreservingRasterTarget(
    { width: sourceWidth, height: sourceHeight },
    targetWidth,
    targetHeight,
    true,
  );
  const width = target?.width ?? sourceWidth;
  const height = target?.height ?? sourceHeight;
  const largestDimension = Math.max(width, height);
  if (largestDimension > MAX_RASTER_DIMENSION) {
    throw new OoxmlDecodedImageLimitError(
      'image-dimension',
      MAX_RASTER_DIMENSION,
      largestDimension,
    );
  }

  const requestedLimit = options.maxRetainedPixels ?? MAX_RASTER_PIXELS;
  if (!Number.isFinite(requestedLimit) || requestedLimit < 0) {
    fail('Invalid TIFF maxRetainedPixels: expected a non-negative finite number');
  }
  const pixelLimit = Math.min(MAX_RASTER_PIXELS, Math.floor(requestedLimit));
  const pixels = width * height;
  if (pixels > pixelLimit) {
    throw new OoxmlDecodedImageLimitError('image-pixels', pixelLimit, pixels);
  }
  return { width, height, pixels };
}

function cmykChannel(channel: number, black: number): number {
  // TIFF 6.0 §16 defines ink coverage but no unique device-independent RGB
  // conversion. This multiplicative subtractive mapping matches Office for
  // unprofiled 8-bit process inks and applies uniformly to the defined class.
  return Math.round(((255 - channel) * (255 - black)) / 255);
}

function unassociate(channel: number, alpha: number): number {
  return alpha === 0 ? 0 : Math.min(255, Math.round((channel * 255) / alpha));
}

/**
 * Decoded source pixels [firstX, endX) of one row. The samples of pixel x start
 * at `bytes[rowOffset + spanBytes(layout, x)]` (1-bit: bit 7 - (x & 7) of
 * `bytes[rowOffset + (x >> 3)]`), so `rowOffset` may be negative for a span
 * that begins inside the row. Pixel x is bounded by MAX_RASTER_SOURCE_DIMENSION
 * (2^27), so its bitwise indexing is safe; strip bit positions have no such bound.
 */
interface SourceSpan {
  bytes: Uint8Array;
  rowOffset: number;
  firstX: number;
  endX: number;
}

/**
 * Visits each source row strictly top to bottom as consecutive spans. A span
 * may be a reused buffer that stays valid only during its visit.
 */
type SourceRowReader = (sourceY: number, visit: (span: SourceSpan) => void) => void;

function uncompressedRowReader(parsed: ParsedTiff, layout: ChunkyLayout): SourceRowReader {
  const span: SourceSpan = {
    bytes: parsed.directory.reader.bytes,
    rowOffset: 0,
    firstX: 0,
    endX: parsed.width,
  };
  return (sourceY, visit) => {
    const strip = Math.floor(sourceY / parsed.rowsPerStrip);
    const rowInStrip = sourceY - strip * parsed.rowsPerStrip;
    span.rowOffset = parsed.stripOffsets.at(strip) + rowInStrip * layout.bytesPerRow;
    visit(span);
  };
}

/**
 * Pixels decoded per compressed span. A multiple of 8 keeps 1-bit spans byte-aligned.
 * The span buffer (at most 4 bytes per pixel) bounds decode memory
 * independently of the source width, which may far exceed the retained width
 * when a compact strip is downsampled.
 */
const COMPRESSED_SPAN_PIXELS = 4096;

/** Bytes holding `pixels` consecutive pixels from a byte-aligned row start. */
function spanBytes(layout: ChunkyLayout, pixels: number): number {
  switch (layout.kind) {
    case 'gray-1':
      return Math.ceil(pixels / 8);
    case 'gray-8':
      return pixels;
    case 'rgb':
      return pixels * 3;
    default:
      return pixels * 4;
  }
}

/** A per-strip byte-stream decompressor (TIFF 6.0 §9 PackBits, §13 LZW). */
interface StripByteDecoder {
  beginStrip(byteOffset: number, byteLength: number): void;
  /** Fill `bytes` completely with the next decoded bytes of the active strip. */
  fill(bytes: Uint8Array): void;
  finishStrip(): void;
}

/**
 * Decode compressed strips span by span into one reusable buffer, so memory
 * stays one span plus the decoder's fixed state instead of a decompressed
 * row, strip or frame.
 */
function compressedRowReader(
  parsed: ParsedTiff,
  layout: ChunkyLayout,
  decoder: StripByteDecoder,
): SourceRowReader {
  const spanPixels = Math.min(parsed.width, COMPRESSED_SPAN_PIXELS);
  const buffer = new Uint8Array(spanBytes(layout, spanPixels));
  const stride = parsed.differencingStride;
  const previousPixel = new Uint8Array(4);
  const span: SourceSpan = { bytes: buffer, rowOffset: 0, firstX: 0, endX: 0 };
  let activeStrip = -1;
  let nextSourceY = 0;
  return (sourceY, visit) => {
    if (sourceY !== nextSourceY) fail('Internal TIFF compressed row sequence failure');
    nextSourceY++;
    const strip = Math.floor(sourceY / parsed.rowsPerStrip);
    if (strip !== activeStrip) {
      // TIFF 6.0 §9 and §13: every strip is compressed independently.
      decoder.beginStrip(parsed.stripOffsets.at(strip), parsed.stripByteCounts.at(strip));
      activeStrip = strip;
    }
    for (let firstX = 0; firstX < parsed.width; firstX += spanPixels) {
      const endX = Math.min(parsed.width, firstX + spanPixels);
      // TIFF rows are padded to a byte boundary, so only a row's final 1-bit
      // span can end in padding bits.
      const bytes = buffer.subarray(0, spanBytes(layout, endX - firstX));
      decoder.fill(bytes);
      if (stride !== null) {
        undoHorizontalDifferencing(bytes, stride, previousPixel, firstX > 0);
      }
      span.firstX = firstX;
      span.endX = endX;
      span.rowOffset = -spanBytes(layout, firstX);
      visit(span);
    }
    if (nextSourceY === parsed.height || nextSourceY % parsed.rowsPerStrip === 0) {
      decoder.finishStrip();
    }
  };
}

function lzwRowReader(parsed: ParsedTiff, layout: ChunkyLayout): SourceRowReader {
  // Library work policy derived from the source-pixel envelope: at most four
  // decoded bytes per pixel (including byte-padded bilevel rows). Data codes
  // cannot outnumber decoded bytes; allow even one clear per data code plus
  // an initial clear and EOI per strip. Expansion + reads thus fit 3B + 2S,
  // bounded by 14 * MAX_RASTER_SOURCE_PIXELS, without a compressed-size gate.
  const workLimit = 3 * layout.bytesPerRow * parsed.height + 2 * parsed.stripCount;
  return compressedRowReader(
    parsed,
    layout,
    new TiffLzwDecoder(parsed.directory.reader.bytes, workLimit),
  );
}

function packBitsRowReader(parsed: ParsedTiff, layout: ChunkyLayout): SourceRowReader {
  // Library work policy derived from the source-pixel envelope: every literal
  // or replicate header yields at least one decoded byte, so data headers
  // cannot outnumber the B decoded bytes. Strips may share byte ranges, so the
  // otherwise legal no-op header (-128) is charged too, with an allowance of
  // another B headers, bounded by 8 * MAX_RASTER_SOURCE_PIXELS, without a
  // compressed-size gate.
  const workLimit = 2 * layout.bytesPerRow * parsed.height;
  return compressedRowReader(
    parsed,
    layout,
    new TiffPackBitsDecoder(parsed.directory.reader.bytes, workLimit),
  );
}

/**
 * TIFF 6.0 §14: each stored sample after the first pixel of a row is the
 * difference from the same component of the previous pixel, modulo 2^8 for
 * the 8-bit samples this decoder accepts with the predictor. `previousPixel`
 * holds the last pixel of the preceding span; it is added only when the span
 * continues a row, and is then replaced by this span's last pixel.
 */
function undoHorizontalDifferencing(
  samples: Uint8Array,
  stride: number,
  previousPixel: Uint8Array,
  continuesRow: boolean,
): void {
  if (continuesRow) {
    for (let index = 0; index < stride; index++) {
      samples[index] = (samples[index] + previousPixel[index]) & 0xff;
    }
  }
  for (let index = stride; index < samples.length; index++) {
    samples[index] = (samples[index] + samples[index - stride]) & 0xff;
  }
  previousPixel.set(samples.subarray(samples.length - stride));
}

function chunkyRowReader(parsed: ParsedTiff, layout: ChunkyLayout): SourceRowReader {
  switch (parsed.compression) {
    case COMPRESSION_LZW:
      return lzwRowReader(parsed, layout);
    case COMPRESSION_PACKBITS:
      return packBitsRowReader(parsed, layout);
    default:
      return uncompressedRowReader(parsed, layout);
  }
}

function decodeChunkyExact(
  parsed: ParsedTiff,
  output: Uint8ClampedArray,
): void {
  const layout = parsed.layout as ChunkyLayout;
  const readRow = chunkyRowReader(parsed, layout);
  let y = 0;
  const writeSpan = (span: SourceSpan) => writeExactSpan(parsed, layout, span, y, output);
  for (; y < parsed.height; y++) readRow(y, writeSpan);
}

function writeExactSpan(
  parsed: ParsedTiff,
  layout: ChunkyLayout,
  { bytes: sourceBytes, rowOffset, firstX, endX }: SourceSpan,
  y: number,
  output: Uint8ClampedArray,
): void {
  for (let x = firstX; x < endX; x++) {
    const destination = (y * parsed.width + x) * 4;

    if (layout.kind === 'gray-1') {
      const sample = (sourceBytes[rowOffset + (x >> 3)] >> (7 - (x & 7))) & 1;
      const gray = layout.photometric === 0
        ? (sample === 0 ? 255 : 0)
        : (sample === 0 ? 0 : 255);
      output[destination] = gray;
      output[destination + 1] = gray;
      output[destination + 2] = gray;
      output[destination + 3] = 255;
      continue;
    }

    if (layout.kind === 'gray-8') {
      const sample = sourceBytes[rowOffset + x];
      const gray = layout.photometric === 0 ? 255 - sample : sample;
      output[destination] = gray;
      output[destination + 1] = gray;
      output[destination + 2] = gray;
      output[destination + 3] = 255;
      continue;
    }

    if (layout.kind === 'rgb') {
      const source = rowOffset + x * 3;
      output[destination] = sourceBytes[source];
      output[destination + 1] = sourceBytes[source + 1];
      output[destination + 2] = sourceBytes[source + 2];
      output[destination + 3] = 255;
      continue;
    }

    if (layout.kind === 'rgba-associated' || layout.kind === 'rgba-unassociated') {
      const source = rowOffset + x * 4;
      const alpha = sourceBytes[source + 3];
      output[destination] = layout.kind === 'rgba-associated'
        ? unassociate(sourceBytes[source], alpha)
        : sourceBytes[source];
      output[destination + 1] = layout.kind === 'rgba-associated'
        ? unassociate(sourceBytes[source + 1], alpha)
        : sourceBytes[source + 1];
      output[destination + 2] = layout.kind === 'rgba-associated'
        ? unassociate(sourceBytes[source + 2], alpha)
        : sourceBytes[source + 2];
      output[destination + 3] = alpha;
      continue;
    }

    const source = rowOffset + x * 4;
    const black = sourceBytes[source + 3];
    output[destination] = cmykChannel(sourceBytes[source], black);
    output[destination + 1] = cmykChannel(sourceBytes[source + 1], black);
    output[destination + 2] = cmykChannel(sourceBytes[source + 2], black);
    output[destination + 3] = 255;
  }
}

function clampByte(value: number): number {
  return Math.max(0, Math.min(255, Math.round(value)));
}

/**
 * Accumulate one source span into one retained row using exact integer box-area
 * overlap weights. Color sums use a channel×alpha representation: associated
 * samples multiply their stored premultiplied channel by 255, unassociated
 * samples multiply straight color by alpha, and opaque samples use alpha 255.
 * Dividing color sums by the accumulated alpha therefore emits straight Canvas
 * RGBA without color bleeding from transparent pixels. The sums are additive,
 * so the spans of a row may each contribute part of an output column.
 */
function accumulateAreaSpan(
  parsed: ParsedTiff,
  plan: OutputPlan,
  { bytes: sourceBytes, rowOffset, firstX, endX }: SourceSpan,
  yWeight: number,
  accumulator: Float64Array,
): void {
  const layout = parsed.layout as ChunkyLayout;
  const firstOutputX = Math.floor((firstX * plan.width) / parsed.width);
  const lastOutputX = Math.floor((endX * plan.width - 1) / parsed.width);

  for (let outputX = firstOutputX; outputX <= lastOutputX; outputX++) {
    // Work in coordinates scaled by outputWidth. All boundaries and overlap
    // weights remain exact safe integers under the source/retained caps.
    const destinationLeft = outputX * parsed.width;
    const destinationRight = (outputX + 1) * parsed.width;
    const firstSourceX = Math.max(firstX, Math.floor(destinationLeft / plan.width));
    const lastSourceX = Math.min(endX - 1, Math.floor((destinationRight - 1) / plan.width));
    let redAlpha = 0;
    let greenAlpha = 0;
    let blueAlpha = 0;
    let alpha = 0;

    for (let sourceX = firstSourceX; sourceX <= lastSourceX; sourceX++) {
      const sourceLeft = sourceX * plan.width;
      const sourceRight = (sourceX + 1) * plan.width;
      const xWeight = Math.min(destinationRight, sourceRight)
        - Math.max(destinationLeft, sourceLeft);
      const weight = xWeight * yWeight;
      let red: number;
      let green: number;
      let blue: number;
      let sampleAlpha = 255;

      if (layout.kind === 'gray-1') {
        const sample = (
          sourceBytes[rowOffset + (sourceX >> 3)]
          >> (7 - (sourceX & 7))
        ) & 1;
        red = layout.photometric === 0
          ? (sample === 0 ? 255 : 0)
          : (sample === 0 ? 0 : 255);
        green = red;
        blue = red;
      } else if (layout.kind === 'gray-8') {
        const sample = sourceBytes[rowOffset + sourceX];
        red = layout.photometric === 0 ? 255 - sample : sample;
        green = red;
        blue = red;
      } else if (layout.kind === 'rgb') {
        const source = rowOffset + sourceX * 3;
        red = sourceBytes[source];
        green = sourceBytes[source + 1];
        blue = sourceBytes[source + 2];
      } else if (layout.kind === 'rgba-associated') {
        const source = rowOffset + sourceX * 4;
        sampleAlpha = sourceBytes[source + 3];
        red = sampleAlpha === 0 ? 0 : sourceBytes[source] * 255;
        green = sampleAlpha === 0 ? 0 : sourceBytes[source + 1] * 255;
        blue = sampleAlpha === 0 ? 0 : sourceBytes[source + 2] * 255;
      } else if (layout.kind === 'rgba-unassociated') {
        const source = rowOffset + sourceX * 4;
        sampleAlpha = sourceBytes[source + 3];
        red = sourceBytes[source] * sampleAlpha;
        green = sourceBytes[source + 1] * sampleAlpha;
        blue = sourceBytes[source + 2] * sampleAlpha;
      } else {
        const source = rowOffset + sourceX * 4;
        const black = sourceBytes[source + 3];
        red = cmykChannel(sourceBytes[source], black);
        green = cmykChannel(sourceBytes[source + 1], black);
        blue = cmykChannel(sourceBytes[source + 2], black);
      }

      if (layout.kind !== 'rgba-associated' && layout.kind !== 'rgba-unassociated') {
        red *= 255;
        green *= 255;
        blue *= 255;
      }
      redAlpha += red * weight;
      greenAlpha += green * weight;
      blueAlpha += blue * weight;
      alpha += sampleAlpha * weight;
    }

    const destination = outputX * 4;
    accumulator[destination] += redAlpha;
    accumulator[destination + 1] += greenAlpha;
    accumulator[destination + 2] += blueAlpha;
    accumulator[destination + 3] += alpha;
  }
}

function writeAreaRow(
  accumulator: Float64Array,
  output: Uint8ClampedArray,
  outputWidth: number,
  outputY: number,
  normalization: number,
): void {
  for (let x = 0; x < outputWidth; x++) {
    const source = x * 4;
    const destination = (outputY * outputWidth + x) * 4;
    const alpha = accumulator[source + 3];
    output[destination] = alpha === 0 ? 0 : clampByte(accumulator[source] / alpha);
    output[destination + 1] = alpha === 0
      ? 0
      : clampByte(accumulator[source + 1] / alpha);
    output[destination + 2] = alpha === 0
      ? 0
      : clampByte(accumulator[source + 2] / alpha);
    output[destination + 3] = clampByte(alpha / normalization);
  }
}

function decodeChunkyArea(
  parsed: ParsedTiff,
  plan: OutputPlan,
  output: Uint8ClampedArray,
): void {
  const readRow = chunkyRowReader(parsed, parsed.layout as ChunkyLayout);
  // The target never upscales, so one source row overlaps at most two
  // consecutive output rows. Rows are visited once (compressed spans cannot be
  // replayed), so each in-flight output row keeps its own accumulator.
  const accumulators = [
    new Float64Array(plan.width * 4),
    new Float64Array(plan.width * 4),
  ] as const;
  const normalization = parsed.width * parsed.height;
  let writtenOutputY = -1;
  let sourceTop = 0;
  let sourceBottom = 0;
  let firstOutputY = 0;
  let lastOutputY = 0;
  const flushThrough = (outputY: number) => {
    while (writtenOutputY < outputY) {
      writtenOutputY++;
      const accumulator = accumulators[writtenOutputY & 1];
      writeAreaRow(accumulator, output, plan.width, writtenOutputY, normalization);
      accumulator.fill(0);
    }
  };
  const accumulateSpan = (span: SourceSpan) => {
    for (let outputY = firstOutputY; outputY <= lastOutputY; outputY++) {
      const destinationTop = outputY * parsed.height;
      const destinationBottom = (outputY + 1) * parsed.height;
      const yWeight = Math.min(sourceBottom, destinationBottom)
        - Math.max(sourceTop, destinationTop);
      accumulateAreaSpan(parsed, plan, span, yWeight, accumulators[outputY & 1]);
    }
  };

  for (let sourceY = 0; sourceY < parsed.height; sourceY++) {
    sourceTop = sourceY * plan.height;
    sourceBottom = (sourceY + 1) * plan.height;
    firstOutputY = Math.floor(sourceTop / parsed.height);
    lastOutputY = Math.floor((sourceBottom - 1) / parsed.height);
    if (firstOutputY <= writtenOutputY || lastOutputY > firstOutputY + 1) {
      fail('Internal TIFF area-row sequence failure');
    }
    flushThrough(firstOutputY - 1);
    readRow(sourceY, accumulateSpan);
  }
  if (lastOutputY !== plan.height - 1) fail('Internal TIFF area-row coverage failure');
  flushThrough(lastOutputY);
}

function decodeChunky(
  parsed: ParsedTiff,
  plan: OutputPlan,
  output: Uint8ClampedArray,
): void {
  if (plan.width === parsed.width && plan.height === parsed.height) {
    decodeChunkyExact(parsed, output);
  } else {
    decodeChunkyArea(parsed, plan, output);
  }
}

const LZW_CLEAR_CODE = 256;
const LZW_END_OF_INFORMATION = 257;
const LZW_FIRST_STRING_CODE = 258;
const LZW_MIN_CODE_WIDTH = 9;
const LZW_MAX_CODE_WIDTH = 12;
const LZW_TABLE_SIZE = 1 << LZW_MAX_CODE_WIDTH;

/**
 * TIFF 6.0 §13 LZW decoder for one strip at a time.
 *
 * - Codes are packed most-significant bit first. They start 9 bits wide and
 *   widen one code early: to 10, 11 and 12 bits once the next free table entry
 *   is 511, 1023 and 2047. 12 bits is the maximum, so the table never exceeds
 *   4096 entries; a conforming writer emits ClearCode (256) before that.
 * - Code 256 resets the table and width; code 257 (EndOfInformation) ends the
 *   strip. Strings start at code 258.
 * - §13's encoder writes ClearCode first. A strip that does not begin with it is
 *   refused, which also rejects the pre-6.0 LSB-first variant that this decoder
 *   does not implement.
 *
 * The table stores each string as (prefix code, final byte, first byte,
 * length), so expansion costs one write per output byte. Code reads and
 * expansion bytes share an image-wide budget, including strip finalization.
 */
class TiffLzwDecoder implements StripByteDecoder {
  private readonly prefix = new Uint16Array(LZW_TABLE_SIZE);
  private readonly suffix = new Uint8Array(LZW_TABLE_SIZE);
  private readonly first = new Uint8Array(LZW_TABLE_SIZE);
  private readonly length = new Uint16Array(LZW_TABLE_SIZE);
  private readonly pending = new Uint8Array(LZW_TABLE_SIZE);
  private pendingStart = 0;
  private pendingEnd = 0;
  private readonly reader: TiffBitReader;
  private codeWidth = LZW_MIN_CODE_WIDTH;
  private nextCode = LZW_FIRST_STRING_CODE;
  private previousCode = -1;
  private ended = false;
  private work = 0;

  constructor(
    bytes: Uint8Array,
    private readonly workLimit: number,
  ) {
    this.reader = new TiffBitReader(bytes, 0, 0, 'TIFF LZW data is truncated');
    for (let code = 0; code < 256; code++) {
      this.suffix[code] = code;
      this.first[code] = code;
      this.length[code] = 1;
    }
  }

  beginStrip(byteOffset: number, byteLength: number): void {
    this.reader.beginStrip(byteOffset, byteLength);
    this.pendingStart = 0;
    this.pendingEnd = 0;
    this.ended = false;
    this.clearTable();
    if (this.readCode() !== LZW_CLEAR_CODE) fail('TIFF LZW strip must begin with ClearCode');
  }

  /** Fill `row` completely with the next decoded bytes of the active strip. */
  fill(row: Uint8Array): void {
    let written = 0;
    while (written < row.length) {
      if (this.pendingStart === this.pendingEnd) this.decodeNextString();
      if (this.ended) fail('TIFF LZW strip ended before its declared rows');
      const count = Math.min(row.length - written, this.pendingEnd - this.pendingStart);
      row.set(this.pending.subarray(this.pendingStart, this.pendingStart + count), written);
      this.pendingStart += count;
      written += count;
    }
  }

  finishStrip(): void {
    if (this.pendingStart !== this.pendingEnd) fail('TIFF LZW data exceeds its declared rows');
    // TIFF 6.0 §13 requires EOI at strip end. libtiff's LZWPostEncode may
    // emit a final ClearCode before EOI, then only byte-alignment padding.
    // Stop at EOI; padding is not another code or decoded data.
    this.decodeNextString(false);
  }

  private chargeWork(amount: number): void {
    this.work += amount;
    if (this.work > this.workLimit) {
      throw new OoxmlDecodedImageLimitError('tiff-lzw-work', this.workLimit, this.work);
    }
  }

  private clearTable(): void {
    this.codeWidth = LZW_MIN_CODE_WIDTH;
    this.nextCode = LZW_FIRST_STRING_CODE;
    this.previousCode = -1;
  }

  private readCode(): number {
    const code = this.reader.readBits(this.codeWidth);
    this.chargeWork(1);
    return code;
  }

  private decodeNextString(allowData = true): void {
    for (;;) {
      if (this.ended) fail('TIFF LZW strip ended before its declared rows');
      const code = this.readCode();
      if (code === LZW_CLEAR_CODE) {
        this.clearTable();
        continue;
      }
      if (code === LZW_END_OF_INFORMATION) {
        this.ended = true;
        return;
      }
      if (this.previousCode < 0) {
        if (code >= LZW_CLEAR_CODE) fail(`Invalid TIFF LZW code: ${code}`);
      } else {
        this.addString(code);
      }
      if (!allowData) fail('TIFF LZW data exceeds its declared rows');
      this.expand(code);
      this.previousCode = code;
      return;
    }
  }

  /** Add previous string + first byte of `code` (KwKwK when `code` is that new entry). */
  private addString(code: number): void {
    if (code > this.nextCode) fail(`Invalid TIFF LZW code: ${code}`);
    if (this.nextCode >= LZW_TABLE_SIZE) fail('TIFF LZW code table overflow');
    const previous = this.previousCode;
    const entry = this.nextCode;
    this.prefix[entry] = previous;
    this.suffix[entry] = code === entry ? this.first[previous] : this.first[code];
    this.first[entry] = this.first[previous];
    this.length[entry] = this.length[previous] + 1;
    this.nextCode++;
    if (this.nextCode >= (1 << this.codeWidth) - 1 && this.codeWidth < LZW_MAX_CODE_WIDTH) {
      this.codeWidth++;
    }
  }

  private expand(code: number): void {
    let current = code;
    const length = this.length[current];
    this.chargeWork(length);
    for (let index = length - 1; index >= 0; index--) {
      this.pending[index] = this.suffix[current];
      current = this.prefix[current];
    }
    this.pendingStart = 0;
    this.pendingEnd = length;
  }
}

/**
 * TIFF 6.0 §9 PackBits decoder for one strip at a time. Each header byte n,
 * read as a signed 8-bit value, introduces
 * - 0..127: a literal run of the next n + 1 bytes;
 * - -127..-1: a replicate run of the next byte, repeated -n + 1 times;
 * - -128: no operation.
 *
 * §9 asks writers to pack each row separately. A run that continues into the
 * next row of the same strip decodes to the same bytes either way, stays
 * within the strip's validated range and is decoded as a byte stream, like
 * LZW. Fail closed instead when a header or literal needs bytes past the strip
 * byte count (truncation), or a run continues past the strip's declared rows
 * (overrun). Unread bytes after the strip's final run are padding.
 */
class TiffPackBitsDecoder implements StripByteDecoder {
  private cursor = 0;
  private end = 0;
  /** Literal bytes still to copy from `cursor`, already range-checked. */
  private literal = 0;
  private repeat = 0;
  private repeatByte = 0;
  private work = 0;

  constructor(
    private readonly bytes: Uint8Array,
    private readonly workLimit: number,
  ) {}

  beginStrip(byteOffset: number, byteLength: number): void {
    this.cursor = byteOffset;
    this.end = byteOffset + byteLength;
    this.literal = 0;
    this.repeat = 0;
  }

  fill(row: Uint8Array): void {
    let written = 0;
    while (written < row.length) {
      if (this.literal > 0) {
        const count = Math.min(this.literal, row.length - written);
        row.set(this.bytes.subarray(this.cursor, this.cursor + count), written);
        this.cursor += count;
        this.literal -= count;
        written += count;
      } else if (this.repeat > 0) {
        const count = Math.min(this.repeat, row.length - written);
        row.fill(this.repeatByte, written, written + count);
        this.repeat -= count;
        written += count;
      } else {
        this.readHeader();
      }
    }
  }

  finishStrip(): void {
    if (this.literal > 0 || this.repeat > 0) {
      fail('TIFF PackBits run exceeds its declared rows');
    }
  }

  private readHeader(): void {
    if (this.cursor >= this.end) fail('TIFF PackBits data is truncated');
    const header = this.bytes[this.cursor++];
    this.work++;
    if (this.work > this.workLimit) {
      throw new OoxmlDecodedImageLimitError('tiff-packbits-work', this.workLimit, this.work);
    }
    if (header < 128) {
      const count = header + 1;
      if (count > this.end - this.cursor) fail('TIFF PackBits data is truncated');
      this.literal = count;
    } else if (header > 128) {
      if (this.cursor >= this.end) fail('TIFF PackBits data is truncated');
      this.repeatByte = this.bytes[this.cursor++];
      this.repeat = 257 - header;
    }
  }
}

type Code = readonly [bits: string, value: number];

const SHARED_MAKEUP_CODES: readonly Code[] = [
  ['00000001000', 1792], ['00000001100', 1856], ['00000001101', 1920],
  ['000000010010', 1984], ['000000010011', 2048], ['000000010100', 2112],
  ['000000010101', 2176], ['000000010110', 2240], ['000000010111', 2304],
  ['000000011100', 2368], ['000000011101', 2432], ['000000011110', 2496],
  ['000000011111', 2560],
];

const WHITE_CODES: readonly Code[] = [
  ['00110101', 0], ['000111', 1], ['0111', 2], ['1000', 3],
  ['1011', 4], ['1100', 5], ['1110', 6], ['1111', 7],
  ['10011', 8], ['10100', 9], ['00111', 10], ['01000', 11],
  ['001000', 12], ['000011', 13], ['110100', 14], ['110101', 15],
  ['101010', 16], ['101011', 17], ['0100111', 18], ['0001100', 19],
  ['0001000', 20], ['0010111', 21], ['0000011', 22], ['0000100', 23],
  ['0101000', 24], ['0101011', 25], ['0010011', 26], ['0100100', 27],
  ['0011000', 28], ['00000010', 29], ['00000011', 30], ['00011010', 31],
  ['00011011', 32], ['00010010', 33], ['00010011', 34], ['00010100', 35],
  ['00010101', 36], ['00010110', 37], ['00010111', 38], ['00101000', 39],
  ['00101001', 40], ['00101010', 41], ['00101011', 42], ['00101100', 43],
  ['00101101', 44], ['00000100', 45], ['00000101', 46], ['00001010', 47],
  ['00001011', 48], ['01010010', 49], ['01010011', 50], ['01010100', 51],
  ['01010101', 52], ['00100100', 53], ['00100101', 54], ['01011000', 55],
  ['01011001', 56], ['01011010', 57], ['01011011', 58], ['01001010', 59],
  ['01001011', 60], ['00110010', 61], ['00110011', 62], ['00110100', 63],
  ['11011', 64], ['10010', 128], ['010111', 192], ['0110111', 256],
  ['00110110', 320], ['00110111', 384], ['01100100', 448], ['01100101', 512],
  ['01101000', 576], ['01100111', 640], ['011001100', 704], ['011001101', 768],
  ['011010010', 832], ['011010011', 896], ['011010100', 960], ['011010101', 1024],
  ['011010110', 1088], ['011010111', 1152], ['011011000', 1216], ['011011001', 1280],
  ['011011010', 1344], ['011011011', 1408], ['010011000', 1472], ['010011001', 1536],
  ['010011010', 1600], ['011000', 1664], ['010011011', 1728],
  ...SHARED_MAKEUP_CODES,
];

const BLACK_CODES: readonly Code[] = [
  ['0000110111', 0], ['010', 1], ['11', 2], ['10', 3],
  ['011', 4], ['0011', 5], ['0010', 6], ['00011', 7],
  ['000101', 8], ['000100', 9], ['0000100', 10], ['0000101', 11],
  ['0000111', 12], ['00000100', 13], ['00000111', 14], ['000011000', 15],
  ['0000010111', 16], ['0000011000', 17], ['0000001000', 18], ['00001100111', 19],
  ['00001101000', 20], ['00001101100', 21], ['00000110111', 22], ['00000101000', 23],
  ['00000010111', 24], ['00000011000', 25], ['000011001010', 26], ['000011001011', 27],
  ['000011001100', 28], ['000011001101', 29], ['000001101000', 30], ['000001101001', 31],
  ['000001101010', 32], ['000001101011', 33], ['000011010010', 34], ['000011010011', 35],
  ['000011010100', 36], ['000011010101', 37], ['000011010110', 38], ['000011010111', 39],
  ['000001101100', 40], ['000001101101', 41], ['000011011010', 42], ['000011011011', 43],
  ['000001010100', 44], ['000001010101', 45], ['000001010110', 46], ['000001010111', 47],
  ['000001100100', 48], ['000001100101', 49], ['000001010010', 50], ['000001010011', 51],
  ['000000100100', 52], ['000000110111', 53], ['000000111000', 54], ['000000100111', 55],
  ['000000101000', 56], ['000001011000', 57], ['000001011001', 58], ['000000101011', 59],
  ['000000101100', 60], ['000001011010', 61], ['000001100110', 62], ['000001100111', 63],
  ['0000001111', 64], ['000011001000', 128], ['000011001001', 192], ['000001011011', 256],
  ['000000110011', 320], ['000000110100', 384], ['000000110101', 448], ['0000001101100', 512],
  ['0000001101101', 576], ['0000001001010', 640], ['0000001001011', 704], ['0000001001100', 768],
  ['0000001001101', 832], ['0000001110010', 896], ['0000001110011', 960], ['0000001110100', 1024],
  ['0000001110101', 1088], ['0000001110110', 1152], ['0000001110111', 1216], ['0000001010010', 1280],
  ['0000001010011', 1344], ['0000001010100', 1408], ['0000001010101', 1472], ['0000001011010', 1536],
  ['0000001011011', 1600], ['0000001100100', 1664], ['0000001100101', 1728],
  ...SHARED_MAKEUP_CODES,
];

function codeTable(codes: readonly Code[]): ReadonlyMap<number, number> {
  const table = new Map<number, number>();
  for (const [bits, value] of codes) {
    table.set((1 << bits.length) | Number.parseInt(bits, 2), value);
  }
  return table;
}

const WHITE_TABLE = codeTable(WHITE_CODES);
const BLACK_TABLE = codeTable(BLACK_CODES);

type TwoDimensionalMode = 'pass' | 'horizontal' | -3 | -2 | -1 | 0 | 1 | 2 | 3;

const MODE_TABLE = new Map<number, TwoDimensionalMode>([
  [(1 << 1) | 0b1, 0],
  [(1 << 3) | 0b001, 'horizontal'],
  [(1 << 3) | 0b010, -1],
  [(1 << 3) | 0b011, 1],
  [(1 << 4) | 0b0001, 'pass'],
  [(1 << 6) | 0b000010, -2],
  [(1 << 6) | 0b000011, 2],
  [(1 << 7) | 0b0000010, -3],
  [(1 << 7) | 0b0000011, 3],
]);

function decodeMode(reader: TiffBitReader): TwoDimensionalMode {
  let code = 0;
  for (let length = 1; length <= 7; length++) {
    code = (code << 1) | reader.readBit();
    const mode = MODE_TABLE.get((1 << length) | code);
    if (mode !== undefined) return mode;
  }
  fail('Invalid or unsupported CCITT Group 4 two-dimensional mode');
}

function decodeRun(reader: TiffBitReader, black: boolean, maximum: number): number {
  const table = black ? BLACK_TABLE : WHITE_TABLE;
  let total = 0;
  const maximumCodes = Math.ceil(maximum / 64) + 64;
  for (let item = 0; item < maximumCodes; item++) {
    let code = 0;
    let value: number | undefined;
    for (let length = 1; length <= 13; length++) {
      code = (code << 1) | reader.readBit();
      value = table.get((1 << length) | code);
      if (value !== undefined) break;
    }
    if (value === undefined) fail('Invalid CCITT Group 4 run-length code');
    total += value;
    if (total > maximum) fail('CCITT Group 4 run exceeds the scanline width');
    if (value < 64) return total;
  }
  fail('CCITT Group 4 run-length code limit exceeded');
}

/**
 * Ordered changing-element positions for one T.6 coding line. Runs always begin
 * white, alternate color at every entry, and end with the width sentinel. A
 * zero first entry therefore represents the legal zero-length leading white
 * run used by an all-black line.
 */
type Group4Row = number[];

/**
 * A retained scanline cannot independently represent more than one transition
 * per pixel plus its terminal sentinel. Bound the sparse run form to that same
 * maximum-axis complexity so a much wider downsampled source cannot amplify a
 * compact strip into an unbounded JavaScript number array.
 */
const MAX_GROUP4_CHANGING_ELEMENTS = MAX_RASTER_DIMENSION + 1;

interface Group4WorkDiagnostics {
  modeCount: number;
  referenceProbeCount: number;
  areaSegmentCount: number;
}

interface Group4ReferenceCursor {
  index: number;
}

function toggleGroup4Boundary(row: Group4Row, position: number): void {
  const last = row[row.length - 1];
  if (last !== undefined && last > position) {
    fail('Internal TIFF Group 4 changing-element order failure');
  }
  // Consecutive zero-length runs create two transitions at the same position.
  // They cancel in the visible reference line and must not survive as phantom
  // changing elements for the next row.
  if (last === position) row.pop();
  else {
    if (row.length >= MAX_GROUP4_CHANGING_ELEMENTS) {
      fail('CCITT Group 4 changing-element limit exceeded');
    }
    row.push(position);
  }
}

function referenceChanges(
  reference: readonly number[],
  width: number,
  a0: number,
  color: number,
  initial: boolean,
  cursor: Group4ReferenceCursor,
  diagnostics?: Group4WorkDiagnostics,
): readonly [b1: number, b2: number] {
  const firstPosition = initial ? a0 : a0 + 1;
  // a0 never moves left, so the lower-bound cursor visits each reference run
  // at most once during this coding line. Keep this cursor at the lower bound
  // rather than b1: a negative vertical mode can leave a skipped, wrong-color
  // boundary relevant to the next operation.
  while (cursor.index < reference.length) {
    if (diagnostics) diagnostics.referenceProbeCount++;
    if (reference[cursor.index] >= firstPosition) break;
    cursor.index++;
  }
  let index = cursor.index;
  // The color after changing element i is black for even i and white for odd
  // i because every row starts white. At most one alternating boundary needs
  // to be skipped to find b1's required opposite color.
  if (index < reference.length && ((index + 1) & 1) !== (color ^ 1)) index++;
  return [reference[index] ?? width, reference[index + 1] ?? width];
}

function decodeGroup4Row(
  reader: TiffBitReader,
  reference: readonly number[],
  coding: Group4Row,
  width: number,
  diagnostics?: Group4WorkDiagnostics,
): void {
  coding.length = 0;
  let a0 = 0;
  let color = 0; // Every T.6 coding line begins with a (possibly zero) white run.
  let initial = true;
  const referenceCursor: Group4ReferenceCursor = { index: 0 };
  const operationLimit = width * 4 + 32;

  for (let operation = 0; a0 < width; operation++) {
    if (operation >= operationLimit) fail('CCITT Group 4 scanline operation limit exceeded');
    if (diagnostics) diagnostics.modeCount++;
    const mode = decodeMode(reader);

    if (mode === 'horizontal') {
      const firstRun = decodeRun(reader, color === 1, width - a0);
      const a1 = a0 + firstRun;
      const secondRun = decodeRun(reader, color === 0, width - a1);
      const a2 = a1 + secondRun;
      if (a2 <= a0) fail('Invalid CCITT Group 4 horizontal mode with no progress');
      toggleGroup4Boundary(coding, a1);
      toggleGroup4Boundary(coding, a2);
      a0 = a2;
      initial = false;
      continue;
    }

    const [b1, b2] = referenceChanges(
      reference,
      width,
      a0,
      color,
      initial,
      referenceCursor,
      diagnostics,
    );
    if (mode === 'pass') {
      if (b2 <= a0) fail('Invalid CCITT Group 4 pass mode with no progress');
      a0 = b2;
      initial = false;
      continue;
    }

    const a1 = b1 + mode;
    if (a1 < a0 || a1 > width) fail('Invalid CCITT Group 4 vertical mode position');
    toggleGroup4Boundary(coding, a1);
    a0 = a1;
    color ^= 1;
    initial = false;
  }
  if (coding[coding.length - 1] !== width) {
    if (coding.length >= MAX_GROUP4_CHANGING_ELEMENTS) {
      fail('CCITT Group 4 changing-element limit exceeded');
    }
    coding.push(width);
  }
}

function writeGroup4ExactRow(
  row: readonly number[],
  output: Uint8ClampedArray,
  width: number,
  outputY: number,
  photometric: 0 | 1,
): void {
  let start = 0;
  for (let run = 0; run < row.length; run++) {
    const end = row[run];
    const sample = run & 1;
    const gray = photometric === 0
      ? (sample === 0 ? 255 : 0)
      : (sample === 0 ? 0 : 255);
    for (let x = start; x < end; x++) {
      const destination = (outputY * width + x) * 4;
      output[destination] = gray;
      output[destination + 1] = gray;
      output[destination + 2] = gray;
      output[destination + 3] = 255;
    }
    start = end;
  }
}

/** Accumulate one Group 4 run row without expanding it to source-width pixels. */
function accumulateGroup4AreaRow(
  parsed: ParsedTiff,
  plan: OutputPlan,
  row: readonly number[],
  photometric: 0 | 1,
  yWeight: number,
  accumulator: Float64Array,
  diagnostics?: Group4WorkDiagnostics,
): void {
  let run = 0;
  let runEnd = row[0] * plan.width;

  for (let outputX = 0; outputX < plan.width; outputX++) {
    const destinationLeft = outputX * parsed.width;
    const destinationRight = (outputX + 1) * parsed.width;
    let position = destinationLeft;

    while (position < destinationRight) {
      if (diagnostics) diagnostics.areaSegmentCount++;
      while (runEnd <= position && run + 1 < row.length) {
        run++;
        runEnd = row[run] * plan.width;
      }
      const overlapEnd = Math.min(destinationRight, runEnd);
      if (overlapEnd <= position) {
        fail('Internal TIFF Group 4 area-run coverage failure');
      }
      const sample = run & 1;
      const gray = photometric === 0
        ? (sample === 0 ? 255 : 0)
        : (sample === 0 ? 0 : 255);
      const weight = (overlapEnd - position) * yWeight;
      const destination = outputX * 4;
      const grayAlpha = gray * 255 * weight;
      accumulator[destination] += grayAlpha;
      accumulator[destination + 1] += grayAlpha;
      accumulator[destination + 2] += grayAlpha;
      accumulator[destination + 3] += 255 * weight;
      position = overlapEnd;
    }
  }
}

function decodeGroup4(
  parsed: ParsedTiff,
  plan: OutputPlan,
  output: Uint8ClampedArray,
  diagnostics?: Group4WorkDiagnostics,
): void {
  const layout = parsed.layout as Extract<PixelLayout, { kind: 'group4' }>;
  let reference: Group4Row = [parsed.width];
  let coding: Group4Row = [];
  const areaFilter = plan.width !== parsed.width || plan.height !== parsed.height;
  const accumulator = areaFilter ? new Float64Array(plan.width * 4) : null;
  const normalization = parsed.width * parsed.height;
  let activeOutputY = -1;

  for (let strip = 0; strip < parsed.stripCount; strip++) {
    // TIFF 6.0: every strip starts with an imaginary white reference line.
    reference.length = 1;
    reference[0] = parsed.width;
    const firstRow = strip * parsed.rowsPerStrip;
    const rowCount = Math.min(parsed.rowsPerStrip, parsed.height - firstRow);
    const reader = new TiffBitReader(
      parsed.directory.reader.bytes,
      parsed.stripOffsets.at(strip),
      parsed.stripByteCounts.at(strip),
      'CCITT Group 4 data is truncated',
    );
    for (let row = 0; row < rowCount; row++) {
      const sourceY = firstRow + row;
      decodeGroup4Row(reader, reference, coding, parsed.width, diagnostics);
      if (!accumulator) {
        writeGroup4ExactRow(coding, output, parsed.width, sourceY, layout.photometric);
      } else {
        const sourceTop = sourceY * plan.height;
        const sourceBottom = (sourceY + 1) * plan.height;
        const firstOutputY = Math.floor(sourceTop / parsed.height);
        const lastOutputY = Math.floor((sourceBottom - 1) / parsed.height);
        for (let outputY = firstOutputY; outputY <= lastOutputY; outputY++) {
          if (outputY !== activeOutputY) {
            if (activeOutputY >= 0) {
              writeAreaRow(accumulator, output, plan.width, activeOutputY, normalization);
              accumulator.fill(0);
            }
            if (outputY !== activeOutputY + 1) {
              fail('Internal TIFF Group 4 area-row sequence failure');
            }
            activeOutputY = outputY;
          }
          const destinationTop = outputY * parsed.height;
          const destinationBottom = (outputY + 1) * parsed.height;
          const yWeight = Math.min(sourceBottom, destinationBottom)
            - Math.max(sourceTop, destinationTop);
          accumulateGroup4AreaRow(
            parsed,
            plan,
            coding,
            layout.photometric,
            yWeight,
            accumulator,
            diagnostics,
          );
        }
      }
      [reference, coding] = [coding, reference];
    }
  }
  if (accumulator) {
    if (activeOutputY !== plan.height - 1) fail('Internal TIFF Group 4 area-row coverage failure');
    writeAreaRow(accumulator, output, plan.width, activeOutputY, normalization);
  }
}

function decodeTiffRgbaInternal(
  bytes: Uint8Array,
  options: Readonly<TiffRenderOptions>,
  diagnostics?: Group4WorkDiagnostics,
): DecodedTiff | null {
  const parsed = parseTiff(bytes);
  if (!parsed) return null;
  const plan = outputPlan(parsed.width, parsed.height, options);
  const output = new Uint8ClampedArray(plan.pixels * 4);
  if (parsed.layout.kind === 'group4') decodeGroup4(parsed, plan, output, diagnostics);
  else decodeChunky(parsed, plan, output);
  return { width: plan.width, height: plan.height, data: output };
}

/** Decode a supported classic TIFF directly into its retained-size Canvas RGBA surface. */
export function decodeTiffRgba(
  bytes: Uint8Array,
  options: Readonly<TiffRenderOptions> = {},
): DecodedTiff | null {
  return decodeTiffRgbaInternal(bytes, options);
}

/** Internal module test seam; intentionally not re-exported by the package root. */
export function __test_decodeTiffRgbaWithGroup4Diagnostics(
  bytes: Uint8Array,
  options: Readonly<TiffRenderOptions> = {},
): Readonly<{
  decoded: DecodedTiff | null;
  modeCount: number;
  referenceProbeCount: number;
  areaSegmentCount: number;
}> {
  const diagnostics: Group4WorkDiagnostics = {
    modeCount: 0,
    referenceProbeCount: 0,
    areaSegmentCount: 0,
  };
  const decoded = decodeTiffRgbaInternal(bytes, options, diagnostics);
  return Object.freeze({ decoded, ...diagnostics });
}

/** Rasterize a supported TIFF without asking the browser to decode TIFF bytes. */
export async function renderTiffToBitmap(
  bytes: Uint8Array,
  options: Readonly<TiffRenderOptions> = {},
): Promise<ImageBitmap | null> {
  const decoded = decodeTiffRgba(bytes, options);
  if (!decoded) return null;

  if (typeof ImageData !== 'function') {
    fail('TIFF ImageData construction is unavailable');
  }
  let imageData: ImageData;
  try {
    // The provided Uint8ClampedArray becomes the ImageData backing store; this
    // avoids a second retained-size RGBA copy and a transient Canvas surface.
    imageData = new ImageData(
      // This decoder always constructs a fresh ArrayBuffer-backed view; the
      // cast excludes SharedArrayBuffer from the public typed-array surface.
      decoded.data as ImageData['data'],
      decoded.width,
      decoded.height,
    );
  } catch (error) {
    fail('TIFF ImageData construction failed', error);
  }

  if (typeof createImageBitmap !== 'function') {
    fail('TIFF ImageBitmap creation is unavailable');
  }
  let bitmap: ImageBitmap | null;
  try {
    bitmap = await createImageBitmap(imageData);
  } catch (error) {
    fail('TIFF ImageBitmap creation failed', error);
  }
  if (!bitmap) fail('TIFF ImageBitmap creation returned no bitmap');
  return bitmap;
}
