// TIFF Revision 6.0 decoder for OOXML image parts. ECMA-376 Part 1
// §15.2.14 permits TIFF image parts, while browser bitmap decoders do not
// consistently support them. Keep this compatibility codec deliberately
// tag-defined: classic TIFF, first IFD, stripped top-left chunky pixels, and the
// baseline bilevel/grayscale/RGB/process-CMYK classes listed below.

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

interface ParsedTiff {
  directory: Directory;
  width: number;
  height: number;
  rowsPerStrip: number;
  stripCount: number;
  stripOffsets: FieldValues;
  stripByteCounts: FieldValues;
  layout: PixelLayout;
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
  if (compression !== 1 && compression !== 4) {
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

  if (compression === 4) {
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
  const predictor = scalar(directory, TAG.predictor, 'Predictor', [TYPE_SHORT], 1);
  if (predictor !== 1) fail(`Unsupported TIFF Predictor: ${predictor}`);

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

  // Validate every declared range before reserving the target RGBA surface.
  // In particular, validate the full declared count, not merely the bytes the
  // supported decoder happens to consume.
  for (let strip = 0; strip < stripCount; strip++) {
    const offset = stripOffsets.at(strip);
    const byteCount = stripByteCounts.at(strip);
    if (!directory.reader.contains(offset, byteCount)) {
      fail(`Malformed TIFF strip ${strip} byte range`);
    }
    if (compression === 1) {
      const firstRow = strip * rowsPerStrip;
      const rowCount = Math.min(rowsPerStrip, height - firstRow);
      const bytesPerRow = (layout as Exclude<PixelLayout, { kind: 'group4' }>).bytesPerRow;
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
    layout,
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

function decodeUncompressedExact(
  parsed: ParsedTiff,
  output: Uint8ClampedArray,
): void {
  const layout = parsed.layout as Exclude<PixelLayout, { kind: 'group4' }>;
  const sourceBytes = parsed.directory.reader.bytes;

  for (let y = 0; y < parsed.height; y++) {
    const strip = Math.floor(y / parsed.rowsPerStrip);
    const rowInStrip = y - strip * parsed.rowsPerStrip;
    const rowOffset = parsed.stripOffsets.at(strip) + rowInStrip * layout.bytesPerRow;
    for (let x = 0; x < parsed.width; x++) {
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
}

function clampByte(value: number): number {
  return Math.max(0, Math.min(255, Math.round(value)));
}

/**
 * Accumulate one source row into one retained row using exact integer box-area
 * overlap weights. Color sums use a channel×alpha representation: associated
 * samples multiply their stored premultiplied channel by 255, unassociated
 * samples multiply straight color by alpha, and opaque samples use alpha 255.
 * Dividing color sums by the accumulated alpha therefore emits straight Canvas
 * RGBA without color bleeding from transparent pixels.
 */
function accumulateAreaRow(
  parsed: ParsedTiff,
  plan: OutputPlan,
  rowOffset: number,
  yWeight: number,
  accumulator: Float64Array,
): void {
  const layout = parsed.layout as Exclude<PixelLayout, { kind: 'group4' }>;
  const sourceBytes = parsed.directory.reader.bytes;

  for (let outputX = 0; outputX < plan.width; outputX++) {
    // Work in coordinates scaled by outputWidth. All boundaries and overlap
    // weights remain exact safe integers under the source/retained caps.
    const destinationLeft = outputX * parsed.width;
    const destinationRight = (outputX + 1) * parsed.width;
    const firstSourceX = Math.floor(destinationLeft / plan.width);
    const lastSourceX = Math.floor((destinationRight - 1) / plan.width);
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

function decodeUncompressedArea(
  parsed: ParsedTiff,
  plan: OutputPlan,
  output: Uint8ClampedArray,
): void {
  const layout = parsed.layout as Exclude<PixelLayout, { kind: 'group4' }>;
  const accumulator = new Float64Array(plan.width * 4);
  const normalization = parsed.width * parsed.height;
  let activeOutputY = -1;

  for (let sourceY = 0; sourceY < parsed.height; sourceY++) {
    const strip = Math.floor(sourceY / parsed.rowsPerStrip);
    const rowInStrip = sourceY - strip * parsed.rowsPerStrip;
    const rowOffset = parsed.stripOffsets.at(strip) + rowInStrip * layout.bytesPerRow;
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
        if (outputY !== activeOutputY + 1) fail('Internal TIFF area-row sequence failure');
        activeOutputY = outputY;
      }
      const destinationTop = outputY * parsed.height;
      const destinationBottom = (outputY + 1) * parsed.height;
      const yWeight = Math.min(sourceBottom, destinationBottom)
        - Math.max(sourceTop, destinationTop);
      accumulateAreaRow(parsed, plan, rowOffset, yWeight, accumulator);
    }
  }
  if (activeOutputY !== plan.height - 1) fail('Internal TIFF area-row coverage failure');
  writeAreaRow(accumulator, output, plan.width, activeOutputY, normalization);
}

function decodeUncompressed(
  parsed: ParsedTiff,
  plan: OutputPlan,
  output: Uint8ClampedArray,
): void {
  if (plan.width === parsed.width && plan.height === parsed.height) {
    decodeUncompressedExact(parsed, output);
  } else {
    decodeUncompressedArea(parsed, plan, output);
  }
}

class StripBitReader {
  private bitOffset = 0;

  constructor(
    private readonly bytes: Uint8Array,
    private readonly byteOffset: number,
    private readonly byteLength: number,
  ) {}

  readBit(): number {
    if (this.bitOffset >= this.byteLength * 8) {
      fail('CCITT Group 4 data is truncated');
    }
    const bit = (
      this.bytes[this.byteOffset + (this.bitOffset >> 3)]
      >> (7 - (this.bitOffset & 7))
    ) & 1;
    this.bitOffset++;
    return bit;
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

function decodeMode(reader: StripBitReader): TwoDimensionalMode {
  let code = 0;
  for (let length = 1; length <= 7; length++) {
    code = (code << 1) | reader.readBit();
    const mode = MODE_TABLE.get((1 << length) | code);
    if (mode !== undefined) return mode;
  }
  fail('Invalid or unsupported CCITT Group 4 two-dimensional mode');
}

function decodeRun(reader: StripBitReader, black: boolean, maximum: number): number {
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
  reader: StripBitReader,
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
    const reader = new StripBitReader(
      parsed.directory.reader.bytes,
      parsed.stripOffsets.at(strip),
      parsed.stripByteCounts.at(strip),
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
  else decodeUncompressed(parsed, plan, output);
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
