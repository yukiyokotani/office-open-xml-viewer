import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  __test_decodeTiffRgbaWithGroup4Diagnostics,
  decodeTiffRgba,
  renderTiffToBitmap,
  TiffDecodeError,
} from './tiff.js';
import { isTiff } from './tiff-contract.js';
import { getOoxmlDecodedImageLimitDetails, OoxmlDecodedImageLimitError } from './pixel-budget.js';

type ByteOrder = 'little' | 'big';

interface TiffFixtureOptions {
  byteOrder?: ByteOrder;
  width: number;
  height: number;
  bitsPerSample: readonly number[];
  compression?: number;
  photometric: number;
  samplesPerPixel?: number;
  rowsPerStrip?: number;
  strips: readonly Uint8Array[];
  orientation?: number;
  planarConfiguration?: number;
  inkSet?: number;
  extraSamples?: readonly number[];
  fillOrder?: number;
  predictor?: number;
}

interface FixtureField {
  tag: number;
  type: 3 | 4;
  values: number[];
  payloadOffset?: number;
}

/** Build a small classic TIFF with explicit strip payloads. The synthetic bytes
 * are public test data and contain no document-derived content. */
function classicTiff({
  byteOrder = 'little',
  width,
  height,
  bitsPerSample,
  compression = 1,
  photometric,
  samplesPerPixel = bitsPerSample.length,
  rowsPerStrip = height,
  strips,
  orientation = 1,
  planarConfiguration = 1,
  inkSet,
  extraSamples,
  fillOrder,
  predictor,
}: TiffFixtureOptions): Uint8Array {
  const stripCount = Math.ceil(height / rowsPerStrip);
  if (strips.length !== stripCount) throw new Error('fixture strip count mismatch');
  const fields: FixtureField[] = [
    { tag: 256, type: 4 as const, values: [width] },
    { tag: 257, type: 4 as const, values: [height] },
    { tag: 258, type: 3 as const, values: [...bitsPerSample] },
    { tag: 259, type: 3 as const, values: [compression] },
    { tag: 262, type: 3 as const, values: [photometric] },
    { tag: 273, type: 4 as const, values: new Array<number>(stripCount).fill(0) },
    { tag: 274, type: 3 as const, values: [orientation] },
    { tag: 277, type: 3 as const, values: [samplesPerPixel] },
    { tag: 278, type: 4 as const, values: [rowsPerStrip] },
    { tag: 279, type: 4 as const, values: strips.map((strip) => strip.length) },
    { tag: 284, type: 3 as const, values: [planarConfiguration] },
    ...(fillOrder === undefined ? [] : [{ tag: 266, type: 3 as const, values: [fillOrder] }]),
    ...(predictor === undefined ? [] : [{ tag: 317, type: 3 as const, values: [predictor] }]),
    ...(inkSet === undefined ? [] : [{ tag: 332, type: 3 as const, values: [inkSet] }]),
    ...(extraSamples === undefined
      ? []
      : [{ tag: 338, type: 3 as const, values: [...extraSamples] }]),
  ].sort((left, right) => left.tag - right.tag);
  const ifdOffset = 8;
  const ifdBytes = 2 + fields.length * 12 + 4;
  let payloadOffset = ifdOffset + ifdBytes;
  for (const field of fields) {
    const byteLength = field.values.length * (field.type === 3 ? 2 : 4);
    if (byteLength <= 4) continue;
    if (payloadOffset % 2 !== 0) payloadOffset++;
    field.payloadOffset = payloadOffset;
    payloadOffset += byteLength;
  }
  if (payloadOffset % 2 !== 0) payloadOffset++;
  let stripOffset = payloadOffset;
  const stripOffsets = fields.find((field) => field.tag === 273) as FixtureField;
  stripOffsets.values = strips.map((strip) => {
    const offset = stripOffset;
    stripOffset += strip.length;
    return offset;
  });

  const bytes = new Uint8Array(stripOffset);
  const little = byteOrder === 'little';
  const view = new DataView(bytes.buffer);
  const u16 = (offset: number, value: number) => view.setUint16(offset, value, little);
  const u32 = (offset: number, value: number) => view.setUint32(offset, value, little);
  const writeValues = (offset: number, field: FixtureField) => {
    const stride = field.type === 3 ? 2 : 4;
    for (let index = 0; index < field.values.length; index++) {
      if (field.type === 3) u16(offset + index * stride, field.values[index]);
      else u32(offset + index * stride, field.values[index]);
    }
  };

  bytes.set(little ? [0x49, 0x49] : [0x4d, 0x4d], 0);
  u16(2, 42);
  u32(4, ifdOffset);
  u16(ifdOffset, fields.length);
  for (let index = 0; index < fields.length; index++) {
    const field = fields[index];
    const entry = ifdOffset + 2 + index * 12;
    u16(entry, field.tag);
    u16(entry + 2, field.type);
    u32(entry + 4, field.values.length);
    if (field.payloadOffset === undefined) writeValues(entry + 8, field);
    else {
      u32(entry + 8, field.payloadOffset);
      writeValues(field.payloadOffset, field);
    }
  }
  u32(ifdOffset + 2 + fields.length * 12, 0);
  for (let index = 0; index < strips.length; index++) {
    bytes.set(strips[index], stripOffsets.values[index]);
  }
  return bytes;
}

function cmykTiff(byteOrder: ByteOrder): Uint8Array {
  return classicTiff({
    byteOrder,
    width: 2,
    height: 2,
    bitsPerSample: [8, 8, 8, 8],
    photometric: 5,
    samplesPerPixel: 4,
    rowsPerStrip: 1,
    inkSet: 1,
    strips: [
      new Uint8Array([0, 0, 0, 0, 255, 0, 0, 0]),
      new Uint8Array([0, 255, 0, 128, 0, 0, 255, 255]),
    ],
  });
}

function bits(value: string): Uint8Array {
  const compact = value.replace(/\s+/g, '');
  const output = new Uint8Array(Math.ceil(compact.length / 8));
  for (let index = 0; index < compact.length; index++) {
    if (compact[index] === '1') output[index >> 3] |= 1 << (7 - (index & 7));
  }
  return output;
}

/**
 * Independent TIFF 6.0 §13 LZW encoder for synthetic fixtures: a leading
 * ClearCode, MSB-first 9..12-bit codes whose width grows one code early, a
 * ClearCode when the table reaches 4094 entries, and a closing
 * EndOfInformation code.
 */
function tiffLzw(data: Uint8Array): Uint8Array {
  const output: number[] = [];
  let buffer = 0;
  let bufferBits = 0;
  let width = 9;
  let nextCode = 258;
  let table = new Map<number, number>();
  const write = (code: number) => {
    buffer = (buffer << width) | code;
    bufferBits += width;
    while (bufferBits >= 8) {
      bufferBits -= 8;
      output.push((buffer >>> bufferBits) & 0xff);
    }
    buffer &= (1 << bufferBits) - 1;
  };
  const advance = () => {
    nextCode++;
    if (nextCode === 4094) {
      write(256);
      table = new Map();
      nextCode = 258;
      width = 9;
    } else if (nextCode > (1 << width) - 1) {
      width++;
    }
  };

  write(256);
  let prefix = -1;
  for (const byte of data) {
    if (prefix < 0) {
      prefix = byte;
      continue;
    }
    const code = table.get(prefix * 256 + byte);
    if (code !== undefined) {
      prefix = code;
      continue;
    }
    write(prefix);
    table.set(prefix * 256 + byte, nextCode);
    advance();
    prefix = byte;
  }
  if (prefix >= 0) {
    write(prefix);
    advance();
  }
  write(257);
  if (bufferBits > 0) output.push((buffer << (8 - bufferBits)) & 0xff);
  return new Uint8Array(output);
}

/**
 * Independent TIFF 6.0 §9 PackBits encoder for synthetic fixtures. Each row of
 * `bytesPerRow` bytes is packed separately: replicate runs of 2..128 equal
 * bytes (header 1 - n) and literal runs of 1..128 bytes (header n - 1).
 */
function tiffPackBits(data: Uint8Array, bytesPerRow: number): Uint8Array {
  const output: number[] = [];
  for (let rowStart = 0; rowStart < data.length; rowStart += bytesPerRow) {
    const rowEnd = Math.min(data.length, rowStart + bytesPerRow);
    let index = rowStart;
    while (index < rowEnd) {
      let run = 1;
      while (index + run < rowEnd && run < 128 && data[index + run] === data[index]) run++;
      if (run >= 2) {
        output.push((257 - run) & 0xff, data[index]);
        index += run;
        continue;
      }
      let literalEnd = index + 1;
      while (
        literalEnd < rowEnd
        && literalEnd - index < 128
        && !(literalEnd + 1 < rowEnd && data[literalEnd] === data[literalEnd + 1])
      ) literalEnd++;
      output.push(literalEnd - index - 1, ...data.subarray(index, literalEnd));
      index = literalEnd;
    }
  }
  return new Uint8Array(output);
}

/** TIFF 6.0 §14 encoder side: replace each sample by its difference from the
 * same component of the previous pixel in the row, modulo 2^8. */
function horizontallyDifferenced(
  samples: Uint8Array,
  bytesPerRow: number,
  samplesPerPixel: number,
): Uint8Array {
  const output = samples.slice();
  for (let row = 0; row < samples.length; row += bytesPerRow) {
    for (let index = samplesPerPixel; index < bytesPerRow; index++) {
      output[row + index] = (samples[row + index] - samples[row + index - samplesPerPixel]) & 0xff;
    }
  }
  return output;
}

/** Deterministic pseudo-random bytes, so LZW sees few repeated strings. */
function noise(length: number, seed: number): Uint8Array {
  const output = new Uint8Array(length);
  let state = seed >>> 0;
  for (let index = 0; index < length; index++) {
    state = (Math.imul(state, 1_664_525) + 1_013_904_223) >>> 0;
    output[index] = state >>> 24;
  }
  return output;
}

function rgbaRows(decoded: ReturnType<typeof decodeTiffRgba>): number[][] {
  if (!decoded) return [];
  const rows: number[][] = [];
  for (let y = 0; y < decoded.height; y++) {
    const values: number[] = [];
    for (let x = 0; x < decoded.width; x++) values.push(decoded.data[(y * decoded.width + x) * 4]);
    rows.push(values);
  }
  return rows;
}

describe('TIFF 6.0 decoder', () => {
  afterEach(() => vi.unstubAllGlobals());

  it('distinguishes non-TIFF bytes from unsupported TIFF container versions', () => {
    expect(decodeTiffRgba(new Uint8Array([1, 2, 3, 4]))).toBeNull();
    expect(() => decodeTiffRgba(new Uint8Array([0x49, 0x49])))
      .toThrowError(/Malformed TIFF header.*version/i);
    expect(() => decodeTiffRgba(new Uint8Array([0x49, 0x49, 0x2b, 0x00])))
      .toThrowError(/Unsupported TIFF version: 43/i);
  });

  it.each<ByteOrder>(['little', 'big'])('decodes multi-strip chunky CMYK (%s endian)', (byteOrder) => {
    const bytes = cmykTiff(byteOrder);
    expect(isTiff(bytes)).toBe(true);
    const decoded = decodeTiffRgba(bytes);
    expect(decoded && { width: decoded.width, height: decoded.height }).toEqual({ width: 2, height: 2 });
    expect(Array.from(decoded?.data ?? [])).toEqual([
      255, 255, 255, 255,
      0, 255, 255, 255,
      127, 0, 127, 255,
      0, 0, 0, 255,
    ]);
    expect(Array.from(decodeTiffRgba(bytes, {
      targetWidthPx: 1,
      targetHeightPx: 1,
    })?.data ?? [])).toEqual([96, 128, 159, 255]);
  });

  it('decodes uncompressed 1-bit and 8-bit WhiteIsZero/BlackIsZero grayscale', () => {
    const bilevel = classicTiff({
      width: 8,
      height: 1,
      bitsPerSample: [1],
      photometric: 0,
      strips: [new Uint8Array([0b01011000])],
    });
    expect(rgbaRows(decodeTiffRgba(bilevel))).toEqual([[255, 0, 255, 0, 0, 255, 255, 255]]);

    const blackIsZero = classicTiff({
      width: 3,
      height: 1,
      bitsPerSample: [8],
      photometric: 1,
      strips: [new Uint8Array([0, 64, 255])],
    });
    expect(rgbaRows(decodeTiffRgba(blackIsZero))).toEqual([[0, 64, 255]]);

    const whiteIsZero = classicTiff({
      width: 3,
      height: 1,
      bitsPerSample: [8],
      photometric: 0,
      strips: [new Uint8Array([0, 64, 255])],
    });
    expect(rgbaRows(decodeTiffRgba(whiteIsZero))).toEqual([[255, 191, 0]]);

    const checkerboard = classicTiff({
      width: 2,
      height: 2,
      bitsPerSample: [1],
      photometric: 0,
      strips: [new Uint8Array([0b01000000, 0b10000000])],
    });
    expect(Array.from(decodeTiffRgba(checkerboard, {
      targetWidthPx: 1,
      targetHeightPx: 1,
    })?.data ?? [])).toEqual([128, 128, 128, 255]);
  });

  it('decodes RGB plus associated and unassociated alpha with straight Canvas RGBA output', () => {
    const rgb = classicTiff({
      width: 1,
      height: 1,
      bitsPerSample: [8, 8, 8],
      photometric: 2,
      samplesPerPixel: 3,
      strips: [new Uint8Array([10, 20, 30])],
    });
    expect(Array.from(decodeTiffRgba(rgb)?.data ?? [])).toEqual([10, 20, 30, 255]);

    const associated = classicTiff({
      width: 1,
      height: 1,
      bitsPerSample: [8, 8, 8, 8],
      photometric: 2,
      samplesPerPixel: 4,
      extraSamples: [1],
      strips: [new Uint8Array([64, 32, 16, 128])],
    });
    expect(Array.from(decodeTiffRgba(associated)?.data ?? [])).toEqual([128, 64, 32, 128]);

    const unassociated = classicTiff({
      width: 1,
      height: 1,
      bitsPerSample: [8, 8, 8, 8],
      photometric: 2,
      samplesPerPixel: 4,
      extraSamples: [2],
      strips: [new Uint8Array([64, 32, 16, 128])],
    });
    expect(Array.from(decodeTiffRgba(unassociated)?.data ?? [])).toEqual([64, 32, 16, 128]);

    const transparentAssociated = classicTiff({
      width: 1,
      height: 1,
      bitsPerSample: [8, 8, 8, 8],
      photometric: 2,
      samplesPerPixel: 4,
      extraSamples: [1],
      strips: [new Uint8Array([123, 45, 67, 0])],
    });
    expect(Array.from(decodeTiffRgba(transparentAssociated)?.data ?? [])).toEqual([0, 0, 0, 0]);

    const blendedRgb = classicTiff({
      width: 2,
      height: 2,
      bitsPerSample: [8, 8, 8],
      photometric: 2,
      samplesPerPixel: 3,
      strips: [new Uint8Array([
        255, 0, 0, 0, 0, 255,
        255, 0, 0, 0, 0, 255,
      ])],
    });
    expect(Array.from(decodeTiffRgba(blendedRgb, {
      targetWidthPx: 1,
      targetHeightPx: 1,
    })?.data ?? [])).toEqual([128, 0, 128, 255]);

    const alphaPixels = [
      255, 0, 0, 255, 0, 0, 255, 0,
      255, 0, 0, 255, 0, 0, 255, 0,
    ];
    const blendedUnassociated = classicTiff({
      width: 2,
      height: 2,
      bitsPerSample: [8, 8, 8, 8],
      photometric: 2,
      samplesPerPixel: 4,
      extraSamples: [2],
      strips: [new Uint8Array(alphaPixels)],
    });
    expect(Array.from(decodeTiffRgba(blendedUnassociated, {
      targetWidthPx: 1,
      targetHeightPx: 1,
    })?.data ?? [])).toEqual([255, 0, 0, 128]);

    const blendedAssociated = classicTiff({
      width: 2,
      height: 2,
      bitsPerSample: [8, 8, 8, 8],
      photometric: 2,
      samplesPerPixel: 4,
      extraSamples: [1],
      strips: [new Uint8Array([
        255, 0, 0, 255, 123, 45, 67, 0,
        255, 0, 0, 255, 123, 45, 67, 0,
      ])],
    });
    expect(Array.from(decodeTiffRgba(blendedAssociated, {
      targetWidthPx: 1,
      targetHeightPx: 1,
    })?.data ?? [])).toEqual([255, 0, 0, 128]);
  });

  it.each<ByteOrder>(['little', 'big'])(
    'undoes horizontal differencing per component in LZW strips (%s endian)',
    (byteOrder) => {
      // Two strips of a 3×3 unassociated-alpha image, each LZW-compressed
      // after TIFF 6.0 §14 horizontal differencing with a 4-sample stride.
      const rgba = new Uint8Array([
        10, 200, 30, 255, 250, 5, 40, 128, 7, 9, 11, 0,
        1, 2, 3, 4, 255, 255, 255, 255, 0, 0, 0, 255,
        90, 80, 70, 60, 50, 40, 30, 20, 10, 0, 250, 240,
      ]);
      const differenced = horizontallyDifferenced(rgba, 12, 4);
      const tiff = classicTiff({
        byteOrder,
        width: 3,
        height: 3,
        rowsPerStrip: 2,
        bitsPerSample: [8, 8, 8, 8],
        compression: 5,
        predictor: 2,
        photometric: 2,
        samplesPerPixel: 4,
        extraSamples: [2],
        strips: [tiffLzw(differenced.subarray(0, 24)), tiffLzw(differenced.subarray(24))],
      });
      expect(Array.from(decodeTiffRgba(tiff)?.data ?? [])).toEqual(Array.from(rgba));

      const gray = new Uint8Array([0, 255, 1, 6, 128, 3]);
      const grayTiff = classicTiff({
        byteOrder,
        width: 3,
        height: 2,
        bitsPerSample: [8],
        compression: 5,
        predictor: 2,
        photometric: 1,
        strips: [tiffLzw(horizontallyDifferenced(gray, 3, 1))],
      });
      expect(rgbaRows(decodeTiffRgba(grayTiff))).toEqual([[0, 255, 1], [6, 128, 3]]);
      expect(rgbaRows(decodeTiffRgba(grayTiff, { targetHeightPx: 1 }))).toEqual([[66, 65]]);
    },
  );

  it('decodes LZW code-width growth, table resets and self-referencing codes', () => {
    // 48×48 RGB noise needs well over 4094 table entries, so the stream
    // crosses all three code-width changes and at least one ClearCode reset;
    // the constant tail produces KwKwK codes (a code equal to the next entry).
    const width = 48;
    const height = 48;
    const rgb = noise(width * height * 3, 0x5eed);
    rgb.fill(77, rgb.length - width * 3 * 4);
    const tiff = classicTiff({
      width,
      height,
      bitsPerSample: [8, 8, 8],
      compression: 5,
      photometric: 2,
      samplesPerPixel: 3,
      strips: [tiffLzw(rgb)],
    });
    const decoded = decodeTiffRgba(tiff);
    const expected = new Uint8Array(width * height * 4);
    for (let pixel = 0; pixel < width * height; pixel++) {
      expected.set(rgb.subarray(pixel * 3, pixel * 3 + 3), pixel * 4);
      expected[pixel * 4 + 3] = 255;
    }
    expect(Array.from(decoded?.data ?? [])).toEqual(Array.from(expected));
  });

  it('decodes wide LZW rows in bounded spans instead of allocating a source row', () => {
    // A compact LZW strip can describe a source row far wider than any
    // retained surface. Downsampling it must not buffer the whole row.
    const width = 65_536;
    const rgb = new Uint8Array(width * 3);
    for (let x = 0; x < width; x++) rgb.set([x & 0xff, 255 - (x & 0xff), 7], x * 3);
    const strip = tiffLzw(horizontallyDifferenced(rgb, width * 3, 3));
    const tiff = classicTiff({
      width,
      height: 2,
      rowsPerStrip: 1,
      bitsPerSample: [8, 8, 8],
      compression: 5,
      predictor: 2,
      photometric: 2,
      samplesPerPixel: 3,
      strips: [strip, strip],
    });
    expect(tiff.byteLength).toBeLessThan(width);

    const Native = Uint8Array;
    const allocations: number[] = [];
    vi.stubGlobal('Uint8Array', function allocate(length: number) {
      allocations.push(length);
      return new Native(length);
    });
    const decoded = decodeTiffRgba(tiff, { targetWidthPx: 2 });
    vi.unstubAllGlobals();

    expect(decoded && { width: decoded.width, height: decoded.height }).toEqual({ width: 2, height: 1 });
    // Each output pixel averages 128 complete 0..255 ramps.
    expect(Array.from(decoded?.data ?? [])).toEqual([128, 128, 7, 255, 128, 128, 7, 255]);
    expect(Math.max(...allocations)).toBeLessThanOrEqual(16 * 1024);
  });

  it('decodes multi-span LZW rows exactly like the equivalent uncompressed strips', () => {
    // 9000 px rows span three 4096-pixel LZW spans. A 3000×3 target puts an
    // output column across each span boundary and makes source rows straddle
    // output rows, so span carries and partial column sums are both exercised.
    const cases = [
      {
        width: 9000,
        height: 7,
        bitsPerSample: [8, 8, 8],
        photometric: 2,
        samplesPerPixel: 3,
        bytesPerRow: 9000 * 3,
        predictor: 2,
        targets: [{}, { targetWidthPx: 3000 }],
      },
      {
        width: 4096 + 13,
        height: 3,
        bitsPerSample: [1],
        photometric: 0,
        samplesPerPixel: 1,
        bytesPerRow: Math.ceil((4096 + 13) / 8),
        predictor: 1,
        targets: [{}, { targetWidthPx: 1000 }],
      },
    ];
    for (const { bytesPerRow, predictor, targets, ...layout } of cases) {
      const samples = noise(bytesPerRow * layout.height, layout.width);
      const rowsPerStrip = 2;
      const strips = (encode: (bytes: Uint8Array) => Uint8Array) => {
        const result: Uint8Array[] = [];
        for (let row = 0; row < layout.height; row += rowsPerStrip) {
          result.push(encode(samples.subarray(
            row * bytesPerRow,
            Math.min(layout.height, row + rowsPerStrip) * bytesPerRow,
          )));
        }
        return result;
      };
      const uncompressed = classicTiff({ ...layout, rowsPerStrip, strips: strips((bytes) => bytes) });
      const lzw = classicTiff({
        ...layout,
        rowsPerStrip,
        compression: 5,
        predictor,
        strips: strips((bytes) => tiffLzw(predictor === 2
          ? horizontallyDifferenced(bytes, bytesPerRow, layout.samplesPerPixel)
          : bytes)),
      });
      for (const target of targets) {
        const expected = decodeTiffRgba(uncompressed, target);
        const decoded = decodeTiffRgba(lzw, target);
        expect(decoded && { width: decoded.width, height: decoded.height })
          .toEqual(expected && { width: expected.width, height: expected.height });
        expect(Buffer.from(decoded?.data ?? []).equals(Buffer.from(expected?.data ?? []))).toBe(true);
      }
    }
  });

  it.each(['before pixels', 'after pixels'])('bounds repeated LZW ClearCodes %s across strips', (position) => {
    const clear = '100000000';
    const pixel = '001001101'; // 77
    const end = '100000001';
    const strip = bits(position === 'before pixels'
      ? clear.repeat(3) + pixel + end
      : clear + pixel + clear.repeat(2) + end);
    const bytes = classicTiff({
      width: 1,
      height: 2,
      rowsPerStrip: 1,
      bitsPerSample: [8],
      compression: 5,
      photometric: 1,
      strips: [strip, strip],
    });
    // Each strip fits the image allowance; both must charge reads (including
    // EOI) and expanded bytes without resetting the image's work allowance.
    let failure: unknown;
    try { decodeTiffRgba(bytes); } catch (error) { failure = error; }
    expect(failure).toBeInstanceOf(OoxmlDecodedImageLimitError);
    expect(getOoxmlDecodedImageLimitDetails(failure)).toMatchObject({ metric: 'tiff-lzw-work' });
  });

  it.each([
    ['missing EOI', [256, 77], /truncated/],
    ['excess literal', [256, 77, 78, 257], /exceeds its declared rows/],
    ['excess string', [256, 77, 258, 257], /exceeds its declared rows/],
    ['excess pending string', [256, 77, 258, 257], /exceeds its declared rows/, 2],
    ['invalid code after pixels', [256, 77, 300, 257], /Invalid TIFF LZW code: 300/],
  ] as const)('rejects LZW strip termination with %s', (_name, codes, diagnostic, width: number = 1) => {
    const bytes = classicTiff({
      width,
      height: 1,
      bitsPerSample: [8],
      compression: 5,
      photometric: 1,
      strips: [bits(codes.map((code) => code.toString(2).padStart(9, '0')).join(''))],
    });
    expect(() => decodeTiffRgba(bytes)).toThrowError(TiffDecodeError);
    expect(() => decodeTiffRgba(bytes)).toThrowError(diagnostic);
  });

  it('accepts a final LZW ClearCode and padding after EOI', () => {
    const bytes = classicTiff({
      width: 1,
      height: 1,
      bitsPerSample: [8],
      compression: 5,
      photometric: 1,
      strips: [bits('100000000' + '001001101' + '100000000' + '100000001')],
    });
    expect(rgbaRows(decodeTiffRgba(bytes))).toEqual([[77]]);
  });

  it('validates non-final LZW strips when downsampling', () => {
    const bytes = classicTiff({
      width: 1,
      height: 2,
      rowsPerStrip: 1,
      bitsPerSample: [8],
      compression: 5,
      photometric: 1,
      strips: [bits('100000000' + '001001101'), tiffLzw(new Uint8Array([88]))],
    });
    expect(() => decodeTiffRgba(bytes, { targetWidthPx: 1, targetHeightPx: 1 }))
      .toThrowError(/TIFF LZW data is truncated/);
  });

  it('rejects predictors and LZW streams outside the implemented TIFF 6.0 rules', () => {
    const gray = (overrides: Partial<TiffFixtureOptions>) => classicTiff({
      width: 2,
      height: 1,
      bitsPerSample: [8],
      compression: 5,
      photometric: 1,
      strips: [tiffLzw(new Uint8Array([1, 2]))],
      ...overrides,
    });
    expect(() => decodeTiffRgba(gray({ predictor: 3 })))
      .toThrowError(/Unsupported TIFF Predictor: 3/);
    expect(() => decodeTiffRgba(gray({
      compression: 1,
      predictor: 2,
      strips: [new Uint8Array([1, 2])],
    }))).toThrowError(/Unsupported TIFF Predictor 2 with Compression 1/);
    expect(() => decodeTiffRgba(gray({
      bitsPerSample: [1],
      predictor: 2,
      strips: [tiffLzw(new Uint8Array([0x40]))],
    }))).toThrowError(/Unsupported TIFF Predictor 2 with 1-bit samples/);

    const encoded = tiffLzw(new Uint8Array([1, 2]));
    expect(() => decodeTiffRgba(gray({ strips: [encoded.subarray(0, 2)] })))
      .toThrowError(/TIFF LZW data is truncated/);
    expect(() => decodeTiffRgba(gray({ strips: [tiffLzw(new Uint8Array([1]))] })))
      .toThrowError(/TIFF LZW strip ended before its declared rows/);
    // A pre-6.0 (LSB-first) stream does not begin with an MSB-first ClearCode.
    expect(() => decodeTiffRgba(gray({ strips: [new Uint8Array([0x00, 0x01, 0x00])] })))
      .toThrowError(/TIFF LZW strip must begin with ClearCode/);
    // ClearCode, literal 1, then code 300, which is not yet in the table.
    expect(() => decodeTiffRgba(gray({
      strips: [bits('100000000' + '000000001' + '100101100')],
    }))).toThrowError(/Invalid TIFF LZW code: 300/);
  });

  it('decodes PackBits literal, replicate and no-op runs (TIFF 6.0 §9)', () => {
    const bytes = classicTiff({
      width: 9,
      height: 1,
      bitsPerSample: [8],
      compression: 32773,
      photometric: 1,
      // literal 2, no-op (-128), replicate 3 (-2), literal 1, replicate 2 (-1), literal 1.
      strips: [new Uint8Array([0x01, 10, 20, 0x80, 0xfe, 30, 0x00, 40, 0xff, 50, 0x00, 60])],
    });
    expect(rgbaRows(decodeTiffRgba(bytes))).toEqual([[10, 20, 30, 30, 30, 40, 50, 50, 60]]);
  });

  it('decodes PackBits for every supported chunky class exactly like uncompressed strips', () => {
    // Wide rows cross the 4096-pixel span boundary; runs of equal bytes and
    // noise both occur, and the second strip is shorter than RowsPerStrip.
    const cases = [
      { bitsPerSample: [1], photometric: 0, width: 4096 + 21 },
      { bitsPerSample: [1], photometric: 1, width: 37 },
      { bitsPerSample: [8], photometric: 0, width: 4096 + 5 },
      { bitsPerSample: [8], photometric: 1, width: 300 },
      { bitsPerSample: [8, 8, 8], photometric: 2, width: 4096 + 3 },
      { bitsPerSample: [8, 8, 8, 8], photometric: 2, extraSamples: [1], width: 130 },
      { bitsPerSample: [8, 8, 8, 8], photometric: 2, extraSamples: [2], width: 130 },
      { bitsPerSample: [8, 8, 8, 8], photometric: 5, inkSet: 1, width: 129 },
    ];
    for (const { width, ...layout } of cases) {
      const samplesPerPixel = layout.bitsPerSample.length;
      const bytesPerRow = layout.bitsPerSample[0] === 1
        ? Math.ceil(width / 8)
        : width * samplesPerPixel;
      const height = 5;
      const rowsPerStrip = 3;
      const samples = noise(bytesPerRow * height, width);
      // Long replicate runs (more than 128 bytes) inside noise.
      samples.fill(0xa5, 7, 7 + Math.min(300, bytesPerRow - 7));
      const strips = (encode: (bytes: Uint8Array) => Uint8Array) => {
        const result: Uint8Array[] = [];
        for (let row = 0; row < height; row += rowsPerStrip) {
          result.push(encode(samples.subarray(
            row * bytesPerRow,
            Math.min(height, row + rowsPerStrip) * bytesPerRow,
          )));
        }
        return result;
      };
      const common = { ...layout, width, height, rowsPerStrip, samplesPerPixel };
      const uncompressed = classicTiff({ ...common, strips: strips((bytes) => bytes) });
      const packBits = classicTiff({
        ...common,
        compression: 32773,
        strips: strips((bytes) => tiffPackBits(bytes, bytesPerRow)),
      });
      for (const target of [{}, { targetWidthPx: Math.max(1, Math.floor(width / 3)) }]) {
        const expected = decodeTiffRgba(uncompressed, target);
        const decoded = decodeTiffRgba(packBits, target);
        expect(decoded && { width: decoded.width, height: decoded.height })
          .toEqual(expected && { width: expected.width, height: expected.height });
        expect(Buffer.from(decoded?.data ?? []).equals(Buffer.from(expected?.data ?? []))).toBe(true);
      }
    }
  });

  it('decodes a PackBits run that continues into the next row of the same strip', () => {
    const bytes = classicTiff({
      width: 2,
      height: 2,
      bitsPerSample: [8],
      compression: 32773,
      photometric: 1,
      strips: [new Uint8Array([0xfe, 9, 0x00, 4])],
    });
    expect(rgbaRows(decodeTiffRgba(bytes))).toEqual([[9, 9], [9, 4]]);
  });

  it.each([
    ['a missing header', [0x00, 1], /TIFF PackBits data is truncated/],
    ['a short literal', [0x00, 1, 0x02, 2, 3], /TIFF PackBits data is truncated/],
    ['a missing replicate byte', [0x00, 1, 0xff], /TIFF PackBits data is truncated/],
    ['only no-ops', [0x80, 0x80], /TIFF PackBits data is truncated/],
    ['an overrunning literal', [0x02, 1, 2, 3], /TIFF PackBits run exceeds its declared rows/],
    ['an overrunning replicate', [0xfd, 7], /TIFF PackBits run exceeds its declared rows/],
  ] as const)('fails closed on PackBits strips with %s', (_name, strip, diagnostic) => {
    const bytes = classicTiff({
      width: 2,
      height: 1,
      bitsPerSample: [8],
      compression: 32773,
      photometric: 1,
      strips: [new Uint8Array(strip)],
    });
    expect(() => decodeTiffRgba(bytes)).toThrowError(TiffDecodeError);
    expect(() => decodeTiffRgba(bytes)).toThrowError(diagnostic);
  });

  it('fails closed when a PackBits literal would read past its strip into the next one', () => {
    // The first strip's literal claims 3 bytes but its byte count holds 2;
    // the next strip's bytes must not be consumed to satisfy it.
    const bytes = classicTiff({
      width: 3,
      height: 2,
      rowsPerStrip: 1,
      bitsPerSample: [8],
      compression: 32773,
      photometric: 1,
      strips: [new Uint8Array([0x02, 1, 2]), new Uint8Array([0x02, 4, 5, 6])],
    });
    expect(() => decodeTiffRgba(bytes)).toThrowError(/TIFF PackBits data is truncated/);
  });

  it('accepts unread padding after the final PackBits run of a strip', () => {
    const bytes = classicTiff({
      width: 2,
      height: 1,
      bitsPerSample: [8],
      compression: 32773,
      photometric: 1,
      strips: [new Uint8Array([0xff, 6, 0x00])],
    });
    expect(rgbaRows(decodeTiffRgba(bytes))).toEqual([[6, 6]]);
  });

  it('bounds PackBits no-op headers across strips that share one byte range', () => {
    const strip = new Uint8Array([0x80, 0x80, 0x00, 77]);
    const bytes = classicTiff({
      width: 1,
      height: 2,
      rowsPerStrip: 1,
      bitsPerSample: [8],
      compression: 32773,
      photometric: 1,
      strips: [strip, strip],
    });
    // Each strip alone fits the image allowance (2 headers per decoded byte);
    // the allowance is image-wide, so the second strip exceeds it.
    let failure: unknown;
    try { decodeTiffRgba(bytes); } catch (error) { failure = error; }
    expect(failure).toBeInstanceOf(OoxmlDecodedImageLimitError);
    expect(getOoxmlDecodedImageLimitDetails(failure)).toMatchObject({ metric: 'tiff-packbits-work' });
  });

  it('refuses the differencing predictor with PackBits', () => {
    const bytes = classicTiff({
      width: 2,
      height: 1,
      bitsPerSample: [8],
      compression: 32773,
      photometric: 1,
      predictor: 2,
      strips: [new Uint8Array([0x01, 1, 1])],
    });
    expect(() => decodeTiffRgba(bytes)).toThrowError(/Unsupported TIFF Predictor 2 with Compression 32773/);
  });

  it('decodes CCITT Group 4 horizontal, vertical and pass modes', () => {
    // Row 0: H(W2,B4), V0(W2) => WWBBBBWW
    // Row 1: V0,V0,V0             => WWBBBBWW
    // Row 2: VR1,VR1,V0           => WWWBBBBW
    // Row 3: PASS,V0              => WWWWWWWW
    const encoded = bits([
      '001', '0111', '011', '1',
      '111',
      '0110111',
      '00011',
    ].join(''));
    const tiff = classicTiff({
      width: 8,
      height: 4,
      bitsPerSample: [1],
      compression: 4,
      photometric: 0,
      strips: [encoded],
    });

    expect(rgbaRows(decodeTiffRgba(tiff))).toEqual([
      [255, 255, 0, 0, 0, 0, 255, 255],
      [255, 255, 0, 0, 0, 0, 255, 255],
      [255, 255, 255, 0, 0, 0, 0, 255],
      [255, 255, 255, 255, 255, 255, 255, 255],
    ]);
    expect(rgbaRows(decodeTiffRgba(tiff, {
      targetWidthPx: 4,
      targetHeightPx: 2,
    }))).toEqual([
      [255, 0, 0, 255],
      [255, 191, 128, 191],
    ]);

    const blackIsZero = classicTiff({
      width: 8,
      height: 4,
      bitsPerSample: [1],
      compression: 4,
      photometric: 1,
      strips: [encoded],
    });
    expect(rgbaRows(decodeTiffRgba(blackIsZero))).toEqual([
      [0, 0, 255, 255, 255, 255, 0, 0],
      [0, 0, 255, 255, 255, 255, 0, 0],
      [0, 0, 0, 255, 255, 255, 255, 0],
      [0, 0, 0, 0, 0, 0, 0, 0],
    ]);
  });

  it('resets the Group 4 reference line for each strip', () => {
    const whiteThenBlack = [bits('1'), bits('00100110101000101')];
    const tiff = classicTiff({
      width: 8,
      height: 2,
      rowsPerStrip: 1,
      bitsPerSample: [1],
      compression: 4,
      photometric: 0,
      strips: whiteThenBlack,
    });
    expect(rgbaRows(decodeTiffRgba(tiff))).toEqual([
      [255, 255, 255, 255, 255, 255, 255, 255],
      [0, 0, 0, 0, 0, 0, 0, 0],
    ]);
  });

  it('decodes every Group 4 vertical offset', () => {
    const encoded = bits([
      // WWWBBBBWWW: H(W3,B4), V0.
      '001', '1000', '011', '1',
      // WWBBBBBBWW: VL1, VR1, V0.
      '010', '011', '1',
      // WWWWBBWWWW: VR2, VL2, V0.
      '000011', '000010', '1',
      // WBBBBBBBBW: VL3, VR3, V0.
      '0000010', '0000011', '1',
    ].join(''));
    const tiff = classicTiff({
      width: 10,
      height: 4,
      bitsPerSample: [1],
      compression: 4,
      photometric: 0,
      strips: [encoded],
    });
    expect(rgbaRows(decodeTiffRgba(tiff))).toEqual([
      [255, 255, 255, 0, 0, 0, 0, 255, 255, 255],
      [255, 255, 0, 0, 0, 0, 0, 0, 255, 255],
      [255, 255, 255, 255, 0, 0, 255, 255, 255, 255],
      [255, 0, 0, 0, 0, 0, 0, 0, 0, 255],
    ]);
  });

  it('decodes black-color pass mode and a zero-length leading white run', () => {
    const encoded = bits([
      // WWBBWWBBWW: H(W2,B2), H(W2,B2), V0.
      '001', '0111', '11', '001', '0111', '11', '1',
      // WWBBBBBBWW: V0, black PASS, V0, V0.
      '1', '0001', '1', '1',
      // BBBBBBBBBB: H(W0,B10).
      '001', '00110101', '0000100',
    ].join(''));
    const tiff = classicTiff({
      width: 10,
      height: 3,
      bitsPerSample: [1],
      compression: 4,
      photometric: 0,
      strips: [encoded],
    });
    expect(rgbaRows(decodeTiffRgba(tiff))).toEqual([
      [255, 255, 0, 0, 255, 255, 0, 0, 255, 255],
      [255, 255, 0, 0, 0, 0, 0, 0, 255, 255],
      [0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    ]);
  });

  it('bounds maximally compressed Group 4 downsample work by rows and retained width', () => {
    const width = 32_767;
    const height = 4_096;
    // Against the imaginary/all-white reference line, vertical-0 is one bit
    // per row. This source is just below the 128 MP source cap while its TIFF
    // payload remains tiny enough to expose source-grid decode amplification.
    const tiff = classicTiff({
      width,
      height,
      bitsPerSample: [1],
      compression: 4,
      photometric: 0,
      strips: [bits('1'.repeat(height))],
    });
    expect(tiff.byteLength).toBeLessThan(1_024);

    const diagnostics = __test_decodeTiffRgbaWithGroup4Diagnostics(tiff, {
      targetWidthPx: 1,
      targetHeightPx: 1,
    });
    expect(diagnostics.decoded && {
      width: diagnostics.decoded.width,
      height: diagnostics.decoded.height,
    }).toEqual({ width: 8, height: 1 });
    expect(rgbaRows(diagnostics.decoded)).toEqual([
      new Array<number>(8).fill(255),
    ]);

    // One mode and one reference probe per encoded row, plus one run overlap
    // per retained pixel column. None of these terms scale with source width.
    expect({
      modeCount: diagnostics.modeCount,
      referenceProbeCount: diagnostics.referenceProbeCount,
      areaSegmentCount: diagnostics.areaSegmentCount,
    }).toEqual({
      modeCount: height,
      referenceProbeCount: height,
      areaSegmentCount: height * 8,
    });
    expect(
      diagnostics.modeCount
      + diagnostics.referenceProbeCount
      + diagnostics.areaSegmentCount,
    ).toBeLessThanOrEqual(height * ((diagnostics.decoded?.width ?? 0) + 2));
  });

  it('bounds dense Group 4 changing-element storage for a downsampled source row', () => {
    const width = 32_768;
    // Begin with a zero-length white run, then alternate one black/white pixel.
    // The resulting row has width + 1 changing elements (including x=0 and the
    // width sentinel), one beyond the retained-axis-derived sparse-row ceiling.
    const firstBlackPixel = '001' + '00110101' + '010';
    const whiteBlackPair = '001' + '000111' + '010';
    const finalWhitePixel = '001' + '000111' + '0000110111';
    const encoded = bits(
      firstBlackPixel
      + whiteBlackPair.repeat((width - 2) / 2)
      + finalWhitePixel,
    );
    const tiff = classicTiff({
      width,
      height: 2,
      rowsPerStrip: 1,
      bitsPerSample: [1],
      compression: 4,
      photometric: 0,
      strips: [encoded, bits('1')],
    });

    expect(encoded.byteLength).toBeLessThan(32 * 1024);
    expect(() => decodeTiffRgba(tiff, { targetWidthPx: 1 }))
      .toThrowError(/CCITT Group 4 changing-element limit exceeded/i);
  });

  it('area-filters into one direct aspect-preserving target allocation', () => {
    const source = classicTiff({
      width: 4,
      height: 2,
      bitsPerSample: [8, 8, 8],
      photometric: 2,
      samplesPerPixel: 3,
      strips: [new Uint8Array([
        1, 1, 1, 2, 2, 2, 3, 3, 3, 4, 4, 4,
        5, 5, 5, 6, 6, 6, 7, 7, 7, 8, 8, 8,
      ])],
    });
    const Native = Uint8ClampedArray;
    const allocations: number[] = [];
    vi.stubGlobal('Uint8ClampedArray', function allocate(length: number) {
      allocations.push(length);
      return new Native(length);
    });

    const decoded = decodeTiffRgba(source, { targetWidthPx: 1, targetHeightPx: 1 });
    expect(decoded && { width: decoded.width, height: decoded.height }).toEqual({ width: 2, height: 1 });
    expect(allocations).toEqual([2 * 1 * 4]);
    expect(rgbaRows(decoded)).toEqual([[4, 6]]);
  });

  it('does not overshoot an exact integer target through floating-point roundoff', () => {
    const source = classicTiff({
      width: 25,
      height: 25,
      bitsPerSample: [8],
      photometric: 1,
      strips: [new Uint8Array(25 * 25)],
    });

    const decoded = decodeTiffRgba(source, {
      targetWidthPx: 7,
      targetHeightPx: 7,
      maxRetainedPixels: 49,
    });
    expect(decoded && { width: decoded.width, height: decoded.height }).toEqual({
      width: 7,
      height: 7,
    });
  });

  it('preserves the decoded-image error contract when the required target exceeds its budget', () => {
    const source = classicTiff({
      width: 4,
      height: 2,
      bitsPerSample: [8],
      photometric: 1,
      strips: [new Uint8Array(8)],
    });
    expect(() => decodeTiffRgba(source, {
      targetWidthPx: 1,
      targetHeightPx: 1,
      maxRetainedPixels: 1,
    })).toThrow(expect.objectContaining({
      code: 'ooxml-decoded-image-limit',
      metric: 'image-pixels',
      limit: 1,
      observed: 2,
    }));
  });

  it('keeps huge source-budget diagnostics safe for worker serialization', () => {
    const source = classicTiff({
      width: 1 << 27,
      height: 1 << 27,
      bitsPerSample: [8],
      photometric: 1,
      rowsPerStrip: 1 << 27,
      strips: [new Uint8Array([0])],
    });
    expect(() => decodeTiffRgba(source)).toThrow(expect.objectContaining({
      code: 'ooxml-decoded-image-limit',
      metric: 'image-pixels',
      limit: 1 << 27,
      observed: Number.MAX_SAFE_INTEGER,
    }));
  });

  it('throws diagnostics for unsupported recognized TIFF classes and malformed ranges', () => {
    const deflate = classicTiff({
      width: 1,
      height: 1,
      bitsPerSample: [8],
      compression: 8,
      photometric: 1,
      strips: [new Uint8Array([0])],
    });
    expect(() => decodeTiffRgba(deflate)).toThrowError(/unsupported TIFF compression: 8/i);

    const malformed = cmykTiff('little');
    const view = new DataView(malformed.buffer);
    const count = view.getUint16(8, true);
    for (let index = 0; index < count; index++) {
      const entry = 10 + index * 12;
      if (view.getUint16(entry, true) === 273) {
        const valuesOffset = view.getUint32(entry + 8, true);
        view.setUint32(valuesOffset, 0xfffffff0, true);
      }
    }
    const allocate = vi.fn(() => {
      throw new Error('decoded allocation must follow strip validation');
    });
    vi.stubGlobal('Uint8ClampedArray', allocate);
    expect(() => decodeTiffRgba(malformed)).toThrowError(/strip 0.*range/i);
    expect(allocate).not.toHaveBeenCalled();
  });

  it('rejects truncated Group 4 data deterministically', () => {
    const tiff = classicTiff({
      width: 8,
      height: 1,
      bitsPerSample: [1],
      compression: 4,
      photometric: 0,
      strips: [bits('001011')],
    });
    expect(() => decodeTiffRgba(tiff)).toThrowError(/CCITT Group 4.*truncated/i);
  });

  it('creates an ImageBitmap directly from the decoded ImageData backing store', async () => {
    const imageData = class {
      constructor(
        readonly data: Uint8ClampedArray,
        readonly width: number,
        readonly height: number,
      ) {}
    };
    vi.stubGlobal('ImageData', imageData);
    vi.stubGlobal('OffscreenCanvas', class {
      constructor() {
        throw new Error('TIFF bitmap handoff must not allocate a Canvas');
      }
    });
    const bitmap = { width: 2, height: 2, close() {} } as unknown as ImageBitmap;
    const create = vi.fn(async (source: unknown) => {
      expect(source).toBeInstanceOf(imageData);
      expect(Array.from((source as InstanceType<typeof imageData>).data)).toEqual([
        255, 255, 255, 255,
        0, 255, 255, 255,
        127, 0, 127, 255,
        0, 0, 0, 255,
      ]);
      return bitmap;
    });
    vi.stubGlobal('createImageBitmap', create);

    await expect(renderTiffToBitmap(cmykTiff('little'))).resolves.toBe(bitmap);
    expect(create).toHaveBeenCalledTimes(1);
  });

  it('reports unavailable ImageData and failed bitmap creation as codec failures', async () => {
    vi.stubGlobal('ImageData', undefined);
    await expect(renderTiffToBitmap(cmykTiff('little'))).rejects.toBeInstanceOf(TiffDecodeError);

    vi.stubGlobal('ImageData', class {
      constructor(
        readonly data: Uint8ClampedArray,
        readonly width: number,
        readonly height: number,
      ) {}
    });
    vi.stubGlobal('createImageBitmap', vi.fn(async () => null));
    await expect(renderTiffToBitmap(cmykTiff('little'))).rejects.toThrowError(/ImageBitmap/i);

    vi.stubGlobal('createImageBitmap', vi.fn(async () => {
      throw new Error('synthetic bitmap failure');
    }));
    await expect(renderTiffToBitmap(cmykTiff('little'))).rejects.toEqual(expect.objectContaining({
      name: 'TiffDecodeError',
      code: 'ooxml-tiff-decode',
      message: expect.stringMatching(/ImageBitmap creation failed/i),
    }));
  });
});
