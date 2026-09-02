import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  __test_decodeTiffRgbaWithGroup4Diagnostics,
  decodeTiffRgba,
  renderTiffToBitmap,
  TiffDecodeError,
} from './tiff.js';
import { isTiff } from './tiff-contract.js';

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
    const lzw = classicTiff({
      width: 1,
      height: 1,
      bitsPerSample: [8],
      compression: 5,
      photometric: 1,
      strips: [new Uint8Array([0])],
    });
    expect(() => decodeTiffRgba(lzw)).toThrowError(/unsupported TIFF compression: 5/i);

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
