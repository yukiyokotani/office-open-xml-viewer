import { afterEach, describe, expect, it, vi } from 'vitest';
import { decodeTiffRgba, renderTiffToBitmap } from './tiff.js';
import { isTiff } from './tiff-contract.js';

type ByteOrder = 'little' | 'big';

interface TiffFixtureOptions {
  byteOrder?: ByteOrder;
  width?: number;
  height?: number;
  rowsPerStrip?: number;
  compression?: number;
  photometric?: number;
  pixels?: number[];
}

/**
 * Build a synthetic TIFF 6.0 separated-image fixture. It deliberately uses
 * multiple strips so the decoder cannot accidentally assume one contiguous
 * pixel block. No private document bytes are committed in this test.
 */
function cmykTiff({
  byteOrder = 'little',
  width = 2,
  height = 2,
  rowsPerStrip = 1,
  compression = 1,
  photometric = 5,
  pixels = [
    0, 0, 0, 0,
    255, 0, 0, 0,
    0, 255, 0, 128,
    0, 0, 255, 255,
  ],
}: TiffFixtureOptions = {}): Uint8Array {
  const little = byteOrder === 'little';
  const entries = 9;
  const ifdOffset = 8;
  const ifdBytes = 2 + entries * 12 + 4;
  const bitsOffset = ifdOffset + ifdBytes;
  const stripCount = Math.ceil(height / rowsPerStrip);
  const stripOffsetsOffset = bitsOffset + 8;
  const stripByteCountsOffset = stripOffsetsOffset + stripCount * 4;
  const pixelOffset = stripByteCountsOffset + stripCount * 4;
  const bytesPerRow = width * 4;
  const total = pixelOffset + pixels.length;
  const bytes = new Uint8Array(total);
  const view = new DataView(bytes.buffer);
  const u16 = (offset: number, value: number) => view.setUint16(offset, value, little);
  const u32 = (offset: number, value: number) => view.setUint32(offset, value, little);

  bytes.set(little ? [0x49, 0x49] : [0x4d, 0x4d], 0);
  u16(2, 42);
  u32(4, ifdOffset);
  u16(ifdOffset, entries);

  let entry = ifdOffset + 2;
  const scalar = (tag: number, type: 3 | 4, value: number) => {
    u16(entry, tag);
    u16(entry + 2, type);
    u32(entry + 4, 1);
    if (type === 3) u16(entry + 8, value);
    else u32(entry + 8, value);
    entry += 12;
  };
  const offsetArray = (tag: number, count: number, offset: number) => {
    u16(entry, tag);
    u16(entry + 2, 4);
    u32(entry + 4, count);
    u32(entry + 8, offset);
    entry += 12;
  };

  scalar(256, 4, width); // ImageWidth
  scalar(257, 4, height); // ImageLength
  u16(entry, 258); // BitsPerSample = [8,8,8,8]
  u16(entry + 2, 3);
  u32(entry + 4, 4);
  u32(entry + 8, bitsOffset);
  entry += 12;
  scalar(259, 3, compression);
  scalar(262, 3, photometric); // Separated
  offsetArray(273, stripCount, stripOffsetsOffset);
  scalar(277, 3, 4); // SamplesPerPixel
  scalar(278, 4, rowsPerStrip);
  offsetArray(279, stripCount, stripByteCountsOffset);
  u32(entry, 0); // next IFD

  for (let i = 0; i < 4; i++) u16(bitsOffset + i * 2, 8);
  let sourceOffset = 0;
  for (let strip = 0; strip < stripCount; strip++) {
    const rows = Math.min(rowsPerStrip, height - strip * rowsPerStrip);
    const byteCount = rows * bytesPerRow;
    u32(stripOffsetsOffset + strip * 4, pixelOffset + sourceOffset);
    u32(stripByteCountsOffset + strip * 4, byteCount);
    sourceOffset += byteCount;
  }
  bytes.set(pixels, pixelOffset);
  return bytes;
}

describe('TIFF 6.0 decoder', () => {
  afterEach(() => vi.unstubAllGlobals());

  it.each<ByteOrder>(['little', 'big'])('decodes uncompressed chunky CMYK strips (%s endian)', (byteOrder) => {
    const bytes = cmykTiff({ byteOrder });
    expect(isTiff(bytes)).toBe(true);
    const decoded = decodeTiffRgba(bytes);
    expect(decoded).not.toBeNull();
    expect(decoded && { width: decoded.width, height: decoded.height }).toEqual({
      width: 2,
      height: 2,
    });
    expect(Array.from(decoded?.data ?? [])).toEqual([
      255, 255, 255, 255,
      0, 255, 255, 255,
      127, 0, 127, 255,
      0, 0, 0, 255,
    ]);
  });

  it('rejects unsupported compression, colour classes, and malformed strip ranges without guessing', () => {
    expect(decodeTiffRgba(cmykTiff({ compression: 5 }))).toBeNull(); // LZW
    expect(decodeTiffRgba(cmykTiff({ photometric: 2 }))).toBeNull(); // RGB
    const malformed = cmykTiff();
    // First StripOffsets value is outside the file.
    new DataView(malformed.buffer).setUint32(130, 0xfffffff0, true);
    expect(decodeTiffRgba(malformed)).toBeNull();
  });

  it('validates every strip range before allocating the decoded raster', () => {
    const malformed = cmykTiff({
      width: 8192,
      height: 4096,
      rowsPerStrip: 4096,
      pixels: [],
    });
    const allocate = vi.fn(() => {
      throw new Error('decoded raster allocation must follow strip validation');
    });
    vi.stubGlobal('Uint8ClampedArray', allocate);

    expect(decodeTiffRgba(malformed)).toBeNull();
    expect(allocate).not.toHaveBeenCalled();
  });

  it('rasterizes decoded pixels through an auxiliary canvas, not browser TIFF decoding', async () => {
    const captured: number[] = [];
    const canvas = class {
      width: number;
      height: number;
      constructor(width: number, height: number) {
        this.width = width;
        this.height = height;
      }
      getContext() {
        return {
          createImageData: (width: number, height: number) => ({
            width,
            height,
            data: new Uint8ClampedArray(width * height * 4),
          }),
          putImageData: (image: ImageData) => captured.push(...image.data),
        };
      }
    };
    vi.stubGlobal('OffscreenCanvas', canvas);
    const bitmap = { width: 2, height: 2, close() {} } as unknown as ImageBitmap;
    const create = vi.fn(async (source: unknown) => {
      expect(source).toBeInstanceOf(canvas);
      return bitmap;
    });
    vi.stubGlobal('createImageBitmap', create);

    await expect(renderTiffToBitmap(cmykTiff())).resolves.toBe(bitmap);
    expect(captured).toEqual([
      255, 255, 255, 255,
      0, 255, 255, 255,
      127, 0, 127, 255,
      0, 0, 0, 255,
    ]);
    expect(create).toHaveBeenCalledTimes(1);
  });
});
