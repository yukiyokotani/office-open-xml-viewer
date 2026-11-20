import { describe, expect, it } from 'vitest';
import { sniffRasterResolution } from './raster-resolution.js';

const be32 = (value: number) => [(value >>> 24) & 255, (value >>> 16) & 255, (value >>> 8) & 255, value & 255];
const be16 = (value: number) => [(value >>> 8) & 255, value & 255];
const le16 = (value: number) => [value & 255, (value >>> 8) & 255];
const le32 = (value: number) => [value & 255, (value >>> 8) & 255, (value >>> 16) & 255, (value >>> 24) & 255];

function png(physUnit: number, ppmX: number, ppmY: number): Uint8Array {
  const chunk = (type: string, data: number[]) => [...be32(data.length), ...[...type].map((c) => c.charCodeAt(0)), ...data, 0, 0, 0, 0];
  return new Uint8Array([
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a,
    ...chunk('IHDR', [...be32(4), ...be32(2), 8, 6, 0, 0, 0]),
    ...chunk('pHYs', [...be32(ppmX), ...be32(ppmY), physUnit]),
    ...chunk('IDAT', []),
  ]);
}

function jfif(units: number, x: number, y: number): Uint8Array {
  const payload = [0x4a, 0x46, 0x49, 0x46, 0, 1, 2, units, ...be16(x), ...be16(y), 0, 0];
  return new Uint8Array([0xff, 0xd8, 0xff, 0xe0, ...be16(payload.length + 2), ...payload, 0xff, 0xda]);
}

/** Little-endian TIFF IFD0 with X/YResolution (RATIONAL) and ResolutionUnit. */
function tiff(x: number, y: number, unit: number): number[] {
  const entries = 3;
  const ifd = 8;
  const values = ifd + 2 + entries * 12 + 4;
  return [
    0x49, 0x49, ...le16(42), ...le32(ifd), ...le16(entries),
    ...le16(0x011a), ...le16(5), ...le32(1), ...le32(values),
    ...le16(0x011b), ...le16(5), ...le32(1), ...le32(values + 8),
    ...le16(0x0128), ...le16(3), ...le32(1), ...le16(unit), 0, 0,
    ...le32(0),
    ...le32(x), ...le32(1), ...le32(y), ...le32(1),
  ];
}

describe('raster header resolution', () => {
  it('reads PNG pHYs pixels per metre and ignores a unitless aspect ratio', () => {
    const dpi = sniffRasterResolution(png(1, 7559, 3780));
    expect(dpi?.dpiX).toBeCloseTo(192, 0);
    expect(dpi?.dpiY).toBeCloseTo(96, 0);
    expect(sniffRasterResolution(png(0, 1, 1))).toBeNull();
  });

  it('reads JPEG JFIF density in dots per inch or per centimetre', () => {
    expect(sniffRasterResolution(jfif(1, 300, 150))).toEqual({ dpiX: 300, dpiY: 150 });
    expect(sniffRasterResolution(jfif(2, 100, 100))?.dpiX).toBeCloseTo(254);
    expect(sniffRasterResolution(jfif(0, 1, 1))).toBeNull();
  });

  it('reads JPEG EXIF and TIFF IFD0 resolution', () => {
    const exifPayload = [0x45, 0x78, 0x69, 0x66, 0, 0, ...tiff(72, 144, 2)];
    const jpeg = new Uint8Array([0xff, 0xd8, 0xff, 0xe1, ...be16(exifPayload.length + 2), ...exifPayload, 0xff, 0xda]);
    expect(sniffRasterResolution(jpeg)).toEqual({ dpiX: 72, dpiY: 144 });
    expect(sniffRasterResolution(new Uint8Array(tiff(100, 100, 3)))?.dpiX).toBeCloseTo(254);
  });

  it('reads BMP pixels per metre and returns null for formats without a resolution', () => {
    const bmp = new Uint8Array(54);
    bmp.set([0x42, 0x4d]);
    bmp.set(le32(40), 14);
    bmp.set(le32(3780), 38);
    bmp.set(le32(3780), 42);
    expect(sniffRasterResolution(bmp)?.dpiX).toBeCloseTo(96, 0);
    expect(sniffRasterResolution(new TextEncoder().encode('GIF89a\u0001\u0000'))).toBeNull();
  });
});
