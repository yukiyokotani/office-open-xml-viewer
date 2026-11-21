import { describe, expect, it } from 'vitest';
import { inspectRasterBlob } from './raster-blob-inspection.js';

// ECMA-376 §20.1.8.14: "the DPI in the blip", read through the production
// raster inspection (header prefix, bounded deep reads, JPEG streaming).

const be32 = (value: number) => [(value >>> 24) & 255, (value >>> 16) & 255, (value >>> 8) & 255, value & 255];
const be16 = (value: number) => [(value >>> 8) & 255, value & 255];
const le16 = (value: number) => [value & 255, (value >>> 8) & 255];
const le32 = (value: number) => [value & 255, (value >>> 8) & 255, (value >>> 16) & 255, (value >>> 24) & 255];

const resolutionOf = async (bytes: number[] | Uint8Array) =>
  (await inspectRasterBlob(new Blob([new Uint8Array(bytes)]))).resolution ?? null;

const pngChunk = (type: string, data: number[]) =>
  [...be32(data.length), ...[...type].map((c) => c.charCodeAt(0)), ...data, 0, 0, 0, 0];
const PNG_SIGNATURE = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];
const IHDR = pngChunk('IHDR', [...be32(4), ...be32(2), 8, 6, 0, 0, 0]);
const pHYs = (unit: number, x: number, y: number) => pngChunk('pHYs', [...be32(x), ...be32(y), unit]);

function jpegSegment(marker: number, payload: number[]): number[] {
  return [0xff, marker, ...be16(payload.length + 2), ...payload];
}
const JFIF = (units: number, x: number, y: number) =>
  jpegSegment(0xe0, [0x4a, 0x46, 0x49, 0x46, 0, 1, 2, units, ...be16(x), ...be16(y), 0, 0]);
const SOF = (width: number, height: number) => jpegSegment(0xc0, [8, ...be16(height), ...be16(width), 1, 1, 0x11, 0]);
const SOS = [0xff, 0xda];

/** Little-endian TIFF with X/YResolution (RATIONAL), ResolutionUnit and an
 * optional Orientation, with the rational values placed at `valuesAt`. */
function tiff(x: number, y: number, unit: number, orientation?: number, valuesAt?: number): number[] {
  const entries = orientation === undefined ? 3 : 4;
  const ifd = 8;
  const directoryEnd = ifd + 2 + entries * 12 + 4;
  const values = valuesAt ?? directoryEnd;
  const out = [
    0x49, 0x49, ...le16(42), ...le32(ifd), ...le16(entries),
    ...(orientation === undefined ? [] : [...le16(0x0112), ...le16(3), ...le32(1), ...le16(orientation), 0, 0]),
    ...le16(0x011a), ...le16(5), ...le32(1), ...le32(values),
    ...le16(0x011b), ...le16(5), ...le32(1), ...le32(values + 8),
    ...le16(0x0128), ...le16(3), ...le32(1), ...le16(unit), 0, 0,
    ...le32(0),
  ];
  while (out.length < values) out.push(0);
  return [...out, ...le32(x), ...le32(1), ...le32(y), ...le32(1)];
}
const EXIF = (tiffBytes: number[]) => jpegSegment(0xe1, [0x45, 0x78, 0x69, 0x66, 0, 0, ...tiffBytes]);

describe('raster declared resolution', () => {
  it('reads PNG pHYs, including after ancillary chunks beyond the header prefix', async () => {
    const dpi = await resolutionOf([...PNG_SIGNATURE, ...IHDR, ...pHYs(1, 7559, 3780), ...pngChunk('IDAT', [])]);
    expect(dpi?.dpiX).toBeCloseTo(192, 0);
    expect(dpi?.dpiY).toBeCloseTo(96, 0);
    const far = await resolutionOf([...PNG_SIGNATURE, ...IHDR, ...pngChunk('tEXt', new Array(70_000).fill(0x41)),
      ...pHYs(1, 3780, 3780), ...pngChunk('IDAT', [])]);
    expect(far?.dpiX).toBeCloseTo(96, 0);
    // A unitless aspect ratio, a pHYs after IDAT, and a truncated walk declare none.
    expect(await resolutionOf([...PNG_SIGNATURE, ...IHDR, ...pHYs(0, 1, 1)])).toBeNull();
    expect(await resolutionOf([...PNG_SIGNATURE, ...IHDR, ...pngChunk('IDAT', []), ...pHYs(1, 1, 1)])).toBeNull();
    expect(await resolutionOf([...PNG_SIGNATURE, ...IHDR, ...pHYs(1, 3780, 3780).slice(0, 12)])).toBeNull();
  });

  it('reads JPEG JFIF density in dots per inch or per centimetre', async () => {
    expect(await resolutionOf([0xff, 0xd8, ...JFIF(1, 300, 150), ...SOS])).toEqual({ dpiX: 300, dpiY: 150 });
    expect((await resolutionOf([0xff, 0xd8, ...JFIF(2, 100, 100), ...SOS]))?.dpiX).toBeCloseTo(254);
    expect(await resolutionOf([0xff, 0xd8, ...JFIF(0, 1, 1), ...SOS])).toBeNull();
  });

  it('skips JPEG fill bytes and standalone markers before a segment', async () => {
    // FF fill bytes may precede any marker; TEM (FF01) and RSTn stand alone.
    const filled = [0xff, 0xd8, 0xff, 0xff, 0xff, 0x01, 0xff, 0xd0, 0xff, ...JFIF(1, 300, 150), ...SOS];
    expect(await resolutionOf(filled)).toEqual({ dpiX: 300, dpiY: 150 });
  });

  it('reads EXIF resolution after large segments beyond the header prefix', async () => {
    const app2 = jpegSegment(0xe2, new Array(65_000).fill(0));
    const jpeg = [0xff, 0xd8, ...app2, ...app2, ...EXIF(tiff(72, 144, 2)), ...SOS];
    expect(jpeg.length).toBeGreaterThan(64 * 1024);
    expect(await resolutionOf(jpeg)).toEqual({ dpiX: 72, dpiY: 144 });
  });

  it('swaps the resolution axes with an EXIF transposing orientation', async () => {
    const jpeg = [0xff, 0xd8, ...JFIF(1, 72, 144), ...EXIF(tiff(72, 144, 2, 6)), ...SOF(200, 100), ...SOS];
    const inspection = await inspectRasterBlob(new Blob([new Uint8Array(jpeg)]));
    expect(inspection.dimensions).toEqual({ width: 100, height: 200 });
    expect(inspection.resolution).toEqual({ dpiX: 144, dpiY: 72 });
    const upright = [0xff, 0xd8, ...EXIF(tiff(72, 144, 2, 3)), ...SOF(200, 100), ...SOS];
    expect(await resolutionOf(upright)).toEqual({ dpiX: 72, dpiY: 144 });
  });

  it('returns null without throwing for truncated or malformed JPEG metadata', async () => {
    const app0 = JFIF(1, 300, 150);
    // Truncated inside APP0, a declared length below 2, a too-short JFIF
    // payload, and EXIF values pointing outside the segment.
    expect(await resolutionOf([0xff, 0xd8, ...app0.slice(0, 10)])).toBeNull();
    expect(await resolutionOf([0xff, 0xd8, 0xff, 0xe0, 0x00, 0x01, ...SOS])).toBeNull();
    expect(await resolutionOf([0xff, 0xd8, ...jpegSegment(0xe0, [0x4a, 0x46, 0x49, 0x46, 0, 1, 2, 1]), ...SOS])).toBeNull();
    expect(await resolutionOf([0xff, 0xd8, ...EXIF(tiff(72, 72, 2, undefined, 60_000).slice(0, 60)), ...SOS])).toBeNull();
  });

  it('reads TIFF IFD0 resolution values stored beyond the header prefix', async () => {
    expect((await resolutionOf(tiff(100, 100, 3)))?.dpiX).toBeCloseTo(254);
    expect(await resolutionOf(tiff(300, 600, 2, undefined, 80_000))).toEqual({ dpiX: 300, dpiY: 600 });
    expect(await resolutionOf(tiff(300, 600, 2, undefined, 80_000).slice(0, 80_004))).toBeNull();
  });

  it('reads BMP pixels per metre and none for formats without a resolution', async () => {
    const bmp = new Uint8Array(54);
    bmp.set([0x42, 0x4d]);
    bmp.set(le32(40), 14);
    bmp.set(le32(3780), 38);
    bmp.set(le32(3780), 42);
    expect((await resolutionOf(bmp))?.dpiX).toBeCloseTo(96, 0);
    expect(await resolutionOf([...new TextEncoder().encode('GIF89a'), 1, 0, 1, 0])).toBeNull();
  });
});
