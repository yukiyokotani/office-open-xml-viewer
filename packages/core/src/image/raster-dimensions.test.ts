import { describe, it, expect } from 'vitest';
import {
  sniffRasterDimensions,
  rasterExceedsBudget,
  sourceRasterExceedsBudget,
  rasterHeaderExceedsBudget,
} from './raster-dimensions.js';
import {
  MAX_RASTER_DIMENSION,
  MAX_RASTER_PIXELS,
  MAX_RASTER_SOURCE_DIMENSION,
  MAX_RASTER_SOURCE_PIXELS,
} from './pixel-budget.js';

// ── Raster header dimension sniff + budget (RB1 decode-bomb guard) ───────────
//
// These fixtures are SYNTHETIC: each carries a valid header declaring enormous
// pixel dimensions but only a handful of bytes of "pixel data" — the shape of a
// decompression bomb (a tiny compressed blip that decodes to a multi-GB RGBA
// surface). The guard must recognize the declared dimensions from the header
// alone and report the image as over budget, WITHOUT decoding it.

/** Big-endian u32 into a byte array at `o`. */
function putBeU32(b: Uint8Array, o: number, v: number): void {
  b[o] = (v >>> 24) & 0xff;
  b[o + 1] = (v >>> 16) & 0xff;
  b[o + 2] = (v >>> 8) & 0xff;
  b[o + 3] = v & 0xff;
}
/** Little-endian u16. */
function putLeU16(b: Uint8Array, o: number, v: number): void {
  b[o] = v & 0xff;
  b[o + 1] = (v >>> 8) & 0xff;
}
/** Little-endian u32. */
function putLeU32(b: Uint8Array, o: number, v: number): void {
  b[o] = v & 0xff;
  b[o + 1] = (v >>> 8) & 0xff;
  b[o + 2] = (v >>> 16) & 0xff;
  b[o + 3] = (v >>> 24) & 0xff;
}
/** Big-endian u16. */
function putBeU16(b: Uint8Array, o: number, v: number): void {
  b[o] = (v >>> 8) & 0xff;
  b[o + 1] = v & 0xff;
}

// ── Header builders (declare `w × h`, carry almost no payload) ───────────────

function pngHeader(w: number, h: number): Uint8Array {
  // 8-byte signature + IHDR chunk (length + "IHDR" + W + H + …). We only need
  // through offset 24; a couple of trailing bytes stand in for the rest.
  const b = new Uint8Array(26);
  b.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a], 0);
  putBeU32(b, 8, 13); // IHDR length
  b.set([0x49, 0x48, 0x44, 0x52], 12); // "IHDR"
  putBeU32(b, 16, w);
  putBeU32(b, 20, h);
  return b;
}

function gifHeader(w: number, h: number): Uint8Array {
  const b = new Uint8Array(13);
  b.set([0x47, 0x49, 0x46, 0x38, 0x39, 0x61], 0); // "GIF89a"
  putLeU16(b, 6, w);
  putLeU16(b, 8, h);
  return b;
}

function bmpHeader(w: number, h: number): Uint8Array {
  // "BM" + 14-byte file header, then a 40-byte BITMAPINFOHEADER.
  const b = new Uint8Array(54);
  b.set([0x42, 0x4d], 0); // "BM"
  putLeU32(b, 14, 40); // biSize = BITMAPINFOHEADER
  putLeU32(b, 18, w >>> 0); // signed i32, positive here
  putLeU32(b, 22, h >>> 0);
  return b;
}

function webpVp8xHeader(w: number, h: number): Uint8Array {
  // RIFF…WEBP + "VP8X" chunk with 24-bit canvas width-1 / height-1.
  const b = new Uint8Array(30);
  b.set([0x52, 0x49, 0x46, 0x46], 0); // "RIFF"
  putLeU32(b, 4, 0); // file size (ignored by the sniff)
  b.set([0x57, 0x45, 0x42, 0x50], 8); // "WEBP"
  b.set([0x56, 0x50, 0x38, 0x58], 12); // "VP8X"
  putLeU32(b, 16, 10); // VP8X payload size (ignored)
  // flags byte at 20; canvas width-1 (24-bit LE) at 24, height-1 at 27.
  const w1 = w - 1;
  const h1 = h - 1;
  b[24] = w1 & 0xff;
  b[25] = (w1 >>> 8) & 0xff;
  b[26] = (w1 >>> 16) & 0xff;
  b[27] = h1 & 0xff;
  b[28] = (h1 >>> 8) & 0xff;
  b[29] = (h1 >>> 16) & 0xff;
  return b;
}

/** WebP "VP8L" (lossless): a 32-bit LE field at offset 21 packs width-1 (14
 * bits) then height-1 (14 bits), after a 0x2F signature byte at offset 20. */
function webpVp8lHeader(w: number, h: number): Uint8Array {
  const b = new Uint8Array(25);
  b.set([0x52, 0x49, 0x46, 0x46], 0); // "RIFF"
  b.set([0x57, 0x45, 0x42, 0x50], 8); // "WEBP"
  b.set([0x56, 0x50, 0x38, 0x4c], 12); // "VP8L"
  b[20] = 0x2f; // VP8L signature
  const bits = ((w - 1) & 0x3fff) | (((h - 1) & 0x3fff) << 14);
  putLeU32(b, 21, bits >>> 0);
  return b;
}

/** WebP "VP8 " (lossy simple): 14-bit dimensions at offsets 26/28. */
function webpVp8Header(w: number, h: number): Uint8Array {
  const b = new Uint8Array(30);
  b.set([0x52, 0x49, 0x46, 0x46], 0); // "RIFF"
  b.set([0x57, 0x45, 0x42, 0x50], 8); // "WEBP"
  b.set([0x56, 0x50, 0x38, 0x20], 12); // "VP8 "
  putLeU16(b, 26, w & 0x3fff);
  putLeU16(b, 28, h & 0x3fff);
  return b;
}

/** A minimal JPEG: SOI, an APP0 segment, then an SOF0 declaring `w × h`. */
function jpegHeader(w: number, h: number, padApp0 = 0): Uint8Array {
  // SOI (2) + APP0 marker (2) + APP0 length (2) + APP0 payload (padApp0) +
  // SOF0 marker (2) + SOF0 length (2) + [prec(1) h(2) w(2) …].
  const app0Len = 2 + padApp0; // length field counts itself
  const sof0PayloadLen = 2 + 1 + 2 + 2; // len + prec + h + w
  const size = 2 + 2 + app0Len + 2 + sof0PayloadLen;
  const b = new Uint8Array(size);
  let i = 0;
  b[i++] = 0xff;
  b[i++] = 0xd8; // SOI
  b[i++] = 0xff;
  b[i++] = 0xe0; // APP0
  putBeU16(b, i, app0Len);
  i += 2 + padApp0; // skip APP0 payload
  b[i++] = 0xff;
  b[i++] = 0xc0; // SOF0
  putBeU16(b, i, sof0PayloadLen);
  i += 2;
  b[i++] = 8; // precision
  putBeU16(b, i, h);
  i += 2;
  putBeU16(b, i, w);
  i += 2;
  return b;
}

function exifOrientationPayload(
  orientation: number,
  byteOrder: 'little' | 'big' = 'little',
  ifdRelativeOffset = 8,
): Uint8Array {
  const little = byteOrder === 'little';
  const exif = new Uint8Array(6 + ifdRelativeOffset + 2 + 12 + 4);
  exif.set([0x45, 0x78, 0x69, 0x66, 0x00, 0x00], 0); // "Exif\0\0"
  exif.set(little ? [0x49, 0x49] : [0x4d, 0x4d], 6);
  const view = new DataView(exif.buffer);
  view.setUint16(8, 42, little);
  view.setUint32(10, ifdRelativeOffset, little);
  const ifdOffset = 6 + ifdRelativeOffset;
  view.setUint16(ifdOffset, 1, little); // one IFD entry
  view.setUint16(ifdOffset + 2, 0x0112, little); // Orientation
  view.setUint16(ifdOffset + 4, 3, little); // SHORT
  view.setUint32(ifdOffset + 6, 1, little);
  view.setUint16(ifdOffset + 10, orientation, little);
  return exif;
}

function jpegSofSegment(w: number, h: number): Uint8Array {
  const sof = new Uint8Array(9);
  sof.set([0xff, 0xc0, 0x00, 0x07, 0x08], 0);
  putBeU16(sof, 5, h);
  putBeU16(sof, 7, w);
  return sof;
}

function jpegApp1Segment(payload: Uint8Array): Uint8Array {
  const segment = new Uint8Array(4 + payload.length);
  segment.set([0xff, 0xe1], 0);
  putBeU16(segment, 2, payload.length + 2);
  segment.set(payload, 4);
  return segment;
}

function jpegFromSegments(...segments: Uint8Array[]): Uint8Array {
  const length = 2 + segments.reduce((total, segment) => total + segment.length, 0);
  const bytes = new Uint8Array(length);
  bytes.set([0xff, 0xd8], 0);
  let offset = 2;
  for (const segment of segments) {
    bytes.set(segment, offset);
    offset += segment.length;
  }
  return bytes;
}

/** A JPEG with a minimal EXIF APP1 orientation before SOF0. */
function jpegWithExifOrientation(
  w: number,
  h: number,
  orientation: number,
  byteOrder: 'little' | 'big' = 'little',
  ifdRelativeOffset = 8,
): Uint8Array {
  return jpegFromSegments(
    jpegApp1Segment(exifOrientationPayload(orientation, byteOrder, ifdRelativeOffset)),
    jpegSofSegment(w, h),
  );
}

/**
 * A JPEG with a DHT (0xFFC4) segment *before* the SOF0. DHT shares the 0xC0..0xCF
 * range with SOF markers but is NOT a frame; the walker must skip it (using its
 * segment length) and keep scanning to reach the real SOF, not misread the DHT
 * payload as dimensions. `dhtPayload` bytes stand in for Huffman tables.
 */
function jpegWithDhtThenSof(w: number, h: number, dhtPayload = 12): Uint8Array {
  const dhtLen = 2 + dhtPayload; // length field counts itself
  const sofLen = 2 + 1 + 2 + 2;
  const size = 2 /* SOI */ + 2 + dhtLen /* DHT */ + 2 + sofLen; /* SOF0 */
  const b = new Uint8Array(size);
  let i = 0;
  b[i++] = 0xff;
  b[i++] = 0xd8; // SOI
  b[i++] = 0xff;
  b[i++] = 0xc4; // DHT (not a frame)
  putBeU16(b, i, dhtLen);
  i += 2 + dhtPayload; // skip the DHT payload
  b[i++] = 0xff;
  b[i++] = 0xc0; // SOF0
  putBeU16(b, i, sofLen);
  i += 2;
  b[i++] = 8; // precision
  putBeU16(b, i, h);
  i += 2;
  putBeU16(b, i, w);
  i += 2;
  return b;
}

/**
 * A malformed JPEG whose first non-standalone segment declares a length < 2
 * (which would leave the walker unable to advance). The sniff must bail (return
 * null) rather than loop forever or read out of bounds.
 */
function jpegMalformedSegLen(): Uint8Array {
  // SOI, then an APP0 marker with a bogus 1-byte length (< the 2-byte minimum).
  const b = new Uint8Array([0xff, 0xd8, 0xff, 0xe0, 0x00, 0x01, 0x00, 0x00]);
  return b;
}

function tiffHeader(w: number, h: number, byteOrder: 'little' | 'big'): Uint8Array {
  const little = byteOrder === 'little';
  const bytes = new Uint8Array(38);
  const view = new DataView(bytes.buffer);
  bytes.set(little ? [0x49, 0x49] : [0x4d, 0x4d], 0);
  view.setUint16(2, 42, little);
  view.setUint32(4, 8, little);
  view.setUint16(8, 2, little);
  view.setUint16(10, 256, little);
  view.setUint16(12, 4, little);
  view.setUint32(14, 1, little);
  view.setUint32(18, w, little);
  view.setUint16(22, 257, little);
  view.setUint16(24, 4, little);
  view.setUint32(26, 1, little);
  view.setUint32(30, h, little);
  return bytes;
}

describe('sniffRasterDimensions — reads declared pixel dimensions from the header', () => {
  it('reads PNG IHDR dimensions', () => {
    expect(sniffRasterDimensions(pngHeader(1920, 1080))).toEqual({ width: 1920, height: 1080 });
    expect(sniffRasterDimensions(pngHeader(60000, 60000))).toEqual({ width: 60000, height: 60000 });
  });

  it('reads GIF logical-screen dimensions', () => {
    expect(sniffRasterDimensions(gifHeader(800, 600))).toEqual({ width: 800, height: 600 });
  });

  it('reads BMP BITMAPINFOHEADER dimensions (and |height| for top-down)', () => {
    expect(sniffRasterDimensions(bmpHeader(1024, 768))).toEqual({ width: 1024, height: 768 });
    // Negative height (top-down bitmap) is budgeted by magnitude.
    const b = bmpHeader(1024, 0);
    putLeU32(b, 22, (-768 >>> 0) as number);
    expect(sniffRasterDimensions(b)).toEqual({ width: 1024, height: 768 });
  });

  it('reads WebP VP8X canvas dimensions', () => {
    expect(sniffRasterDimensions(webpVp8xHeader(4000, 3000))).toEqual({ width: 4000, height: 3000 });
  });

  it('reads WebP VP8L (lossless) dimensions', () => {
    expect(sniffRasterDimensions(webpVp8lHeader(2000, 1500))).toEqual({ width: 2000, height: 1500 });
    // Signature byte must be 0x2F; a wrong signature fails open (null).
    const bad = webpVp8lHeader(2000, 1500);
    bad[20] = 0x00;
    expect(sniffRasterDimensions(bad)).toBeNull();
  });

  it('reads WebP VP8 (lossy simple) dimensions', () => {
    expect(sniffRasterDimensions(webpVp8Header(640, 480))).toEqual({ width: 640, height: 480 });
  });

  it('reads JPEG SOF dimensions, even past an APP0 segment', () => {
    expect(sniffRasterDimensions(jpegHeader(3264, 2448))).toEqual({ width: 3264, height: 2448 });
    // With a large APP0 (EXIF/ICC stand-in) the walker still reaches the SOF.
    expect(sniffRasterDimensions(jpegHeader(3264, 2448, 400))).toEqual({
      width: 3264,
      height: 2448,
    });
  });

  it.each([5, 6, 7, 8])('swaps JPEG dimensions for EXIF orientation %i', (orientation) => {
    expect(sniffRasterDimensions(jpegWithExifOrientation(400, 100, orientation))).toEqual({
      width: 100,
      height: 400,
    });
  });

  it.each([1, 2, 3, 4])('keeps JPEG dimensions for EXIF orientation %i', (orientation) => {
    expect(sniffRasterDimensions(jpegWithExifOrientation(400, 100, orientation))).toEqual({
      width: 400,
      height: 100,
    });
  });

  it('reads big-endian EXIF at a non-default IFD offset', () => {
    expect(sniffRasterDimensions(jpegWithExifOrientation(400, 100, 8, 'big', 16))).toEqual({
      width: 100,
      height: 400,
    });
  });

  it('applies the first valid EXIF orientation even when it follows SOF', () => {
    const xmp = jpegApp1Segment(new TextEncoder().encode('http://ns.adobe.com/xap/1.0/'));
    const bytes = jpegFromSegments(
      xmp,
      jpegSofSegment(400, 100),
      jpegApp1Segment(exifOrientationPayload(6)),
      new Uint8Array([0xff, 0xda]),
    );
    expect(sniffRasterDimensions(bytes)).toEqual({ width: 100, height: 400 });
  });

  it('does not inspect EXIF-looking bytes after Start-Of-Scan', () => {
    const bytes = jpegFromSegments(
      jpegSofSegment(400, 100),
      new Uint8Array([0xff, 0xda]),
      jpegApp1Segment(exifOrientationPayload(6)),
    );
    expect(sniffRasterDimensions(bytes)).toEqual({ width: 400, height: 100 });
  });

  it.each([0, 9])('ignores invalid EXIF orientation %i', (orientation) => {
    expect(sniffRasterDimensions(jpegWithExifOrientation(400, 100, orientation))).toEqual({
      width: 400,
      height: 100,
    });
  });

  it('treats the first EXIF APP1 as authoritative even when its orientation is invalid', () => {
    const bytes = jpegFromSegments(
      jpegApp1Segment(exifOrientationPayload(0)),
      jpegApp1Segment(exifOrientationPayload(6)),
      jpegSofSegment(400, 100),
    );
    expect(sniffRasterDimensions(bytes)).toEqual({ width: 400, height: 100 });
  });

  it('uses a later valid Orientation entry within the authoritative EXIF IFD', () => {
    const exif = new Uint8Array(44);
    exif.set([0x45, 0x78, 0x69, 0x66, 0x00, 0x00, 0x49, 0x49], 0);
    const view = new DataView(exif.buffer);
    view.setUint16(8, 42, true);
    view.setUint32(10, 8, true);
    view.setUint16(14, 2, true);
    for (const [offset, orientation] of [[16, 0], [28, 6]] as const) {
      view.setUint16(offset, 0x0112, true);
      view.setUint16(offset + 2, 3, true);
      view.setUint32(offset + 4, 1, true);
      view.setUint16(offset + 8, orientation, true);
    }
    expect(sniffRasterDimensions(jpegFromSegments(
      jpegApp1Segment(exif),
      jpegSofSegment(400, 100),
    ))).toEqual({ width: 100, height: 400 });
  });

  it('uses a complete Orientation entry when the declared EXIF table is truncated later', () => {
    const exif = exifOrientationPayload(6).subarray(0, 28);
    new DataView(exif.buffer, exif.byteOffset, exif.byteLength).setUint16(14, 2, true);
    expect(sniffRasterDimensions(jpegFromSegments(
      jpegApp1Segment(exif),
      jpegSofSegment(400, 100),
    ))).toEqual({ width: 100, height: 400 });
  });

  it('does not recover a malformed first EXIF IFD from a later EXIF APP1', () => {
    const malformed = exifOrientationPayload(6);
    new DataView(malformed.buffer).setUint32(10, 0xffff, true);
    const bytes = jpegFromSegments(
      jpegApp1Segment(malformed),
      jpegApp1Segment(exifOrientationPayload(6)),
      jpegSofSegment(400, 100),
    );
    expect(sniffRasterDimensions(bytes)).toEqual({ width: 400, height: 100 });
  });

  it('ignores a truncated or out-of-range EXIF IFD', () => {
    const outOfRange = exifOrientationPayload(6);
    new DataView(outOfRange.buffer).setUint32(10, 0xffff, true);
    expect(sniffRasterDimensions(jpegFromSegments(
      jpegApp1Segment(outOfRange),
      jpegSofSegment(400, 100),
    ))).toEqual({ width: 400, height: 100 });
    expect(sniffRasterDimensions(jpegFromSegments(
      jpegApp1Segment(exifOrientationPayload(6).subarray(0, 15)),
      jpegSofSegment(400, 100),
    ))).toEqual({ width: 400, height: 100 });
  });

  it.each(['little', 'big'] as const)('reads TIFF IFD dimensions (%s endian)', (byteOrder) => {
    expect(sniffRasterDimensions(tiffHeader(776, 31, byteOrder))).toEqual({
      width: 776,
      height: 31,
    });
  });

  it('skips a DHT (0xFFC4) segment and reads the following SOF, not the DHT payload', () => {
    // DHT sits in the 0xC0..0xCF range but is not a frame; misreading it would
    // yield garbage dimensions. The walker must length-skip it to the real SOF.
    expect(sniffRasterDimensions(jpegWithDhtThenSof(3000, 2000))).toEqual({
      width: 3000,
      height: 2000,
    });
  });

  it('bails (null) on a JPEG segment length < 2 without looping', () => {
    // A bogus sub-minimum segment length must terminate the walk, not spin.
    expect(sniffRasterDimensions(jpegMalformedSegLen())).toBeNull();
  });

  it('returns null for unrecognized / too-short / SVG headers (fail-open)', () => {
    expect(sniffRasterDimensions(new Uint8Array([]))).toBeNull();
    expect(sniffRasterDimensions(new Uint8Array([0x00, 0x01, 0x02, 0x03]))).toBeNull();
    // SVG starts with "<?xml" or "<svg" — not a recognized raster.
    expect(sniffRasterDimensions(new TextEncoder().encode('<svg width="9"></svg>'))).toBeNull();
    // PNG signature but a first chunk that is not IHDR → don't trust it.
    const notIhdr = pngHeader(10, 10);
    notIhdr[12] = 0x00; // corrupt "IHDR" → "\0HDR"
    expect(sniffRasterDimensions(notIhdr)).toBeNull();
  });
});

describe('rasterExceedsBudget — enforces the shared pixel budget', () => {
  it('accepts in-budget dimensions', () => {
    expect(rasterExceedsBudget({ width: 1920, height: 1080 })).toBe(false);
    // A 28 MP scan remains within the 32 MP budget and per-axis cap.
    expect(rasterExceedsBudget({ width: 7000, height: 4000 })).toBe(false);
  });

  it('rejects an over-megapixel image within the per-axis cap', () => {
    // Both axes ≤ 32767 but the product blows the 32 MP budget.
    const w = 30000;
    const h = 30000;
    expect(w).toBeLessThanOrEqual(MAX_RASTER_DIMENSION);
    expect(h).toBeLessThanOrEqual(MAX_RASTER_DIMENSION);
    expect(w * h).toBeGreaterThan(MAX_RASTER_PIXELS);
    expect(rasterExceedsBudget({ width: w, height: h })).toBe(true);
  });

  it('rejects an over-dimension image even when it would be within the MP budget', () => {
    // 40000 × 1: only 40 000 px (well under 32 MP) but past the 32767 axis cap,
    // so it could never be drawn to a canvas — reject it.
    expect(rasterExceedsBudget({ width: 40000, height: 1 })).toBe(true);
  });

  it('treats non-positive or non-finite dimensions as over budget (untrustworthy header)', () => {
    expect(rasterExceedsBudget({ width: 0, height: 100 })).toBe(true);
    expect(rasterExceedsBudget({ width: -5, height: 100 })).toBe(true);
    expect(rasterExceedsBudget({ width: NaN, height: 100 })).toBe(true);
  });

  it('accepts exactly at the caps (boundary)', () => {
    expect(rasterExceedsBudget({ width: MAX_RASTER_DIMENSION, height: 1 })).toBe(false);
    // A 1 × MAX_RASTER_PIXELS strip is exactly the MP budget (but > axis cap):
    // use a squarer at-budget case instead.
    expect(rasterExceedsBudget({ width: 4096, height: 8192 })).toBe(false);
    expect(4096 * 8192).toBe(MAX_RASTER_PIXELS);
  });
});

describe('rasterHeaderExceedsBudget — end-to-end neutralization of decode bombs', () => {
  it('flags a 60000×60000 PNG bomb (tiny payload, huge declared size)', () => {
    const bomb = pngHeader(60000, 60000);
    // The fixture is a couple dozen bytes — a real decode would be ~14 GB RGBA.
    expect(bomb.length).toBeLessThan(64);
    expect(rasterHeaderExceedsBudget(bomb)).toBe(true);
  });

  it('flags GIF / BMP / WebP / JPEG bombs alike', () => {
    expect(rasterHeaderExceedsBudget(gifHeader(50000, 50000))).toBe(true);
    expect(rasterHeaderExceedsBudget(bmpHeader(40000, 40000))).toBe(true);
    expect(rasterHeaderExceedsBudget(webpVp8xHeader(16384, 16384))).toBe(true); // 268 MP
    expect(rasterHeaderExceedsBudget(jpegHeader(60000, 60000))).toBe(true);
  });

  it('flags WebP VP8L / VP8 bombs (both non-VP8X lossy/lossless variants)', () => {
    // VP8L/VP8 dimensions are 14-bit (≤16383 per axis), so a 16383² bomb is
    // ~268 MP — well past the 64 MP budget while staying representable.
    expect(rasterHeaderExceedsBudget(webpVp8lHeader(16383, 16383))).toBe(true);
    expect(rasterHeaderExceedsBudget(webpVp8Header(16383, 16383))).toBe(true);
  });

  it('does NOT flag a normal, in-budget image (no false positive)', () => {
    expect(rasterHeaderExceedsBudget(pngHeader(1920, 1080))).toBe(false);
    expect(rasterHeaderExceedsBudget(jpegHeader(4032, 3024))).toBe(false);
    expect(rasterHeaderExceedsBudget(gifHeader(500, 500))).toBe(false);
    expect(rasterHeaderExceedsBudget(webpVp8lHeader(2000, 1500))).toBe(false);
    expect(rasterHeaderExceedsBudget(webpVp8Header(640, 480))).toBe(false);
  });

  it('does NOT flag an unrecognized header (fail-open — the caller decodes normally)', () => {
    expect(rasterHeaderExceedsBudget(new Uint8Array([0xde, 0xad, 0xbe, 0xef]))).toBe(false);
    expect(rasterHeaderExceedsBudget(new TextEncoder().encode('<svg/>'))).toBe(false);
  });
});

describe('sourceRasterExceedsBudget — encoded-source hard ceiling', () => {
  it('admits the reported 109,571,670-pixel poster class for bounded display decode', () => {
    const dimensions = { width: 12_090, height: 9_063 };
    expect(dimensions.width * dimensions.height).toBe(109_571_670);
    expect(dimensions.width * dimensions.height).toBeGreaterThan(MAX_RASTER_PIXELS);
    expect(sourceRasterExceedsBudget(dimensions)).toBe(false);
  });

  it('rejects source grids beyond the separate 128-Mi-pixel hard ceiling', () => {
    expect(sourceRasterExceedsBudget({ width: 30_000, height: 30_000 })).toBe(true);
    expect(30_000 * 30_000).toBeGreaterThan(MAX_RASTER_SOURCE_PIXELS);
  });

  it('allows a source axis above the retained-canvas limit when resizing can bound it', () => {
    expect(40_000).toBeGreaterThan(MAX_RASTER_DIMENSION);
    expect(40_000).toBeLessThan(MAX_RASTER_SOURCE_DIMENSION);
    expect(sourceRasterExceedsBudget({ width: 40_000, height: 2_000 })).toBe(false);
  });

  it('does not impose JPEG\'s 16-bit axis limit on other raster formats', () => {
    expect(MAX_RASTER_SOURCE_DIMENSION).toBe(MAX_RASTER_SOURCE_PIXELS);
    expect(sourceRasterExceedsBudget({ width: 100_000, height: 1_000 })).toBe(false);
  });
});
