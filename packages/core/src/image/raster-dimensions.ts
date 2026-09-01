// ── Raster pixel-dimension sniffing + budget (decode-bomb guard) ─────────────
//
// A blip in an OOXML part is attacker-controllable bytes. `createImageBitmap`
// decodes the compressed image into an uncompressed RGBA surface sized by the
// image *header*, not by the compressed byte length — so a tiny PNG/JPEG whose
// header declares, say, 60000×60000 forces the browser to allocate a multi-GB
// bitmap ("decompression bomb"), which OOMs or hangs the tab. The ZIP-entry
// size cap (RB11) does not help: the compressed part is small; only the decoded
// surface is enormous.
//
// This module sniffs the pixel dimensions straight from the image header (no
// full decode) so an over-budget image can be refused BEFORE it ever reaches
// `createImageBitmap`. It recognizes the raster formats OOXML embeds — PNG,
// JPEG, GIF, BMP, WebP, TIFF — and returns `null` for anything it does not recognize
// (SVG, metafiles, truncated/garbage headers), leaving the caller to fall back
// to its normal path. Recognizing "too big" is a safe superset: an unrecognized
// header simply isn't blocked here.

import {
  MAX_RASTER_DIMENSION,
  MAX_RASTER_PIXELS,
  MAX_RASTER_SOURCE_DIMENSION,
  MAX_RASTER_SOURCE_PIXELS,
} from './pixel-budget.js';
import { sniffTiffDimensions } from './tiff-contract.js';

/** Read a big-endian u16 at `o` (bounds already checked by the caller). */
function beU16(b: Uint8Array, o: number): number {
  return (b[o] << 8) | b[o + 1];
}

/** Read a big-endian u32 at `o`. `>>> 0` keeps it an unsigned 32-bit value. */
function beU32(b: Uint8Array, o: number): number {
  return ((b[o] << 24) | (b[o + 1] << 16) | (b[o + 2] << 8) | b[o + 3]) >>> 0;
}

/** Read a little-endian u16 at `o`. */
function leU16(b: Uint8Array, o: number): number {
  return b[o] | (b[o + 1] << 8);
}

/** Read a little-endian i32 at `o` (BMP dimensions are signed). */
function leI32(b: Uint8Array, o: number): number {
  return (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) | 0;
}

/** Pixel dimensions sniffed from an image header. */
export interface RasterDimensions {
  width: number;
  height: number;
}

/**
 * Sniff the pixel width/height of a raster image from its header bytes, or
 * `null` if the format is not one we recognize (or the header is too short /
 * malformed to read the dimensions).
 *
 * `head` need only contain the leading bytes of the file — 30 bytes covers PNG,
 * GIF, BMP and WebP; JPEG's dimensions live in a later SOF marker, so for JPEG a
 * larger prefix (a few hundred bytes typically, ideally the whole file) yields a
 * result while a short prefix returns `null` (fail-open: not recognized ⇒ not
 * blocked). Never throws.
 */
export function sniffRasterDimensions(head: Uint8Array): RasterDimensions | null {
  const n = head.length;

  // PNG — 8-byte signature, then the IHDR chunk: [len u32][type "IHDR"][W u32]
  // [H u32]… Dimensions are big-endian at offsets 16 and 20. ([ISO/IEC 15948])
  if (
    n >= 24 &&
    head[0] === 0x89 &&
    head[1] === 0x50 && // P
    head[2] === 0x4e && // N
    head[3] === 0x47 && // G
    head[4] === 0x0d &&
    head[5] === 0x0a &&
    head[6] === 0x1a &&
    head[7] === 0x0a
  ) {
    // Verify the first chunk really is IHDR before trusting the dimensions.
    if (head[12] === 0x49 && head[13] === 0x48 && head[14] === 0x44 && head[15] === 0x52) {
      return { width: beU32(head, 16), height: beU32(head, 20) };
    }
    return null;
  }

  // GIF — "GIF87a"/"GIF89a", then the logical-screen descriptor: width, height
  // as little-endian u16 at offsets 6 and 8. (Note: this is the canvas size, an
  // upper bound on any frame — exactly what we want to budget.)
  if (
    n >= 10 &&
    head[0] === 0x47 && // G
    head[1] === 0x49 && // I
    head[2] === 0x46 && // F
    head[3] === 0x38 && // 8
    (head[4] === 0x37 || head[4] === 0x39) && // 7 | 9
    head[5] === 0x61 // a
  ) {
    return { width: leU16(head, 6), height: leU16(head, 8) };
  }

  // BMP — "BM", then (skipping the 14-byte file header) a DIB header whose first
  // u32 is its own size. The two common headers put signed i32 width/height at
  // offsets 18/22 (BITMAPINFOHEADER, 40-byte, and later). The rare 12-byte
  // BITMAPCOREHEADER uses u16 width/height at 18/20. Height may be negative
  // (top-down); budget on its magnitude.
  if (n >= 26 && head[0] === 0x42 && head[1] === 0x4d) {
    const dibSize = beU32Le(head, 14);
    if (dibSize === 12) {
      // BITMAPCOREHEADER: u16 dimensions.
      return { width: leU16(head, 18), height: leU16(head, 20) };
    }
    // BITMAPINFOHEADER and successors: signed i32 dimensions.
    return { width: Math.abs(leI32(head, 18)), height: Math.abs(leI32(head, 22)) };
  }

  // WebP — RIFF container: "RIFF"[size u32]"WEBP", then a chunk FourCC:
  //   VP8  (lossy):   after the 3-byte frame tag, two u16 dimensions (14-bit,
  //                   the top 2 bits are scale) at offsets 26/28.
  //   VP8L (lossless): a 32-bit little-endian field at offset 21 packs
  //                    width-1 (14 bits) then height-1 (14 bits).
  //   VP8X (extended): a 24-bit little-endian canvas width-1 / height-1 at
  //                    offsets 24/27.
  if (
    n >= 16 &&
    head[0] === 0x52 && // R
    head[1] === 0x49 && // I
    head[2] === 0x46 && // F
    head[3] === 0x46 && // F
    head[8] === 0x57 && // W
    head[9] === 0x45 && // E
    head[10] === 0x42 && // B
    head[11] === 0x50 // P
  ) {
    return sniffWebp(head);
  }

  // JPEG — starts with SOI (FFD8). Walk the marker segments to the first
  // Start-Of-Frame (SOF0..SOFF, excluding non-SOF FFC4/FFC8/FFCC), whose payload
  // holds [precision u8][height u16][width u16], big-endian.
  if (n >= 4 && head[0] === 0xff && head[1] === 0xd8) {
    return sniffJpegSof(head);
  }

  // TIFF — classic TIFF starts with byte order (`II`/`MM`), magic 42, then a
  // u32 offset to its first IFD. Width/height are IFD tags 256/257. The IFD may
  // live beyond the caller's header prefix, in which case this safely returns
  // null and the TIFF decoder validates dimensions before allocating pixels.
  if (
    n >= 4
    && ((head[0] === 0x49 && head[1] === 0x49) || (head[0] === 0x4d && head[1] === 0x4d))
  ) {
    return sniffTiffDimensions(head);
  }

  return null;
}

/** Read a little-endian u32 at `o` (used for the BMP DIB-header size). */
function beU32Le(b: Uint8Array, o: number): number {
  return (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0;
}

function sniffWebp(b: Uint8Array): RasterDimensions | null {
  const n = b.length;
  // Chunk FourCC at offset 12.
  const c0 = b[12];
  const c1 = b[13];
  const c2 = b[14];
  const c3 = b[15];
  // "VP8 " (lossy simple)
  if (c0 === 0x56 && c1 === 0x50 && c2 === 0x38 && c3 === 0x20) {
    // frame tag at 20 (3 bytes) + start code (3 bytes) then dimensions at 26/28.
    if (n < 30) return null;
    const width = leU16(b, 26) & 0x3fff;
    const height = leU16(b, 28) & 0x3fff;
    return { width, height };
  }
  // "VP8L" (lossless)
  if (c0 === 0x56 && c1 === 0x50 && c2 === 0x38 && c3 === 0x4c) {
    if (n < 25) return null;
    // Signature byte 0x2F at offset 20, then a 32-bit LE field at 21.
    if (b[20] !== 0x2f) return null;
    const bits = beU32Le(b, 21);
    const width = (bits & 0x3fff) + 1;
    const height = ((bits >>> 14) & 0x3fff) + 1;
    return { width, height };
  }
  // "VP8X" (extended)
  if (c0 === 0x56 && c1 === 0x50 && c2 === 0x38 && c3 === 0x58) {
    if (n < 30) return null;
    // 24-bit LE canvas width-1 at 24, height-1 at 27.
    const width = (b[24] | (b[25] << 8) | (b[26] << 16)) + 1;
    const height = (b[27] | (b[28] << 8) | (b[29] << 16)) + 1;
    return { width, height };
  }
  return null;
}

function sniffJpegSof(b: Uint8Array): RasterDimensions | null {
  const n = b.length;
  let i = 2; // past SOI (FFD8)
  while (i + 1 < n) {
    // Markers are 0xFF followed by a non-0x00, non-0xFF type byte. Fill bytes
    // (0xFF 0xFF) and stuffed 0x00 are skipped by advancing one byte.
    if (b[i] !== 0xff) {
      i += 1;
      continue;
    }
    const marker = b[i + 1];
    if (marker === 0xff) {
      // fill byte
      i += 1;
      continue;
    }
    // Standalone markers with no length: RSTn (D0..D7), SOI (D8), EOI (D9),
    // TEM (01). Advance past them.
    if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) {
      i += 2;
      continue;
    }
    if (marker === 0xd9) {
      // EOI — no frame found.
      return null;
    }
    // All other markers carry a 2-byte big-endian segment length (including the
    // length field itself). Need the length bytes.
    if (i + 3 >= n) return null;
    const segLen = beU16(b, i + 2);
    // SOF markers: 0xC0..0xCF EXCEPT 0xC4 (DHT), 0xC8 (JPG), 0xCC (DAC).
    const isSof =
      marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc;
    if (isSof) {
      // Payload: [precision u8][height u16][width u16] at i+4.
      if (i + 8 >= n) return null;
      const height = beU16(b, i + 5);
      const width = beU16(b, i + 7);
      return { width, height };
    }
    // Skip this segment: 2 (marker) + segLen (length field + payload).
    if (segLen < 2) return null; // malformed
    i += 2 + segLen;
  }
  return null;
}

/**
 * `true` when `dims` exceeds the raster pixel budget — a per-axis cap or the
 * total-megapixel cap (see `./pixel-budget`). Non-positive or non-finite
 * dimensions also count as "over budget" (a header we can't trust). Callers
 * treat `true` as "do not decode this image".
 */
export function rasterExceedsBudget(dims: RasterDimensions): boolean {
  const { width, height } = dims;
  if (!Number.isFinite(width) || !Number.isFinite(height)) return true;
  if (width <= 0 || height <= 0) return true;
  if (width > MAX_RASTER_DIMENSION || height > MAX_RASTER_DIMENSION) return true;
  // width, height ≤ MAX_RASTER_DIMENSION (≤ 32767), so the product is ≤ ~1.07e9,
  // exact in a double — a plain comparison suffices.
  return width * height > MAX_RASTER_PIXELS;
}

/**
 * Whether an encoded raster's declared source grid is too large to hand to a
 * browser or injected decoder, even when the retained output will be resized.
 * This is deliberately separate from {@link rasterExceedsBudget}: the latter
 * bounds a decoded surface retained by the renderer, while this ceiling bounds
 * decoder input complexity and possible implementation-specific temporaries.
 */
export function sourceRasterExceedsBudget(dims: RasterDimensions): boolean {
  const { width, height } = dims;
  if (!Number.isFinite(width) || !Number.isFinite(height)) return true;
  if (width <= 0 || height <= 0) return true;
  if (width > MAX_RASTER_SOURCE_DIMENSION || height > MAX_RASTER_SOURCE_DIMENSION) return true;
  return width * height > MAX_RASTER_SOURCE_PIXELS;
}

/**
 * Convenience: sniff `head` and report whether the image is a recognized raster
 * that exceeds the pixel budget. `false` for unrecognized headers (fail-open) —
 * only a recognized, over-budget raster returns `true`.
 */
export function rasterHeaderExceedsBudget(head: Uint8Array): boolean {
  const dims = sniffRasterDimensions(head);
  return dims !== null && rasterExceedsBudget(dims);
}
