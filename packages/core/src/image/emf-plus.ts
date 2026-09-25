// EMF+ playback ([MS-EMFPLUS]) for the records GDI+ writes when it records a
// bitmap into a metafile: EmfPlusHeader, EmfPlusSetPageTransform, EmfPlusClear,
// EmfPlusSetCompositingMode, EmfPlusObject (ImageAttributes and uncompressed
// 32-bit bitmap Image objects, including continued objects), EmfPlusDrawImage,
// rendering-quality state records, world-transform records, Save/Restore,
// GetDC and EmfPlusEndOfFile.
//
// EMF+ records live in EMR_COMMENT records ("EMF+" identifier). A dual-mode
// metafile also carries a complete GDI rendering, and EMF+-aware players play
// one of the two. `scanEmfPlus` therefore decides up front:
//   - the EMF+ stream is played only when it contains drawing records and
//     every record and object it uses is implemented here;
//   - otherwise a dual-mode file keeps its GDI rendering (the player's
//     established behaviour), and an EMF+-only file is played as far as it is
//     implemented with the gaps reported, never dropped silently.
// Evidence for choosing the EMF+ rendering: Excel's PDF exports of GDI+
// metafiles whose GDI part has no drawing records (a bitmap drawn by
// EmfPlusDrawImage) show that bitmap; PowerPoint's PDF export of dual files
// whose EMF+ part has no drawing records shows the GDI drawing.

import { blitDibToCtx, type DecodedDib } from './dib.js';

const PLUS = {
  HEADER: 0x4001,
  EOF: 0x4002,
  COMMENT: 0x4003,
  GET_DC: 0x4004,
  OBJECT: 0x4008,
  CLEAR: 0x4009,
  DRAW_IMAGE: 0x401a,
  SET_ANTI_ALIAS_MODE: 0x401e,
  SET_TEXT_RENDERING_HINT: 0x401f,
  SET_TEXT_CONTRAST: 0x4020,
  SET_INTERPOLATION_MODE: 0x4021,
  SET_PIXEL_OFFSET_MODE: 0x4022,
  SET_COMPOSITING_MODE: 0x4023,
  SET_COMPOSITING_QUALITY: 0x4024,
  SAVE: 0x4025,
  RESTORE: 0x4026,
  SET_WORLD_TRANSFORM: 0x402a,
  RESET_WORLD_TRANSFORM: 0x402b,
  MULTIPLY_WORLD_TRANSFORM: 0x402c,
  TRANSLATE_WORLD_TRANSFORM: 0x402d,
  SCALE_WORLD_TRANSFORM: 0x402e,
  ROTATE_WORLD_TRANSFORM: 0x402f,
  SET_PAGE_TRANSFORM: 0x4030,
} as const;

/** Records this player implements (state records with no visible effect on
 *  the implemented drawing are accepted as no-ops). */
const IMPLEMENTED = new Set<number>(Object.values(PLUS));

/** [MS-EMFPLUS] 2.1.1.1 RecordType drawing records. */
const DRAWING = new Set<number>([
  0x400a, 0x400b, 0x400c, 0x400d, 0x400e, 0x400f, 0x4010, 0x4011, 0x4012, 0x4013,
  0x4014, 0x4015, 0x4016, 0x4017, 0x4018, 0x4019, 0x401a, 0x401b, 0x401c, 0x4036,
]);

/** [MS-EMFPLUS] 2.1.1.22 ObjectType. */
const OBJECT_IMAGE = 5;
const OBJECT_IMAGE_ATTRIBUTES = 8;
/** [MS-EMFPLUS] 2.1.1.25 PixelFormat: 32bpp ARGB and premultiplied ARGB. */
const PIXEL_32BPP_ARGB = 0x0026200a;
const PIXEL_32BPP_PARGB = 0x000e200b;
/** [MS-EMFPLUS] 2.1.1.33 UnitType. */
const UNIT_PIXEL = 2;
// Resource policy, not a format limit.
const MAX_BITMAP_PIXELS = 40_000_000;
const MAX_OBJECT_BYTES = 256 * 1024 * 1024;

interface PlusRecord {
  type: number;
  flags: number;
  /** Absolute byte range of the record data. */
  start: number;
  end: number;
}

/** Walk the EMF+ records of one EMR_COMMENT (record starting at `pos`). */
function* plusRecords(dv: DataView, pos: number, recEnd: number): Generator<PlusRecord> {
  if (recEnd - pos < 16) return;
  const dataSize = dv.getUint32(pos + 8, true);
  if (dv.getUint32(pos + 12, true) !== 0x2b464d45) return; // 'EMF+'
  const end = Math.min(recEnd, pos + 12 + dataSize);
  let p = pos + 16;
  while (p + 12 <= end) {
    const type = dv.getUint16(p, true);
    const flags = dv.getUint16(p + 2, true);
    const size = dv.getUint32(p + 4, true);
    const dataLength = dv.getUint32(p + 8, true);
    if (size < 12 || p + size > end || dataLength > size - 12) return;
    yield { type, flags, start: p + 12, end: p + 12 + dataLength };
    p += size;
  }
}

export interface EmfPlusScan {
  /** The file carries an EmfPlusHeader. */
  present: boolean;
  /** EmfPlusHeader Flags bit 0: the GDI records are a complete alternative. */
  dual: boolean;
  /** Play the EMF+ records instead of the GDI drawing. */
  play: boolean;
}

/** Decide which of a metafile's renderings to play (see the module note). */
export function scanEmfPlus(bytes: Uint8Array): EmfPlusScan {
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let present = false;
  let dual = false;
  let drawing = false;
  let implemented = true;
  let pos = 0;
  while (pos + 8 <= bytes.length) {
    const type = dv.getUint32(pos, true);
    const size = dv.getUint32(pos + 4, true);
    if (size < 8 || pos + size > bytes.length) break;
    if (type === 70) {
      for (const record of plusRecords(dv, pos, pos + size)) {
        if (record.type === PLUS.HEADER) {
          present = true;
          dual = (record.flags & 1) !== 0;
        }
        if (DRAWING.has(record.type)) drawing = true;
        if (!IMPLEMENTED.has(record.type)) implemented = false;
        if (record.type === PLUS.OBJECT && !objectImplemented(dv, record)) implemented = false;
      }
    }
    if (type === 14) break;
    pos += size;
  }
  return { present, dual, play: present && drawing && (implemented || !dual) };
}

function objectImplemented(dv: DataView, record: PlusRecord): boolean {
  const kind = (record.flags >> 8) & 0x7f;
  if (kind === OBJECT_IMAGE_ATTRIBUTES) return true;
  if (kind !== OBJECT_IMAGE) return false;
  // A continued object's first record carries TotalObjectSize first; later
  // fragments carry no header, so only the first fragment is inspected.
  const continued = (record.flags & 0x8000) !== 0;
  const at = record.start + (continued ? 4 : 0);
  if (record.end - at < 28) return continued; // a later fragment
  const imageType = dv.getUint32(at + 4, true);
  const pixelFormat = dv.getUint32(at + 20, true);
  const bitmapType = dv.getUint32(at + 24, true);
  return imageType === 1
    && bitmapType === 0
    && (pixelFormat === PIXEL_32BPP_ARGB || pixelFormat === PIXEL_32BPP_PARGB);
}

/** The subset of the GDI player state the EMF+ player draws through. */
export interface EmfPlusTarget {
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;
  W: number;
  H: number;
  /** Device-unit rectangle of the picture frame mapped onto W×H. */
  left: number;
  top: number;
  boundsW: number;
  boundsH: number;
  drew: boolean;
  unsupported: Set<string>;
}

interface Matrix {
  a: number;
  b: number;
  c: number;
  d: number;
  e: number;
  f: number;
}
const IDENTITY: Matrix = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };

/** `m1` then `m2` ([MS-EMFPLUS] 2.2.2.25 EmfPlusTransformMatrix order). */
function multiply(m1: Matrix, m2: Matrix): Matrix {
  return {
    a: m1.a * m2.a + m1.b * m2.c,
    b: m1.a * m2.b + m1.b * m2.d,
    c: m1.c * m2.a + m1.d * m2.c,
    d: m1.c * m2.b + m1.d * m2.d,
    e: m1.e * m2.a + m1.f * m2.c + m2.e,
    f: m1.e * m2.b + m1.f * m2.d + m2.f,
  };
}

interface Bitmap extends DecodedDib {}

export class EmfPlusPlayer {
  private world: Matrix = IDENTITY;
  private pageUnit = UNIT_PIXEL;
  private pageScale = 1;
  private sourceCopy = false;
  private readonly images = new Map<number, Bitmap>();
  private readonly saved = new Map<number, { world: Matrix; unit: number; scale: number }>();
  private pending: { id: number; total: number; parts: Uint8Array[]; length: number } | null = null;
  /** Inside EmfPlusGetDC until the next EMF+ record: GDI records draw. */
  gdiAllowed = false;

  constructor(private readonly target: EmfPlusTarget) {}

  /** Play the EMF+ records of one EMR_COMMENT record. */
  playComment(dv: DataView, pos: number, recEnd: number): void {
    for (const record of plusRecords(dv, pos, recEnd)) {
      this.gdiAllowed = false;
      try {
        this.play(dv, record);
      } catch {
        this.target.unsupported.add(`EMF+ record 0x${record.type.toString(16)} (malformed)`);
      }
    }
  }

  private play(dv: DataView, r: PlusRecord): void {
    const f32 = (at: number) => dv.getFloat32(r.start + at, true);
    const need = (bytes: number) => {
      if (r.end - r.start < bytes) throw new RangeError('Truncated EMF+ record');
    };
    switch (r.type) {
      case PLUS.HEADER:
      case PLUS.EOF:
      case PLUS.COMMENT:
      case PLUS.SET_ANTI_ALIAS_MODE:
      case PLUS.SET_TEXT_RENDERING_HINT:
      case PLUS.SET_TEXT_CONTRAST:
      case PLUS.SET_INTERPOLATION_MODE:
      case PLUS.SET_PIXEL_OFFSET_MODE:
      case PLUS.SET_COMPOSITING_QUALITY:
        return;
      case PLUS.GET_DC:
        this.gdiAllowed = true;
        return;
      case PLUS.SET_PAGE_TRANSFORM:
        // [MS-EMFPLUS] 2.3.9.3: Flags low byte is the UnitType; PageScale.
        need(4);
        this.pageUnit = r.flags & 0xff;
        this.pageScale = f32(0);
        return;
      case PLUS.SET_COMPOSITING_MODE:
        // [MS-EMFPLUS] 2.1.1.5: 0 SourceOver, 1 SourceCopy.
        this.sourceCopy = (r.flags & 0xff) === 1;
        return;
      case PLUS.SET_WORLD_TRANSFORM:
        need(24);
        this.world = { a: f32(0), b: f32(4), c: f32(8), d: f32(12), e: f32(16), f: f32(20) };
        return;
      case PLUS.RESET_WORLD_TRANSFORM:
        this.world = IDENTITY;
        return;
      case PLUS.MULTIPLY_WORLD_TRANSFORM: {
        need(24);
        const m = { a: f32(0), b: f32(4), c: f32(8), d: f32(12), e: f32(16), f: f32(20) };
        // Flags 0x2000 (A): post-multiply (append); otherwise pre-multiply.
        this.world = r.flags & 0x2000 ? multiply(this.world, m) : multiply(m, this.world);
        return;
      }
      case PLUS.TRANSLATE_WORLD_TRANSFORM:
      case PLUS.SCALE_WORLD_TRANSFORM:
      case PLUS.ROTATE_WORLD_TRANSFORM: {
        let m: Matrix;
        if (r.type === PLUS.ROTATE_WORLD_TRANSFORM) {
          need(4);
          const angle = (f32(0) * Math.PI) / 180;
          m = { a: Math.cos(angle), b: Math.sin(angle), c: -Math.sin(angle), d: Math.cos(angle), e: 0, f: 0 };
        } else {
          need(8);
          m = r.type === PLUS.TRANSLATE_WORLD_TRANSFORM
            ? { a: 1, b: 0, c: 0, d: 1, e: f32(0), f: f32(4) }
            : { a: f32(0), b: 0, c: 0, d: f32(4), e: 0, f: 0 };
        }
        this.world = r.flags & 0x2000 ? multiply(this.world, m) : multiply(m, this.world);
        return;
      }
      case PLUS.SAVE:
        need(4);
        this.saved.set(dv.getUint32(r.start, true), {
          world: this.world,
          unit: this.pageUnit,
          scale: this.pageScale,
        });
        return;
      case PLUS.RESTORE: {
        need(4);
        const state = this.saved.get(dv.getUint32(r.start, true));
        if (state) {
          this.world = state.world;
          this.pageUnit = state.unit;
          this.pageScale = state.scale;
        }
        return;
      }
      case PLUS.CLEAR:
        need(4);
        this.clear(dv.getUint32(r.start, true));
        return;
      case PLUS.OBJECT:
        this.object(dv, r);
        return;
      case PLUS.DRAW_IMAGE:
        this.drawImage(dv, r);
        return;
      default:
        this.target.unsupported.add(`EMF+ record 0x${r.type.toString(16)}`);
    }
  }

  /** [MS-EMFPLUS] 2.3.4.1 EmfPlusClear: the output area becomes the ARGB colour. */
  private clear(argb: number): void {
    const { ctx, W, H } = this.target;
    const alpha = (argb >>> 24) / 255;
    try {
      ctx.clearRect(0, 0, W, H);
      if (alpha > 0) {
        ctx.fillStyle = `rgba(${(argb >>> 16) & 255},${(argb >>> 8) & 255},${argb & 255},${alpha})`;
        ctx.fillRect(0, 0, W, H);
        this.target.drew = true;
      }
    } catch {
      /* a ctx without clearRect (some mocks) */
    }
  }

  /** [MS-EMFPLUS] 2.3.5.1 EmfPlusObject, with continued-object assembly. */
  private object(dv: DataView, r: PlusRecord): void {
    const id = r.flags & 0xff;
    const kind = (r.flags >> 8) & 0x7f;
    let data = new Uint8Array(dv.buffer, dv.byteOffset + r.start, r.end - r.start);
    if (r.flags & 0x8000) {
      // Continued: TotalObjectSize, then this fragment.
      if (data.length < 4) throw new RangeError('Truncated EMF+ object');
      const total = dv.getUint32(r.start, true);
      if (total > MAX_OBJECT_BYTES) throw new RangeError('EMF+ object too large');
      if (!this.pending || this.pending.id !== id) {
        this.pending = { id, total, parts: [], length: 0 };
      }
      const part = data.subarray(4);
      this.pending.parts.push(part);
      this.pending.length += part.length;
      if (this.pending.length < this.pending.total) return;
      const whole = new Uint8Array(this.pending.length);
      let at = 0;
      for (const piece of this.pending.parts) {
        whole.set(piece, at);
        at += piece.length;
      }
      this.pending = null;
      data = whole;
    }
    if (kind === OBJECT_IMAGE_ATTRIBUTES) return; // defaults suit in-bounds sources
    if (kind !== OBJECT_IMAGE) {
      this.target.unsupported.add(`EMF+ object type ${kind}`);
      return;
    }
    const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
    if (data.length < 28) throw new RangeError('Truncated EMF+ image');
    // EmfPlusImage (2.2.1.4): Version, Type; EmfPlusBitmap (2.2.2.2): Width,
    // Height, Stride, PixelFormat, Type, BitmapData.
    const imageType = view.getUint32(4, true);
    const width = view.getInt32(8, true);
    const height = view.getInt32(12, true);
    const stride = view.getInt32(16, true);
    const pixelFormat = view.getUint32(20, true);
    const bitmapType = view.getUint32(24, true);
    if (imageType !== 1 || bitmapType !== 0
      || (pixelFormat !== PIXEL_32BPP_ARGB && pixelFormat !== PIXEL_32BPP_PARGB)) {
      this.target.unsupported.add('EMF+ image other than an uncompressed 32-bit bitmap');
      return;
    }
    if (width <= 0 || height <= 0 || width * height > MAX_BITMAP_PIXELS
      || stride < width * 4 || 28 + stride * height > data.length) {
      throw new RangeError('Invalid EMF+ bitmap');
    }
    // BitmapData is top-down B,G,R,A per pixel (little-endian ARGB).
    const rgba = new Uint8ClampedArray(width * height * 4);
    const premultiplied = pixelFormat === PIXEL_32BPP_PARGB;
    for (let y = 0; y < height; y++) {
      let src = 28 + y * stride;
      let dst = y * width * 4;
      for (let x = 0; x < width; x++, src += 4, dst += 4) {
        const alpha = data[src + 3];
        const scale = premultiplied && alpha > 0 && alpha < 255 ? 255 / alpha : 1;
        rgba[dst] = data[src + 2] * scale;
        rgba[dst + 1] = data[src + 1] * scale;
        rgba[dst + 2] = data[src] * scale;
        rgba[dst + 3] = alpha;
      }
    }
    this.images.set(id, { width, height, data: rgba });
  }

  /** World → page → device pixels → target. Only UnitPixel pages are
   *  implemented (1 unit = 1 reference-device pixel). */
  private toTarget(x: number, y: number): [number, number] {
    const w = this.world;
    const dx = (w.a * x + w.c * y + w.e) * this.pageScale;
    const dy = (w.b * x + w.d * y + w.f) * this.pageScale;
    const t = this.target;
    return [((dx - t.left) * t.W) / t.boundsW, ((dy - t.top) * t.H) / t.boundsH];
  }

  /** [MS-EMFPLUS] 2.3.4.8 EmfPlusDrawImage: ImageAttributesID, SrcUnit,
   *  SrcRect (RectF), then the destination as RectF or, with flag C, Rect. */
  private drawImage(dv: DataView, r: PlusRecord): void {
    const compressed = (r.flags & 0x4000) !== 0;
    if (r.end - r.start < 24 + (compressed ? 8 : 16)) throw new RangeError('Truncated EMF+ DrawImage');
    const image = this.images.get(r.flags & 0xff);
    if (!image) {
      this.target.unsupported.add('EMF+ DrawImage of an unavailable image');
      return;
    }
    const srcUnit = dv.getInt32(r.start + 4, true);
    const f = (at: number) => dv.getFloat32(r.start + at, true);
    const [sx, sy, sw, sh] = [f(8), f(12), f(16), f(20)];
    const dest = compressed
      ? [0, 2, 4, 6].map((at) => dv.getInt16(r.start + 24 + at, true))
      : [f(24), f(28), f(32), f(36)];
    const [dx, dy, dw, dh] = dest;
    const axisAligned = Math.abs(this.world.b) < 1e-9 && Math.abs(this.world.c) < 1e-9;
    if (srcUnit !== UNIT_PIXEL || this.pageUnit !== UNIT_PIXEL || !axisAligned
      || !(dw > 0) || !(dh > 0) || !(sw > 0) || !(sh > 0)
      || this.world.a <= 0 || this.world.d <= 0) {
      this.target.unsupported.add('EMF+ DrawImage placement (units, rotation or mirroring)');
      return;
    }
    // Crop the source rectangle (pixel units) out of the decoded bitmap.
    const x0 = Math.max(0, Math.round(sx));
    const y0 = Math.max(0, Math.round(sy));
    const x1 = Math.min(image.width, Math.round(sx + sw));
    const y1 = Math.min(image.height, Math.round(sy + sh));
    if (x1 <= x0 || y1 <= y0) return;
    let source: Bitmap = image;
    if (x0 !== 0 || y0 !== 0 || x1 !== image.width || y1 !== image.height) {
      const width = x1 - x0;
      const data = new Uint8ClampedArray(width * (y1 - y0) * 4);
      for (let y = y0; y < y1; y++) {
        data.set(image.data.subarray((y * image.width + x0) * 4, (y * image.width + x1) * 4), (y - y0) * width * 4);
      }
      source = { width, height: y1 - y0, data };
    }
    const [tx0, ty0] = this.toTarget(dx, dy);
    const [tx1, ty1] = this.toTarget(dx + dw, dy + dh);
    const { ctx } = this.target;
    if (this.sourceCopy) {
      try {
        ctx.clearRect(Math.min(tx0, tx1), Math.min(ty0, ty1), Math.abs(tx1 - tx0), Math.abs(ty1 - ty0));
      } catch {
        /* a ctx without clearRect (some mocks) */
      }
    }
    if (blitDibToCtx(ctx, source, tx0, ty0, tx1, ty1)) this.target.drew = true;
  }
}
