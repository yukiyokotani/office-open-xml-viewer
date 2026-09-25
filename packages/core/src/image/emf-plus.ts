// EMF+ playback ([MS-EMFPLUS]) for the records GDI+ writes when it records a
// bitmap into a metafile: EmfPlusHeader, EmfPlusSetPageTransform, EmfPlusClear,
// EmfPlusSetCompositingMode, EmfPlusObject (ImageAttributes and uncompressed
// 32-bit bitmap Image objects, including continued objects), EmfPlusDrawImage,
// rendering-quality state records, world-transform records, Save/Restore,
// GetDC and EmfPlusEndOfFile.
//
// EMF+ records live in EMR_COMMENT records ("EMF+" identifier). A dual-mode
// metafile also carries a complete GDI rendering, and EMF+-aware players play
// one of the two. `scanEmfPlus` therefore decides up front, semantically: it
// plays the whole EMF+ stream through this same player in a dry run (every
// record, object, unit, transform, rectangle and image-attribute check, no
// drawing) and plays the EMF+ stream only when it contains drawing records
// and the dry run reports no failure. Structural damage counts as a failure
// exactly like an unsupported record: an EMF record or EMF+ comment whose
// size leaves its container, an EMF+ record whose Size/DataSize is invalid
// ([MS-EMFPLUS] 2.3 EmfPlusRecord: Size covers the 12-byte header and the
// data, DataSize ≤ Size − 12), a trailing fragment shorter than a record
// header, and a continued EmfPlusObject ([MS-EMFPLUS] 2.3.5.1 flag C) that
// is interleaved, inconsistent or never completed. Otherwise
//   - a dual-mode file keeps its complete GDI rendering, so a GDI alternative
//     is never discarded for EMF+ content that playback would reject or only
//     partly draw;
//   - an EMF+-only file has no complete alternative: the failures are
//     returned (`EmfPlusScan.failures`) for the caller to surface as an
//     incomplete picture, and the GDI records are played as they were before
//     EMF+ support existed (see emf.ts).
// Evidence for choosing the EMF+ rendering: Excel's PDF exports of GDI+
// metafiles whose GDI part has no drawing records (a bitmap drawn by
// EmfPlusDrawImage) show that bitmap; PowerPoint's PDF export of dual files
// whose EMF+ part has no drawing records shows the GDI drawing.

import { blitDibToCtx, type DecodedDib } from './dib.js';
import { HARD_MAX_DECODED_IMAGE_BYTES, MAX_RASTER_DIMENSION, MAX_RASTER_PIXELS } from './pixel-budget.js';

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
/** [MS-EMFPLUS] 2.1.1.34 WrapMode: Tile .. Clamp. */
const WRAP_MODE_MAX = 4;
// Resource policy, not a format limit (see pixel-budget.ts). One bitmap
// object obeys the shared per-surface limits (MAX_RASTER_DIMENSION per axis,
// MAX_RASTER_PIXELS in total), like the DIBs of GDI records. The player's
// whole decoded footprint — the RGBA of every Image object it retains, the
// assembly buffer of a continued object, and the intermediates of one
// DrawImage (the cropped source copy, the helper canvas backing store and its
// ImageData) — must fit HARD_MAX_DECODED_IMAGE_BYTES: exactly the four
// coexisting surfaces a single maximal bitmap needs to be retained, cropped
// and blitted. The dry run applies the same accounting, so a stream that
// exceeds it is rejected before a GDI alternative is given up.
const MAX_PLAYER_DECODED_BYTES = HARD_MAX_DECODED_IMAGE_BYTES;
const BUDGET_FAILURE = 'EMF+ bitmap (decoded-image budget)';

interface PlusRecord {
  type: number;
  flags: number;
  /** Absolute byte range of the record data. */
  start: number;
  end: number;
}

/** Walk the EMF+ records of one EMR_COMMENT (record starting at `pos`).
 *  A comment that is not an EMF+ comment yields nothing; structural damage in
 *  an EMF+ comment is reported through `fail` and ends the walk. */
function* plusRecords(
  dv: DataView,
  pos: number,
  recEnd: number,
  fail: (reason: string) => void,
): Generator<PlusRecord> {
  if (recEnd - pos < 16) return;
  if (dv.getUint32(pos + 12, true) !== 0x2b464d45) return; // 'EMF+'
  // EMR_COMMENT DataSize counts the identifier and the EMF+ records.
  const dataSize = dv.getUint32(pos + 8, true);
  if (dataSize < 4 || pos + 12 + dataSize > recEnd) {
    fail('EMF+ comment (invalid DataSize)');
    return;
  }
  const end = pos + 12 + dataSize;
  let p = pos + 16;
  while (p < end) {
    if (end - p < 12) {
      fail('EMF+ record (truncated header)');
      return;
    }
    const type = dv.getUint16(p, true);
    const flags = dv.getUint16(p + 2, true);
    const size = dv.getUint32(p + 4, true);
    const dataLength = dv.getUint32(p + 8, true);
    if (size < 12 || size > end - p || dataLength > size - 12) {
      fail(`EMF+ record 0x${type.toString(16)} (invalid Size or DataSize)`);
      return;
    }
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
  /** Why the EMF+ stream is not played: unsupported content and structural
   *  damage found by the dry run (empty when it is fully playable, and for a
   *  file without EMF+ comments). A file whose EMF+ comments are damaged
   *  before any EmfPlusHeader has failures with `present: false`. */
  failures: readonly string[];
}

/** Decide which of a metafile's renderings to play (see the module note). */
export function scanEmfPlus(bytes: Uint8Array): EmfPlusScan {
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let present = false;
  let seen = false;
  let dual = false;
  let drawing = false;
  const dry: EmfPlusTarget = {
    ctx: null,
    W: 1,
    H: 1,
    left: 0,
    top: 0,
    boundsW: 1,
    boundsH: 1,
    drew: false,
    unsupported: new Set(),
  };
  const player = new EmfPlusPlayer(dry, true);
  const ignore = () => {};
  let pos = 0;
  while (pos + 8 <= bytes.length) {
    const type = dv.getUint32(pos, true);
    const size = dv.getUint32(pos + 4, true);
    // The EMF record walk of playEmf: a 4-aligned size inside the file.
    if (size < 8 || (size & 3) !== 0 || pos + size > bytes.length) {
      dry.unsupported.add('EMF record stream (truncated or invalid record size)');
      break;
    }
    if (type === 70) {
      if (size >= 16 && dv.getUint32(pos + 12, true) === 0x2b464d45) seen = true;
      // Structural failures are recorded once, by the dry-run player below.
      for (const record of plusRecords(dv, pos, pos + size, ignore)) {
        if (record.type === PLUS.HEADER) {
          present = true;
          dual = (record.flags & 1) !== 0;
        }
        if (DRAWING.has(record.type)) drawing = true;
      }
      player.playComment(dv, pos, pos + size);
    }
    if (type === 14) break;
    pos += size;
  }
  player.finish();
  // Failures count once the file carries EMF+ at all: a damaged first comment
  // can hide the EmfPlusHeader, and then there is no Flags bit asserting that
  // the GDI records are a complete alternative.
  const failures = seen ? [...dry.unsupported] : [];
  return { present, dual, play: present && drawing && failures.length === 0, failures };
}

/** The subset of the GDI player state the EMF+ player draws through. */
export interface EmfPlusTarget {
  /** `null` only for the dry run of `scanEmfPlus`. */
  ctx: CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D | null;
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

/** `m1` then `m2` ([MS-EMFPLUS] 2.2.2.47 EmfPlusTransformMatrix order). */
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

/** The EMF+ Object Table entry kinds this player keeps ([MS-EMFPLUS] 2.3.5.1:
 *  a later object with the same index replaces the entry). */
type TableObject =
  | { kind: 'image'; bitmap: Bitmap }
  | { kind: 'attributes'; wrapMode: number };

/** The implemented graphics state saved by EmfPlusSave ([MS-EMFPLUS] 2.3.7.5). */
interface GraphicsState {
  world: Matrix;
  unit: number;
  scale: number;
  sourceCopy: boolean;
}

export class EmfPlusPlayer {
  private world: Matrix = IDENTITY;
  private pageUnit = UNIT_PIXEL;
  private pageScale = 1;
  private sourceCopy = false;
  private readonly objects = new Map<number, TableObject>();
  private readonly saved = new Map<number, GraphicsState>();
  private pending: { id: number; total: number; parts: Uint8Array[]; length: number } | null = null;
  /** Inside EmfPlusGetDC until the next EMF+ record: GDI records draw. */
  gdiAllowed = false;
  /** A record the dry run admitted failed in real playback (the failure is
   *  also in the target's report), e.g. a refused blit surface. The dry run
   *  cannot see such a failure, so the GDI player decides what replaces the
   *  EMF+ rendering (see emf.ts). */
  drawFailed = false;

  /** `dry`: validate everything playback checks, decode and draw nothing. */
  constructor(private readonly target: EmfPlusTarget, private readonly dry = false) {}

  /** Play the EMF+ records of one EMR_COMMENT record. */
  playComment(dv: DataView, pos: number, recEnd: number): void {
    const fail = (reason: string) => this.target.unsupported.add(reason);
    for (const record of plusRecords(dv, pos, recEnd, fail)) {
      this.gdiAllowed = false;
      // [MS-EMFPLUS] 2.3.5.1: the fragments of a continued object are
      // consecutive EmfPlusObject records; anything else in between leaves
      // the pending object unfinished.
      if (this.pending && !(record.type === PLUS.OBJECT && (record.flags & 0x8000) !== 0)) {
        this.abandonPending();
      }
      const reported = this.target.unsupported.size;
      try {
        this.play(dv, record);
      } catch {
        this.target.unsupported.add(`EMF+ record 0x${record.type.toString(16)} (malformed)`);
      }
      if (!this.dry && this.target.unsupported.size > reported) this.drawFailed = true;
    }
  }

  /** End of the metafile: a continued object that never completed is a
   *  failure, not a silently missing object. */
  finish(): void {
    if (this.pending) this.abandonPending();
  }

  private abandonPending(): void {
    this.target.unsupported.add('EMF+ continued object (unfinished)');
    this.pending = null;
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
          sourceCopy: this.sourceCopy,
        });
        return;
      case PLUS.RESTORE: {
        need(4);
        const state = this.saved.get(dv.getUint32(r.start, true));
        if (!state) {
          this.target.unsupported.add('EMF+ Restore of an unsaved graphics state');
          return;
        }
        this.world = state.world;
        this.pageUnit = state.unit;
        this.pageScale = state.scale;
        this.sourceCopy = state.sourceCopy;
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
    if (this.dry || !ctx) return;
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
    // Bytes of an assembled continued object, alive while it is decoded.
    let assembled = 0;
    if (r.flags & 0x8000) {
      // Continued: TotalObjectSize, then this fragment.
      if (data.length < 4) throw new RangeError('Truncated EMF+ object');
      const total = dv.getUint32(r.start, true);
      if (this.pending && (this.pending.id !== id || this.pending.total !== total)) {
        // A different object (or a changed TotalObjectSize) before the
        // pending one completed.
        this.abandonPending();
      }
      if (!this.pending) {
        // The fragments are views of the file; only the assembled copy
        // allocates, so it is admitted before the first fragment is kept.
        if (!this.admits(total)) {
          this.target.unsupported.add(BUDGET_FAILURE);
          this.objects.delete(id);
          return;
        }
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
      assembled = whole.length;
    }
    const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
    this.objects.delete(id);
    if (kind === OBJECT_IMAGE_ATTRIBUTES) {
      // EmfPlusImageAttributes (2.2.1.5): Version, Reserved1, WrapMode,
      // ClampColor, ObjectClamp, Reserved2. It serializes no colour
      // adjustment; WrapMode, ClampColor and ObjectClamp only govern samples
      // outside the image, which DrawImage below never admits.
      if (data.length < 24) throw new RangeError('Truncated EMF+ image attributes');
      const wrapMode = view.getUint32(8, true);
      if (wrapMode > WRAP_MODE_MAX || view.getUint32(16, true) > 1) {
        this.target.unsupported.add('EMF+ image attributes (invalid wrap or clamp)');
        return;
      }
      this.objects.set(id, { kind: 'attributes', wrapMode });
      return;
    }
    if (kind !== OBJECT_IMAGE) {
      this.target.unsupported.add(`EMF+ object type ${kind}`);
      return;
    }
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
    if (width <= 0 || height <= 0 || stride < width * 4 || 28 + stride * height > data.length) {
      throw new RangeError('Invalid EMF+ bitmap');
    }
    if (width > MAX_RASTER_DIMENSION || height > MAX_RASTER_DIMENSION
      || width * height > MAX_RASTER_PIXELS
      || !this.admits(assembled + width * height * 4, id)) {
      this.target.unsupported.add(BUDGET_FAILURE);
      return;
    }
    if (this.dry) {
      this.objects.set(id, { kind: 'image', bitmap: { width, height, data: new Uint8ClampedArray(0) } });
      return;
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
    this.objects.set(id, { kind: 'image', bitmap: { width, height, data: rgba } });
  }

  /** Decoded bytes the object table retains, without the entry `except`
   *  (about to be replaced). The dry run counts the bitmaps it only sized. */
  private retainedBytes(except?: number): number {
    let bytes = 0;
    for (const [id, entry] of this.objects) {
      if (id !== except && entry.kind === 'image') bytes += entry.bitmap.width * entry.bitmap.height * 4;
    }
    return bytes;
  }

  /** Whether `transient` more decoded bytes fit MAX_PLAYER_DECODED_BYTES next
   *  to what the table retains (see the resource note at the top). */
  private admits(transient: number, replacing?: number): boolean {
    return this.retainedBytes(replacing) + transient <= MAX_PLAYER_DECODED_BYTES;
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
    const entry = this.objects.get(r.flags & 0xff);
    if (entry?.kind !== 'image') {
      this.target.unsupported.add('EMF+ DrawImage of an unavailable image');
      return;
    }
    const image = entry.bitmap;
    // ImageAttributesID names the EmfPlusImageAttributes object of the draw.
    if (this.objects.get(dv.getUint32(r.start, true))?.kind !== 'attributes') {
      this.target.unsupported.add('EMF+ DrawImage without its image attributes');
      return;
    }
    const srcUnit = dv.getInt32(r.start + 4, true);
    const f = (at: number) => dv.getFloat32(r.start + at, true);
    const [sx, sy, sw, sh] = [f(8), f(12), f(16), f(20)];
    const dest = compressed
      ? [0, 2, 4, 6].map((at) => dv.getInt16(r.start + 24 + at, true))
      : [f(24), f(28), f(32), f(36)];
    const [dx, dy, dw, dh] = dest;
    // Every number that places the image must be finite: the float fields of
    // SrcRect/RectF, the world matrix, PageScale ([MS-EMFPLUS] 2.3.9.3) and
    // the mapped corners (a finite chain of transforms can still overflow).
    // A NaN or infinite placement draws nothing a player can reproduce, so
    // the dry run rejects it and a GDI alternative is kept.
    const w = this.world;
    const [tx0, ty0] = this.toTarget(dx, dy);
    const [tx1, ty1] = this.toTarget(dx + dw, dy + dh);
    if (![sx, sy, sw, sh, dx, dy, dw, dh, w.a, w.b, w.c, w.d, w.e, w.f, this.pageScale, tx0, ty0, tx1, ty1]
      .every(Number.isFinite)) {
      this.target.unsupported.add('EMF+ DrawImage placement (non-finite value)');
      return;
    }
    const axisAligned = Math.abs(w.b) < 1e-9 && Math.abs(w.c) < 1e-9;
    if (srcUnit !== UNIT_PIXEL || this.pageUnit !== UNIT_PIXEL || !axisAligned
      || !(dw > 0) || !(dh > 0) || !(sw > 0) || !(sh > 0)
      || w.a <= 0 || w.d <= 0 || !(this.pageScale > 0)) {
      this.target.unsupported.add('EMF+ DrawImage placement (units, rotation or mirroring)');
      return;
    }
    // A source rectangle reaching outside the image samples the wrap mode
    // and clamp colour (2.1.1.34), which are not drawn here.
    const x0 = Math.round(sx);
    const y0 = Math.round(sy);
    const x1 = Math.round(sx + sw);
    const y1 = Math.round(sy + sh);
    if (x0 < 0 || y0 < 0 || x1 > image.width || y1 > image.height || x1 <= x0 || y1 <= y0) {
      this.target.unsupported.add('EMF+ DrawImage source outside the image');
      return;
    }
    // Intermediates of the blit: the cropped copy (when cropping), then the
    // helper canvas backing store and its ImageData (dib.ts blitDibToCtx).
    const cropped = x0 !== 0 || y0 !== 0 || x1 !== image.width || y1 !== image.height;
    const sourceBytes = (x1 - x0) * (y1 - y0) * 4;
    if (!this.admits((cropped ? sourceBytes : 0) + 2 * sourceBytes)) {
      this.target.unsupported.add(BUDGET_FAILURE);
      return;
    }
    if (this.dry) return;
    let source: Bitmap = image;
    if (cropped) {
      const width = x1 - x0;
      const data = new Uint8ClampedArray(width * (y1 - y0) * 4);
      for (let y = y0; y < y1; y++) {
        data.set(image.data.subarray((y * image.width + x0) * 4, (y * image.width + x1) * 4), (y - y0) * width * 4);
      }
      source = { width, height: y1 - y0, data };
    }
    const { ctx } = this.target;
    if (!ctx) return;
    if (this.sourceCopy) {
      try {
        ctx.clearRect(Math.min(tx0, tx1), Math.min(ty0, ty1), Math.abs(tx1 - tx0), Math.abs(ty1 - ty0));
      } catch {
        /* a ctx without clearRect (some mocks) */
      }
    }
    if (blitDibToCtx(ctx, source, tx0, ty0, tx1, ty1)) {
      this.target.drew = true;
    } else {
      // No helper surface or a refused blit: the image is missing, which the
      // result must say (the GDI rendering was set aside for this stream).
      this.target.unsupported.add('EMF+ DrawImage (the bitmap could not be drawn)');
    }
  }
}
