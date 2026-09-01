// ── WMF (Windows Metafile) player ───────────────────────────────────────────
//
// Browsers cannot decode WMF/EMF via `createImageBitmap`, so the renderer falls
// back to this player for metafile blips. It is a *minimal* WMF interpreter:
// just enough to rasterize the vector-graphics metafiles that Office embeds for
// charts and diagrams (e.g. sample-10.docx `word/media/image1.emf`, which —
// despite the `.emf` extension — is a standard non-placeable WMF whose labels
// are POLYPOLYGON glyph outlines, not text-out records).
//
// Format reference: ECMA-376 references WMF/EMF; the byte layout below follows
// the [MS-WMF] Windows Metafile Format spec.
//   - Header: 18 bytes (u16 type, u16 headerSizeWords=9, u16 version, u32
//     fileSizeWords, u16 numObjects, u32 maxRecordWords, u16 numMembers).
//   - Record: u32 recordSizeWords (INCLUDING the 6-byte size+function header,
//     counted in 16-bit words), u16 function, then (recordSizeWords*2 − 6)
//     param bytes. Loop until function==0x0000 (META_EOF) or bytes exhausted.
//   - All values little-endian. COLORREF = u32 0x00BBGGRR (low byte = R).
//
// Implemented records: SETWINDOWORG, SETWINDOWEXT, SETPOLYFILLMODE,
// SETTEXTCOLOR, SETTEXTALIGN, CREATEPENINDIRECT, CREATEBRUSHINDIRECT,
// CREATEFONTINDIRECT, SELECTOBJECT, DELETEOBJECT, POLYLINE, POLYGON,
// POLYPOLYGON, RECTANGLE, TEXTOUT, STRETCHDIBITS (embedded raster DIB via the
// shared decoder in ./dib.ts), EOF.
// Ignored (no-op, skipped by size): ESCAPE, SETROP2, SETBKMODE,
// SETSTRETCHBLTMODE, SETMAPMODE, DIBBITBLT/DIBSTRETCHBLT (their exact param
// layout is not decoded here — skipped rather than mis-parsed), and any
// unrecognized record.
//
// Shared across the docx, pptx and xlsx renderers (originally docx-only). The
// device-boundary edge-suppression heuristic (see `deviceInteriorEdges`) is a
// docx-specific cosmetic-frame workaround, so it is gated behind the
// `suppressBoundaryFrame` option (default OFF = spec-clean, draws every edge).
// docx will opt IN (`suppressBoundaryFrame: true`) when it re-points to core
// (deferred), to preserve its current behavior.
//
// True EMF (a separate, larger 32-bit format — see {@link isEmf}) is rasterized
// by the sibling player {@link ./emf.ts}#renderEmfToBitmap, routed from
// {@link decodeRasterOrMetafile}.

import { decodePackedDib, blitDibToCtx } from './dib.js';
import { renderEmfToBitmap } from './emf.js';
import {
  rasterExceedsBudget,
  sniffRasterDimensions,
  sourceRasterExceedsBudget,
  type RasterDimensions,
} from './raster-dimensions.js';
import {
  MAX_RASTER_DIMENSION,
  MAX_RASTER_PIXELS,
  MAX_RASTER_SOURCE_DIMENSION,
  MAX_RASTER_SOURCE_PIXELS,
  OoxmlDecodedImageLimitError,
} from './pixel-budget.js';
import { createAuxCanvas } from '../canvas/aux-canvas.js';
import { closeImageBitmapIfSupported } from './image-bitmap-lifecycle.js';
import { isTiff, type TiffRenderer } from './tiff-contract.js';

// WMF record function codes (the subset we act on; others are skipped by size).
const META = {
  EOF: 0x0000,
  SETBKMODE: 0x0102,
  SETTEXTALIGN: 0x012e,
  SETTEXTCOLOR: 0x0209,
  SETPOLYFILLMODE: 0x0106,
  SETWINDOWORG: 0x020b,
  SETWINDOWEXT: 0x020c,
  SELECTOBJECT: 0x012d,
  DELETEOBJECT: 0x01f0,
  TEXTOUT: 0x0521,
  POLYGON: 0x0324,
  POLYLINE: 0x0325,
  POLYPOLYGON: 0x0538,
  RECTANGLE: 0x041b,
  CREATEPENINDIRECT: 0x02fa,
  CREATEFONTINDIRECT: 0x02fb,
  CREATEBRUSHINDIRECT: 0x02fc,
  DIBBITBLT: 0x0940,
  DIBSTRETCHBLT: 0x0b41,
  STRETCHDIBITS: 0x0f43,
} as const;

const PLACEABLE_MAGIC = 0x9ac6cdd7; // little-endian bytes D7 CD C6 9A
const PLACEABLE_HEADER_BYTES = 22;
const WMF_HEADER_BYTES = 18;
const EMF_SIGNATURE = 0x464d4520; // " EMF" (bytes 20 45 4D 46)

// ── detection ───────────────────────────────────────────────────────────────

/** Reads the standard 18-byte WMF header at `off` and validates type∈{1,2} and
 *  headerSize==9 words. */
function looksLikeStandardHeader(b: Uint8Array, off: number): boolean {
  if (b.length < off + WMF_HEADER_BYTES) return false;
  const type = b[off] | (b[off + 1] << 8);
  const headerSize = b[off + 2] | (b[off + 3] << 8);
  return (type === 1 || type === 2) && headerSize === 9;
}

/** True for a placeable (`D7 CD C6 9A`) or standard (type∈{1,2}, headerSize==9)
 *  WMF. A placeable file prepends a 22-byte header before the standard one. */
export function isWmf(bytes: Uint8Array): boolean {
  if (bytes.length < 4) return false;
  const magic =
    bytes[0] | (bytes[1] << 8) | (bytes[2] << 16) | (bytes[3] << 24);
  if ((magic >>> 0) === PLACEABLE_MAGIC) {
    // Be lenient: a placeable file is a WMF even if we can't re-validate the
    // inner header (some toolchains emit slightly off inner fields).
    return true;
  }
  return looksLikeStandardHeader(bytes, 0);
}

/** True for a true EMF (ENHMETAHEADER): u32@0 == 1 (EMR_HEADER) AND u32@40 ==
 *  0x464D4520 (" EMF"). True EMF is a different, larger 32-bit format than WMF;
 *  it is rasterized by the sibling player {@link ./emf.ts}#renderEmfToBitmap. */
export function isEmf(bytes: Uint8Array): boolean {
  if (bytes.length < 44) return false;
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  return dv.getUint32(0, true) === 1 && dv.getUint32(40, true) === EMF_SIGNATURE;
}

/** True for the WMF/EMF metafile MIME types (`image/wmf`, `image/emf`). Used by
 *  the shared `<a:srcRect>` crop (ECMA-376 §20.1.8.55): a metafile carries no
 *  native pixel grid, so a cropped one must be rasterized at its FULL picture
 *  frame before the fractional crop applies (see {@link ./crop.ts}'s
 *  `metafileRasterSize`); raster blips skip that scale-up. The MIME is
 *  extension-derived by the parsers (`mime_from_ext`), so a metafile mislabeled
 *  with a raster extension would be treated as a raster — acceptable because
 *  authored crops on metafiles are rare. */
export function isMetafileMime(mime: string | undefined): boolean {
  return mime === 'image/wmf' || mime === 'image/emf';
}

// ── color ─────────────────────────────────────────────────────────────────

/** COLORREF (u32 0x00BBGGRR) → CSS `#rrggbb`. Shared with the EMF player
 *  ({@link ./emf.ts}); COLORREF has the same byte layout in [MS-EMF]. */
export function colorRefToCss(c: number): string {
  const r = c & 0xff;
  const g = (c >>> 8) & 0xff;
  const b = (c >>> 16) & 0xff;
  const hex = (n: number) => n.toString(16).padStart(2, '0');
  return `#${hex(r)}${hex(g)}${hex(b)}`;
}

// ── object table ─────────────────────────────────────────────────────────

interface Pen {
  kind: 'pen';
  stroke: string | null; // null = PS_NULL (no stroke)
  width: number; // device-independent logical width; mapped to ≥1 device px
}
interface Brush {
  kind: 'brush';
  fill: string | null; // null = BS_NULL / hollow (no fill)
}
interface Font {
  kind: 'font';
  height: number; // |lfHeight|, logical units
  weight: number; // lfWeight (400 normal, 700 bold)
  italic: boolean;
  face: string;
}
type WmfObject = Pen | Brush | Font;

/** Inserts an object at the FIRST free slot (lowest index whose slot is empty
 *  or was deleted), mirroring the WMF object-table allocation rule. */
function insertObject(table: (WmfObject | null)[], obj: WmfObject): void {
  for (let i = 0; i < table.length; i++) {
    if (table[i] == null) {
      table[i] = obj;
      return;
    }
  }
  table.push(obj);
}

// ── little-endian cursor over the param region of one record ───────────────

class Cursor {
  private p = 0;
  constructor(
    private readonly b: Uint8Array,
    start: number,
    private readonly end: number, // exclusive
  ) {
    this.p = start;
  }
  get remaining(): number {
    return this.end - this.p;
  }
  i16(): number {
    const v = this.u16();
    return v >= 0x8000 ? v - 0x10000 : v;
  }
  u16(): number {
    const v = this.b[this.p] | (this.b[this.p + 1] << 8);
    this.p += 2;
    return v;
  }
  u8(): number {
    return this.b[this.p++];
  }
  u32(): number {
    const v =
      (this.b[this.p] |
        (this.b[this.p + 1] << 8) |
        (this.b[this.p + 2] << 16) |
        (this.b[this.p + 3] << 24)) >>>
      0;
    this.p += 4;
    return v;
  }
  bytes(n: number): Uint8Array {
    const end = Math.min(this.p + Math.max(0, n), this.end);
    const out = this.b.subarray(this.p, end);
    this.p = end;
    return out;
  }
  skip(n: number): void {
    this.p = Math.min(this.p + Math.max(0, n), this.end);
  }
}

// ── any 2D context we can replay onto (Offscreen or HTMLCanvas) ─────────────

type AnyCtx = CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;

interface PlayState {
  ctx: AnyCtx;
  W: number;
  H: number;
  // window mapping
  orgX: number;
  orgY: number;
  extX: number;
  extY: number;
  haveExt: boolean;
  // GDI state
  objects: (WmfObject | null)[];
  curPen: Pen | null;
  curBrush: Brush | null;
  curFont: Font | null;
  textColor: string;
  textAlign: number;
  fillRule: CanvasFillRule; // from SETPOLYFILLMODE
  drew: boolean;
  // When true, strokes whose edge lies on the window/device boundary are
  // suppressed (docx cosmetic-frame heuristic; see deviceInteriorEdges). When
  // false (spec-clean default), every edge is drawn.
  suppressBoundaryFrame: boolean;
}

/** logical → device along X. */
function mapX(s: PlayState, x: number): number {
  return (x - s.orgX) * (s.W / s.extX);
}
/** logical → device along Y. */
function mapY(s: PlayState, y: number): number {
  return (y - s.orgY) * (s.H / s.extY);
}

/** Device line width: scale the logical pen width by |W/extX| and clamp to ≥1
 *  so hairlines stay visible after mapping. */
function deviceLineWidth(s: PlayState, logicalWidth: number): number {
  const scale = Math.abs(s.W / s.extX);
  const w = logicalWidth * scale;
  return w >= 1 ? w : 1;
}

// Tolerance (device px) for treating a mapped coordinate as lying ON a device
// boundary line. The window→device mapping is a float multiply, so an edge that
// is logically on the window frame can land at e.g. 0.0000001 or W-0.0000002.
const BOUNDARY_EPS = 1e-3;

/** Whether a device coordinate lies on either of the two boundary lines `lo`/`hi`
 *  (within BOUNDARY_EPS). Used to detect polygon edges that coincide with the
 *  metafile window/device boundary. */
function onBoundaryLine(v: number, lo: number, hi: number): boolean {
  return Math.abs(v - lo) <= BOUNDARY_EPS || Math.abs(v - hi) <= BOUNDARY_EPS;
}

/**
 * HEURISTIC (not a clean ECMA/GDI rule): a cosmetic (hairline) stroke whose edge
 * coincides with the metafile window/device boundary (x∈{0,W} or y∈{0,H}) is
 * suppressed, because Word does not render a visible frame there — the common
 * case being the full-window "frame rectangle" many authoring tools emit with a
 * 1-device-pixel cosmetic pen.
 *
 * This is NOT GDI's actual clip behavior: GDI's `Rectangle` is half-open and
 * excludes only the RIGHT and BOTTOM edges, not all four; here we drop edges on
 * ALL four boundary lines (x=0, x=W, y=0, y=H). That over-broad drop is the
 * pragmatic point — it removes the cosmetic frame — but it may also drop a
 * legitimate full-window border.
 * TODO: replace with a proper WMF window/viewport clip model ([MS-WMF]); tracked
 * as a follow-up.
 *
 * Returns the list of stroke segments to draw — every polygon edge EXCEPT those
 * whose two endpoints both lie on one boundary line (a vertical edge on x=0 or
 * x=W, or a horizontal edge on y=0 or y=H). A rectangle one pixel inside the
 * surface keeps all four edges; only edges literally on the boundary are dropped.
 */
function deviceInteriorEdges(
  s: PlayState,
  pts: Array<[number, number]>,
  closed: boolean,
): Array<[[number, number], [number, number]]> {
  const segs: Array<[[number, number], [number, number]]> = [];
  const last = closed ? pts.length : pts.length - 1;
  for (let i = 0; i < last; i++) {
    const a = pts[i];
    const b = pts[(i + 1) % pts.length];
    // A vertical edge sitting on x=0 or x=W, or a horizontal edge on y=0 or y=H,
    // lies on the window/device boundary → suppressed (see the heuristic above).
    const verticalOnBoundary =
      Math.abs(a[0] - b[0]) <= BOUNDARY_EPS &&
      onBoundaryLine(a[0], 0, s.W) &&
      onBoundaryLine(b[0], 0, s.W);
    const horizontalOnBoundary =
      Math.abs(a[1] - b[1]) <= BOUNDARY_EPS &&
      onBoundaryLine(a[1], 0, s.H) &&
      onBoundaryLine(b[1], 0, s.H);
    if (verticalOnBoundary || horizontalOnBoundary) continue;
    segs.push([a, b]);
  }
  return segs;
}

/** Stroke a polygon/polyline's edges with the current pen. When
 *  `suppressBoundaryFrame` is set, applies the window/device-boundary
 *  suppression heuristic (see deviceInteriorEdges): edges coincident with the
 *  boundary are dropped, and consecutive KEPT edges are emitted as one
 *  continuous sub-path (a dropped edge breaks the chain). When it is NOT set
 *  (the spec-clean default), the full polygon/polyline strokes as a single path
 *  with no edge dropped. */
function strokeEdges(s: PlayState, pts: Array<[number, number]>, closed: boolean): void {
  if (!s.curPen || s.curPen.stroke == null) return;
  if (pts.length < 2) return;
  const { ctx } = s;
  ctx.strokeStyle = s.curPen.stroke;
  ctx.lineWidth = deviceLineWidth(s, s.curPen.width);

  if (!s.suppressBoundaryFrame) {
    // Spec-clean path: draw every edge as one continuous sub-path.
    ctx.beginPath();
    ctx.moveTo(pts[0][0], pts[0][1]);
    for (let i = 1; i < pts.length; i++) ctx.lineTo(pts[i][0], pts[i][1]);
    if (closed) ctx.closePath();
    ctx.stroke();
    s.drew = true;
    return;
  }

  const segs = deviceInteriorEdges(s, pts, closed);
  if (segs.length === 0) return;
  ctx.beginPath();
  // Chain consecutive edges that share an endpoint into one sub-path; restart
  // when the previous edge was dropped (endpoints no longer contiguous).
  let prevEnd: [number, number] | null = null;
  for (const [a, b] of segs) {
    if (!prevEnd || prevEnd[0] !== a[0] || prevEnd[1] !== a[1]) {
      ctx.moveTo(a[0], a[1]);
    }
    ctx.lineTo(b[0], b[1]);
    prevEnd = b;
  }
  ctx.stroke();
  s.drew = true;
}

// ── per-record handlers ─────────────────────────────────────────────────────

function readPoints(s: PlayState, c: Cursor, count: number): Array<[number, number]> {
  const pts: Array<[number, number]> = [];
  for (let i = 0; i < count; i++) {
    if (c.remaining < 4) break; // malformed; bail with what we have
    const x = c.i16();
    const y = c.i16();
    pts.push([mapX(s, x), mapY(s, y)]);
  }
  return pts;
}

function strokePolyline(s: PlayState, pts: Array<[number, number]>): void {
  if (pts.length < 2 || !s.curPen || s.curPen.stroke == null) return;
  // Open polyline: routed through the same edge stroker as polygons so the
  // boundary-suppression heuristic applies when enabled (see deviceInteriorEdges).
  strokeEdges(s, pts, false);
}

/** Fill (current brush) + stroke (current pen) a single closed polygon. The
 *  stroke routes through {@link strokeEdges}, which applies the
 *  window/device-boundary suppression heuristic only when
 *  `suppressBoundaryFrame` is set. The FILL is unaffected — a brush-filled
 *  rectangle at the window bounds still fills the surface. */
function fillStrokePolygon(s: PlayState, pts: Array<[number, number]>): void {
  if (pts.length < 2) return;
  const { ctx } = s;
  if (s.curBrush && s.curBrush.fill != null) {
    ctx.beginPath();
    ctx.moveTo(pts[0][0], pts[0][1]);
    for (let i = 1; i < pts.length; i++) ctx.lineTo(pts[i][0], pts[i][1]);
    ctx.closePath();
    ctx.fillStyle = s.curBrush.fill;
    ctx.fill(s.fillRule);
    s.drew = true;
  }
  strokeEdges(s, pts, true);
}

/** POLYPOLYGON: one path spanning every sub-polygon so the fill rule applies
 *  across them as a unit (correct glyph holes). */
function fillStrokePolyPolygon(s: PlayState, c: Cursor): void {
  const numPolys = c.u16();
  if (numPolys <= 0 || numPolys > 0x10000) return;
  const counts: number[] = [];
  for (let i = 0; i < numPolys; i++) {
    if (c.remaining < 2) return;
    counts.push(c.u16());
  }
  const { ctx } = s;
  ctx.beginPath();
  let any = false;
  for (const count of counts) {
    if (count < 2) {
      // still consume this sub-poly's points to stay aligned
      for (let i = 0; i < count && c.remaining >= 4; i++) {
        c.i16();
        c.i16();
      }
      continue;
    }
    const pts = readPoints(s, c, count);
    if (pts.length < 2) continue;
    ctx.moveTo(pts[0][0], pts[0][1]);
    for (let i = 1; i < pts.length; i++) ctx.lineTo(pts[i][0], pts[i][1]);
    ctx.closePath();
    any = true;
  }
  if (!any) return;
  if (s.curBrush && s.curBrush.fill != null) {
    ctx.fillStyle = s.curBrush.fill;
    ctx.fill(s.fillRule);
    s.drew = true;
  }
  if (s.curPen && s.curPen.stroke != null) {
    ctx.strokeStyle = s.curPen.stroke;
    ctx.lineWidth = deviceLineWidth(s, s.curPen.width);
    ctx.stroke();
    s.drew = true;
  }
}

function createPen(c: Cursor): Pen {
  const style = c.u16();
  const widthX = c.i16();
  c.i16(); // widthY (unused — WMF pens are isotropic in practice)
  const color = c.u32();
  const lowStyle = style & 0xff;
  // PS_NULL (5) → no stroke. Dash/dot styles (1..4) are rendered solid (we do
  // not synthesize a dash pattern); see module note.
  const stroke = lowStyle === 5 ? null : colorRefToCss(color);
  return { kind: 'pen', stroke, width: Math.abs(widthX) };
}

function createBrush(c: Cursor): Brush {
  const style = c.u16();
  const color = c.u32();
  c.u16(); // hatch (HATCHED brushes are rendered as solid fills)
  // BS_NULL / BS_HOLLOW (1) → no fill. SOLID (0) and HATCHED (2) fill solid.
  const fill = style === 1 ? null : colorRefToCss(color);
  return { kind: 'brush', fill };
}

function decodeSingleByteText(bytes: Uint8Array): string {
  const end = bytes.indexOf(0);
  const body = end >= 0 ? bytes.subarray(0, end) : bytes;
  if (body.length === 0) return '';
  try {
    return new TextDecoder('shift_jis').decode(body);
  } catch {
    return String.fromCharCode(...body);
  }
}

function createFont(c: Cursor): Font {
  const height = Math.abs(c.i16());
  c.i16(); // lfWidth
  c.i16(); // lfEscapement (rotation is not replayed by the WMF player yet)
  c.i16(); // lfOrientation
  const weight = c.i16();
  const italic = c.u8() !== 0;
  c.u8(); // lfUnderline
  c.u8(); // lfStrikeOut
  c.u8(); // lfCharSet
  c.u8(); // lfOutPrecision
  c.u8(); // lfClipPrecision
  c.u8(); // lfQuality
  c.u8(); // lfPitchAndFamily
  const face = decodeSingleByteText(c.bytes(Math.min(32, c.remaining)));
  return { kind: 'font', height, weight, italic, face };
}

function drawTextOut(s: PlayState, text: string, x: number, y: number): void {
  if (text.length === 0) return;
  const font = s.curFont;
  const logicalHeight = font?.height || 12;
  const px = Math.abs(mapY(s, s.orgY + logicalHeight) - mapY(s, s.orgY));
  if (!Number.isFinite(px) || px < 1) return;
  const { ctx } = s;
  try {
    ctx.fillStyle = s.textColor;
    const weight = font && font.weight >= 700 ? 'bold ' : '';
    const italic = font?.italic ? 'italic ' : '';
    ctx.font = `${italic}${weight}${px}px ${font?.face || 'sans-serif'}`;
    const horiz = s.textAlign & 0x6;
    ctx.textAlign = horiz === 0x2 ? 'right' : horiz === 0x6 ? 'center' : 'left';
    ctx.textBaseline = (s.textAlign & 0x18) === 0x18 ? 'alphabetic' : 'top';
    ctx.fillText(text, mapX(s, x), mapY(s, y));
    s.drew = true;
  } catch {
    // Some test or server-side contexts may not implement fillText.
  }
}

// ── core record-replay loop (pure; testable with a mock ctx) ────────────────

/**
 * Replay a WMF byte buffer onto a 2D context, mapping logical coordinates to a
 * `W`×`H` device space. Returns `true` if anything was drawn, `false` for a
 * non-WMF buffer or a metafile that produced no geometry.
 *
 * `suppressBoundaryFrame` (default false) enables the docx cosmetic-frame
 * heuristic (see deviceInteriorEdges); leave it false for a spec-clean render.
 *
 * Pure with respect to the injected `ctx`, so it is unit-testable against a
 * recording mock — no OffscreenCanvas required.
 */
export function playWmf(
  bytes: Uint8Array,
  ctx: AnyCtx,
  W: number,
  H: number,
  suppressBoundaryFrame = false,
): boolean {
  if (!isWmf(bytes)) return false;

  // Skip the placeable header if present, then the 18-byte standard header.
  let base = 0;
  const magic =
    bytes.length >= 4
      ? (bytes[0] | (bytes[1] << 8) | (bytes[2] << 16) | (bytes[3] << 24)) >>> 0
      : 0;
  if (magic === PLACEABLE_MAGIC) base = PLACEABLE_HEADER_BYTES;
  let pos = base + WMF_HEADER_BYTES;
  if (pos > bytes.length) return false;

  const s: PlayState = {
    ctx,
    W,
    H,
    orgX: 0,
    orgY: 0,
    extX: W || 1,
    extY: H || 1,
    haveExt: false,
    objects: [],
    curPen: null,
    curBrush: null,
    curFont: null,
    textColor: '#000000',
    textAlign: 0,
    fillRule: 'nonzero',
    drew: false,
    suppressBoundaryFrame,
  };

  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);

  while (pos + 6 <= bytes.length) {
    const sizeWords = dv.getUint32(pos, true);
    const fn = dv.getUint16(pos + 4, true);
    // Validate: a record is ≥3 words (u32 size + u16 fn) and in-bounds.
    if (sizeWords < 3) break;
    const recordBytes = sizeWords * 2;
    const recEnd = pos + recordBytes;
    if (recEnd > bytes.length) break; // truncated/malformed → partial render
    if (fn === META.EOF) break;

    const paramStart = pos + 6;
    const c = new Cursor(bytes, paramStart, recEnd);

    switch (fn) {
      case META.SETWINDOWORG: {
        // params: i16 yOrg, i16 xOrg (Y FIRST)
        s.orgY = c.i16();
        s.orgX = c.i16();
        break;
      }
      case META.SETWINDOWEXT: {
        // params: i16 yExt, i16 xExt (Y FIRST)
        const yExt = c.i16();
        const xExt = c.i16();
        s.extY = yExt || 1;
        s.extX = xExt || 1;
        s.haveExt = true;
        break;
      }
      case META.SETPOLYFILLMODE: {
        const mode = c.u16(); // 1=ALTERNATE→evenodd, 2=WINDING→nonzero
        s.fillRule = mode === 1 ? 'evenodd' : 'nonzero';
        break;
      }
      case META.SETTEXTCOLOR: {
        s.textColor = colorRefToCss(c.u32());
        break;
      }
      case META.SETTEXTALIGN: {
        s.textAlign = c.u16();
        break;
      }
      case META.CREATEPENINDIRECT: {
        insertObject(s.objects, createPen(c));
        break;
      }
      case META.CREATEBRUSHINDIRECT: {
        insertObject(s.objects, createBrush(c));
        break;
      }
      case META.CREATEFONTINDIRECT: {
        insertObject(s.objects, createFont(c));
        break;
      }
      case META.SELECTOBJECT: {
        const idx = c.u16();
        const obj = s.objects[idx];
        if (obj?.kind === 'pen') s.curPen = obj;
        else if (obj?.kind === 'brush') s.curBrush = obj;
        else if (obj?.kind === 'font') s.curFont = obj;
        break;
      }
      case META.DELETEOBJECT: {
        const idx = c.u16();
        const obj = s.objects[idx];
        if (obj) {
          if (obj === s.curPen) s.curPen = null;
          if (obj === s.curBrush) s.curBrush = null;
          if (obj === s.curFont) s.curFont = null;
          s.objects[idx] = null;
        }
        break;
      }
      case META.POLYLINE: {
        const count = c.i16();
        strokePolyline(s, readPoints(s, c, count));
        break;
      }
      case META.POLYGON: {
        const count = c.i16();
        fillStrokePolygon(s, readPoints(s, c, count));
        break;
      }
      case META.POLYPOLYGON: {
        fillStrokePolyPolygon(s, c);
        break;
      }
      case META.RECTANGLE: {
        // params: i16 bottom, i16 right, i16 top, i16 left
        const bottom = c.i16();
        const right = c.i16();
        const top = c.i16();
        const left = c.i16();
        fillStrokePolygon(s, [
          [mapX(s, left), mapY(s, top)],
          [mapX(s, right), mapY(s, top)],
          [mapX(s, right), mapY(s, bottom)],
          [mapX(s, left), mapY(s, bottom)],
        ]);
        break;
      }
      case META.TEXTOUT: {
        // META_TEXTOUT params: u16 StringLength, String bytes padded to a WORD
        // boundary, then i16 yStart and i16 xStart.
        const len = c.u16();
        const text = decodeSingleByteText(c.bytes(len));
        if (len % 2 !== 0) c.skip(1);
        const y = c.i16();
        const x = c.i16();
        drawTextOut(s, text, x, y);
        break;
      }
      case META.STRETCHDIBITS: {
        // META_STRETCHDIBITS ([MS-WMF] 2.3.1.6). Params after the 6-byte
        // size+function header, all little-endian:
        //   u32 RasterOperation, i16 SrcHeight, i16 SrcWidth, i16 YSrc, i16 XSrc,
        //   u16 UsageSrc, i16 DestHeight, i16 DestWidth, i16 YDest, i16 XDest,
        //   then the packed DIB (BITMAPINFOHEADER + palette + pixel bits).
        // We ignore the raster-op (draw plainly, like the EMF player) and draw
        // the whole DIB, so Src{Height,Width,X,Y} and UsageSrc are read but unused.
        c.u32(); // RasterOperation (ignored)
        c.i16(); // SrcHeight (unused — whole DIB drawn)
        c.i16(); // SrcWidth
        c.i16(); // YSrc
        c.i16(); // XSrc
        c.u16(); // UsageSrc (unused)
        const destHeight = c.i16();
        const destWidth = c.i16();
        const yDest = c.i16();
        const xDest = c.i16();
        // The packed DIB begins at the current cursor position:
        //   paramStart + 22  (22 = u32 RasterOperation + 9 × i16/u16 = 4 + 18).
        const dibOff = paramStart + 22;
        const dibLen = recEnd - dibOff;
        const dib = decodePackedDib(dv, dibOff, dibLen);
        if (dib) {
          const x0 = mapX(s, xDest);
          const y0 = mapY(s, yDest);
          const x1 = mapX(s, xDest + destWidth);
          const y1 = mapY(s, yDest + destHeight);
          if (blitDibToCtx(s.ctx, dib, x0, y0, x1, y1)) s.drew = true;
        }
        break;
      }
      case META.DIBSTRETCHBLT:
      case META.DIBBITBLT:
      case META.SETBKMODE:
        // META_DIBBITBLT / META_DIBSTRETCHBLT ([MS-WMF] 2.3.1.2 / 2.3.1.3) also
        // carry a packed DIB, but with a different (raster-op-dependent) preamble
        // whose exact layout we do not decode here. Skipping is safer than
        // guessing an offset that could mis-parse; STRETCHDIBITS covers the
        // common embedded-raster case.
        break;
      default:
        // ESCAPE, SETROP2, SETBKMODE, SETTEXTALIGN, SETSTRETCHBLTMODE,
        // SETMAPMODE, and anything unrecognized: skip by record size.
        break;
    }

    pos = recEnd;
  }

  return s.drew;
}

// ── raster sizing ───────────────────────────────────────────────────────────

/** Upper bound for a metafile raster dimension (px). Keeps memory bounded for
 *  large intended draw sizes while staying sharp at typical chart sizes. */
const WMF_RASTER_MAX_PX = 2000;
/** Supersampling factor for metafile rasterization (≈retina). The draw site
 *  scales the bitmap to the resolved box via `drawImage`, so a higher intrinsic
 *  resolution just buys sharper vector edges. */
const WMF_RASTER_SCALE = 2;

/** Pick a raster target size (px) for a vector metafile from its intended draw
 *  size (pt), supersampled and capped. Falls back to a sane square when the
 *  intended size is unknown (0). */
export function wmfRasterTarget(widthPt: number, heightPt: number): { w: number; h: number } {
  const fallbackPt = 300; // ~4 inch — only used when no size is surfaced
  const wPt = widthPt > 0 ? widthPt : fallbackPt;
  const hPt = heightPt > 0 ? heightPt : fallbackPt;
  const clamp = (n: number) => Math.max(1, Math.min(WMF_RASTER_MAX_PX, Math.round(n)));
  return { w: clamp(wPt * WMF_RASTER_SCALE), h: clamp(hPt * WMF_RASTER_SCALE) };
}

// ── async OffscreenCanvas wrapper ───────────────────────────────────────────

/**
 * Rasterize a WMF metafile to an `ImageBitmap` of `targetW`×`targetH`, replaying
 * onto an `OffscreenCanvas` 2D context. Returns `null` if the bytes are not a
 * parseable WMF or nothing drew (so the caller can fall back to the existing
 * "missing image" behavior without crashing).
 *
 * `suppressBoundaryFrame` (default false) enables the docx cosmetic-frame
 * heuristic; pptx/xlsx leave it false for a spec-clean render.
 */
export async function renderWmfToBitmap(
  bytes: Uint8Array,
  targetW: number,
  targetH: number,
  suppressBoundaryFrame = false,
): Promise<ImageBitmap | null> {
  if (!isWmf(bytes)) return null;
  if (targetW <= 0 || targetH <= 0) return null;
  // Prefer OffscreenCanvas but fall back to a detached <canvas> (or null in a
  // headless env) via the shared allocator, matching emf.ts / dib.ts — a WMF
  // blip then rasterizes on the main thread too instead of hard-requiring
  // OffscreenCanvas. null ⇒ degrade to the caller's "missing image" path.
  const canvas = createAuxCanvas(targetW, targetH);
  if (!canvas) return null;
  const ctx = canvas.getContext('2d') as AnyCtx | null;
  if (!ctx) return null;
  ctx.lineJoin = 'round';
  ctx.lineCap = 'round';
  const drew = playWmf(bytes, ctx, targetW, targetH, suppressBoundaryFrame);
  if (!drew) return null;
  return createImageBitmap(canvas);
}

// ── shared raster/metafile decoder ──────────────────────────────────────────

/** Options for {@link decodeRasterOrMetafile}. */
export interface DecodeRasterOptions {
  /** Intended draw width in points; sizes the metafile raster target. 0/omitted
   *  falls back to a sane square (see {@link wmfRasterTarget}). */
  widthPt?: number;
  /** Intended draw height in points; see `widthPt`. */
  heightPt?: number;
  /** Enable the docx cosmetic window/device-frame suppression heuristic (see
   *  the HEURISTIC note on {@link playWmf}/deviceInteriorEdges). Default false =
   *  spec-clean (every edge drawn). docx opts IN when it re-points to core. */
  suppressBoundaryFrame?: boolean;
  /** Optional TIFF codec. Omitted TIFF parts take the null-bitmap path. */
  tiff?: TiffRenderer;
  /** Desired width of the full source grid in retained device pixels. The
   * decoder preserves source aspect ratio and never upscales. */
  targetWidthPx?: number;
  /** Desired height of the full source grid in retained device pixels. */
  targetHeightPx?: number;
}

function decodeResizeOptions(
  source: RasterDimensions,
  targetWidthPx: number | undefined,
  targetHeightPx: number | undefined,
): ImageBitmapOptions | null {
  // Preserve the browser's established native-decode + drawImage sampling for
  // ordinary, already-safe sources. Target-sized decoding is the admission path
  // for a source that cannot be retained natively, not a fidelity rewrite for
  // every existing picture.
  if (!rasterExceedsBudget(source)) return null;
  if (typeof targetWidthPx !== 'number' || typeof targetHeightPx !== 'number') return null;
  if (!Number.isFinite(targetWidthPx) || !Number.isFinite(targetHeightPx)) return null;
  if (!(targetWidthPx > 0) || !(targetHeightPx > 0)) return null;
  const requestedScale = Math.min(1, Math.max(
    targetWidthPx / source.width,
    targetHeightPx / source.height,
  ));
  let scale = requestedScale;
  scale = Math.min(
    scale,
    MAX_RASTER_DIMENSION / source.width,
    MAX_RASTER_DIMENSION / source.height,
    Math.sqrt(MAX_RASTER_PIXELS / (source.width * source.height)),
  );
  if (!(scale < 1)) return null;
  const resizeWidth = Math.max(1, Math.floor(source.width * scale));
  // Specify one axis only. The HTML algorithm derives the other from the
  // oriented source rectangle, preserving JPEG EXIF rotation/aspect instead of
  // forcing coded-header W×H onto a 90°-oriented image.
  return { resizeWidth, resizeQuality: 'high' };
}

function rasterLimitError(
  dimensions: RasterDimensions,
  dimensionLimit: number,
  pixelLimit: number,
): OoxmlDecodedImageLimitError {
  const observedDimension = Math.max(dimensions.width, dimensions.height);
  if (!Number.isFinite(observedDimension) || observedDimension > dimensionLimit) {
    return new OoxmlDecodedImageLimitError(
      'image-dimension',
      dimensionLimit,
      Number.isFinite(observedDimension) ? observedDimension : Number.MAX_SAFE_INTEGER,
    );
  }
  const observedPixels = dimensions.width * dimensions.height;
  return new OoxmlDecodedImageLimitError(
    'image-pixels',
    pixelLimit,
    Number.isSafeInteger(observedPixels) && observedPixels >= 0
      ? observedPixels
      : Number.MAX_SAFE_INTEGER,
  );
}

/**
 * Decode an embedded raster-or-metafile blip to an `ImageBitmap`, content-
 * sniffing the bytes first (extension/MIME are unreliable — sample-10's chart is
 * a standard WMF mislabeled `.emf`). Shared by the pptx and xlsx renderers (and,
 * deferred, docx) so a WMF blip rasterizes through the minimal player instead of
 * throwing in `createImageBitmap` and vanishing. SVG stays each package's
 * concern (decoded via the `<img>` path), so it is intentionally NOT handled
 * here.
 *
 * Branch logic:
 *   - {@link isWmf} → rasterize via {@link renderWmfToBitmap} at
 *     {@link wmfRasterTarget}(widthPt, heightPt). A WMF that produces no geometry
 *     returns `null`.
 *   - {@link isEmf} → rasterize via {@link ./emf.ts}#renderEmfToBitmap at
 *     {@link wmfRasterTarget}(widthPt, heightPt). An EMF that produces no
 *     geometry returns `null`.
 *   - {@link isTiff} → decode the supported TIFF 6.0 class through the shared
 *     software rasterizer because browsers generally cannot decode TIFF.
 *   - otherwise → `createImageBitmap(blob)`, with a bounded display-sized
 *     decode when an over-budget source has a target (PNG/JPEG/GIF/BMP/WEBP…).
 *
 * Returns `null` (never throws on an unsupported metafile) so every caller can
 * treat the result like the existing "missing image" / null-bitmap path and
 * skip the picture rather than crash.
 *
 * `data` is taken as a `Blob` (what `fetchImage` yields and what
 * `createImageBitmap` consumes); the bytes are read once for sniffing.
 */
export async function decodeRasterOrMetafile(
  data: Blob,
  opts: DecodeRasterOptions = {},
): Promise<ImageBitmap | null> {
  const {
    widthPt = 0,
    heightPt = 0,
    suppressBoundaryFrame = false,
    tiff,
    targetWidthPx,
    targetHeightPx,
  } = opts;

  // Sniff a header prefix, not the whole blob. `isEmf` reads bytes 40-43 (u32@40
  // == " EMF") and `isWmf` at most the 18-byte standard header, and the raster
  // pixel-dimension sniff (below) needs ≤30 bytes for PNG/GIF/BMP/WEBP but must
  // walk JPEG marker segments to reach the Start-Of-Frame — which can sit a few
  // hundred bytes in past EXIF/ICC metadata. A 64 KiB prefix covers the SOF of
  // essentially every real JPEG while staying far smaller than a full-image copy;
  // if the SOF is even deeper the sniff simply fails open (not recognized ⇒ not
  // blocked). The common raster case then hands the Blob straight to
  // createImageBitmap without materializing the full bytes in JS. `Blob.slice`
  // clamps to the blob's actual length, so a tiny image just yields a short head.
  const RASTER_SNIFF_BYTES = 64 * 1024;
  const head = new Uint8Array(await data.slice(0, RASTER_SNIFF_BYTES).arrayBuffer());

  if (isWmf(head)) {
    const { w, h } = wmfRasterTarget(widthPt, heightPt);
    return enforceDecodedBitmapBudget(
      await renderWmfToBitmap(new Uint8Array(await data.arrayBuffer()), w, h, suppressBoundaryFrame),
    );
  }
  if (isEmf(head)) {
    const { w, h } = wmfRasterTarget(widthPt, heightPt);
    return enforceDecodedBitmapBudget(
      await renderEmfToBitmap(new Uint8Array(await data.arrayBuffer()), w, h),
    );
  }
  // Decode-bomb guard: if the header declares a recognized raster (PNG/JPEG/GIF/
  // BMP/WEBP/TIFF) whose pixel dimensions exceed the shared budget, refuse it
  // before either a browser decoder or an injected codec can allocate a large
  // surface. An unrecognized header is validated immediately after decode.
  const rasterDimensions = sniffRasterDimensions(head);
  if (rasterDimensions && sourceRasterExceedsBudget(rasterDimensions)) {
    throw rasterLimitError(
      rasterDimensions,
      MAX_RASTER_SOURCE_DIMENSION,
      MAX_RASTER_SOURCE_PIXELS,
    );
  }
  const resizeOptions = rasterDimensions
    ? decodeResizeOptions(rasterDimensions, targetWidthPx, targetHeightPx)
    : null;
  const tiffInput = isTiff(head);
  if (rasterDimensions && rasterExceedsBudget(rasterDimensions) && (!resizeOptions || tiffInput)) {
    throw rasterLimitError(
      rasterDimensions,
      MAX_RASTER_DIMENSION,
      MAX_RASTER_PIXELS,
    );
  }
  if (tiffInput) {
    if (!tiff) return null;
    return enforceDecodedBitmapBudget(
      await tiff.render(new Uint8Array(await data.arrayBuffer())),
    );
  }
  return enforceDecodedBitmapBudget(
    resizeOptions
      ? await createImageBitmap(data, resizeOptions)
      : await createImageBitmap(data),
  );
}

/** Validate at the decoder boundary so direct callers and the shared cache have
 * identical hard-quota semantics. The allocation has already happened when a
 * header was unrecognized, but an oversized surface is closed immediately and
 * never reaches a renderer or cache. */
function enforceDecodedBitmapBudget(bitmap: ImageBitmap | null): ImageBitmap | null {
  if (!bitmap) return null;
  const width = Number(bitmap.width);
  const height = Number(bitmap.height);
  if (!rasterExceedsBudget({ width, height })) return bitmap;

  closeImageBitmapIfSupported(bitmap);
  throw rasterLimitError(
    { width, height },
    MAX_RASTER_DIMENSION,
    MAX_RASTER_PIXELS,
  );
}
