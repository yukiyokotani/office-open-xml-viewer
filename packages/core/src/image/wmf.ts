// ── WMF (Windows Metafile) player ───────────────────────────────────────────
//
// Browsers cannot decode WMF/EMF via `createImageBitmap`, so the renderer falls
// back to this player for metafile blips. It is a *minimal* WMF interpreter:
// just enough to rasterize the vector-graphics metafiles that Office embeds for
// charts and diagrams.
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
// POLYPOLYGON, RECTANGLE, TEXTOUT, a bounded ANSI EXTTEXTOUT subset,
// STRETCHDIBITS (embedded raster DIB via the
// shared decoder in ./dib.ts), EOF.
// Ignored (no-op, skipped by size): SETROP2,
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
// by the sibling player {@link ./emf.ts}#renderEmfToBitmap, routed from the
// shared admission boundary in `raster-or-metafile.ts`.

import { decodePackedDib, blitDibToCtx } from './dib.js';
import { createAuxCanvas } from '../canvas/aux-canvas.js';

// WMF record function codes (the subset we act on; others are skipped by size).
const META = {
  EOF: 0x0000,
  SETBKMODE: 0x0102,
  SETBKCOLOR: 0x0201,
  SETTEXTALIGN: 0x012e,
  SETTEXTCOLOR: 0x0209,
  SETPOLYFILLMODE: 0x0106,
  SETWINDOWORG: 0x020b,
  SETWINDOWEXT: 0x020c,
  SELECTOBJECT: 0x012d,
  DELETEOBJECT: 0x01f0,
  TEXTOUT: 0x0521,
  EXTTEXTOUT: 0x0a32,
  ESCAPE: 0x0626,
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
  signedHeight: number;
  width: number;
  escapement: number;
  orientation: number;
  weight: number; // lfWeight (400 normal, 700 bold)
  italic: boolean;
  underline: boolean;
  strikeout: boolean;
  charset: number;
  face: string;
  ansiFace: string;
  validForExt: boolean;
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
  bkMode: number;
  extTextSafe: boolean;
  extTextMappingSafe: boolean;
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

// WHATWG Encoding §9.1, index-windows-1252.txt (identifier e56d49d9,
// 2024-09-18). Decode directly so WMF ANSI text is deterministic even on
// hosts whose TextDecoder delegates this label to an ISO-8859-1 codec.
const WINDOWS_1252_C1 = [
  0x20ac, 0x0081, 0x201a, 0x0192, 0x201e, 0x2026, 0x2020, 0x2021,
  0x02c6, 0x2030, 0x0160, 0x2039, 0x0152, 0x008d, 0x017d, 0x008f,
  0x0090, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014,
  0x02dc, 0x2122, 0x0161, 0x203a, 0x0153, 0x009d, 0x017e, 0x0178,
] as const;

function decodeWindows1252(bytes: Uint8Array): string {
  return Array.from(bytes, (byte) =>
    String.fromCharCode(byte >= 0x80 && byte <= 0x9f ? WINDOWS_1252_C1[byte - 0x80] : byte),
  ).join('');
}

function createFont(c: Cursor): Font {
  const validLength = c.remaining === 50;
  const signedHeight = c.i16();
  const height = Math.abs(signedHeight);
  const width = c.i16();
  const escapement = c.i16();
  const orientation = c.i16();
  const weight = c.i16();
  const italic = c.u8() !== 0;
  const underline = c.u8() !== 0;
  const strikeout = c.u8() !== 0;
  const charset = c.u8();
  c.u8(); // lfOutPrecision
  c.u8(); // lfClipPrecision
  c.u8(); // lfQuality
  c.u8(); // lfPitchAndFamily
  const faceBytes = c.bytes(Math.min(32, c.remaining));
  const face = decodeSingleByteText(faceBytes);
  const nul = faceBytes.indexOf(0);
  const ansiFace = String.fromCharCode(...faceBytes.subarray(0, nul < 0 ? faceBytes.length : nul));
  return { kind: 'font', height, signedHeight, width, escapement, orientation, weight, italic, underline, strikeout, charset, face, ansiFace, validForExt: validLength && nul >= 0 };
}

/** MS-WMF 2.3.3.5, 2.2.1.2 and 2.1.2.3. This implementation intentionally
 * omits valid variants outside the measured ANSI/transparent/baseline subset;
 * omission is implementation policy, not a claim that those records are invalid. */
function drawExtTextOut(s: PlayState, c: Cursor): void {
  if (c.remaining < 8) return;
  const y = c.i16();
  const x = c.i16();
  const length = c.i16();
  const options = c.u16();
  if (length < 0) return;
  if (options !== 0) return; // Rectangle-bearing/clipped/opaque variants are outside this subset.
  const padded = length + (length & 1);
  if (c.remaining !== padded) return; // Reject truncation and any optional Dx array.
  const raw = c.bytes(length);
  if (length & 1) c.skip(1);
  const font = s.curFont;
  const horizontal = s.textAlign & 0x6;
  const unsupportedAlignment = (s.textAlign & ~0x1e) !== 0 || (s.textAlign & 0x18) !== 0x18 || ![0, 2, 6].includes(horizontal);
  if (
    !s.extTextSafe ||
    !s.extTextMappingSafe ||
    s.bkMode !== 1 ||
    unsupportedAlignment ||
    !s.haveExt ||
    !(s.extX > 0 && s.extY > 0 && s.W > 0 && s.H > 0) ||
    ![s.extX, s.extY, s.W, s.H].every(Number.isFinite) ||
    !font ||
    font.charset !== 0 ||
    font.signedHeight >= 0 ||
    font.width !== 0 ||
    font.escapement !== 0 ||
    font.orientation !== 0 ||
    font.underline ||
    font.strikeout ||
    !font.ansiFace ||
    /[\x00-\x1f\x7f]/.test(font.ansiFace) ||
    !font.validForExt ||
    font.weight < 1 ||
    font.weight > 1000
  ) return;
  const text = decodeWindows1252(raw);
  if (!text.length) return;
  const px = font.height * (s.H / s.extY);
  if (!Number.isFinite(px) || px <= 0) return;
  try {
    s.ctx.fillStyle = s.textColor;
    const family = `"${font.ansiFace.replaceAll('\\', '\\\\').replaceAll('"', '\\"')}"`;
    s.ctx.font = `${font.italic ? 'italic ' : ''}${font.weight} ${px}px ${family}`;
    const horiz = s.textAlign & 0x6;
    s.ctx.textAlign = horiz === 2 ? 'right' : horiz === 6 ? 'center' : 'left';
    s.ctx.textBaseline = 'alphabetic';
    s.ctx.fillText(text, mapX(s, x), mapY(s, y));
    s.drew = true;
  } catch { /* Context may not implement text. */ }
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
    bkMode: 2,
    extTextSafe: true,
    extTextMappingSafe: false,
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
        if (c.remaining !== 4) { s.extTextSafe = false; break; }
        s.orgY = c.i16(); s.orgX = c.i16();
        break;
      }
      case META.SETWINDOWEXT: {
        // params: i16 yExt, i16 xExt (Y FIRST)
        if (c.remaining !== 4) { s.extTextSafe = false; break; }
        const yExt = c.i16();
        const xExt = c.i16();
        s.extTextMappingSafe = xExt > 0 && yExt > 0;
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
        if (c.remaining !== 4) { s.extTextSafe = false; break; }
        s.textColor = colorRefToCss(c.u32());
        break;
      }
      case META.SETTEXTALIGN: {
        if (c.remaining !== 2 && c.remaining !== 4) { s.extTextSafe = false; break; }
        s.textAlign = c.u16();
        if (c.remaining === 2) c.u16(); // Optional Reserved; MUST be ignored.
        break;
      }
      case META.SETBKMODE: {
        if (c.remaining !== 2 && c.remaining !== 4) { s.extTextSafe = false; break; }
        s.bkMode = c.u16();
        if (c.remaining === 2) c.u16(); // Optional Reserved; MUST be ignored.
        break;
      }
      case META.SETBKCOLOR: {
        // Background color cannot affect this supported transparent-text subset.
        if (c.remaining !== 4) s.extTextSafe = false;
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
        if (c.remaining < 18) { s.extTextSafe = false; break; }
        insertObject(s.objects, createFont(c));
        break;
      }
      case META.SELECTOBJECT: {
        if (c.remaining !== 2) { s.extTextSafe = false; break; }
        const idx = c.u16();
        const obj = s.objects[idx];
        if (obj?.kind === 'pen') s.curPen = obj;
        else if (obj?.kind === 'brush') s.curBrush = obj;
        else if (obj?.kind === 'font') s.curFont = obj;
        else s.extTextSafe = false;
        break;
      }
      case META.DELETEOBJECT: {
        if (c.remaining !== 2) { s.extTextSafe = false; break; }
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
      case META.EXTTEXTOUT: {
        drawExtTextOut(s, c);
        break;
      }
      case META.ESCAPE: {
        const escape = c.remaining >= 4 ? c.u16() : -1;
        const count = c.remaining >= 2 ? c.u16() : -1;
        if (escape !== 15 || count < 0 || c.remaining !== count + (count & 1)) {
          s.extTextSafe = false;
          break;
        }
        const data = c.bytes(count);
        if (count & 1) c.skip(1);
        // META_ESCAPE_ENHANCED_METAFILE begins with the little-endian WMFC
        // CommentIdentifier (MS-WMF 2.3.6.25); nested graphics state is outside
        // this player. Historical Microsoft SDK Q81497 documents non-WMFC
        // private comments as ignored during ordinary playback (no DC change).
        if (data.length >= 4 && new DataView(data.buffer, data.byteOffset, data.byteLength).getUint32(0, true) === 0x43464d57) s.extTextSafe = false;
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
        // META_DIBBITBLT / META_DIBSTRETCHBLT ([MS-WMF] 2.3.1.2 / 2.3.1.3) also
        // carry a packed DIB, but with a different (raster-op-dependent) preamble
        // whose exact layout we do not decode here. Skipping is safer than
        // guessing an offset that could mis-parse; STRETCHDIBITS covers the
        // common embedded-raster case.
        break;
      default:
        // Unsupported graphics-state changes make later EXTTEXTOUT unsafe for
        // this bounded player, although established geometry handlers remain
        // unaffected and the record is still skipped by its declared size.
        s.extTextSafe = false;
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
