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
// Implemented records: window/viewport org/ext/offset/scale, SAVEDC/RESTOREDC,
// INTERSECTCLIPRECT/EXCLUDECLIPRECT/OFFSETCLIPRGN, CREATEREGION/SELECTCLIPREGION,
// SETROP2 (COPYPEN/NOP), SETPOLYFILLMODE,
// SETTEXTCOLOR, SETTEXTALIGN, CREATEPENINDIRECT, CREATEBRUSHINDIRECT,
// CREATEFONTINDIRECT, SELECTOBJECT, DELETEOBJECT, POLYLINE, POLYGON,
// POLYPOLYGON, RECTANGLE, MOVETO, LINETO, TEXTOUT and EXTTEXTOUT (with
// SETBKCOLOR/SETBKMODE/SETTEXTCHAREXTRA/SETTEXTJUSTIFICATION; see ./wmf-text.ts),
// STRETCHDIBITS (embedded raster DIB via the
// shared decoder in ./dib.ts), EOF.
// Ignored (no-op, skipped by size):
// SETSTRETCHBLTMODE, DIBBITBLT/DIBSTRETCHBLT (their exact param
// layout is not decoded here — skipped rather than mis-parsed), and any
// unrecognized record.
//
// Shared across the DOCX, PPTX and XLSX renderers.
//
// True EMF (a separate, larger 32-bit format — see {@link isEmf}) is rasterized
// by the sibling player {@link ./emf.ts}#renderEmfToBitmap, routed from the
// shared admission boundary in `raster-or-metafile.ts`.

import { decodePackedDib, blitDibToCtx } from './dib.js';
import { createAuxCanvas } from '../canvas/aux-canvas.js';
import { drawWmfText, parseWmfFont, ETO_CLIPPED, ETO_OPAQUE, TA_UPDATECP, type WmfFont } from './wmf-text.js';

// WMF record function codes (the subset we act on; others are skipped by size).
const META = {
  EOF: 0x0000,
  SETBKMODE: 0x0102,
  SETBKCOLOR: 0x0201,
  SETTEXTCHAREXTRA: 0x0108,
  SETTEXTJUSTIFICATION: 0x020a,
  CREATEPALETTE: 0x00f7,
  CREATEPATTERNBRUSH: 0x01f9,
  DIBCREATEPATTERNBRUSH: 0x0142,
  SETTEXTALIGN: 0x012e,
  SETTEXTCOLOR: 0x0209,
  SETPOLYFILLMODE: 0x0106,
  SETROP2: 0x0104,
  SETMAPMODE: 0x0103,
  SETWINDOWORG: 0x020b,
  SETWINDOWEXT: 0x020c,
  SETVIEWPORTORG: 0x020d,
  SETVIEWPORTEXT: 0x020e,
  OFFSETWINDOWORG: 0x020f,
  OFFSETVIEWPORTORG: 0x0211,
  SCALEWINDOWEXT: 0x0410,
  SCALEVIEWPORTEXT: 0x0412,
  INTERSECTCLIPRECT: 0x0416,
  EXCLUDECLIPRECT: 0x0415,
  OFFSETCLIPRGN: 0x0220,
  CREATEREGION: 0x06ff,
  SELECTCLIPREGION: 0x012c,
  SAVEDC: 0x001e,
  RESTOREDC: 0x0127,
  SELECTOBJECT: 0x012d,
  DELETEOBJECT: 0x01f0,
  LINETO: 0x0213,
  MOVETO: 0x0214,
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
  lineWidth: number;
  lineEligible: boolean;
}
interface Brush {
  kind: 'brush';
  fill: string | null; // null = BS_NULL / hollow (no fill)
}
type Font = WmfFont;
/** An object a record creates but this player does not interpret (palette,
 *  pattern brush). It occupies its object-table slot so later indices align. */
interface OtherObject { kind: 'other' }
type RectBounds = [number, number, number, number];
interface Region { kind: 'region'; tree: ClipNode | null; valid: boolean }
type WmfObject = Pen | Brush | Font | Region | OtherObject;
// The DC region is independent of both the raster surface and caller clip.
// An immutable, height-balanced spatial tree owns disjoint rectangle pieces.
// Saved DCs share untouched subtrees; offsets are O(1), never history replays.
interface ClipNode {
  bounds: RectBounds;
  height: number;
  children?: readonly [ClipNode, ClipNode];
}
interface WmfClip {
  tree: ClipNode | null;
  dx: number;
  dy: number;
  /** Borrow selected scan bands without copying them for every SELECT/SAVEDC. */
  mapping?: readonly [number, number, number, number];
}
const UNBOUNDED_CLIP: WmfClip = {
  tree: { bounds: [-Infinity, -Infinity, Infinity, Infinity], height: 1 }, dx: 0, dy: 0,
};

function clipBranch(a: ClipNode, b: ClipNode): ClipNode {
  return { bounds: [Math.min(a.bounds[0], b.bounds[0]), Math.min(a.bounds[1], b.bounds[1]),
    Math.max(a.bounds[2], b.bounds[2]), Math.max(a.bounds[3], b.bounds[3])],
    height: Math.max(a.height, b.height) + 1, children: [a, b] };
}

/** Persistent AVL concatenation: retain spatial neighbors and logarithmic
 * tree depth after a rectangle is split or removed. No saved tree is mutated. */
function joinClip(a: ClipNode | null, b: ClipNode | null): ClipNode | null {
  if (!a) return b;
  if (!b) return a;
  if (a.height > b.height + 1 && a.children) {
    const [l, r] = a.children;
    const next = joinClip(r, b) as ClipNode;
    if (next.height <= l.height + 1) return clipBranch(l, next);
    const [nl, nr] = next.children as readonly [ClipNode, ClipNode];
    if (nr.height >= nl.height) return clipBranch(clipBranch(l, nl), nr);
    const [nll, nlr] = nl.children as readonly [ClipNode, ClipNode];
    return clipBranch(clipBranch(l, nll), clipBranch(nlr, nr));
  }
  if (b.height > a.height + 1 && b.children) {
    const [l, r] = b.children;
    const next = joinClip(a, l) as ClipNode;
    if (next.height <= r.height + 1) return clipBranch(next, r);
    const [nl, nr] = next.children as readonly [ClipNode, ClipNode];
    if (nl.height >= nr.height) return clipBranch(nl, clipBranch(nr, r));
    const [nrl, nrr] = nr.children as readonly [ClipNode, ClipNode];
    return clipBranch(clipBranch(nl, nrl), clipBranch(nrr, r));
  }
  return clipBranch(a, b);
}

function clipTree(rects: readonly RectBounds[], start = 0, end = rects.length): ClipNode | null {
  if (start === end) return null;
  if (end === start + 1) return { bounds: rects[start], height: 1 };
  const mid = Math.floor((start + end) / 2);
  return clipBranch(clipTree(rects, start, mid) as ClipNode, clipTree(rects, mid, end) as ClipNode);
}

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
  // Window and viewport mapping (MS-WMF 2.3.5.28–31). The caller supplies
  // the initial target viewport when the metafile leaves it unspecified.
  orgX: number;
  orgY: number;
  extX: number;
  extY: number;
  haveExt: boolean;
  vpX: number;
  vpY: number;
  vpW: number;
  vpH: number;
  clip: WmfClip;
  clipDirty: boolean;
  mapMode: number;
  // GDI state
  objects: (WmfObject | null)[];
  curPen: Pen | null;
  curBrush: Brush | null;
  curFont: Font | null;
  textColor: string;
  bkColor: string;
  textAlign: number;
  charExtra: number;
  breakCount: number;
  breakQuotientPx: number;
  breakErrorPx: number;
  fillRule: CanvasFillRule; // from SETPOLYFILLMODE
  rop2: number;
  bkMode: number;
  unsupported: Set<string>;
  extTextSafe: boolean;
  extTextMappingSafe: boolean;
  currentX: number;
  currentY: number;
  currentPositionKnown: boolean;
  lineStateSafe: boolean;
  drew: boolean;
}

/** logical → device along X. */
function mapX(s: PlayState, x: number): number {
  return s.vpX + (x - s.orgX) * (s.vpW / s.extX);
}
/** logical → device along Y. */
function mapY(s: PlayState, y: number): number {
  return s.vpY + (y - s.orgY) * (s.vpH / s.extY);
}

/** MS-WMF 2.1.1.16 / 2.3.5.17: fixed-unit modes ignore window/viewport
 * extent setters and scalers; origins remain independently adjustable.
 * Library device policy: Canvas pixels are square, 96 pixels per inch (CSS
 * reference pixels), not the original metafile author's printer/display DPI.
 * MM_TEXT is exactly one device pixel. Physical modes have upward-positive Y.
 * The initial mode is caller-defined; only an authored SETMAPMODE selects one
 * of these fixed mappings. No machine/OS-dependent device capabilities enter.
 */
function setMapMode(s: PlayState, mode: number): void {
  s.mapMode = mode;
  if (mode <= 6) {
    const unitsPerInch = [0, 1, 254, 2540, 100, 1000, 1440][mode];
    s.extX = unitsPerInch; s.extY = unitsPerInch;
    s.vpW = mode === 1 ? 1 : 96;
    s.vpH = mode === 1 ? 1 : -96;
    s.haveExt = true;
    s.extTextMappingSafe = true;
  }
  adjustIsotropic(s);
}

/** MS-WMF 2.1.1.16 requires equal physical units in MM_ISOTROPIC and
 * recommends adjusting the viewport when extents change. On the square-pixel
 * Canvas device, reduce the larger scale to fit both requested extents, while
 * preserving each authored axis direction. MM_ANISOTROPIC keeps both scales.
 */
function adjustIsotropic(s: PlayState): void {
  if (s.mapMode !== 7) return;
  const scale = Math.min(Math.abs(s.vpW / s.extX), Math.abs(s.vpH / s.extY));
  s.vpW = Math.sign(s.vpW) * Math.abs(s.extX) * scale;
  s.vpH = Math.sign(s.vpH) * Math.abs(s.extY) * scale;
}

/** Device line width: scale the logical pen width by |W/extX| and clamp to ≥1
 *  so hairlines stay visible after mapping. */
function deviceLineWidth(s: PlayState, logicalWidth: number): number {
  const scale = Math.abs(s.vpW / s.extX);
  const w = logicalWidth * scale;
  return w >= 1 ? w : 1;
}

/** Draw the whole authored boundary. MS-WMF 3.1.5 defines the clip separately
 * from the window/viewport mapping; boundary coincidence never removes edges. */
function strokeEdges(s: PlayState, pts: Array<[number, number]>, closed: boolean): void {
  if (s.rop2 !== 13 || !s.curPen || s.curPen.stroke == null || pts.length < 2) return;
  const { ctx } = s;
  ctx.strokeStyle = s.curPen.stroke;
  ctx.lineWidth = deviceLineWidth(s, s.curPen.width);
  ctx.beginPath();
  ctx.moveTo(pts[0][0], pts[0][1]);
  for (let i = 1; i < pts.length; i++) ctx.lineTo(pts[i][0], pts[i][1]);
  if (closed) ctx.closePath();
  ctx.stroke();
  s.drew = true;
}

/** Map a selected scan band only when it is consumed; subsequent mapping
 * changes do not move a region already in device space (MS-WMF 2.3.4.9). */
function clipRect(clip: WmfClip, rect: RectBounds): RectBounds {
  if (!clip.mapping) return rect;
  const [sx, sy, tx, ty] = clip.mapping;
  const x0 = rect[0] * sx + tx; const x1 = rect[2] * sx + tx;
  const y0 = rect[1] * sy + ty; const y1 = rect[3] * sy + ty;
  return [Math.min(x0, x1), Math.min(y0, y1), Math.max(x0, x1), Math.max(y0, y1)];
}

/** MS-WMF 2.3.5.2–3: Boolean operations affect the actual device region,
 * including off-surface pieces that OFFSETCLIPRGN can later bring into view.
 * Prune disjoint/contained subtrees; copy only affected branches. Work follows
 * visited live geometry, with at most four pieces per split leaf, rather than
 * scanning/copying every saved region or replaying a growing operation log.
 * The implicit region is the plane; the surface/caller clip is applied only at
 * Canvas emission, never used as a subtraction universe. */
function cutClip(node: ClipNode | null, bounds: RectBounds, exclude: boolean): ClipNode | null {
  if (!node) return null;
  const [pl, pt, pr, pb] = node.bounds;
  const il = Math.max(pl, bounds[0]); const it = Math.max(pt, bounds[1]);
  const ir = Math.min(pr, bounds[2]); const ib = Math.min(pb, bounds[3]);
  if (il >= ir || it >= ib) return exclude ? node : null;
  if (il === pl && it === pt && ir === pr && ib === pb) return exclude ? null : node;
  if (node.children) {
    const [a, b] = node.children;
    const l = cutClip(a, bounds, exclude); const r = cutClip(b, bounds, exclude);
    return l === a && r === b ? node : joinClip(l, r);
  }
  if (!exclude) return { bounds: [il, it, ir, ib], height: 1 };
  const pieces: RectBounds[] = [];
  if (pt < it) pieces.push([pl, pt, pr, it]);
  if (pl < il) pieces.push([pl, it, il, ib]);
  if (ir < pr) pieces.push([ir, it, pr, ib]);
  if (ib < pb) pieces.push([pl, ib, pr, pb]);
  return clipTree(pieces);
}

function combineClip(s: PlayState, bounds: RectBounds, exclude: boolean): void {
  const { dx, dy, mapping } = s.clip;
  let local: RectBounds = [bounds[0] - dx, bounds[1] - dy, bounds[2] - dx, bounds[3] - dy];
  // Keep a selected region in its captured coordinate space. Inverse-map the
  // new operand, so a cut copies only affected branches even for a large region
  // shared by many saved DCs. Signed scales preserve mirrored region membership.
  if (mapping) {
    const [sx, sy, tx, ty] = mapping;
    const x0 = (local[0] - tx) / sx; const x1 = (local[2] - tx) / sx;
    const y0 = (local[1] - ty) / sy; const y1 = (local[3] - ty) / sy;
    local = [Math.min(x0, x1), Math.min(y0, y1), Math.max(x0, x1), Math.max(y0, y1)];
  }
  s.clip = { ...s.clip, tree: cutClip(s.clip.tree, local, exclude) };
  s.clipDirty = true;
}

function applyClip(s: PlayState): void {
  if (!s.clipDirty) return;
  // Own one Canvas save frame, independent of WMF DC levels. Replacement
  // preserves the caller's pre-existing clip. Emit the current region once,
  // just before drawing; exclude/offset-only files do no Canvas clip work.
  s.ctx.restore();
  s.ctx.save();
  s.ctx.beginPath();
  const pending = s.clip.tree ? [s.clip.tree] : [];
  while (pending.length) {
    const node = pending.pop() as ClipNode;
    const rect = clipRect(s.clip, node.bounds);
    const l = Math.max(0, rect[0] + s.clip.dx); const t = Math.max(0, rect[1] + s.clip.dy);
    const r = Math.min(s.W, rect[2] + s.clip.dx); const b = Math.min(s.H, rect[3] + s.clip.dy);
    if (l >= r || t >= b) continue;
    if (node.children) pending.push(...node.children);
    else s.ctx.rect(l, t, r - l, b - t);
  }
  s.ctx.clip();
  s.clipDirty = false;
}

function mappedRect(s: PlayState, l: number, t: number, r: number, b: number): RectBounds {
  const x0 = mapX(s, l); const x1 = mapX(s, r);
  const y0 = mapY(s, t); const y1 = mapY(s, b);
  return [Math.min(x0, x1), Math.min(y0, y1), Math.max(x0, x1), Math.max(y0, y1)];
}

/** MS-WMF 2.2.1.5 and 2.2.2.21: a region is the union of every interval in
 * every scan band, not its bounding rectangle. Validate counts before reads. */
function createRegion(c: Cursor): Region {
  const region: Region = { kind: 'region', tree: null, valid: false };
  const rects: RectBounds[] = [];
  if (c.remaining < 22) return region;
  c.u16(); const type = c.i16(); c.u32(); c.i16();
  const scans = c.i16(); c.i16(); c.skip(8);
  if (type !== 6 || scans < 0) return region;
  for (let i = 0; i < scans; i++) {
    if (c.remaining < 8) return region;
    const count = c.u16(); const top = c.u16(); const bottom = c.u16();
    if (count % 2 !== 0 || c.remaining < count * 2 + 2) return region;
    for (let j = 0; j < count; j += 2) {
      const left = c.u16(); const right = c.u16();
      if (left < right && top < bottom) rects.push([left, top, right, bottom]);
    }
    if (c.u16() !== count) return region;
  }
  region.valid = c.remaining === 0;
  if (region.valid) region.tree = clipTree(rects);
  return region;
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
  strokeEdges(s, pts, false);
}

/** MS-WMF 2.3.3.10: stroke from the playback DC drawing position and then
 * update it to the destination. The specified endpoint is excluded by GDI;
 * Canvas has no equivalent endpoint-exclusion primitive, so the geometric
 * segment is emitted without an arbitrary device-pixel correction. */
function strokeLineTo(s: PlayState, x: number, y: number): void {
  const fromX = s.currentX;
  const fromY = s.currentY;
  const known = s.currentPositionKnown;
  s.currentX = x;
  s.currentY = y;
  s.currentPositionKnown = true;
  if (
    s.rop2 !== 13 ||
    !known ||
    !s.lineStateSafe ||
    !s.haveExt ||
    !s.extTextMappingSafe ||
    ![s.extX, s.extY, s.W, s.H].every(Number.isFinite) ||
    !(s.W > 0 && s.H > 0) ||
    !s.curPen ||
    !s.curPen.lineEligible ||
    s.curPen.stroke == null
  ) return;

  const { ctx } = s;
  ctx.save();
  ctx.beginPath();
  ctx.moveTo(mapX(s, fromX), mapY(s, fromY));
  ctx.lineTo(mapX(s, x), mapY(s, y));
  ctx.strokeStyle = s.curPen.stroke;
  // A zero-width cosmetic pen is one device pixel. Positive logical widths use
  // the window mapping directly; unlike the legacy polygon path, this does not
  // empirically clamp positive subpixel widths to one pixel.
  ctx.lineWidth = s.curPen.lineWidth === 0
    ? 1
    : Math.abs(s.curPen.lineWidth * (s.vpW / s.extX));
  // This bounded path supports the normative PS_SOLID default combination:
  // cosmetic pen, round end caps, and round joins (MS-WMF 2.1.1.23).
  ctx.lineCap = 'round';
  ctx.lineJoin = 'round';
  ctx.setLineDash([]);
  ctx.stroke();
  ctx.restore();
  s.drew = true;
}

/** Fill and stroke under the current playback DC clip. */
function fillStrokePolygon(s: PlayState, pts: Array<[number, number]>): void {
  if (s.rop2 !== 13 || pts.length < 2) return;
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
  if (s.rop2 !== 13) return;
  const numPolys = c.u16();
  // MS-WMF 2.2.2.17: NumberOfPolygons is u16, already bounded by its
  // representation. There is no additional library polygon-count ceiling.
  if (numPolys === 0) return;
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
  const validLength = c.remaining === 10;
  const style = c.u16();
  const widthX = c.i16();
  c.i16(); // widthY (unused — WMF pens are isotropic in practice)
  const color = c.u32();
  const lowStyle = style & 0xff;
  // PS_NULL (5) → no stroke. Dash/dot styles (1..4) are rendered solid (we do
  // not synthesize a dash pattern); see module note.
  const stroke = lowStyle === 5 ? null : colorRefToCss(color);
  // Existing polygon/polyline playback retains its historical solid fallback.
  // LINETO is stricter: only complete PS_SOLID or PS_NULL records with the
  // default cosmetic/round-cap/round-join combination are eligible.
  const lineEligible = validLength && widthX >= 0 && (style === 0 || style === 5);
  return { kind: 'pen', stroke, width: Math.abs(widthX), lineWidth: widthX, lineEligible };
}

function createBrush(c: Cursor): Brush {
  const style = c.u16();
  const color = c.u32();
  c.u16(); // hatch (HATCHED brushes are rendered as solid fills)
  // BS_NULL / BS_HOLLOW (1) → no fill. SOLID (0) and HATCHED (2) fill solid.
  const fill = style === 1 ? null : colorRefToCss(color);
  return { kind: 'brush', fill };
}

/** META_TEXTOUT / META_EXTTEXTOUT through the shared text model. */
function drawText(s: PlayState, record: string, output: Parameters<typeof drawWmfText>[1]): void {
  if (!s.extTextSafe) {
    s.unsupported.add(`${record} (text state undefined after an unsupported record)`);
    if (s.textAlign & TA_UPDATECP) s.currentPositionKnown = false;
    return;
  }
  try {
    const result = drawWmfText(s, output);
    if ('problem' in result) {
      s.unsupported.add(`${record} (${result.problem})`);
      if (s.textAlign & TA_UPDATECP) s.currentPositionKnown = false;
    } else if (result.drew) s.drew = true;
  } catch {
    // A context without text support (some test or server-side contexts).
    s.unsupported.add(`${record} (text output failed)`);
    if (s.textAlign & TA_UPDATECP) s.currentPositionKnown = false;
  }
}

/** MS-WMF 2.3.3.5 META_EXTTEXTOUT. */
function drawExtTextOut(s: PlayState, c: Cursor): void {
  const record = 'META_EXTTEXTOUT';
  if (c.remaining < 8) { s.unsupported.add(`${record} (malformed)`); return; }
  const y = c.i16();
  const x = c.i16();
  const length = c.i16();
  const options = c.u16();
  const hasRect = (options & (ETO_OPAQUE | ETO_CLIPPED)) !== 0;
  if (length < 0 || c.remaining < (hasRect ? 8 : 0) + length) { s.unsupported.add(`${record} (malformed)`); return; }
  let rect: [number, number, number, number] | null = null;
  if (hasRect) rect = [c.i16(), c.i16(), c.i16(), c.i16()];
  const bytes = c.bytes(length);
  if (length & 1) c.skip(1);
  // The optional Dx array has one 16-bit entry per String byte (two with
  // ETO_PDY); absent when the record ends with the string.
  const entries = length * ((options & 0x2000) ? 2 : 1);
  let dx: Int16Array | null = null;
  if (c.remaining > 0 && length > 0) {
    if (c.remaining < entries * 2) { s.unsupported.add(`${record} (malformed Dx array)`); return; }
    dx = new Int16Array(entries);
    for (let i = 0; i < entries; i++) dx[i] = c.i16();
  }
  drawText(s, record, { x, y, bytes, options, rect, dx });
}

const STATE_NEUTRAL_RECORDS: ReadonlySet<number> = new Set([
  0x0035, // META_REALIZEPALETTE
  0x0037, // META_SETPALENTRIES
  0x0105, // META_SETRELABS (reserved)
  0x0107, // META_SETSTRETCHBLTMODE
  0x012a, // META_INVERTREGION
  0x012b, // META_PAINTREGION
  0x0139, // META_RESIZEPALETTE
  0x0228, // META_FILLREGION
  0x0231, // META_SETMAPPERFLAGS (font-mapper aspect matching only)
  0x0234, // META_SELECTPALETTE
  0x0418, // META_ELLIPSE
  0x0419, // META_FLOODFILL
  0x041f, // META_SETPIXEL
  0x0429, // META_FRAMEREGION
  0x0436, // META_ANIMATEPALETTE
  0x0548, // META_EXTFLOODFILL
  0x061c, // META_ROUNDRECT
  0x061d, // META_PATBLT
  0x0817, // META_ARC
  0x081a, // META_PIE
  0x0830, // META_CHORD
  0x0922, // META_BITBLT
  0x0b23, // META_STRETCHBLT
  0x0d33, // META_SETDIBTODEV
]);

// ── core record-replay loop (pure; testable with a mock ctx) ────────────────

/**
 * Replay a WMF byte buffer onto a 2D context, mapping logical coordinates to a
 * `W`×`H` device space. Returns `true` if anything was drawn, `false` for a
 * non-WMF buffer or a metafile that produced no geometry.
 *
 * Pure with respect to the injected `ctx`, so it is unit-testable against a
 * recording mock — no OffscreenCanvas required.
 */
export function playWmf(
  bytes: Uint8Array,
  ctx: AnyCtx,
  W: number,
  H: number,
  onUnsupported?: (records: readonly string[]) => void,
  fontFamilies?: ReadonlyMap<string, string>,
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
    vpX: 0, vpY: 0, vpW: W, vpH: H,
    clip: UNBOUNDED_CLIP,
    clipDirty: false,
    // MS-WMF 3.1.5 leaves initial DC attributes device-dependent. Preserve
    // this player's caller-sized anisotropic viewport until SETMAPMODE.
    mapMode: 8,
    objects: [],
    curPen: null,
    curBrush: null,
    curFont: null,
    ...(fontFamilies ? { fontFamilies } : {}),
    textColor: '#000000',
    // Win32 DC defaults: white background, OPAQUE mode (SetBkColor/SetBkMode).
    bkColor: '#ffffff',
    textAlign: 0,
    charExtra: 0,
    breakCount: 0,
    breakQuotientPx: 0,
    breakErrorPx: 0,
    fillRule: 'nonzero',
    rop2: 13,
    bkMode: 2,
    unsupported: new Set(),
    extTextSafe: true,
    extTextMappingSafe: false,
    currentX: 0,
    currentY: 0,
    // MS-WMF 3.1.5 makes initial playback-DC state device-dependent. Do not
    // invent an initial drawing position; wait for META_MOVETO.
    currentPositionKnown: false,
    lineStateSafe: true,
    drew: false,
  };

  ctx.save();
  try {
    replayWmfRecords(bytes, s, pos);
  } finally {
    ctx.restore();
  }
  if (s.unsupported.size > 0) onUnsupported?.([...s.unsupported]);
  return s.drew;
}

/** Replay owns the WMF DC stack; the public wrapper owns Canvas state cleanup. */
function replayWmfRecords(bytes: Uint8Array, s: PlayState, pos: number): void {
  // Object table, diagnostics and output are player-owned, not saved DC state.
  type SavedDc = Omit<PlayState, 'objects' | 'unsupported' | 'drew' | 'ctx' | 'clipDirty'>;
  const stack: SavedDc[] = [];
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

    // All drawing records share the same current device-region clip.
    if (fn === META.LINETO || fn === META.POLYLINE || fn === META.POLYGON ||
        fn === META.POLYPOLYGON || fn === META.RECTANGLE || fn === META.TEXTOUT ||
        fn === META.EXTTEXTOUT || fn === META.STRETCHDIBITS) applyClip(s);

    switch (fn) {
      case META.SETMAPMODE: {
        if (c.remaining !== 2) { s.unsupported.add('META_SETMAPMODE malformed'); break; }
        const mode = c.u16();
        if (mode < 1 || mode > 8) { s.unsupported.add('META_SETMAPMODE invalid mode'); break; }
        setMapMode(s, mode);
        break;
      }
      case META.SAVEDC: {
        if (c.remaining !== 0) { s.unsupported.add('META_SAVEDC malformed'); break; }
        const { objects: _objects, unsupported: _unsupported, drew: _drew, ctx: _ctx, clipDirty: _clipDirty, ...dc } = s;
        stack.push(dc);
        break;
      }
      case META.RESTOREDC: {
        if (c.remaining !== 2) { s.unsupported.add('META_RESTOREDC malformed'); break; }
        const level = c.i16();
        const index = level < 0 ? stack.length + level : level - 1;
        if (level === 0 || index < 0 || index >= stack.length) {
          s.unsupported.add('META_RESTOREDC invalid level'); break;
        }
        Object.assign(s, stack[index]);
        stack.length = index;
        s.clipDirty = true;
        break;
      }
      case META.CREATEREGION: {
        const region = createRegion(c);
        insertObject(s.objects, region);
        if (!region.valid) s.unsupported.add('META_CREATEREGION malformed');
        break;
      }
      case META.SELECTCLIPREGION: {
        const region = c.remaining === 2 ? s.objects[c.u16()] : undefined;
        if (region?.kind !== 'region' || !region.valid) {
          s.unsupported.add('META_SELECTCLIPREGION invalid region'); break;
        }
        s.clip = { tree: region.tree,
          mapping: [s.vpW / s.extX, s.vpH / s.extY, mapX(s, 0), mapY(s, 0)], dx: 0, dy: 0 };
        s.clipDirty = true;
        break;
      }
      case META.INTERSECTCLIPRECT:
      case META.EXCLUDECLIPRECT: {
        if (c.remaining !== 8) { s.unsupported.add('WMF clip rectangle malformed'); break; }
        const b = c.i16(); const r = c.i16(); const t = c.i16(); const l = c.i16();
        combineClip(s, mappedRect(s, l, t, r, b), fn === META.EXCLUDECLIPRECT);
        break;
      }
      case META.OFFSETCLIPRGN: {
        if (c.remaining !== 4) { s.unsupported.add('META_OFFSETCLIPRGN malformed'); break; }
        const y = c.i16(); const x = c.i16();
        if (s.clip !== UNBOUNDED_CLIP) {
          s.clip = { ...s.clip, dx: s.clip.dx + x * s.vpW / s.extX, dy: s.clip.dy + y * s.vpH / s.extY };
          s.clipDirty = true;
        }
        break;
      }
      case META.SETVIEWPORTORG:
      case META.OFFSETVIEWPORTORG:
      case META.OFFSETWINDOWORG:
      case META.SETVIEWPORTEXT: {
        if (c.remaining !== 4) { s.unsupported.add('WMF viewport mapping malformed'); break; }
        const y = c.i16(); const x = c.i16();
        if (fn === META.SETVIEWPORTORG) { s.vpX = x; s.vpY = y; }
        else if (fn === META.OFFSETVIEWPORTORG) { s.vpX += x; s.vpY += y; }
        else if (fn === META.OFFSETWINDOWORG) { s.orgX += x; s.orgY += y; }
        else if (s.mapMode >= 7 && x !== 0 && y !== 0) {
          s.vpW = x; s.vpH = y; adjustIsotropic(s);
        }
        break;
      }
      case META.SCALEWINDOWEXT:
      case META.SCALEVIEWPORTEXT: {
        if (c.remaining !== 8) { s.unsupported.add('WMF mapping scale malformed'); break; }
        const yd = c.i16(); const yn = c.i16(); const xd = c.i16(); const xn = c.i16();
        if (s.mapMode <= 6) break;
        if (xd === 0 || yd === 0 || xn === 0 || yn === 0) { s.unsupported.add('WMF mapping scale invalid'); break; }
        if (fn === META.SCALEWINDOWEXT) { s.extX *= xn / xd; s.extY *= yn / yd; }
        else { s.vpW *= xn / xd; s.vpH *= yn / yd; }
        adjustIsotropic(s);
        break;
      }
      case META.SETWINDOWORG: {
        // params: i16 yOrg, i16 xOrg (Y FIRST)
        if (c.remaining !== 4) {
          s.extTextSafe = false;
          s.currentPositionKnown = false;
          s.lineStateSafe = false;
          break;
        }
        s.orgY = c.i16(); s.orgX = c.i16();
        break;
      }
      case META.SETWINDOWEXT: {
        // params: i16 yExt, i16 xExt (Y FIRST)
        if (c.remaining !== 4) {
          s.extTextSafe = false;
          s.currentPositionKnown = false;
          s.lineStateSafe = false;
          break;
        }
        const yExt = c.i16();
        const xExt = c.i16();
        if (s.mapMode <= 6) break;
        s.extTextMappingSafe = xExt > 0 && yExt > 0;
        s.extY = yExt || 1;
        s.extX = xExt || 1;
        s.haveExt = true;
        adjustIsotropic(s);
        break;
      }
      case META.SETROP2: {
        // MS-WMF 2.3.5.22 / 2.1.1.2: R2_NOP is D (destination unchanged),
        // including pens and filled interiors. It is an authored no-op, not a
        // boundary classifier. MOVETO/LINETO still update the DC position.
        if (c.remaining !== 2 && c.remaining !== 4) {
          s.rop2 = 0;
          s.unsupported.add('META_SETROP2 malformed');
          break;
        }
        s.rop2 = c.u16();
        if (c.remaining === 2) c.u16(); // Optional Reserved MUST be ignored.
        // Canvas source-over supplies COPYPEN, while NOP needs no raster work.
        // Other Boolean RGB mixes need a bitwise compositor (Canvas alpha XOR
        // is not RGB XOR). Diagnose that existing limitation instead of silently
        // painting them as COPYPEN; strict metafile callers can reject it.
        if (s.rop2 !== 11 && s.rop2 !== 13) s.unsupported.add('META_SETROP2 (unsupported binary mix)');
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
        if (c.remaining !== 2 && c.remaining !== 4) {
          s.extTextSafe = false;
          s.lineStateSafe = false;
          break;
        }
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
        if (c.remaining !== 4) { s.extTextSafe = false; break; }
        s.bkColor = colorRefToCss(c.u32());
        break;
      }
      case META.SETTEXTCHAREXTRA: {
        // MS-WMF 2.3.5.25: u16 CharExtra, logical units added per character.
        if (c.remaining !== 2) { s.extTextSafe = false; break; }
        s.charExtra = c.u16();
        break;
      }
      case META.SETTEXTJUSTIFICATION: {
        // MS-WMF 2.3.5.27: u16 BreakCount, then u16 BreakExtra (logical units
        // shared by the break characters of following text output).
        if (c.remaining !== 4) { s.extTextSafe = false; break; }
        // GDI transforms BreakExtra to device units, rounds it, and shares it
        // over the break characters: each gets the quotient, and the error
        // (remainder) is paid one unit at a time by successive break
        // characters, across text outputs, until the next justification.
        s.breakCount = c.u16();
        const breakExtraPx = Math.round(c.u16() * Math.abs(s.vpW / s.extX));
        s.breakQuotientPx = s.breakCount > 0 && Number.isFinite(breakExtraPx) ? Math.trunc(breakExtraPx / s.breakCount) : 0;
        s.breakErrorPx = s.breakCount > 0 && Number.isFinite(breakExtraPx) ? breakExtraPx - s.breakQuotientPx * s.breakCount : 0;
        break;
      }
      case META.CREATEPALETTE:
      case META.CREATEPATTERNBRUSH:
      case META.DIBCREATEPATTERNBRUSH: {
        // MS-WMF 3.1.4.1: every object record takes the lowest free slot.
        insertObject(s.objects, { kind: 'other' });
        break;
      }
      case META.CREATEPENINDIRECT: {
        if (c.remaining !== 10) s.lineStateSafe = false;
        insertObject(s.objects, createPen(c));
        break;
      }
      case META.CREATEBRUSHINDIRECT: {
        if (c.remaining !== 8) s.lineStateSafe = false;
        insertObject(s.objects, createBrush(c));
        break;
      }
      case META.CREATEFONTINDIRECT: {
        // MS-WMF 2.2.1.2: FaceName is NUL-terminated within at most 32 bytes,
        // so writers may store fewer than 50 parameter bytes. Every record
        // takes an object slot; a malformed font is reported when used.
        const created = parseWmfFont(c.bytes(c.remaining));
        if (!created.valid) s.lineStateSafe = false;
        insertObject(s.objects, created);
        break;
      }
      case META.SELECTOBJECT: {
        if (c.remaining !== 2) {
          s.extTextSafe = false;
          s.lineStateSafe = false;
          break;
        }
        const idx = c.u16();
        const obj = s.objects[idx];
        if (obj?.kind === 'pen') s.curPen = obj;
        else if (obj?.kind === 'brush') s.curBrush = obj;
        else if (obj?.kind === 'font') s.curFont = obj;
        else if (obj?.kind === 'other') s.lineStateSafe = false;
        else {
          s.extTextSafe = false;
          s.lineStateSafe = false;
        }
        break;
      }
      case META.DELETEOBJECT: {
        if (c.remaining !== 2) {
          s.extTextSafe = false;
          s.lineStateSafe = false;
          break;
        }
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
      case META.MOVETO: {
        // MS-WMF 2.3.5.4: params are i16 Y followed by i16 X.
        if (c.remaining !== 4) {
          s.extTextSafe = false;
          s.currentPositionKnown = false;
          break;
        }
        s.currentY = c.i16();
        s.currentX = c.i16();
        s.currentPositionKnown = true;
        break;
      }
      case META.LINETO: {
        // MS-WMF 2.3.3.10: params are i16 Y followed by i16 X.
        if (c.remaining !== 4) {
          s.extTextSafe = false;
          s.currentPositionKnown = false;
          break;
        }
        const y = c.i16();
        const x = c.i16();
        strokeLineTo(s, x, y);
        break;
      }
      case META.TEXTOUT: {
        // META_TEXTOUT params: u16 StringLength, String bytes padded to a WORD
        // boundary, then i16 yStart and i16 xStart.
        if (c.remaining < 2) { s.unsupported.add('META_TEXTOUT (malformed)'); break; }
        const len = c.u16();
        if (c.remaining < len + (len & 1) + 4) { s.unsupported.add('META_TEXTOUT (malformed)'); break; }
        const bytes = c.bytes(len);
        if (len % 2 !== 0) c.skip(1);
        const y = c.i16();
        const x = c.i16();
        drawText(s, 'META_TEXTOUT', { x, y, bytes, options: 0, rect: null, dx: null });
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
          s.currentPositionKnown = false;
          s.lineStateSafe = false;
          break;
        }
        const data = c.bytes(count);
        if (count & 1) c.skip(1);
        // META_ESCAPE_ENHANCED_METAFILE begins with the little-endian WMFC
        // CommentIdentifier (MS-WMF 2.3.6.25); nested graphics state is outside
        // this player. Historical Microsoft SDK Q81497 documents non-WMFC
        // private comments as ignored during ordinary playback (no DC change).
        if (data.length >= 4 && new DataView(data.buffer, data.byteOffset, data.byteLength).getUint32(0, true) === 0x43464d57) {
          s.extTextSafe = false;
          s.currentPositionKnown = false;
          s.lineStateSafe = false;
        }
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
        // MS-WMF 2.1.1.1 records that neither create objects nor change the DC
        // text, line or mapping state (drawing-only operations, palette
        // realization, font-mapper hints, the reserved SETRELABS) are skipped
        // by size without affecting later records.
        if (STATE_NEUTRAL_RECORDS.has(fn)) break;
        // Unsupported graphics-state changes make later text output unsafe for
        // this bounded player, although established geometry handlers remain
        // unaffected and the record is still skipped by its declared size.
        s.extTextSafe = false;
        s.currentPositionKnown = false;
        s.lineStateSafe = false;
        break;
    }

    pos = recEnd;
  }
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
 */
export async function renderWmfToBitmap(
  bytes: Uint8Array,
  targetW: number,
  targetH: number,
  onUnsupported?: (records: readonly string[]) => void,
  fontFamilies?: ReadonlyMap<string, string>,
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
  const drew = playWmf(bytes, ctx, targetW, targetH, onUnsupported, fontFamilies);
  if (!drew) return null;
  return createImageBitmap(canvas);
}
