// ── EMF (Enhanced Metafile) player ───────────────────────────────────────────
//
// Browsers cannot decode EMF via `createImageBitmap`, so the renderer falls back
// to this player for `.emf` blips — the EMF twin of {@link ./wmf.ts}. It is a
// *minimal but spec-faithful* [MS-EMF] interpreter: enough to rasterize the
// vector charts/diagrams Office embeds (e.g. sample-13.docx `word/media/
// image3.emf` / `image4.emf`, which carry their bars/axes as POLYGON16 /
// POLYLINE16 records and their labels as EXTTEXTOUTW text-out records, all
// scaled by a long run of MODIFYWORLDTRANSFORM affines).
//
// Format reference: ECMA-376 references WMF/EMF; the byte layout below follows
// the [MS-EMF] Enhanced Metafile Format spec.
//   - File = a sequence of records, each `u32 iType, u32 nSize`, then nSize−8
//     data bytes. nSize is 4-byte aligned; walk `offset += nSize`. The first
//     record is EMR_HEADER (iType=1); the last is EMR_EOF (iType=14).
//   - All values little-endian. COLORREF = u32 0x00BBGGRR (low byte = R), shared
//     byte layout with WMF (we reuse {@link colorRefToCss}).
//   - The full [MS-EMF] coordinate pipeline is world → page → device → target:
//       1. WORLD → PAGE: the world transform (EMR_SETWORLDTRANSFORM /
//          EMR_MODIFYWORLDTRANSFORM, [MS-EMF] 2.3.12), a 2×3 affine. Getting the
//          multiply order right ([MS-EMF] 2.3.12: MWT_LEFTMULTIPLY ⇒ the
//          supplied XFORM is the LEFT operand, `newWT = xform × WT`) is what
//          makes bars/axes/labels land in the right place in world-scaled files.
//       2. PAGE → DEVICE: the window→viewport mapping (SETMAPMODE +
//          SET{WINDOW,VIEWPORT}{ORG,EXT}EX + SCALE{WINDOW,VIEWPORT}EXTEX,
//          [MS-EMF] 2.3.11 / 2.1.21 MapMode). GDI computes device coords as
//          `D = (P − winOrg)·(vpExt/winExt) + vpOrg`. Under the default MM_TEXT
//          (winOrg=vpOrg=0, winExt=vpExt=1) this is the identity, so files that
//          scale purely with the world transform are unaffected; files that set
//          MM_ANISOTROPIC/ISOTROPIC + explicit window/viewport extents (common
//          for charts pasted as EMF, e.g. from Excel/Visio into pptx) map
//          correctly instead of collapsing to a blank/off-canvas raster.
//       3. DEVICE → TARGET: the EMR_HEADER rclFrame/rclBounds rectangle scaled
//          onto the W×H raster (see the HEADER case).
//
// Implemented records: HEADER, SETMAPMODE, SETWINDOWORGEX/SETWINDOWEXTEX,
// SETVIEWPORTORGEX/SETVIEWPORTEXTEX, SCALEWINDOWEXTEX/SCALEVIEWPORTEXTEX,
// SETWORLDTRANSFORM, MODIFYWORLDTRANSFORM,
// SAVEDC/RESTOREDC, SELECTOBJECT (incl. stock objects), DELETEOBJECT,
// CREATEPEN, EXTCREATEPEN, CREATEBRUSHINDIRECT, CREATEMONOBRUSH /
// CREATEDIBPATTERNBRUSHPT (→ average solid color), EXTCREATEFONTINDIRECTW,
// POLYLINE16/POLYGON16/POLYBEZIER16/POLYLINETO16/POLYBEZIERTO16 (+ their 32-bit
// twins), POLYPOLYGON16/POLYPOLYLINE16 (+ 32-bit twins), MOVETOEX, LINETO,
// RECTANGLE, ELLIPSE, SETPOLYFILLMODE, EXTTEXTOUTW, SETTEXTCOLOR, SETTEXTALIGN,
// SETBKMODE, BITBLT, STRETCHDIBITS (minimal DIB decoder), EOF.
// GDI path brackets are handled: BEGINPATH/ENDPATH/CLOSEFIGURE accumulate
// geometry for FILLPATH/STROKEPATH/STROKEANDFILLPATH or SELECTCLIPPATH (the
// clip remains scoped by SAVEDC/RESTOREDC).
// Ignored (no-op, skipped by nSize): GDICOMMENT (may hold EMF+, out of scope),
// SETICMMODE, SETMITERLIMIT, SETROP2, SETSTRETCHBLTMODE, INTERSECTCLIPRECT,
// and any unrecognized iType.
//
// Shared across the docx, pptx and xlsx renderers via
// {@link ./raster-or-metafile.ts}#decodeRasterOrMetafile, which sniffs the bytes and routes
// true EMF here.

import { decodeDib, blitDibToCtx, type DecodedDib } from './dib.js';
import { colorRefToCss, isEmf } from './wmf.js';
import { createAuxCanvas } from '../canvas/aux-canvas.js';

// EMF record type codes ([MS-EMF] 2.1.1 EMR enumeration; the subset we act on,
// others are skipped by nSize).
const EMR = {
  HEADER: 1,
  POLYBEZIER: 2,
  POLYGON: 3,
  POLYLINE: 4,
  POLYBEZIERTO: 5,
  POLYLINETO: 6,
  POLYPOLYLINE: 7,
  POLYPOLYGON: 8,
  SETWINDOWEXTEX: 9,
  SETWINDOWORGEX: 10,
  SETVIEWPORTEXTEX: 11,
  SETVIEWPORTORGEX: 12,
  EOF: 14,
  SETMAPMODE: 17,
  SETPOLYFILLMODE: 19,
  SETBKMODE: 18,
  SETTEXTALIGN: 22,
  SETTEXTCOLOR: 24,
  MOVETOEX: 27,
  SCALEVIEWPORTEXTEX: 31,
  SCALEWINDOWEXTEX: 32,
  SAVEDC: 33,
  RESTOREDC: 34,
  SETWORLDTRANSFORM: 35,
  MODIFYWORLDTRANSFORM: 36,
  SELECTOBJECT: 37,
  CREATEPEN: 38,
  CREATEBRUSHINDIRECT: 39,
  DELETEOBJECT: 40,
  ELLIPSE: 42,
  RECTANGLE: 43,
  LINETO: 54,
  BEGINPATH: 59,
  ENDPATH: 60,
  CLOSEFIGURE: 61,
  FILLPATH: 62,
  STROKEANDFILLPATH: 63,
  STROKEPATH: 64,
  SELECTCLIPPATH: 67,
  EXTCREATEFONTINDIRECTW: 82,
  EXTTEXTOUTW: 84,
  POLYBEZIER16: 85,
  POLYGON16: 86,
  POLYLINE16: 87,
  POLYBEZIERTO16: 88,
  POLYLINETO16: 89,
  POLYPOLYLINE16: 90,
  POLYPOLYGON16: 91,
  CREATEMONOBRUSH: 93,
  CREATEDIBPATTERNBRUSHPT: 94,
  EXTCREATEPEN: 95,
  BITBLT: 76,
  STRETCHDIBITS: 81,
} as const;

// Stock object handle ids ([MS-EMF] 2.1.31 StockObject) — high bit 0x80000000
// set in a SELECTOBJECT handle.
const STOCK = {
  WHITE_BRUSH: 0x80000000,
  LTGRAY_BRUSH: 0x80000001,
  GRAY_BRUSH: 0x80000002,
  DKGRAY_BRUSH: 0x80000003,
  BLACK_BRUSH: 0x80000004,
  NULL_BRUSH: 0x80000005,
  WHITE_PEN: 0x80000006,
  BLACK_PEN: 0x80000007,
  NULL_PEN: 0x80000008,
  DC_BRUSH: 0x80000012,
  DC_PEN: 0x8000000e,
} as const;

// MapMode enumeration ([MS-EMF] 2.1.21). The metric modes below express a fixed
// physical size per logical unit; GDI derives their window/viewport extents from
// the reference-device resolution in the EMR_HEADER. MM_TEXT is 1 device px per
// logical unit (y-down); the metric modes are y-UP (viewport Y extent negative).
const MM = {
  TEXT: 1,
  LOMETRIC: 2, // 0.1 mm per unit
  HIMETRIC: 3, // 0.01 mm per unit
  LOENGLISH: 4, // 0.01 inch per unit
  HIENGLISH: 5, // 0.001 inch per unit
  TWIPS: 6, // 1/1440 inch per unit
  ISOTROPIC: 7,
  ANISOTROPIC: 8,
} as const;

// ── color ─────────────────────────────────────────────────────────────────
// COLORREF → CSS is shared with WMF (identical 0x00BBGGRR layout); imported as
// `colorRefToCss`.

// ── object table ─────────────────────────────────────────────────────────

interface Pen {
  kind: 'pen';
  stroke: string | null; // null = PS_NULL (no stroke)
  width: number; // logical width; mapped to device px via world+device scale
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
  escapement: number; // lfEscapement — tenths of a degree, counterclockwise
}
type EmfObject = Pen | Brush | Font;

// ── 2×3 affine world transform ([MS-EMF] 2.2.28 XFORM) ──────────────────────

interface Xform {
  m11: number;
  m12: number;
  m21: number;
  m22: number;
  dx: number;
  dy: number;
}

const identity = (): Xform => ({ m11: 1, m12: 0, m21: 0, m22: 1, dx: 0, dy: 0 });

/**
 * Affine product `A × B`, treating each 2×3 as a 3×3 with bottom row [0,0,1]
 * ([MS-EMF] 2.3.12). For MWT_LEFTMULTIPLY the supplied XFORM is the LEFT
 * operand (`newWT = xform × WT`); for MWT_RIGHTMULTIPLY it is the RIGHT operand.
 */
function mulXform(A: Xform, B: Xform): Xform {
  return {
    m11: A.m11 * B.m11 + A.m21 * B.m12,
    m12: A.m12 * B.m11 + A.m22 * B.m12,
    m21: A.m11 * B.m21 + A.m21 * B.m22,
    m22: A.m12 * B.m21 + A.m22 * B.m22,
    dx: A.m11 * B.dx + A.m21 * B.dy + A.dx,
    dy: A.m12 * B.dx + A.m22 * B.dy + A.dy,
  };
}

// ── little-endian cursor over a record's data region ────────────────────────
//
// EMF is 32-bit, so the primitives are i32/u32/f32 (plus i16 for POINTS).

class EmfCursor {
  private p: number;
  constructor(
    private readonly dv: DataView,
    start: number,
    private readonly end: number, // exclusive
  ) {
    this.p = start;
  }
  get pos(): number {
    return this.p;
  }
  set pos(v: number) {
    this.p = v;
  }
  get remaining(): number {
    return this.end - this.p;
  }
  i16(): number {
    const v = this.dv.getInt16(this.p, true);
    this.p += 2;
    return v;
  }
  i32(): number {
    const v = this.dv.getInt32(this.p, true);
    this.p += 4;
    return v;
  }
  u32(): number {
    const v = this.dv.getUint32(this.p, true);
    this.p += 4;
    return v;
  }
  f32(): number {
    const v = this.dv.getFloat32(this.p, true);
    this.p += 4;
    return v;
  }
  /** Read a 2×3 XFORM (6×f32). */
  xform(): Xform {
    return {
      m11: this.f32(),
      m12: this.f32(),
      m21: this.f32(),
      m22: this.f32(),
      dx: this.f32(),
      dy: this.f32(),
    };
  }
  skip(n: number): void {
    this.p += n;
  }
}

// ── any 2D context we can replay onto (Offscreen or HTMLCanvas) ─────────────

type AnyCtx = CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;

interface PlayState {
  ctx: AnyCtx;
  W: number; // target raster width (px)
  H: number; // target raster height (px)
  // device extent from EMR_HEADER rclBounds ([MS-EMF] 2.2.9)
  left: number;
  top: number;
  boundsW: number;
  boundsH: number;
  // GDI state
  wt: Xform; // current world transform (world → page)
  // window→viewport mapping ([MS-EMF] 2.3.11): page → device via
  // `D = (P − winOrg)·(vpExt/winExt) + vpOrg`. MM_TEXT default is the identity
  // (winOrg=vpOrg=0, winExt=vpExt=1), so world-scaled files are unchanged.
  mapMode: number;
  winOrgX: number;
  winOrgY: number;
  winExtX: number;
  winExtY: number;
  vpOrgX: number;
  vpOrgY: number;
  vpExtX: number;
  vpExtY: number;
  // Reference-device resolution from the EMR_HEADER, used to derive fixed
  // window/viewport extents for the metric map modes ([MS-EMF] 2.1.21).
  devPxPerMmX: number;
  devPxPerMmY: number;
  objects: Map<number, EmfObject>; // indexed by ihObject
  curPen: Pen | null;
  curBrush: Brush | null;
  curFont: Font | null;
  textColor: string;
  bkMode: number; // 1 = TRANSPARENT
  textAlign: number;
  fillRule: CanvasFillRule;
  curX: number; // current position (logical)
  curY: number;
  stack: SavedDc[]; // SAVEDC/RESTOREDC graphics-state stack
  drew: boolean;
  inPath: boolean; // between BEGINPATH and ENDPATH — geometry builds a path, no draw
  pathCommandCount: number;
  pathDiscarded: boolean;
  maxPathCommands: number;
}

/**
 * Maximum Canvas path commands retained by one EMF path bracket. This preserves
 * every previously admitted single geometry record: POLYPOLYGON may contain up
 * to 0x200000 points in at most 0x10000 closed sub-paths. The cumulative bound
 * prevents many individually valid records from multiplying that retained work.
 */
const MAX_EMF_PATH_COMMANDS = 0x200000 + 0x10000;

interface EmfReplayLimits {
  /** Package-internal override used by focused boundary tests. */
  maxPathCommands?: number;
}

function reservePathCommands(s: PlayState, additional: number): boolean {
  if (!s.inPath) return true;
  if (s.pathDiscarded) return false;
  const next = s.pathCommandCount + additional;
  if (!Number.isSafeInteger(next) || additional < 0 || next > s.maxPathCommands) {
    // Never leave a paintable prefix of an over-budget attacker-controlled path.
    s.ctx.beginPath();
    s.pathDiscarded = true;
    return false;
  }
  s.pathCommandCount = next;
  return true;
}

function readablePointCount(c: EmfCursor, count: number, rp: PointReader): number {
  const bytesPerPoint = rp === readPoint16 ? 4 : 8;
  return Math.min(count, Math.floor(c.remaining / bytesPerPoint));
}

/** Snapshot of the graphics state pushed by EMR_SAVEDC. */
interface SavedDc {
  wt: Xform;
  mapMode: number;
  winOrgX: number;
  winOrgY: number;
  winExtX: number;
  winExtY: number;
  vpOrgX: number;
  vpOrgY: number;
  vpExtX: number;
  vpExtY: number;
  curPen: Pen | null;
  curBrush: Brush | null;
  curFont: Font | null;
  textColor: string;
  bkMode: number;
  textAlign: number;
  fillRule: CanvasFillRule;
  curX: number;
  curY: number;
}

// ── coordinate pipeline ─────────────────────────────────────────────────────
//
// [MS-EMF] maps world → page → device → target:
//   1. world → PAGE via the world transform ([MS-EMF] 2.2.28 XFORM):
//        `Xp = m11·Xw + m21·Yw + dx`, `Yp = m12·Xw + m22·Yw + dy`
//   2. page → DEVICE via the window→viewport mapping ([MS-EMF] 2.3.11):
//        `Xd = (Xp − winOrgX)·(vpExtX/winExtX) + vpOrgX` (and likewise Y)
//   3. device → TARGET via the header frame/bounds → W×H raster scale.

/** world → page X (world transform only). */
function pageX(s: PlayState, xl: number, yl: number): number {
  return s.wt.m11 * xl + s.wt.m21 * yl + s.wt.dx;
}
/** world → page Y (world transform only). */
function pageY(s: PlayState, xl: number, yl: number): number {
  return s.wt.m12 * xl + s.wt.m22 * yl + s.wt.dy;
}

/** Raw page → device scale along X (`vpExtX/winExtX`), before any isotropic
 *  aspect correction; 1 under MM_TEXT. */
function rawPageScaleX(s: PlayState): number {
  return s.winExtX !== 0 ? s.vpExtX / s.winExtX : 1;
}
/** Raw page → device scale along Y (`vpExtY/winExtY`); 1 under MM_TEXT. */
function rawPageScaleY(s: PlayState): number {
  return s.winExtY !== 0 ? s.vpExtY / s.winExtY : 1;
}
/**
 * The common magnitude both axes use under MM_ISOTROPIC ([MS-EMF] 2.1.21):
 * "one unit along the x-axis is equal to one unit along the y-axis". GDI
 * realises that by shrinking the axis with the larger extent ratio down to the
 * SMALLER of the two |vpExt/winExt| ratios, so the picture fits its viewport
 * without distortion. Returns the shared |scale|; the callers reapply each
 * axis's sign so the y-flip (negative viewport extent) is preserved.
 */
function isotropicMagnitude(s: PlayState): number {
  return Math.min(Math.abs(rawPageScaleX(s)), Math.abs(rawPageScaleY(s)));
}
/** page → device scale along X. Under MM_ISOTROPIC both axes share the smaller
 *  |ratio| (equal-aspect, §2.1.21); otherwise the raw x ratio. */
function pageScaleX(s: PlayState): number {
  if (s.mapMode === MM.ISOTROPIC) {
    const raw = rawPageScaleX(s);
    return raw < 0 ? -isotropicMagnitude(s) : isotropicMagnitude(s);
  }
  return rawPageScaleX(s);
}
/** page → device scale along Y. Under MM_ISOTROPIC both axes share the smaller
 *  |ratio| (equal-aspect, §2.1.21); otherwise the raw y ratio. */
function pageScaleY(s: PlayState): number {
  if (s.mapMode === MM.ISOTROPIC) {
    const raw = rawPageScaleY(s);
    return raw < 0 ? -isotropicMagnitude(s) : isotropicMagnitude(s);
  }
  return rawPageScaleY(s);
}
/** page → device X ([MS-EMF] 2.3.11). */
function deviceX(s: PlayState, xp: number): number {
  return (xp - s.winOrgX) * pageScaleX(s) + s.vpOrgX;
}
/** page → device Y ([MS-EMF] 2.3.11). */
function deviceY(s: PlayState, yp: number): number {
  return (yp - s.winOrgY) * pageScaleY(s) + s.vpOrgY;
}

/** logical (world) point → target px: world transform, then window→viewport
 *  page→device, then device→target. */
function toPx(s: PlayState, xl: number, yl: number): [number, number] {
  const Xd = deviceX(s, pageX(s, xl, yl));
  const Yd = deviceY(s, pageY(s, xl, yl));
  const px = ((Xd - s.left) * s.W) / s.boundsW;
  const py = ((Yd - s.top) * s.H) / s.boundsH;
  return [px, py];
}

/** Average world scale magnitude (column-vector lengths) — the world→page part
 *  of the pen/font scale. */
function worldScale(s: PlayState): number {
  const sx = Math.hypot(s.wt.m11, s.wt.m12);
  const sy = Math.hypot(s.wt.m21, s.wt.m22);
  return (sx + sy) / 2;
}
/** Average page→device scale magnitude (window→viewport); 1 under MM_TEXT. */
function pageScale(s: PlayState): number {
  return (Math.abs(pageScaleX(s)) + Math.abs(pageScaleY(s))) / 2;
}
/** Average device→target scale (px per device unit). */
function deviceScale(s: PlayState): number {
  return (s.W / s.boundsW + s.H / s.boundsH) / 2;
}
/** Y-only world scale magnitude (Y column length) — for font px sizing. */
function worldScaleY(s: PlayState): number {
  return Math.hypot(s.wt.m21, s.wt.m22);
}
/** Y-only device→target scale — for font px sizing. */
function deviceScaleY(s: PlayState): number {
  return s.H / s.boundsH;
}

/** Device line width: scale the pen's logical width through world → page →
 *  device → target and clamp to ≥0.75 so hairlines stay visible. */
function deviceLineWidth(s: PlayState, logicalWidth: number): number {
  const w = logicalWidth * worldScale(s) * pageScale(s) * deviceScale(s);
  return Math.max(0.75, w);
}

/**
 * Apply EMR_SETMAPMODE ([MS-EMF] 2.3.11 / MapMode enumeration 2.1.21).
 *
 * MM_TEXT / MM_ANISOTROPIC / MM_ISOTROPIC leave the current window & viewport
 * origins/extents in place (the app sets them with the SET*ORGEX/SET*EXTEX
 * records). The five metric modes impose a fixed physical scale — one logical
 * unit is a fixed fraction of an inch or millimetre — with a y-UP axis; GDI
 * realizes that by fixed window/viewport extents derived from the reference
 * device resolution (px per mm from the EMR_HEADER). We compute those extents
 * so the page→device stage reproduces the physical scale. Origins reset to 0.
 */
function applyMapMode(s: PlayState, mode: number): void {
  s.mapMode = mode;
  if (mode === MM.TEXT) {
    // 1 logical unit = 1 device px, y-down. Reset to the identity mapping.
    s.winOrgX = 0;
    s.winOrgY = 0;
    s.vpOrgX = 0;
    s.vpOrgY = 0;
    s.winExtX = 1;
    s.winExtY = 1;
    s.vpExtX = 1;
    s.vpExtY = 1;
    return;
  }
  if (mode === MM.ANISOTROPIC || mode === MM.ISOTROPIC) {
    // The application supplies window/viewport extents explicitly; keep the
    // current ones (default 1:1 until a SET*EXTEX record arrives). ANISOTROPIC
    // scales each axis independently; ISOTROPIC forces both axes to the SMALLER
    // |vpExt/winExt| ratio (equal-aspect, [MS-EMF] 2.1.21) — that correction is
    // applied in pageScaleX/pageScaleY (gated on `mapMode === MM.ISOTROPIC`), so
    // it tracks any later SET*EXTEX / SCALE*EXTEX without recomputation here.
    return;
  }
  // Metric modes: a fixed physical unit per logical unit, y-UP. Realize as a
  // window extent of `unitMm` logical units mapping to `devPxPerMm` device px,
  // i.e. vpExt/winExt = px per logical unit. Needs the header device resolution;
  // if it is unknown, leave the mapping untouched (degrade to world-transform).
  if (s.devPxPerMmX <= 0 || s.devPxPerMmY <= 0) return;
  const MM_PER_INCH = 25.4;
  // Physical size, in millimetres, of ONE logical unit for each metric mode.
  const unitMm =
    mode === MM.LOMETRIC
      ? 0.1
      : mode === MM.HIMETRIC
        ? 0.01
        : mode === MM.LOENGLISH
          ? 0.01 * MM_PER_INCH
          : mode === MM.HIENGLISH
            ? 0.001 * MM_PER_INCH
            : mode === MM.TWIPS
              ? MM_PER_INCH / 1440
              : 0;
  if (unitMm <= 0) return;
  s.winOrgX = 0;
  s.winOrgY = 0;
  s.vpOrgX = 0;
  s.vpOrgY = 0;
  // 1 logical unit → unitMm mm → unitMm·devPxPerMm device px. Encode as
  // winExt=1, vpExt=px-per-unit. Viewport Y is negated: metric modes are y-UP
  // while device space is y-down ([MS-EMF] 2.1.21).
  s.winExtX = 1;
  s.winExtY = 1;
  s.vpExtX = unitMm * s.devPxPerMmX;
  s.vpExtY = -(unitMm * s.devPxPerMmY);
}

// ── stock objects ([MS-EMF] 2.1.31) ─────────────────────────────────────────

const STOCK_BRUSH: Record<number, Brush> = {
  [STOCK.WHITE_BRUSH]: { kind: 'brush', fill: '#ffffff' },
  [STOCK.LTGRAY_BRUSH]: { kind: 'brush', fill: '#c0c0c0' },
  [STOCK.GRAY_BRUSH]: { kind: 'brush', fill: '#808080' },
  [STOCK.DKGRAY_BRUSH]: { kind: 'brush', fill: '#404040' },
  [STOCK.BLACK_BRUSH]: { kind: 'brush', fill: '#000000' },
  [STOCK.NULL_BRUSH]: { kind: 'brush', fill: null },
};
const STOCK_PEN: Record<number, Pen> = {
  [STOCK.WHITE_PEN]: { kind: 'pen', stroke: '#ffffff', width: 1 },
  [STOCK.BLACK_PEN]: { kind: 'pen', stroke: '#000000', width: 1 },
  [STOCK.NULL_PEN]: { kind: 'pen', stroke: null, width: 1 },
  [STOCK.DC_PEN]: { kind: 'pen', stroke: '#000000', width: 1 },
};

/** Apply a stock-object handle to the current pen/brush. Unknown stock ids are
 *  a no-op (leave current). */
function selectStock(s: PlayState, handle: number): void {
  const brush = STOCK_BRUSH[handle];
  if (brush) {
    s.curBrush = brush;
    return;
  }
  const pen = STOCK_PEN[handle];
  if (pen) {
    s.curPen = pen;
    return;
  }
  if (handle === STOCK.DC_BRUSH) {
    // DC_BRUSH defaults to white; keep current if any, else fall back to black.
    s.curBrush = s.curBrush ?? { kind: 'brush', fill: '#000000' };
  }
  // Anything else: no-op.
}

// ── DIB decoder ([MS-WMF] 2.2.2.9 DeviceIndependentBitmap, BITMAPINFOHEADER) ──
//
// The decode + blit live in the shared {@link ./dib.ts} module (used by both the
// EMF and WMF players); imported here as {@link decodeDib} / {@link blitDibToCtx}
// / {@link DecodedDib}. `dibAverageColor` (EMF-only, for CREATEDIBPATTERNBRUSHPT
// → average solid color) stays local.

/** Average RGB of a decoded DIB (skipping fully transparent pixels) → CSS. */
function dibAverageColor(dib: DecodedDib): string {
  let r = 0;
  let g = 0;
  let b = 0;
  let n = 0;
  for (let i = 0; i < dib.data.length; i += 4) {
    if (dib.data[i + 3] === 0) continue;
    r += dib.data[i];
    g += dib.data[i + 1];
    b += dib.data[i + 2];
    n++;
  }
  if (n === 0) return '#808080';
  const hex = (v: number) =>
    Math.round(v / n)
      .toString(16)
      .padStart(2, '0');
  return `#${hex(r)}${hex(g)}${hex(b)}`;
}

// ── point readers (16-bit vs 32-bit) ────────────────────────────────────────

type PointReader = (c: EmfCursor) => [number, number];
const readPoint16: PointReader = (c) => [c.i16(), c.i16()];
const readPoint32: PointReader = (c) => [c.i32(), c.i32()];

// ── poly drawing ────────────────────────────────────────────────────────────

/** EMR_POLYLINE(16): open path stroked with the current pen. */
function strokePolyline(s: PlayState, c: EmfCursor, rp: PointReader): void {
  c.skip(16); // RECTL rclBounds — drawing uses world transform, not bounds
  const count = c.u32();
  if (count < 2 || count > 0x100000) return;
  if (!s.inPath && (!s.curPen || s.curPen.stroke == null)) {
    // still drop current position to the last point for ...TO continuity callers
    return;
  }
  const { ctx } = s;
  const appendPath = reservePathCommands(s, readablePointCount(c, count, rp));
  if (!s.inPath) ctx.beginPath();
  let lx = 0;
  let ly = 0;
  for (let i = 0; i < count; i++) {
    if (c.remaining < 4) break;
    const [xl, yl] = rp(c);
    const [px, py] = toPx(s, xl, yl);
    if (appendPath) {
      if (i === 0) ctx.moveTo(px, py);
      else ctx.lineTo(px, py);
    }
    lx = xl;
    ly = yl;
  }
  if (!s.inPath && s.curPen?.stroke != null) {
    ctx.strokeStyle = s.curPen.stroke;
    ctx.lineWidth = deviceLineWidth(s, s.curPen.width);
    ctx.stroke();
    s.drew = true;
  }
  s.curX = lx;
  s.curY = ly;
}

/** EMR_POLYLINETO(16): like POLYLINE but the implicit first point is the
 *  current position; updates current position. */
function strokePolylineTo(s: PlayState, c: EmfCursor, rp: PointReader): void {
  c.skip(16);
  const count = c.u32();
  if (count < 1 || count > 0x100000) return;
  const { ctx } = s;
  const draw = s.inPath || (s.curPen != null && s.curPen.stroke != null);
  const appendPath = !s.inPath || reservePathCommands(s, readablePointCount(c, count, rp));
  if (draw) {
    if (!s.inPath) {
      ctx.beginPath();
      const [px0, py0] = toPx(s, s.curX, s.curY);
      ctx.moveTo(px0, py0);
    }
  }
  for (let i = 0; i < count; i++) {
    if (c.remaining < 4) break;
    const [xl, yl] = rp(c);
    if (draw && appendPath) {
      const [px, py] = toPx(s, xl, yl);
      ctx.lineTo(px, py);
    }
    s.curX = xl;
    s.curY = yl;
  }
  if (draw && !s.inPath && s.curPen) {
    ctx.strokeStyle = s.curPen.stroke as string;
    ctx.lineWidth = deviceLineWidth(s, s.curPen.width);
    ctx.stroke();
    s.drew = true;
  }
}

/** EMR_POLYGON(16): closed path filled (brush) + stroked (pen). */
function fillStrokePolygon(s: PlayState, c: EmfCursor, rp: PointReader): void {
  c.skip(16);
  const count = c.u32();
  if (count < 2 || count > 0x100000) return;
  const { ctx } = s;
  const appendPath = reservePathCommands(s, readablePointCount(c, count, rp) + 1);
  if (!s.inPath) ctx.beginPath();
  let started = false;
  for (let i = 0; i < count; i++) {
    if (c.remaining < 4) break;
    const [xl, yl] = rp(c);
    const [px, py] = toPx(s, xl, yl);
    if (!started) {
      if (appendPath) ctx.moveTo(px, py);
      started = true;
    } else if (appendPath) ctx.lineTo(px, py);
  }
  if (!started) return;
  if (appendPath) ctx.closePath();
  if (s.inPath) return; // path bracket: defer fill/stroke
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

/** EMR_POLYBEZIER(16) / ...TO: cubic Bézier — start (or current pos for the
 *  ...TO variant), then triples of (control, control, end). Stroked open. */
function strokePolyBezier(
  s: PlayState,
  c: EmfCursor,
  rp: PointReader,
  isTo: boolean,
): void {
  c.skip(16);
  const count = c.u32();
  if (count < 1 || count > 0x100000) return;
  const readableCount = readablePointCount(c, count, rp);
  const appendPath = !s.inPath || reservePathCommands(s, readableCount);
  if (s.inPath && !appendPath) {
    for (let i = 0; i < readableCount; i++) {
      const [xl, yl] = rp(c);
      s.curX = xl;
      s.curY = yl;
    }
    return;
  }
  const pts: Array<[number, number]> = [];
  for (let i = 0; i < count; i++) {
    if (c.remaining < 4) break;
    pts.push(rp(c));
  }
  if (pts.length < (isTo ? 3 : 4)) {
    if (pts.length) {
      s.curX = pts[pts.length - 1][0];
      s.curY = pts[pts.length - 1][1];
    }
    return;
  }
  const draw = s.inPath || (s.curPen != null && s.curPen.stroke != null);
  const { ctx } = s;
  if (draw) {
    if (!s.inPath) ctx.beginPath();
    const start = isTo ? toPx(s, s.curX, s.curY) : toPx(s, pts[0][0], pts[0][1]);
    if (!isTo || !s.inPath) ctx.moveTo(start[0], start[1]);
  }
  let i = isTo ? 0 : 1;
  for (; i + 2 < pts.length + (isTo ? 1 : 0); i += 3) {
    const c1 = pts[i];
    const c2 = pts[i + 1];
    const end = pts[i + 2];
    if (!c1 || !c2 || !end) break;
    if (draw) {
      const p1 = toPx(s, c1[0], c1[1]);
      const p2 = toPx(s, c2[0], c2[1]);
      const pe = toPx(s, end[0], end[1]);
      ctx.bezierCurveTo(p1[0], p1[1], p2[0], p2[1], pe[0], pe[1]);
    }
    s.curX = end[0];
    s.curY = end[1];
  }
  if (draw && !s.inPath && s.curPen) {
    ctx.strokeStyle = s.curPen.stroke as string;
    ctx.lineWidth = deviceLineWidth(s, s.curPen.width);
    ctx.stroke();
    s.drew = true;
  }
}

/** EMR_POLYPOLYGON(16) / POLYPOLYLINE(16): one path spanning all sub-polys so
 *  the fill rule resolves holes correctly. Polygon variant fills+strokes;
 *  polyline variant only strokes. */
function fillStrokePolyPoly(
  s: PlayState,
  c: EmfCursor,
  rp: PointReader,
  isPolygon: boolean,
): void {
  c.skip(16); // RECTL rclBounds
  const numPolys = c.u32();
  const totalPoints = c.u32();
  if (numPolys <= 0 || numPolys > 0x10000) return;
  if (totalPoints <= 0 || totalPoints > 0x200000) return;
  const counts: number[] = [];
  let countedPoints = 0;
  for (let i = 0; i < numPolys; i++) {
    if (c.remaining < 4) return;
    const count = c.u32();
    countedPoints += count;
    if (!Number.isSafeInteger(countedPoints) || countedPoints > 0x200000) return;
    counts.push(count);
  }
  if (countedPoints !== totalPoints) return;
  const { ctx } = s;
  const appendPath = reservePathCommands(
    s,
    readablePointCount(c, totalPoints, rp) + (isPolygon ? numPolys : 0),
  );
  if (!s.inPath) ctx.beginPath(); // in a path bracket: accumulate, don't reset
  let any = false;
  for (const cnt of counts) {
    if (cnt < 2) {
      for (let i = 0; i < cnt && c.remaining >= 4; i++) rp(c);
      continue;
    }
    for (let i = 0; i < cnt; i++) {
      if (c.remaining < 4) break;
      const [xl, yl] = rp(c);
      if (appendPath) {
        const [px, py] = toPx(s, xl, yl);
        if (i === 0) ctx.moveTo(px, py);
        else ctx.lineTo(px, py);
      }
    }
    if (isPolygon && appendPath) ctx.closePath();
    any = true;
  }
  if (!any || s.inPath) return; // path bracket: geometry added, defer fill/stroke
  if (isPolygon && s.curBrush && s.curBrush.fill != null) {
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

/** Fill+stroke an axis-aligned rectangle (EMR_RECTANGLE) given logical corners. */
function fillStrokeRect(s: PlayState, l: number, t: number, r: number, b: number): void {
  const { ctx } = s;
  const c0 = toPx(s, l, t);
  const c1 = toPx(s, r, t);
  const c2 = toPx(s, r, b);
  const c3 = toPx(s, l, b);
  const appendPath = reservePathCommands(s, 5);
  if (!s.inPath) ctx.beginPath();
  if (appendPath) {
    ctx.moveTo(c0[0], c0[1]);
    ctx.lineTo(c1[0], c1[1]);
    ctx.lineTo(c2[0], c2[1]);
    ctx.lineTo(c3[0], c3[1]);
    ctx.closePath();
  }
  if (s.inPath) return; // path bracket: defer fill/stroke
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

/** Consume the current path with the selected brush and/or pen ([MS-EMF] 2.3.10). */
function paintSelectedPath(s: PlayState, fill: boolean, stroke: boolean): void {
  if (!s.pathDiscarded && fill && s.curBrush?.fill != null) {
    s.ctx.fillStyle = s.curBrush.fill;
    s.ctx.fill(s.fillRule);
    s.drew = true;
  }
  if (!s.pathDiscarded && stroke && s.curPen?.stroke != null) {
    s.ctx.strokeStyle = s.curPen.stroke;
    s.ctx.lineWidth = deviceLineWidth(s, s.curPen.width);
    s.ctx.stroke();
    s.drew = true;
  }
  // EMR_FILLPATH / STROKEPATH / STROKEANDFILLPATH close the path bracket.
  // Clear Canvas's persistent current path so later records cannot repaint it.
  s.ctx.beginPath();
  s.pathCommandCount = 0;
  s.pathDiscarded = false;
}

// ── object creators ──────────────────────────────────────────────────────────

/** EMR_CREATEPEN(38): u32 ihObject + LOGPEN{u32 style, POINTL width, COLORREF}. */
function readCreatePen(c: EmfCursor): [number, Pen] {
  const ih = c.u32();
  const style = c.u32();
  const widthX = c.i32();
  c.i32(); // POINTL.y (unused)
  const color = c.u32();
  const stroke = (style & 0xff) === 5 ? null : colorRefToCss(color); // PS_NULL=5
  return [ih, { kind: 'pen', stroke, width: Math.abs(widthX) }];
}

/** EMR_EXTCREATEPEN(95): u32 ihObject, 4×u32 offsets, then ELP {u32 style,
 *  u32 width, u32 brushStyle, COLORREF color, ...}. */
function readExtCreatePen(c: EmfCursor): [number, Pen] {
  const ih = c.u32();
  c.skip(16); // offBmi, cbBmi, offBits, cbBits
  const style = c.u32();
  const width = c.u32();
  c.u32(); // brushStyle
  const color = c.u32();
  const stroke = (style & 0xff) === 5 ? null : colorRefToCss(color); // PS_NULL=5
  return [ih, { kind: 'pen', stroke, width: Math.abs(width) }];
}

/** EMR_CREATEBRUSHINDIRECT(39): u32 ihObject + LOGBRUSH{u32 style, COLORREF,
 *  u32 hatch}. */
function readCreateBrush(c: EmfCursor): [number, Brush] {
  const ih = c.u32();
  const style = c.u32();
  const color = c.u32();
  c.u32(); // hatch (HATCHED → solid)
  const fill = style === 1 ? null : colorRefToCss(color); // BS_NULL=1
  return [ih, { kind: 'brush', fill }];
}

/**
 * EMR_CREATEMONOBRUSH(93) / EMR_CREATEDIBPATTERNBRUSHPT(94): u32 ihObject,
 * u32 iUsage, u32 offBmi, u32 cbBmi, u32 offBits, u32 cbBits, then a DIB.
 *
 * APPROXIMATION: DIB pattern brush ([MS-EMF] 2.3.7) rendered as its average
 * solid color pending true pattern-brush support. If DIB decode fails, fall
 * back to mid-gray so bars still show.
 */
function readDibPatternBrush(c: EmfCursor, dv: DataView, recStart: number): [number, Brush] {
  const ih = c.u32();
  c.u32(); // iUsage
  const offBmi = c.u32();
  const cbBmi = c.u32();
  const offBits = c.u32();
  const cbBits = c.u32();
  let fill = '#808080';
  try {
    const dib = decodeDib(dv, recStart + offBmi, cbBmi, recStart + offBits, cbBits);
    if (dib) fill = dibAverageColor(dib);
  } catch {
    /* keep mid-gray fallback */
  }
  return [ih, { kind: 'brush', fill }];
}

/** EMR_EXTCREATEFONTINDIRECTW(82): u32 ihObject + LOGFONT (lfHeight,…,lfFaceName
 *  UTF-16 at LOGFONT offset 28). */
function readCreateFont(c: EmfCursor, dv: DataView, recStart: number): [number, Font] {
  const ih = c.u32();
  const lfBase = recStart + 12; // ihObject (4) after the 8-byte record header
  const lfHeight = dv.getInt32(lfBase, true);
  // lfWidth(4), lfEscapement(8), lfOrientation(12) — escapement drives rotated
  // axis labels (e.g. a vertical "Dx [mm]" at 900 = 90° CCW).
  const lfEscapement = dv.getInt32(lfBase + 8, true);
  const lfWeight = dv.getInt32(lfBase + 16, true);
  const lfItalic = dv.getUint8(lfBase + 20);
  // lfFaceName: UTF-16, up to 32 code units, at LOGFONT offset 28.
  let face = '';
  for (let i = 0; i < 32; i++) {
    const o = lfBase + 28 + i * 2;
    if (o + 2 > dv.byteLength) break;
    const cu = dv.getUint16(o, true);
    if (cu === 0) break;
    face += String.fromCharCode(cu);
  }
  return [
    ih,
    {
      kind: 'font',
      height: Math.abs(lfHeight),
      weight: lfWeight,
      italic: lfItalic !== 0,
      face,
      escapement: lfEscapement,
    },
  ];
}

// ── text — EMR_EXTTEXTOUTW(84) ([MS-EMF] 2.3.5.2) ────────────────────────────

function drawText(s: PlayState, c: EmfCursor, dv: DataView, recStart: number): void {
  c.skip(16); // RECTL rclBounds (record offset 8..24)
  c.u32(); // iGraphicsMode
  c.f32(); // exScale
  c.f32(); // eyScale
  // EMRTEXT at record offset 36:
  const refX = c.i32();
  const refY = c.i32();
  const nChars = c.u32();
  const offString = c.u32(); // BYTE offset from RECORD START to UTF-16 string
  c.u32(); // fOptions
  if (nChars <= 0 || nChars > 0x10000) return;
  // Read nChars UTF-16LE code units at recStart + offString.
  let str = '';
  for (let i = 0; i < nChars; i++) {
    const o = recStart + offString + i * 2;
    if (o + 2 > dv.byteLength) break;
    str += String.fromCharCode(dv.getUint16(o, true));
  }
  if (str.length === 0) return;

  const font = s.curFont;
  // Font height is in logical units: world→page (worldScaleY), page→device
  // (|pageScaleY|, 1 under MM_TEXT), then device→target (deviceScaleY).
  const px =
    Math.abs(font?.height ?? 0) *
    worldScaleY(s) *
    Math.abs(pageScaleY(s)) *
    deviceScaleY(s);
  if (!Number.isFinite(px) || px < 1) return;

  const { ctx } = s;
  const [dx, dy] = toPx(s, refX, refY);
  ctx.fillStyle = s.textColor;
  const weight = font && font.weight >= 700 ? 'bold ' : '';
  const italic = font?.italic ? 'italic ' : '';
  ctx.font = `${italic}${weight}${px}px ${font?.face || 'sans-serif'}`;

  // SETTEXTALIGN: low 2 bits horizontal. TA_LEFT(0), TA_RIGHT(2), TA_CENTER(6).
  const horiz = s.textAlign & 0x6;
  ctx.textAlign = horiz === 0x2 ? 'right' : horiz === 0x6 ? 'center' : 'left';
  // TA_BASELINE(0x18) → alphabetic; else top.
  ctx.textBaseline = (s.textAlign & 0x18) === 0x18 ? 'alphabetic' : 'top';
  // bkMode TRANSPARENT(1): never paint a background box (always the case here).
  // lfEscapement rotates the text about the reference point (tenths of a degree,
  // counterclockwise from the device x-axis). Canvas angles are clockwise on a
  // y-down surface, so negate. Used for vertical axis labels (e.g. 900 = 90°).
  const escTenths = font?.escapement ?? 0;
  try {
    if (escTenths !== 0) {
      ctx.save();
      try {
        ctx.translate(dx, dy);
        ctx.rotate((-escTenths / 10) * (Math.PI / 180));
        ctx.fillText(str, 0, 0);
      } finally {
        ctx.restore(); // always unwind the save, even if fillText throws
      }
    } else {
      ctx.fillText(str, dx, dy);
    }
    s.drew = true;
  } catch {
    // Some ctx mocks lack fillText; a missing fillText must not abort the render.
  }
}

// ── bitmaps — EMR_BITBLT(76) / EMR_STRETCHDIBITS(81) ([MS-EMF] 2.3.1) ─────────

/** Decode a DIB and blit it into the dest rect (logical corners mapped via
 *  toPx). Skips gracefully on unsupported DIBs or missing OffscreenCanvas. */
function blitDib(
  s: PlayState,
  dv: DataView,
  recStart: number,
  offBmi: number,
  cbBmi: number,
  offBits: number,
  cbBits: number,
  destL: number,
  destT: number,
  destR: number,
  destB: number,
): void {
  if (cbBmi === 0 || cbBits === 0) return; // pattern-only blt → skip
  const dib = decodeDib(dv, recStart + offBmi, cbBmi, recStart + offBits, cbBits);
  if (!dib) return;
  const [x0, y0] = toPx(s, destL, destT);
  const [x1, y1] = toPx(s, destR, destB);
  if (blitDibToCtx(s.ctx, dib, x0, y0, x1, y1)) s.drew = true;
}

/** EMR_BITBLT(76) ([MS-EMF] 2.3.1.2). */
function doBitBlt(s: PlayState, c: EmfCursor, dv: DataView, recStart: number): void {
  c.skip(16); // RECTL rclBounds
  const xDest = c.i32();
  const yDest = c.i32();
  const cxDest = c.i32();
  const cyDest = c.i32();
  c.u32(); // bitBltRasterOp
  c.i32(); // xSrc
  c.i32(); // ySrc
  c.skip(24); // XFORM xformSrc (6×f32)
  c.u32(); // bkColorSrc
  c.u32(); // usageSrc
  const offBmi = c.u32();
  const cbBmi = c.u32();
  const offBits = c.u32();
  const cbBits = c.u32();
  blitDib(
    s, dv, recStart, offBmi, cbBmi, offBits, cbBits,
    xDest, yDest, xDest + cxDest, yDest + cyDest,
  );
}

/** EMR_STRETCHDIBITS(81) ([MS-EMF] 2.3.1.7). */
function doStretchDibits(s: PlayState, c: EmfCursor, dv: DataView, recStart: number): void {
  c.skip(16); // RECTL rclBounds
  const xDest = c.i32();
  const yDest = c.i32();
  c.i32(); // xSrc
  c.i32(); // ySrc
  c.i32(); // cxSrc
  c.i32(); // cySrc
  const offBmi = c.u32();
  const cbBmi = c.u32();
  const offBits = c.u32();
  const cbBits = c.u32();
  c.u32(); // usageSrc
  c.u32(); // bitBltRasterOp
  const cxDest = c.i32();
  const cyDest = c.i32();
  blitDib(
    s, dv, recStart, offBmi, cbBmi, offBits, cbBits,
    xDest, yDest, xDest + cxDest, yDest + cyDest,
  );
}

// ── core record-replay loop (pure; testable with a mock ctx) ────────────────

/**
 * Replay an EMF byte buffer onto a 2D context, mapping logical coordinates
 * through the world transform and then into a `W`×`H` target raster. Returns
 * `true` if anything was drawn, `false` for a non-EMF buffer or a metafile that
 * produced no geometry.
 *
 * Pure with respect to the injected `ctx`, so it is unit-testable against a
 * recording mock — no OffscreenCanvas required (the only exception is the
 * optional BITBLT/STRETCHDIBITS path, which needs a temp OffscreenCanvas and is
 * skipped gracefully when absent).
 */
export function playEmf(bytes: Uint8Array, ctx: AnyCtx, W: number, H: number): boolean;
export function playEmf(
  bytes: Uint8Array,
  ctx: AnyCtx,
  W: number,
  H: number,
  limits: EmfReplayLimits = {},
): boolean {
  if (!isEmf(bytes)) return false;
  if (W <= 0 || H <= 0) return false;

  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);

  const s: PlayState = {
    ctx,
    W,
    H,
    left: 0,
    top: 0,
    boundsW: W,
    boundsH: H,
    wt: identity(),
    // MM_TEXT default: page→device is the identity (winOrg=vpOrg=0,
    // winExt=vpExt=1), so files that scale purely with the world transform are
    // byte-for-byte unaffected by the window→viewport stage.
    mapMode: MM.TEXT,
    winOrgX: 0,
    winOrgY: 0,
    winExtX: 1,
    winExtY: 1,
    vpOrgX: 0,
    vpOrgY: 0,
    vpExtX: 1,
    vpExtY: 1,
    devPxPerMmX: 0,
    devPxPerMmY: 0,
    objects: new Map(),
    curPen: null,
    curBrush: null,
    curFont: null,
    textColor: '#000000',
    bkMode: 1,
    textAlign: 0,
    fillRule: 'nonzero',
    curX: 0,
    curY: 0,
    stack: [],
    drew: false,
    inPath: false,
    pathCommandCount: 0,
    pathDiscarded: false,
    maxPathCommands: Number.isSafeInteger(limits.maxPathCommands)
      && (limits.maxPathCommands as number) > 0
      ? Math.min(limits.maxPathCommands as number, MAX_EMF_PATH_COMMANDS)
      : MAX_EMF_PATH_COMMANDS,
  };

  let pos = 0;
  while (pos + 8 <= bytes.length) {
    const iType = dv.getUint32(pos, true);
    const nSize = dv.getUint32(pos + 4, true);
    // Validate: nSize ≥ 8 (the iType+nSize header) and 4-aligned and in-bounds.
    if (nSize < 8 || (nSize & 3) !== 0) break;
    const recEnd = pos + nSize;
    if (recEnd > bytes.length) break; // truncated → partial render
    if (iType === EMR.EOF) break;

    // A cursor over the data region (starts at record offset 8).
    const c = new EmfCursor(dv, pos + 8, recEnd);

    // Never throw on a malformed record — just advance by nSize.
    try {
      switch (iType) {
        case EMR.HEADER: {
          // ENHMETAHEADER ([MS-EMF] 2.2.9). rclBounds (the INK bounding box, in
          // device units) @ record offset 8; rclFrame (the intended PICTURE
          // FRAME, in .01 mm) @ 24; szlDevice (reference device size, px) @ 72;
          // szlMillimeters (reference device size, mm) @ 80.
          const bLeft = dv.getInt32(pos + 8, true);
          const bTop = dv.getInt32(pos + 12, true);
          const bRight = dv.getInt32(pos + 16, true);
          const bBottom = dv.getInt32(pos + 20, true);
          // Default mapping: the ink bounds fill the target. Used when the frame
          // or reference device size is absent/degenerate (mirrors the
          // LibreOffice/POI fallback to the bounds rectangle).
          s.left = bLeft;
          s.top = bTop;
          s.boundsW = Math.max(1, bRight - bLeft);
          s.boundsH = Math.max(1, bBottom - bTop);
          // GDI `PlayEnhMetaFile` maps the FRAME — not the ink bounds — onto the
          // target rectangle, so whitespace around the ink is preserved and a
          // PowerPoint/Word `<a:srcRect>` crop (defined relative to the frame,
          // ECMA-376 §20.1.8.55) aligns with the picture. The records draw in
          // device units (same units as rclBounds), so convert the frame from
          // .01 mm to those device units via the reference device resolution
          // (px per .01 mm = szlDevice / (szlMillimeters · 100)) and map THAT
          // rectangle instead. (Confirmed against [MS-EMF] 2.2.9 + the GDI
          // PlayEnhMetaFile remarks + LibreOffice emfio / Apache POI HEMF, which
          // both size the picture to the frame.)
          if (recEnd >= pos + 88) {
            const fLeft = dv.getInt32(pos + 24, true);
            const fTop = dv.getInt32(pos + 28, true);
            const fRight = dv.getInt32(pos + 32, true);
            const fBottom = dv.getInt32(pos + 36, true);
            const devCx = dv.getInt32(pos + 72, true);
            const devCy = dv.getInt32(pos + 76, true);
            const mmCx = dv.getInt32(pos + 80, true);
            const mmCy = dv.getInt32(pos + 84, true);
            const fwMm = fRight - fLeft;
            const fhMm = fBottom - fTop;
            if (fwMm > 0 && fhMm > 0 && devCx > 0 && devCy > 0 && mmCx > 0 && mmCy > 0) {
              const sx = devCx / (mmCx * 100); // device px per .01 mm (X)
              const sy = devCy / (mmCy * 100); // device px per .01 mm (Y)
              s.left = fLeft * sx;
              s.top = fTop * sy;
              s.boundsW = Math.max(1, fwMm * sx);
              s.boundsH = Math.max(1, fhMm * sy);
              // Reference-device resolution in px per mm, for the metric map
              // modes ([MS-EMF] 2.1.21). szlDevice is px, szlMillimeters is mm.
              s.devPxPerMmX = devCx / mmCx;
              s.devPxPerMmY = devCy / mmCy;
            }
          }
          break;
        }
        case EMR.SETWORLDTRANSFORM: {
          s.wt = c.xform();
          break;
        }
        case EMR.MODIFYWORLDTRANSFORM: {
          const x = c.xform();
          const iMode = c.u32();
          // [MS-EMF] 2.3.12: 1=IDENTITY, 2=LEFTMULTIPLY (xform × WT),
          // 3=RIGHTMULTIPLY (WT × xform), 4=SET.
          if (iMode === 1) s.wt = identity();
          else if (iMode === 2) s.wt = mulXform(x, s.wt);
          else if (iMode === 3) s.wt = mulXform(s.wt, x);
          else if (iMode === 4) s.wt = x;
          break;
        }
        case EMR.SETMAPMODE: {
          // data: u32 MapMode ([MS-EMF] 2.3.11).
          applyMapMode(s, c.u32());
          break;
        }
        case EMR.SETWINDOWORGEX: {
          // data: POINTL { i32 x, i32 y } ([MS-EMF] 2.3.11).
          s.winOrgX = c.i32();
          s.winOrgY = c.i32();
          break;
        }
        case EMR.SETWINDOWEXTEX: {
          // data: SIZEL { i32 cx, i32 cy } — the window extent, page space.
          const cx = c.i32();
          const cy = c.i32();
          if (cx !== 0) s.winExtX = cx;
          if (cy !== 0) s.winExtY = cy;
          break;
        }
        case EMR.SETVIEWPORTORGEX: {
          s.vpOrgX = c.i32();
          s.vpOrgY = c.i32();
          break;
        }
        case EMR.SETVIEWPORTEXTEX: {
          // data: SIZEL { i32 cx, i32 cy } — the viewport extent, device space.
          const cx = c.i32();
          const cy = c.i32();
          if (cx !== 0) s.vpExtX = cx;
          if (cy !== 0) s.vpExtY = cy;
          break;
        }
        case EMR.SCALEWINDOWEXTEX: {
          // data: i32 xNum, i32 xDenom, i32 yNum, i32 yDenom ([MS-EMF] 2.3.11):
          // window extent ×= num/denom. Divisor 0 leaves that axis unchanged.
          const xNum = c.i32();
          const xDenom = c.i32();
          const yNum = c.i32();
          const yDenom = c.i32();
          if (xDenom !== 0) s.winExtX = (s.winExtX * xNum) / xDenom;
          if (yDenom !== 0) s.winExtY = (s.winExtY * yNum) / yDenom;
          break;
        }
        case EMR.SCALEVIEWPORTEXTEX: {
          const xNum = c.i32();
          const xDenom = c.i32();
          const yNum = c.i32();
          const yDenom = c.i32();
          if (xDenom !== 0) s.vpExtX = (s.vpExtX * xNum) / xDenom;
          if (yDenom !== 0) s.vpExtY = (s.vpExtY * yNum) / yDenom;
          break;
        }
        case EMR.SAVEDC: {
          // Mirror the GDI state push on the canvas too, so a clip set via
          // SELECTCLIPPATH (below) is scoped to the matching RESTOREDC.
          s.ctx.save();
          s.stack.push({
            wt: { ...s.wt },
            mapMode: s.mapMode,
            winOrgX: s.winOrgX,
            winOrgY: s.winOrgY,
            winExtX: s.winExtX,
            winExtY: s.winExtY,
            vpOrgX: s.vpOrgX,
            vpOrgY: s.vpOrgY,
            vpExtX: s.vpExtX,
            vpExtY: s.vpExtY,
            curPen: s.curPen,
            curBrush: s.curBrush,
            curFont: s.curFont,
            textColor: s.textColor,
            bkMode: s.bkMode,
            textAlign: s.textAlign,
            fillRule: s.fillRule,
            curX: s.curX,
            curY: s.curY,
          });
          break;
        }
        case EMR.RESTOREDC: {
          // data: i32 iRelative (e.g. -1 = pop one). Pop |iRelative| times,
          // clamped to the stack size.
          const iRelative = c.i32();
          const times = Math.min(Math.abs(iRelative) || 1, s.stack.length);
          let saved: SavedDc | undefined;
          for (let i = 0; i < times; i++) {
            saved = s.stack.pop();
            s.ctx.restore(); // unwind the matching canvas save (clip/state)
          }
          if (saved) {
            s.wt = saved.wt;
            s.mapMode = saved.mapMode;
            s.winOrgX = saved.winOrgX;
            s.winOrgY = saved.winOrgY;
            s.winExtX = saved.winExtX;
            s.winExtY = saved.winExtY;
            s.vpOrgX = saved.vpOrgX;
            s.vpOrgY = saved.vpOrgY;
            s.vpExtX = saved.vpExtX;
            s.vpExtY = saved.vpExtY;
            s.curPen = saved.curPen;
            s.curBrush = saved.curBrush;
            s.curFont = saved.curFont;
            s.textColor = saved.textColor;
            s.bkMode = saved.bkMode;
            s.textAlign = saved.textAlign;
            s.fillRule = saved.fillRule;
            s.curX = saved.curX;
            s.curY = saved.curY;
          }
          break;
        }
        case EMR.BEGINPATH: {
          // Start a path bracket ([MS-EMF] 2.3.10): subsequent geometry records
          // build the path instead of drawing it, until ENDPATH.
          s.ctx.beginPath();
          s.inPath = true;
          s.pathCommandCount = 0;
          s.pathDiscarded = false;
          break;
        }
        case EMR.CLOSEFIGURE: {
          if (s.inPath && reservePathCommands(s, 1)) s.ctx.closePath();
          break;
        }
        case EMR.ENDPATH: {
          s.inPath = false;
          break;
        }
        case EMR.FILLPATH: {
          paintSelectedPath(s, true, false);
          break;
        }
        case EMR.STROKEANDFILLPATH: {
          paintSelectedPath(s, true, true);
          break;
        }
        case EMR.STROKEPATH: {
          paintSelectedPath(s, false, true);
          break;
        }
        case EMR.SELECTCLIPPATH: {
          // Use the path just defined as the clip region (intersecting the
          // current clip — the common RGN_AND case, and what a following blit
          // relies on, e.g. sample-13 Fig.3 clips a bar-chart DIB to the bar
          // shapes so its background is masked out). Scoped by the enclosing
          // SAVEDC/RESTOREDC.
          if (!s.pathDiscarded) {
            try {
              s.ctx.clip(s.fillRule);
            } catch {
              /* a ctx without clip() (some mocks): leave unclipped */
            }
          }
          s.pathCommandCount = 0;
          s.pathDiscarded = false;
          break;
        }
        case EMR.SELECTOBJECT: {
          const ih = c.u32();
          if ((ih & 0x80000000) !== 0) {
            selectStock(s, ih >>> 0);
          } else {
            const obj = s.objects.get(ih);
            if (obj?.kind === 'pen') s.curPen = obj;
            else if (obj?.kind === 'brush') s.curBrush = obj;
            else if (obj?.kind === 'font') s.curFont = obj;
          }
          break;
        }
        case EMR.DELETEOBJECT: {
          const ih = c.u32();
          const obj = s.objects.get(ih);
          if (obj) {
            if (obj === s.curPen) s.curPen = null;
            if (obj === s.curBrush) s.curBrush = null;
            if (obj === s.curFont) s.curFont = null;
            s.objects.delete(ih);
          }
          break;
        }
        case EMR.CREATEPEN: {
          const [ih, pen] = readCreatePen(c);
          s.objects.set(ih, pen);
          break;
        }
        case EMR.EXTCREATEPEN: {
          const [ih, pen] = readExtCreatePen(c);
          s.objects.set(ih, pen);
          break;
        }
        case EMR.CREATEBRUSHINDIRECT: {
          const [ih, brush] = readCreateBrush(c);
          s.objects.set(ih, brush);
          break;
        }
        case EMR.CREATEMONOBRUSH:
        case EMR.CREATEDIBPATTERNBRUSHPT: {
          const [ih, brush] = readDibPatternBrush(c, dv, pos);
          s.objects.set(ih, brush);
          break;
        }
        case EMR.EXTCREATEFONTINDIRECTW: {
          const [ih, font] = readCreateFont(c, dv, pos);
          s.objects.set(ih, font);
          break;
        }
        case EMR.POLYLINE16:
          strokePolyline(s, c, readPoint16);
          break;
        case EMR.POLYLINE:
          strokePolyline(s, c, readPoint32);
          break;
        case EMR.POLYLINETO16:
          strokePolylineTo(s, c, readPoint16);
          break;
        case EMR.POLYLINETO:
          strokePolylineTo(s, c, readPoint32);
          break;
        case EMR.POLYGON16:
          fillStrokePolygon(s, c, readPoint16);
          break;
        case EMR.POLYGON:
          fillStrokePolygon(s, c, readPoint32);
          break;
        case EMR.POLYBEZIER16:
          strokePolyBezier(s, c, readPoint16, false);
          break;
        case EMR.POLYBEZIER:
          strokePolyBezier(s, c, readPoint32, false);
          break;
        case EMR.POLYBEZIERTO16:
          strokePolyBezier(s, c, readPoint16, true);
          break;
        case EMR.POLYBEZIERTO:
          strokePolyBezier(s, c, readPoint32, true);
          break;
        case EMR.POLYPOLYGON16:
          fillStrokePolyPoly(s, c, readPoint16, true);
          break;
        case EMR.POLYPOLYGON:
          fillStrokePolyPoly(s, c, readPoint32, true);
          break;
        case EMR.POLYPOLYLINE16:
          fillStrokePolyPoly(s, c, readPoint16, false);
          break;
        case EMR.POLYPOLYLINE:
          fillStrokePolyPoly(s, c, readPoint32, false);
          break;
        case EMR.MOVETOEX: {
          s.curX = c.i32();
          s.curY = c.i32();
          if (s.inPath && reservePathCommands(s, 1)) {
            const [px, py] = toPx(s, s.curX, s.curY);
            s.ctx.moveTo(px, py);
          }
          break;
        }
        case EMR.LINETO: {
          const xl = c.i32();
          const yl = c.i32();
          if (s.inPath && reservePathCommands(s, 1)) {
            const [px1, py1] = toPx(s, xl, yl);
            s.ctx.lineTo(px1, py1);
          } else if (s.curPen && s.curPen.stroke != null) {
            const [px0, py0] = toPx(s, s.curX, s.curY);
            const [px1, py1] = toPx(s, xl, yl);
            s.ctx.beginPath();
            s.ctx.moveTo(px0, py0);
            s.ctx.lineTo(px1, py1);
            s.ctx.strokeStyle = s.curPen.stroke;
            s.ctx.lineWidth = deviceLineWidth(s, s.curPen.width);
            s.ctx.stroke();
            s.drew = true;
          }
          s.curX = xl;
          s.curY = yl;
          break;
        }
        case EMR.RECTANGLE: {
          const left = c.i32();
          const top = c.i32();
          const right = c.i32();
          const bottom = c.i32();
          fillStrokeRect(s, left, top, right, bottom);
          break;
        }
        case EMR.ELLIPSE: {
          const left = c.i32();
          const top = c.i32();
          const right = c.i32();
          const bottom = c.i32();
          const [cxl, cyl] = [(left + right) / 2, (top + bottom) / 2];
          const [cx, cy] = toPx(s, cxl, cyl);
          const [ex] = toPx(s, right, cyl);
          const [, ey] = toPx(s, cxl, bottom);
          const rx = Math.abs(ex - cx);
          const ry = Math.abs(ey - cy);
          ctx.beginPath();
          ctx.ellipse(cx, cy, rx, ry, 0, 0, Math.PI * 2);
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
          break;
        }
        case EMR.SETPOLYFILLMODE: {
          const mode = c.u32(); // 1=ALTERNATE→evenodd, 2=WINDING→nonzero
          s.fillRule = mode === 1 ? 'evenodd' : 'nonzero';
          break;
        }
        case EMR.SETTEXTCOLOR: {
          s.textColor = colorRefToCss(c.u32());
          break;
        }
        case EMR.SETTEXTALIGN: {
          s.textAlign = c.u32();
          break;
        }
        case EMR.SETBKMODE: {
          s.bkMode = c.u32(); // 1 = TRANSPARENT
          break;
        }
        case EMR.EXTTEXTOUTW:
          drawText(s, c, dv, pos);
          break;
        case EMR.BITBLT:
          doBitBlt(s, c, dv, pos);
          break;
        case EMR.STRETCHDIBITS:
          doStretchDibits(s, c, dv, pos);
          break;
        default:
          // GDICOMMENT (may hold EMF+, out of scope), SETICMMODE,
          // SETMITERLIMIT, SETROP2, SETSTRETCHBLTMODE, INTERSECTCLIPRECT, and any
          // unrecognized iType: skip by nSize. (Path/clip records ARE handled
          // above.)
          break;
      }
    } catch {
      // A malformed record must never abort the whole render — just advance.
    }

    pos = recEnd;
  }

  return s.drew;
}

// ── async OffscreenCanvas wrapper ───────────────────────────────────────────

/**
 * Rasterize an EMF metafile to an `ImageBitmap` of `targetW`×`targetH`, replaying
 * onto an `OffscreenCanvas` 2D context. Returns `null` if the bytes are not a
 * parseable EMF or nothing drew (so the caller can fall back to the existing
 * "missing image" behavior without crashing). Mirrors
 * {@link ./wmf.ts}#renderWmfToBitmap.
 */
export async function renderEmfToBitmap(
  bytes: Uint8Array,
  targetW: number,
  targetH: number,
): Promise<ImageBitmap | null> {
  if (!isEmf(bytes)) return null;
  if (targetW <= 0 || targetH <= 0) return null;
  // Rasterize on a shared aux canvas (OffscreenCanvas, else a detached <canvas>).
  // Absent both (e.g. a headless test / SSR runtime without either) ⇒ degrade
  // gracefully to null, exactly as the caller already handles an unsupported
  // metafile — never throw.
  const canvas = createAuxCanvas(targetW, targetH);
  if (!canvas) return null;
  const ctx = canvas.getContext('2d') as AnyCtx | null;
  if (!ctx) return null;
  ctx.lineJoin = 'round';
  ctx.lineCap = 'round';
  const drew = playEmf(bytes, ctx, targetW, targetH);
  if (!drew) return null;
  return createImageBitmap(canvas);
}
