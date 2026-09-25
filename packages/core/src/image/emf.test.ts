import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { isEmf } from './wmf.js';
import { playEmf, renderEmfToBitmap } from './emf.js';
import { scanEmfPlus } from './emf-plus.js';
import {
  decodeRasterOrMetafile,
  getIncompleteMetafileReport,
  isOoxmlIncompleteMetafileError,
} from './raster-or-metafile.js';
import { dropBitmapCacheByPath, getCachedBitmapByPath } from './bitmap-image-by-path.js';
import { getCachedDuotoneBitmapByPath } from './duotone-bitmap-by-path.js';

// ── EMF (Enhanced Metafile) player unit tests ───────────────────────────────
// The renderer falls back to this player for true `.emf` blips the browser can't
// decode (createImageBitmap throws on metafiles, and EMF is a different, larger
// 32-bit format than WMF). sample-13.docx embeds two EMF charts (`image3.emf`
// Fig.2, `image4.emf` Fig.3) whose bars/axes are POLYGON16/POLYLINE16 records
// and whose labels are EXTTEXTOUTW text-out records, all scaled by a long run of
// MODIFYWORLDTRANSFORM affines — so the player needs the world transform, an
// object table (pens+brushes+fonts), polygon/polyline drawing, and text-out.
//
// `playEmf(bytes, ctx, W, H)` is the pure record-replay core: it issues
// moveTo/lineTo/stroke/fill/fillText calls onto an injected ctx, so a recording
// mock pins coordinate mapping + state without needing OffscreenCanvas (absent
// in the node test env). The two sample EMFs themselves live only inside the
// gitignored `sample-13.docx` (docx/pptx/xlsx files are not committed), so these
// tests craft synthetic EMF byte buffers rather than read the private samples.

// ── byte builders ───────────────────────────────────────────────────────────

/** Little-endian byte writer for crafting EMF records (32-bit / IEEE-754). */
class Writer {
  private bytes: number[] = [];
  u16(v: number) {
    this.bytes.push(v & 0xff, (v >>> 8) & 0xff);
    return this;
  }
  i16(v: number) {
    return this.u16(v & 0xffff);
  }
  u32(v: number) {
    this.bytes.push(v & 0xff, (v >>> 8) & 0xff, (v >>> 16) & 0xff, (v >>> 24) & 0xff);
    return this;
  }
  i32(v: number) {
    return this.u32(v >>> 0);
  }
  f32(v: number) {
    const buf = new ArrayBuffer(4);
    new DataView(buf).setFloat32(0, v, true);
    const u8 = new Uint8Array(buf);
    this.bytes.push(u8[0], u8[1], u8[2], u8[3]);
    return this;
  }
  utf16(s: string) {
    for (let i = 0; i < s.length; i++) this.u16(s.charCodeAt(i));
    return this;
  }
  raw(...vals: number[]) {
    for (const v of vals) this.bytes.push(v & 0xff);
    return this;
  }
  get length(): number {
    return this.bytes.length;
  }
  build(): Uint8Array {
    return new Uint8Array(this.bytes);
  }
}

// EMF record type ids ([MS-EMF] 2.1.1 RecordType).
const EMR = {
  HEADER: 1,
  SETWINDOWEXTEX: 9,
  SETWINDOWORGEX: 10,
  SETVIEWPORTEXTEX: 11,
  SETVIEWPORTORGEX: 12,
  EOF: 14,
  SETMAPMODE: 17,
  SETPOLYFILLMODE: 19,
  SETTEXTALIGN: 22,
  SETTEXTCOLOR: 24,
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
  MOVETOEX: 27,
  LINETO: 54,
  BEGINPATH: 59,
  ENDPATH: 60,
  CLOSEFIGURE: 61,
  FILLPATH: 62,
  SELECTCLIPPATH: 67,
  EXTCREATEFONTINDIRECTW: 82,
  EXTTEXTOUTW: 84,
  POLYGON16: 86,
  POLYLINE16: 87,
  POLYBEZIERTO16: 88,
  POLYPOLYGON16: 91,
  CREATEDIBPATTERNBRUSHPT: 94,
} as const;

/** An EMF record: u32 iType, u32 nSize (incl. the 8-byte header), then data.
 *  nSize is padded to a 4-byte boundary. */
function record(iType: number, data: (w: Writer) => void): Uint8Array {
  const pw = new Writer();
  data(pw);
  let body = pw.build();
  // 4-byte align the record body.
  if (body.length % 4 !== 0) {
    const pad = new Uint8Array(body.length + (4 - (body.length % 4)));
    pad.set(body, 0);
    body = pad;
  }
  const nSize = 8 + body.length;
  const head = new Writer().u32(iType).u32(nSize).build();
  const out = new Uint8Array(head.length + body.length);
  out.set(head, 0);
  out.set(body, head.length);
  return out;
}

function concat(...parts: Uint8Array[]): Uint8Array {
  const total = parts.reduce((n, p) => n + p.length, 0);
  const out = new Uint8Array(total);
  let off = 0;
  for (const p of parts) {
    out.set(p, off);
    off += p.length;
  }
  return out;
}

/** EMR_HEADER with the given inclusive device bounds (left,top,right,bottom).
 *  The signature " EMF" (0x464D4520) must land at byte offset 40 of the whole
 *  file → record offset 40 (the header record starts at file offset 0).
 *
 *  `opts.frame` (.01 mm) + `opts.dev`/`opts.mm` (reference device px/mm) drive
 *  the picture-frame mapping (`playEmf` maps the frame, not the ink bounds, onto
 *  the target). The default leaves the frame degenerate (0,0,0,0) so the player
 *  falls back to mapping the ink bounds — keeping the coordinate-pipeline tests
 *  below focused on the world transform + scaling. The frame path has its own
 *  test (`maps the picture frame … not the ink bounds`). */
function emfHeader(
  left = 0,
  top = 0,
  right = 100,
  bottom = 100,
  opts: {
    frame?: { l: number; t: number; r: number; b: number };
    dev?: { cx: number; cy: number };
    mm?: { cx: number; cy: number };
  } = {},
): Uint8Array {
  const f = opts.frame ?? { l: 0, t: 0, r: 0, b: 0 };
  const dev = opts.dev ?? { cx: 1920, cy: 1080 };
  const mm = opts.mm ?? { cx: 508, cy: 286 };
  return record(EMR.HEADER, (w) => {
    // data starts at record offset 8:
    w.i32(left).i32(top).i32(right).i32(bottom); // rclBounds   (off 8..24)
    w.i32(f.l).i32(f.t).i32(f.r).i32(f.b); //        rclFrame    (off 24..40)
    w.u32(0x464d4520); //                            dSignature  (off 40) " EMF"
    w.u32(0x00010000); //                            nVersion    (off 44)
    w.u32(0); //                                     nBytes      (off 48)
    w.u32(0); //                                     nRecords    (off 52)
    w.u16(0).u16(0); //                              nHandles/sReserved
    w.u32(0).u32(0); //                              nDescription/offDescription
    w.u32(0); //                                     nPalEntries
    w.i32(dev.cx).i32(dev.cy); //                    szlDevice (px)     (off 72)
    w.i32(mm.cx).i32(mm.cy); //                      szlMillimeters     (off 80)
  });
}

// ── recording mock ctx (records the draw calls + style mutations) ───────────

interface Call {
  op: string;
  args: (number | string)[];
}
interface MockCtx {
  ctx: CanvasRenderingContext2D;
  calls: Call[];
  styles: {
    fill: string[];
    stroke: string[];
    text: string[];
    fillRules: (string | undefined)[];
  };
}

function makeRecordingCtx(): MockCtx {
  const calls: Call[] = [];
  const styles = {
    fill: [] as string[],
    stroke: [] as string[],
    text: [] as string[],
    fillRules: [] as (string | undefined)[],
  };
  let _fill = '#000';
  let _stroke = '#000';
  let _lw = 1;
  const ctx = {
    get fillStyle() {
      return _fill;
    },
    set fillStyle(v: string) {
      _fill = v;
    },
    get strokeStyle() {
      return _stroke;
    },
    set strokeStyle(v: string) {
      _stroke = v;
    },
    get lineWidth() {
      return _lw;
    },
    set lineWidth(v: number) {
      _lw = v;
    },
    font: '10px sans-serif',
    textAlign: 'left' as CanvasTextAlign,
    textBaseline: 'top' as CanvasTextBaseline,
    lineJoin: 'miter' as CanvasLineJoin,
    lineCap: 'butt' as CanvasLineCap,
    save() {
      calls.push({ op: 'save', args: [] });
    },
    restore() {
      calls.push({ op: 'restore', args: [] });
    },
    beginPath() {
      calls.push({ op: 'beginPath', args: [] });
    },
    closePath() {
      calls.push({ op: 'closePath', args: [] });
    },
    moveTo(x: number, y: number) {
      calls.push({ op: 'moveTo', args: [x, y] });
    },
    lineTo(x: number, y: number) {
      calls.push({ op: 'lineTo', args: [x, y] });
    },
    bezierCurveTo(...a: number[]) {
      calls.push({ op: 'bezierCurveTo', args: a });
    },
    ellipse(...a: number[]) {
      calls.push({ op: 'ellipse', args: a });
    },
    rect(x: number, y: number, w: number, h: number) {
      calls.push({ op: 'rect', args: [x, y, w, h] });
    },
    stroke() {
      calls.push({ op: 'stroke', args: [] });
      styles.stroke.push(_stroke);
    },
    fill(rule?: string) {
      calls.push({ op: 'fill', args: [] });
      styles.fill.push(_fill);
      styles.fillRules.push(rule);
    },
    fillText(t: string, x: number, y: number) {
      calls.push({ op: 'fillText', args: [t, x, y] });
      styles.text.push(_fill);
    },
    translate(x: number, y: number) {
      calls.push({ op: 'translate', args: [x, y] });
    },
    rotate(a: number) {
      calls.push({ op: 'rotate', args: [a] });
    },
    clip(...a: unknown[]) {
      // clip(rule) or clip(path, rule): a Path2D region is recorded by name.
      calls.push({ op: 'clip', args: a.filter((v) => v !== undefined).map((v) => (typeof v === 'string' ? v : 'Path2D')) });
    },
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, calls, styles };
}

// ── isEmf detection ─────────────────────────────────────────────────────────

describe('isEmf detection (shared with the WMF sniffer)', () => {
  it('detects a synthetic EMF header (EMR_HEADER + " EMF" signature@40)', () => {
    expect(isEmf(emfHeader())).toBe(true);
  });

  it('rejects a non-EMF buffer (PNG magic)', () => {
    const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 0]);
    expect(isEmf(png)).toBe(false);
  });
});

// ── playEmf: world transform + polyline + pen ────────────────────────────────

describe('playEmf — world transform, pen, polyline16', () => {
  it('maps a polyline through the world transform then device→target px', () => {
    // bounds 0..100 × 0..100, target 200×200 → device→px scale ×2.
    // World transform: scale ×0.5 (m11=m22=0.5). So logical (40,60) → device
    // (20,30) → px (40,60); logical (80,100) → device (40,50) → px (80,100).
    const file = concat(
      emfHeader(0, 0, 100, 100),
      // MODIFYWORLDTRANSFORM, iMode=4 (MWT_SET): WT = the supplied xform.
      record(EMR.MODIFYWORLDTRANSFORM, (w) =>
        w.f32(0.5).f32(0).f32(0).f32(0.5).f32(0).f32(0).u32(4),
      ),
      // CREATEPEN ih=1, style=0 (solid), width (1,0), color blue 0x00FF0000.
      record(EMR.CREATEPEN, (w) => w.u32(1).u32(0).i32(1).i32(0).u32(0x00ff0000)),
      record(EMR.SELECTOBJECT, (w) => w.u32(1)),
      // POLYLINE16: RECTL bounds (skipped), count=2, then 2× POINTS(i16,i16).
      record(EMR.POLYLINE16, (w) =>
        w.i32(0).i32(0).i32(100).i32(100).u32(2).i16(40).i16(60).i16(80).i16(100),
      ),
      record(EMR.EOF, () => {}),
    );

    const m = makeRecordingCtx();
    expect(playEmf(file, m.ctx, 200, 200)).toBe(true);

    const moves = m.calls.filter((c) => c.op === 'moveTo');
    const lines = m.calls.filter((c) => c.op === 'lineTo');
    expect(moves.length).toBe(1);
    expect(moves[0].args).toEqual([40, 60]); // (40,60)·0.5·2
    expect(lines.length).toBe(1);
    expect(lines[0].args).toEqual([80, 100]); // (80,100)·0.5·2

    const strokes = m.calls.filter((c) => c.op === 'stroke');
    expect(strokes.length).toBe(1);
    expect(m.styles.stroke.at(-1)?.toLowerCase()).toBe('#0000ff'); // blue pen
    expect(m.calls.some((c) => c.op === 'fill')).toBe(false); // polyline never fills
  });

  it('MWT_LEFTMULTIPLY (iMode=2) composes xform × WT (scale-down then draw)', () => {
    // First SET ×16, then LEFTMULTIPLY ×(1/16). LEFT-multiplying 1/16 onto a ×16
    // WT yields identity, so logical (10,20) → device (10,20) → px (10,20) at a
    // 1:1 device mapping (bounds == target).
    const file = concat(
      emfHeader(0, 0, 100, 100),
      record(EMR.MODIFYWORLDTRANSFORM, (w) =>
        w.f32(16).f32(0).f32(0).f32(16).f32(0).f32(0).u32(4),
      ),
      record(EMR.MODIFYWORLDTRANSFORM, (w) =>
        w.f32(1 / 16).f32(0).f32(0).f32(1 / 16).f32(0).f32(0).u32(2),
      ),
      record(EMR.CREATEPEN, (w) => w.u32(1).u32(0).i32(1).i32(0).u32(0x00000000)),
      record(EMR.SELECTOBJECT, (w) => w.u32(1)),
      record(EMR.POLYLINE16, (w) =>
        w.i32(0).i32(0).i32(100).i32(100).u32(2).i16(10).i16(20).i16(30).i16(40),
      ),
      record(EMR.EOF, () => {}),
    );
    const m = makeRecordingCtx();
    playEmf(file, m.ctx, 100, 100);
    const moves = m.calls.filter((c) => c.op === 'moveTo');
    const lines = m.calls.filter((c) => c.op === 'lineTo');
    expect(moves[0].args[0]).toBeCloseTo(10, 4);
    expect(moves[0].args[1]).toBeCloseTo(20, 4);
    expect(lines[0].args[0]).toBeCloseTo(30, 4);
    expect(lines[0].args[1]).toBeCloseTo(40, 4);
  });

  it('a PS_NULL pen (style 5) does not stroke', () => {
    const file = concat(
      emfHeader(0, 0, 100, 100),
      record(EMR.CREATEPEN, (w) => w.u32(1).u32(5).i32(1).i32(0).u32(0)), // PS_NULL
      record(EMR.SELECTOBJECT, (w) => w.u32(1)),
      record(EMR.POLYLINE16, (w) =>
        w.i32(0).i32(0).i32(100).i32(100).u32(2).i16(0).i16(0).i16(10).i16(10),
      ),
      record(EMR.EOF, () => {}),
    );
    const m = makeRecordingCtx();
    playEmf(file, m.ctx, 100, 100);
    expect(m.calls.some((c) => c.op === 'stroke')).toBe(false);
  });
});

// ── playEmf: window/viewport mapping (MS-EMF 2.3.11 page→device) ─────────────

describe('playEmf — window/viewport mapping (page → device)', () => {
  it('MM_TEXT default leaves geometry at the world-transform mapping', () => {
    // No window/viewport records → the page→device stage is the identity, so a
    // 1:1 device mapping (bounds == target) draws logical points unchanged.
    const file = concat(
      emfHeader(0, 0, 100, 100),
      record(EMR.CREATEPEN, (w) => w.u32(1).u32(0).i32(1).i32(0).u32(0)),
      record(EMR.SELECTOBJECT, (w) => w.u32(1)),
      record(EMR.POLYLINE16, (w) =>
        w.i32(0).i32(0).i32(100).i32(100).u32(2).i16(10).i16(20).i16(30).i16(40),
      ),
      record(EMR.EOF, () => {}),
    );
    const m = makeRecordingCtx();
    expect(playEmf(file, m.ctx, 100, 100)).toBe(true);
    const moves = m.calls.filter((c) => c.op === 'moveTo');
    const lines = m.calls.filter((c) => c.op === 'lineTo');
    expect(moves[0].args).toEqual([10, 20]);
    expect(lines[0].args).toEqual([30, 40]);
  });

  it('maps page → device via window origin/extent and viewport origin/extent', () => {
    // Window {org 0,0; ext 1000×1000} → Viewport {org 0,0; ext 100×100} on a
    // 100×100 device (bounds 0..100), target 100×100 (device→target ×1).
    // page→device factor = vpExt/winExt = 100/1000 = 0.1, so logical (500,800)
    // → device (50,80) → px (50,80).
    const file = concat(
      emfHeader(0, 0, 100, 100),
      record(EMR.SETMAPMODE, (w) => w.u32(8)), // MM_ANISOTROPIC
      record(EMR.SETWINDOWEXTEX, (w) => w.i32(1000).i32(1000)),
      record(EMR.SETVIEWPORTEXTEX, (w) => w.i32(100).i32(100)),
      record(EMR.CREATEPEN, (w) => w.u32(1).u32(0).i32(1).i32(0).u32(0)),
      record(EMR.SELECTOBJECT, (w) => w.u32(1)),
      record(EMR.POLYLINE16, (w) =>
        w.i32(0).i32(0).i32(1000).i32(1000).u32(2).i16(500).i16(800).i16(100).i16(200),
      ),
      record(EMR.EOF, () => {}),
    );
    const m = makeRecordingCtx();
    expect(playEmf(file, m.ctx, 100, 100)).toBe(true);
    const moves = m.calls.filter((c) => c.op === 'moveTo');
    const lines = m.calls.filter((c) => c.op === 'lineTo');
    expect(moves[0].args[0]).toBeCloseTo(50, 4);
    expect(moves[0].args[1]).toBeCloseTo(80, 4);
    expect(lines[0].args[0]).toBeCloseTo(10, 4);
    expect(lines[0].args[1]).toBeCloseTo(20, 4);
  });

  it('honors window and viewport origins (page → device offset)', () => {
    // winOrg (100,100), vpOrg (10,20), ext 1:1. device = (page−winOrg)+vpOrg.
    // logical (150,160) → device (150−100+10, 160−100+20) = (60,80) → px.
    const file = concat(
      emfHeader(0, 0, 100, 100),
      record(EMR.SETMAPMODE, (w) => w.u32(8)),
      record(EMR.SETWINDOWORGEX, (w) => w.i32(100).i32(100)),
      record(EMR.SETVIEWPORTORGEX, (w) => w.i32(10).i32(20)),
      record(EMR.SETWINDOWEXTEX, (w) => w.i32(100).i32(100)),
      record(EMR.SETVIEWPORTEXTEX, (w) => w.i32(100).i32(100)),
      record(EMR.CREATEPEN, (w) => w.u32(1).u32(0).i32(1).i32(0).u32(0)),
      record(EMR.SELECTOBJECT, (w) => w.u32(1)),
      record(EMR.POLYLINE16, (w) =>
        w.i32(0).i32(0).i32(100).i32(100).u32(2).i16(150).i16(160).i16(110).i16(120),
      ),
      record(EMR.EOF, () => {}),
    );
    const m = makeRecordingCtx();
    playEmf(file, m.ctx, 100, 100);
    const moves = m.calls.filter((c) => c.op === 'moveTo');
    const lines = m.calls.filter((c) => c.op === 'lineTo');
    expect(moves[0].args[0]).toBeCloseTo(60, 4);
    expect(moves[0].args[1]).toBeCloseTo(80, 4);
    expect(lines[0].args[0]).toBeCloseTo(20, 4);
    expect(lines[0].args[1]).toBeCloseTo(40, 4);
  });

  it('brings large window-space geometry back on-canvas (pptx null-drop fix)', () => {
    // Regression for the pptx EMF null-drop: a chart drawn in a big logical
    // window (0..10000) with NO window→viewport mapping would land ~100× off
    // the device bounds (0..100) → clipped to nothing → drew=false → null.
    // With the mapping (window 10000 → viewport 100) it scales onto the raster.
    const file = concat(
      emfHeader(0, 0, 100, 100),
      record(EMR.SETMAPMODE, (w) => w.u32(8)),
      record(EMR.SETWINDOWEXTEX, (w) => w.i32(10000).i32(10000)),
      record(EMR.SETVIEWPORTEXTEX, (w) => w.i32(100).i32(100)),
      // brush ih=1 SOLID blue; polygon fills a big window-space rectangle.
      record(EMR.CREATEBRUSHINDIRECT, (w) => w.u32(1).u32(0).u32(0x00ff0000).u32(0)),
      record(EMR.SELECTOBJECT, (w) => w.u32(1)),
      record(EMR.POLYGON16, (w) =>
        w
          .i32(0)
          .i32(0)
          .i32(10000)
          .i32(10000)
          .u32(4)
          .i16(1000)
          .i16(1000)
          .i16(9000)
          .i16(1000)
          .i16(9000)
          .i16(9000)
          .i16(1000)
          .i16(9000),
      ),
      record(EMR.EOF, () => {}),
    );
    const m = makeRecordingCtx();
    expect(playEmf(file, m.ctx, 100, 100)).toBe(true);
    const moves = m.calls.filter((c) => c.op === 'moveTo');
    const lines = m.calls.filter((c) => c.op === 'lineTo');
    // All four corners land inside 0..100 (10..90), not off at 1000..9000.
    const xs = [moves[0].args[0], ...lines.map((l) => l.args[0])] as number[];
    const ys = [moves[0].args[1], ...lines.map((l) => l.args[1])] as number[];
    for (const v of [...xs, ...ys]) {
      expect(v).toBeGreaterThanOrEqual(0);
      expect(v).toBeLessThanOrEqual(100);
    }
    expect(moves[0].args[0]).toBeCloseTo(10, 4);
    expect(moves[0].args[1]).toBeCloseTo(10, 4);
    expect(m.calls.some((c) => c.op === 'fill')).toBe(true);
  });

  it('SCALEVIEWPORTEXTEX scales the viewport extent by num/denom', () => {
    // viewport ext 100, then SCALE ×(1/2) → 50. window 1000 → factor 50/1000.
    const file = concat(
      emfHeader(0, 0, 100, 100),
      record(EMR.SETMAPMODE, (w) => w.u32(8)),
      record(EMR.SETWINDOWEXTEX, (w) => w.i32(1000).i32(1000)),
      record(EMR.SETVIEWPORTEXTEX, (w) => w.i32(100).i32(100)),
      record(EMR.SCALEVIEWPORTEXTEX, (w) => w.i32(1).i32(2).i32(1).i32(2)),
      record(EMR.CREATEPEN, (w) => w.u32(1).u32(0).i32(1).i32(0).u32(0)),
      record(EMR.SELECTOBJECT, (w) => w.u32(1)),
      record(EMR.POLYLINE16, (w) =>
        w.i32(0).i32(0).i32(1000).i32(1000).u32(2).i16(1000).i16(1000).i16(0).i16(0),
      ),
      record(EMR.EOF, () => {}),
    );
    const m = makeRecordingCtx();
    playEmf(file, m.ctx, 100, 100);
    const moves = m.calls.filter((c) => c.op === 'moveTo');
    // (1000)·(50/1000) = 50.
    expect(moves[0].args[0]).toBeCloseTo(50, 4);
    expect(moves[0].args[1]).toBeCloseTo(50, 4);
  });

  it('scopes the window/viewport mapping to SAVEDC/RESTOREDC', () => {
    // SAVEDC, set a window/viewport mapping, RESTOREDC → the mapping reverts to
    // the identity, so the second polyline draws at the world mapping again.
    const file = concat(
      emfHeader(0, 0, 100, 100),
      record(EMR.CREATEPEN, (w) => w.u32(1).u32(0).i32(1).i32(0).u32(0)),
      record(EMR.SELECTOBJECT, (w) => w.u32(1)),
      record(EMR.SAVEDC, () => {}),
      record(EMR.SETMAPMODE, (w) => w.u32(8)),
      record(EMR.SETWINDOWEXTEX, (w) => w.i32(1000).i32(1000)),
      record(EMR.SETVIEWPORTEXTEX, (w) => w.i32(100).i32(100)),
      record(EMR.POLYLINE16, (w) =>
        w.i32(0).i32(0).i32(1000).i32(1000).u32(2).i16(500).i16(500).i16(0).i16(0),
      ),
      record(EMR.RESTOREDC, (w) => w.i32(-1)),
      record(EMR.POLYLINE16, (w) =>
        w.i32(0).i32(0).i32(100).i32(100).u32(2).i16(50).i16(50).i16(0).i16(0),
      ),
      record(EMR.EOF, () => {}),
    );
    const m = makeRecordingCtx();
    playEmf(file, m.ctx, 100, 100);
    const moves = m.calls.filter((c) => c.op === 'moveTo');
    // First polyline (scaled): 500·0.1 = 50. Second (post-restore, identity): 50.
    expect(moves[0].args[0]).toBeCloseTo(50, 4);
    expect(moves[1].args[0]).toBeCloseTo(50, 4);
  });
});

// ── playEmf: metric map modes + MM_ISOTROPIC ([MS-EMF] 2.1.21) ───────────────
//
// The five metric map modes (LOMETRIC/HIMETRIC/LOENGLISH/HIENGLISH/TWIPS) impose
// a fixed physical size per logical unit, derived from the EMR_HEADER reference-
// device resolution (px per mm). MM_ISOTROPIC additionally forces both axes to
// the SMALLER |vpExt/winExt| ratio so the picture keeps its aspect ratio.
//
// A shared header maps device→target 1:1 so the metric page→device scale is the
// only factor under test: dev 1000px / mm 100 ⇒ 10 px/mm; a 1000×1000 .01-mm
// frame ⇒ boundsW/H = 1000·(1000/(100·100)) = 100 on a 100×100 raster.
describe('playEmf — metric map modes + MM_ISOTROPIC ([MS-EMF] 2.1.21)', () => {
  /** Header whose device→target mapping is exactly 1:1 on a 100×100 target and
   *  whose reference device is 10 px/mm (drives the metric-mode extents). */
  const metricHeader = () =>
    emfHeader(0, 0, 100, 100, {
      frame: { l: 0, t: 0, r: 1000, b: 1000 }, // .01 mm → 10 mm square
      dev: { cx: 1000, cy: 1000 }, // px
      mm: { cx: 100, cy: 100 }, // mm ⇒ 10 px/mm, sx = 1000/(100·100) = 0.1
    });

  // Each metric mode: one logical unit is a fixed physical length. At 10 px/mm
  // the expected page→device (= target, since device→target is 1:1) px-per-unit
  // is `unitMm · 10`. A viewport origin lifts the y-UP metric axis back on-canvas.
  const MM_PER_INCH = 25.4;
  const cases: Array<{ name: string; mode: number; unitMm: number }> = [
    { name: 'MM_LOMETRIC (0.1 mm/unit)', mode: 2, unitMm: 0.1 },
    { name: 'MM_HIMETRIC (0.01 mm/unit)', mode: 3, unitMm: 0.01 },
    { name: 'MM_LOENGLISH (0.01 in/unit)', mode: 4, unitMm: 0.01 * MM_PER_INCH },
    { name: 'MM_HIENGLISH (0.001 in/unit)', mode: 5, unitMm: 0.001 * MM_PER_INCH },
    { name: 'MM_TWIPS (1/1440 in/unit)', mode: 6, unitMm: MM_PER_INCH / 1440 },
  ];

  for (const { name, mode, unitMm } of cases) {
    it(`${name}: scales logical units to px via the reference-device resolution`, () => {
      const pxPerUnit = unitMm * 10; // 10 px/mm reference device
      // A viewport origin of (0,100) cancels the y-UP flip so the drawn point
      // lands on-canvas: device Y = -yl·pxPerUnit + 100.
      const yl = 40;
      const xl = 60;
      const file = concat(
        metricHeader(),
        record(EMR.SETMAPMODE, (w) => w.u32(mode)),
        record(EMR.SETVIEWPORTORGEX, (w) => w.i32(0).i32(100)),
        record(EMR.CREATEPEN, (w) => w.u32(1).u32(0).i32(1).i32(0).u32(0)),
        record(EMR.SELECTOBJECT, (w) => w.u32(1)),
        record(EMR.POLYLINE16, (w) =>
          w.i32(0).i32(0).i32(1000).i32(1000).u32(2).i16(xl).i16(yl).i16(0).i16(0),
        ),
        record(EMR.EOF, () => {}),
      );
      const m = makeRecordingCtx();
      expect(playEmf(file, m.ctx, 100, 100)).toBe(true);
      const move = m.calls.find((c) => c.op === 'moveTo');
      // X: xl·pxPerUnit (x viewport origin 0). Y: y-UP ⇒ -yl·pxPerUnit + 100.
      expect(move?.args[0] as number).toBeCloseTo(xl * pxPerUnit, 3);
      expect(move?.args[1] as number).toBeCloseTo(-yl * pxPerUnit + 100, 3);
    });
  }

  it('metric mode with a viewport origin brings the y-UP axis on-canvas', () => {
    // Without a viewport origin, MM_LOMETRIC's y-UP flip drives y negative
    // (off-canvas, above the top); with a viewport origin the same point lands
    // inside the raster. Assert the sign flips as the origin is introduced.
    const withoutOrg = concat(
      metricHeader(),
      record(EMR.SETMAPMODE, (w) => w.u32(2)),
      record(EMR.CREATEPEN, (w) => w.u32(1).u32(0).i32(1).i32(0).u32(0)),
      record(EMR.SELECTOBJECT, (w) => w.u32(1)),
      record(EMR.POLYLINE16, (w) =>
        w.i32(0).i32(0).i32(1000).i32(1000).u32(2).i16(50).i16(50).i16(0).i16(0),
      ),
      record(EMR.EOF, () => {}),
    );
    const a = makeRecordingCtx();
    playEmf(withoutOrg, a.ctx, 100, 100);
    const yNoOrg = a.calls.find((c) => c.op === 'moveTo')?.args[1] as number;
    expect(yNoOrg).toBeLessThan(0); // y-UP ⇒ above the canvas top

    const withOrg = concat(
      metricHeader(),
      record(EMR.SETMAPMODE, (w) => w.u32(2)),
      record(EMR.SETVIEWPORTORGEX, (w) => w.i32(0).i32(100)),
      record(EMR.CREATEPEN, (w) => w.u32(1).u32(0).i32(1).i32(0).u32(0)),
      record(EMR.SELECTOBJECT, (w) => w.u32(1)),
      record(EMR.POLYLINE16, (w) =>
        w.i32(0).i32(0).i32(1000).i32(1000).u32(2).i16(50).i16(50).i16(0).i16(0),
      ),
      record(EMR.EOF, () => {}),
    );
    const b = makeRecordingCtx();
    playEmf(withOrg, b.ctx, 100, 100);
    const yOrg = b.calls.find((c) => c.op === 'moveTo')?.args[1] as number;
    expect(yOrg).toBeGreaterThan(0); // origin lifts it back on-canvas
    expect(yOrg).toBeCloseTo(yNoOrg + 100, 3); // shifted by exactly the origin
  });

  it('MM_ISOTROPIC forces both axes to the SMALLER |vpExt/winExt| ratio (equal aspect)', () => {
    // Window 1000×1000, viewport 500×100 ⇒ raw ratios x=0.5, y=0.1. ANISOTROPIC
    // would use those independently; ISOTROPIC forces BOTH to the smaller (0.1),
    // so a square logical box stays square (no x/y distortion). Compare the two
    // modes on the same geometry to prove the correction only applies to 7.
    const geom = (mode: number) =>
      concat(
        emfHeader(0, 0, 100, 100),
        record(EMR.SETMAPMODE, (w) => w.u32(mode)),
        record(EMR.SETWINDOWEXTEX, (w) => w.i32(1000).i32(1000)),
        record(EMR.SETVIEWPORTEXTEX, (w) => w.i32(500).i32(100)),
        record(EMR.CREATEPEN, (w) => w.u32(1).u32(0).i32(1).i32(0).u32(0)),
        record(EMR.SELECTOBJECT, (w) => w.u32(1)),
        record(EMR.POLYLINE16, (w) =>
          w.i32(0).i32(0).i32(1000).i32(1000).u32(2).i16(200).i16(200).i16(0).i16(0),
        ),
        record(EMR.EOF, () => {}),
      );

    const aniso = makeRecordingCtx();
    playEmf(geom(8), aniso.ctx, 100, 100); // MM_ANISOTROPIC
    const anisoMove = aniso.calls.find((c) => c.op === 'moveTo');
    // Independent axes: x = 200·0.5 = 100, y = 200·0.1 = 20 (distorted).
    expect(anisoMove?.args[0] as number).toBeCloseTo(100, 3);
    expect(anisoMove?.args[1] as number).toBeCloseTo(20, 3);

    const iso = makeRecordingCtx();
    playEmf(geom(7), iso.ctx, 100, 100); // MM_ISOTROPIC
    const isoMove = iso.calls.find((c) => c.op === 'moveTo');
    // Both axes use min(0.5, 0.1) = 0.1 ⇒ x = y = 200·0.1 = 20 (square kept).
    expect(isoMove?.args[0] as number).toBeCloseTo(20, 3);
    expect(isoMove?.args[1] as number).toBeCloseTo(20, 3);
  });

  it('MM_ISOTROPIC preserves the viewport y-flip sign while equalizing magnitude', () => {
    // Negative viewport Y extent (y-UP) with unequal magnitudes: the shared
    // isotropic |scale| is the smaller magnitude, but Y keeps its negative sign.
    const file = concat(
      emfHeader(0, 0, 100, 100),
      record(EMR.SETMAPMODE, (w) => w.u32(7)), // MM_ISOTROPIC
      record(EMR.SETWINDOWEXTEX, (w) => w.i32(1000).i32(1000)),
      record(EMR.SETVIEWPORTEXTEX, (w) => w.i32(500).i32(-100)), // y-UP, |0.1| < 0.5
      record(EMR.SETVIEWPORTORGEX, (w) => w.i32(0).i32(50)),
      record(EMR.CREATEPEN, (w) => w.u32(1).u32(0).i32(1).i32(0).u32(0)),
      record(EMR.SELECTOBJECT, (w) => w.u32(1)),
      record(EMR.POLYLINE16, (w) =>
        w.i32(0).i32(0).i32(1000).i32(1000).u32(2).i16(200).i16(200).i16(0).i16(0),
      ),
      record(EMR.EOF, () => {}),
    );
    const m = makeRecordingCtx();
    playEmf(file, m.ctx, 100, 100);
    const move = m.calls.find((c) => c.op === 'moveTo');
    // x = 200·0.1 = 20; y = 200·(−0.1) + 50 = 30 (magnitude equalized, sign kept).
    expect(move?.args[0] as number).toBeCloseTo(20, 3);
    expect(move?.args[1] as number).toBeCloseTo(30, 3);
  });

  it('composes the world transform with the window/viewport mapping (pin the order)', () => {
    // world → PAGE (world transform) THEN page → DEVICE (window/viewport), per
    // [MS-EMF] 2.3.11/2.3.12. World scales ×2 + translates (+10,+20); the mapping
    // then scales ×0.5 (window 1000 → viewport 500). The ORDER matters: applying
    // the device scale to the translation too (0.5·(2·xl+10)) differs from
    // translating in device space. Pin the world-then-device composition.
    const xl = 30;
    const yl = 40;
    const file = concat(
      emfHeader(0, 0, 100, 100),
      // world transform: m11=m22=2, dx=10, dy=20 (MWT_SET).
      record(EMR.MODIFYWORLDTRANSFORM, (w) =>
        w.f32(2).f32(0).f32(0).f32(2).f32(10).f32(20).u32(4),
      ),
      record(EMR.SETMAPMODE, (w) => w.u32(8)), // MM_ANISOTROPIC
      record(EMR.SETWINDOWEXTEX, (w) => w.i32(1000).i32(1000)),
      record(EMR.SETVIEWPORTEXTEX, (w) => w.i32(500).i32(500)), // ×0.5
      record(EMR.CREATEPEN, (w) => w.u32(1).u32(0).i32(1).i32(0).u32(0)),
      record(EMR.SELECTOBJECT, (w) => w.u32(1)),
      record(EMR.POLYLINE16, (w) =>
        w.i32(0).i32(0).i32(1000).i32(1000).u32(2).i16(xl).i16(yl).i16(0).i16(0),
      ),
      record(EMR.EOF, () => {}),
    );
    const m = makeRecordingCtx();
    expect(playEmf(file, m.ctx, 100, 100)).toBe(true);
    const move = m.calls.find((c) => c.op === 'moveTo');
    // page = world(xl,yl) = (2·30+10, 2·40+20) = (70,100).
    // device = page·0.5 = (35,50). device→target is 1:1 (bounds==target).
    expect(move?.args[0] as number).toBeCloseTo((2 * xl + 10) * 0.5, 3);
    expect(move?.args[1] as number).toBeCloseTo((2 * yl + 20) * 0.5, 3);
  });
});

// ── playEmf: polygon fill + object table ─────────────────────────────────────

describe('playEmf — polygon16 fill + brush/pen select', () => {
  it('fills + strokes a POLYGON16 with the current brush + pen', () => {
    const file = concat(
      emfHeader(0, 0, 10, 10),
      // brush ih=1 SOLID(0) green 0x0000FF00; pen ih=2 SOLID red 0x000000FF.
      record(EMR.CREATEBRUSHINDIRECT, (w) => w.u32(1).u32(0).u32(0x0000ff00).u32(0)),
      record(EMR.CREATEPEN, (w) => w.u32(2).u32(0).i32(1).i32(0).u32(0x000000ff)),
      record(EMR.SELECTOBJECT, (w) => w.u32(1)), // brush
      record(EMR.SELECTOBJECT, (w) => w.u32(2)), // pen
      record(EMR.POLYGON16, (w) =>
        w.i32(0).i32(0).i32(10).i32(10).u32(3).i16(0).i16(0).i16(10).i16(0).i16(5).i16(10),
      ),
      record(EMR.EOF, () => {}),
    );
    const m = makeRecordingCtx();
    expect(playEmf(file, m.ctx, 10, 10)).toBe(true);
    expect(m.styles.fill.at(-1)?.toLowerCase()).toBe('#00ff00'); // green brush
    expect(m.styles.stroke.at(-1)?.toLowerCase()).toBe('#ff0000'); // red pen
  });

  it('a BS_NULL brush (style 1) does not fill', () => {
    const file = concat(
      emfHeader(0, 0, 10, 10),
      record(EMR.CREATEBRUSHINDIRECT, (w) => w.u32(1).u32(1).u32(0).u32(0)), // BS_NULL
      record(EMR.CREATEPEN, (w) => w.u32(2).u32(0).i32(1).i32(0).u32(0)),
      record(EMR.SELECTOBJECT, (w) => w.u32(1)),
      record(EMR.SELECTOBJECT, (w) => w.u32(2)),
      record(EMR.POLYGON16, (w) =>
        w.i32(0).i32(0).i32(10).i32(10).u32(3).i16(0).i16(0).i16(10).i16(0).i16(5).i16(10),
      ),
      record(EMR.EOF, () => {}),
    );
    const m = makeRecordingCtx();
    playEmf(file, m.ctx, 10, 10);
    expect(m.calls.some((c) => c.op === 'fill')).toBe(false);
    expect(m.calls.some((c) => c.op === 'stroke')).toBe(true);
  });

  it('DELETEOBJECT of the selected pen clears it and frees the slot (no later stroke)', () => {
    // Mirrors the WMF twin: deleting the object currently selected into the DC
    // un-selects it (curPen → null), and a later SELECTOBJECT of the now-empty
    // slot is a no-op. So the polyline issues no stroke — and never throws.
    const file = concat(
      emfHeader(0, 0, 10, 10),
      record(EMR.CREATEPEN, (w) => w.u32(1).u32(0).i32(1).i32(0).u32(0x000000ff)), // red
      record(EMR.SELECTOBJECT, (w) => w.u32(1)),
      record(EMR.DELETEOBJECT, (w) => w.u32(1)), // clears curPen + frees slot 1
      record(EMR.SELECTOBJECT, (w) => w.u32(1)), // slot now empty → no-op
      record(EMR.POLYLINE16, (w) =>
        w.i32(0).i32(0).i32(10).i32(10).u32(2).i16(0).i16(0).i16(5).i16(5),
      ),
      record(EMR.EOF, () => {}),
    );
    const m = makeRecordingCtx();
    expect(() => playEmf(file, m.ctx, 10, 10)).not.toThrow();
    expect(m.calls.some((c) => c.op === 'stroke')).toBe(false);
  });

  it('reuses a freed slot when an index is recreated after DELETEOBJECT', () => {
    // Create red pen at ih=1, select+delete it, recreate ih=1 as green, select →
    // the stroke uses the recreated green pen (Map slot reused by index).
    const file = concat(
      emfHeader(0, 0, 10, 10),
      record(EMR.CREATEPEN, (w) => w.u32(1).u32(0).i32(1).i32(0).u32(0x000000ff)), // red
      record(EMR.SELECTOBJECT, (w) => w.u32(1)),
      record(EMR.DELETEOBJECT, (w) => w.u32(1)),
      record(EMR.CREATEPEN, (w) => w.u32(1).u32(0).i32(1).i32(0).u32(0x0000ff00)), // green
      record(EMR.SELECTOBJECT, (w) => w.u32(1)),
      record(EMR.POLYLINE16, (w) =>
        w.i32(0).i32(0).i32(10).i32(10).u32(2).i16(0).i16(0).i16(5).i16(5),
      ),
      record(EMR.EOF, () => {}),
    );
    const m = makeRecordingCtx();
    playEmf(file, m.ctx, 10, 10);
    expect(m.styles.stroke.at(-1)?.toLowerCase()).toBe('#00ff00'); // recreated green
  });
});

// ── playEmf: text-out ────────────────────────────────────────────────────────

describe('playEmf — EXTTEXTOUTW text', () => {
  it('draws the UTF-16 string with the selected font color at the mapped ref point', () => {
    // offString = byte offset from the RECORD start to the string. The EXTTEXTOUTW
    // data layout up to the string: header(8) + RECTL(16) + iGraphicsMode(4) +
    // exScale(4) + eyScale(4) + ptlReference(8) + nChars(4) + offString(4) +
    // fOptions(4) + rcl RECTL(16) + offDx(4) = 76 bytes → the string starts at 76.
    const text = 'F1';
    const file = concat(
      emfHeader(0, 0, 100, 100),
      // font ih=1, lfHeight=-12 (negative = char height), weight=400, not italic.
      record(EMR.EXTCREATEFONTINDIRECTW, (w) => {
        w.u32(1); // ihObject
        // LOGFONT starts at record offset 12 (data offset 4):
        w.i32(-12).i32(0).i32(0).i32(0).i32(400); // lfHeight..lfWeight
        w.raw(0, 0, 0, 0); // lfItalic, lfUnderline, lfStrikeOut, lfCharSet
        w.raw(0, 0, 0, 0); // lfOutPrecision..lfPitchAndFamily (4 bytes)
        // lfFaceName (UTF-16, 32 code units) at LOGFONT offset 28:
        const face = 'Arial';
        w.utf16(face);
        for (let i = face.length; i < 32; i++) w.u16(0);
      }),
      record(EMR.SETTEXTCOLOR, (w) => w.u32(0x000000ff)), // red
      record(EMR.SELECTOBJECT, (w) => w.u32(1)), // font
      record(EMR.EXTTEXTOUTW, (w) => {
        w.i32(0).i32(0).i32(100).i32(100); // RECTL rclBounds
        w.u32(1); // iGraphicsMode
        w.f32(1).f32(1); // exScale, eyScale
        w.i32(20).i32(30); // ptlReference (logical) → identity WT → device (20,30)
        w.u32(text.length); // nChars
        w.u32(76); // offString (computed above)
        w.u32(0); // fOptions
        w.i32(0).i32(0).i32(0).i32(0); // rcl RECTL
        w.u32(0); // offDx
        w.utf16(text); // the string at record offset 76
      }),
      record(EMR.EOF, () => {}),
    );
    const m = makeRecordingCtx();
    expect(playEmf(file, m.ctx, 100, 100)).toBe(true);
    const texts = m.calls.filter((c) => c.op === 'fillText');
    expect(texts.length).toBe(1);
    expect(texts[0].args[0]).toBe('F1');
    expect(texts[0].args.slice(1)).toEqual([20, 30]); // identity WT, ×1 device
    expect(m.styles.text.at(-1)?.toLowerCase()).toBe('#ff0000'); // red text color
  });

  it('rotates text by lfEscapement (vertical axis labels)', () => {
    // lfEscapement = 900 → 90° counterclockwise. The draw becomes
    // translate(refPx) + rotate(−90°) + fillText at the origin.
    const text = 'Dx';
    const file = concat(
      emfHeader(0, 0, 100, 100),
      record(EMR.EXTCREATEFONTINDIRECTW, (w) => {
        w.u32(1);
        // lfHeight, lfWidth, lfEscapement=900, lfOrientation, lfWeight
        w.i32(-12).i32(0).i32(900).i32(0).i32(400);
        w.raw(0, 0, 0, 0);
        w.raw(0, 0, 0, 0);
        const face = 'Arial';
        w.utf16(face);
        for (let i = face.length; i < 32; i++) w.u16(0);
      }),
      record(EMR.SELECTOBJECT, (w) => w.u32(1)),
      record(EMR.EXTTEXTOUTW, (w) => {
        w.i32(0).i32(0).i32(100).i32(100);
        w.u32(1);
        w.f32(1).f32(1);
        w.i32(20).i32(30);
        w.u32(text.length);
        w.u32(76);
        w.u32(0);
        w.i32(0).i32(0).i32(0).i32(0);
        w.u32(0);
        w.utf16(text);
      }),
      record(EMR.EOF, () => {}),
    );
    const m = makeRecordingCtx();
    expect(playEmf(file, m.ctx, 100, 100)).toBe(true);
    expect(m.calls.find((c) => c.op === 'translate')?.args).toEqual([20, 30]);
    expect(m.calls.find((c) => c.op === 'rotate')?.args[0]).toBeCloseTo(-Math.PI / 2, 6);
    expect(m.calls.find((c) => c.op === 'fillText')?.args).toEqual(['Dx', 0, 0]);
  });
});

describe('playEmf — picture frame mapping + path clip', () => {
  it('maps the picture frame onto the target — ink fills a sub-rectangle, not the whole raster', () => {
    // Ink bounds 0..100 (device px); the picture FRAME is twice as large
    // (rclFrame 0..200 .01 mm with a 1 px/.01 mm reference device ⇒ frame device
    // extent 200). GDI maps the FRAME to the target, so on a 200×200 raster the
    // ink corner (100,100) lands at (100,100) — half the frame — NOT (200,200) as
    // a bounds-fill mapping would give. This is what lets an `<a:srcRect>` crop
    // (relative to the frame) select the ink region.
    const file = concat(
      emfHeader(0, 0, 100, 100, {
        frame: { l: 0, t: 0, r: 200, b: 200 },
        dev: { cx: 50800, cy: 28600 }, // = mm × 100 ⇒ 1 device px per .01 mm
        mm: { cx: 508, cy: 286 },
      }),
      record(EMR.CREATEPEN, (w) => w.u32(1).u32(0).i32(1).i32(0).u32(0x00ff0000)),
      record(EMR.SELECTOBJECT, (w) => w.u32(1)),
      record(EMR.POLYLINE16, (w) =>
        w.i32(0).i32(0).i32(100).i32(100).u32(2).i16(0).i16(0).i16(100).i16(100),
      ),
      record(EMR.EOF, () => {}),
    );
    const m = makeRecordingCtx();
    expect(playEmf(file, m.ctx, 200, 200)).toBe(true);
    expect(m.calls.find((c) => c.op === 'moveTo')?.args).toEqual([0, 0]);
    expect(m.calls.find((c) => c.op === 'lineTo')?.args).toEqual([100, 100]);
  });

  it('BEGINPATH…ENDPATH + SELECTCLIPPATH sets a clip; the path geometry is not filled', () => {
    // A polygon between BEGINPATH and ENDPATH defines the clip shape — it must
    // build the path (no fill/stroke) and SELECTCLIPPATH applies it as a clip,
    // bracketed by SAVEDC/RESTOREDC on the canvas (sample-13 Fig.3 clips a DIB to
    // the bar shapes). Without this the clip-path polygon would paint a red fill.
    const file = concat(
      emfHeader(0, 0, 100, 100),
      record(EMR.CREATEBRUSHINDIRECT, (w) => w.u32(1).u32(0).u32(0x000000ff).u32(0)), // red solid
      record(EMR.SELECTOBJECT, (w) => w.u32(1)),
      record(EMR.SAVEDC, () => {}),
      record(EMR.BEGINPATH, () => {}),
      record(EMR.POLYGON16, (w) =>
        w.i32(0).i32(0).i32(50).i32(50).u32(3).i16(0).i16(0).i16(50).i16(0).i16(50).i16(50),
      ),
      record(EMR.ENDPATH, () => {}),
      record(EMR.SELECTCLIPPATH, (w) => w.u32(1)), // RGN_AND
      record(EMR.RESTOREDC, (w) => w.i32(-1)),
      record(EMR.EOF, () => {}),
    );
    const m = makeRecordingCtx();
    playEmf(file, m.ctx, 100, 100);
    expect(m.calls.some((c) => c.op === 'clip')).toBe(true);
    expect(m.calls.some((c) => c.op === 'fill')).toBe(false); // in-path polygon not filled
    expect(m.calls.some((c) => c.op === 'save')).toBe(true);
    expect(m.calls.some((c) => c.op === 'restore')).toBe(true);
  });

  it('fills a bracketed Bézier path when EMR_FILLPATH consumes it', () => {
    const file = concat(
      emfHeader(0, 0, 100, 100),
      record(EMR.CREATEBRUSHINDIRECT, (w) => w.u32(1).u32(0).u32(0x0000ff00).u32(0)),
      record(EMR.SELECTOBJECT, (w) => w.u32(1)),
      record(EMR.BEGINPATH, () => {}),
      record(EMR.MOVETOEX, (w) => w.i32(10).i32(50)),
      record(EMR.POLYBEZIERTO16, (w) =>
        w.i32(10).i32(10).i32(90).i32(90).u32(3)
          .i16(25).i16(10).i16(75).i16(10).i16(90).i16(50),
      ),
      record(EMR.CLOSEFIGURE, () => {}),
      record(EMR.ENDPATH, () => {}),
      record(EMR.FILLPATH, (w) => w.i32(10).i32(10).i32(90).i32(90)),
      record(EMR.EOF, () => {}),
    );
    const m = makeRecordingCtx();

    expect(playEmf(file, m.ctx, 100, 100)).toBe(true);
    expect(m.calls.find((call) => call.op === 'moveTo')?.args).toEqual([10, 50]);
    expect(m.calls.some((call) => call.op === 'bezierCurveTo')).toBe(true);
    expect(m.calls.filter((call) => call.op === 'fill')).toHaveLength(1);
    expect(m.styles.fill).toEqual(['#00ff00']);
  });

  it('accumulates a bracketed polyline even when no pen is selected', () => {
    const file = concat(
      emfHeader(0, 0, 100, 100),
      record(EMR.CREATEBRUSHINDIRECT, (w) => w.u32(1).u32(0).u32(0x0000ff00).u32(0)),
      record(EMR.SELECTOBJECT, (w) => w.u32(1)),
      record(EMR.BEGINPATH, () => {}),
      record(EMR.POLYLINE16, (w) =>
        w.i32(10).i32(10).i32(90).i32(10).u32(2)
          .i16(10).i16(10).i16(90).i16(10),
      ),
      record(EMR.POLYLINE16, (w) =>
        w.i32(90).i32(10).i32(90).i32(90).u32(2)
          .i16(90).i16(10).i16(90).i16(90),
      ),
      record(EMR.ENDPATH, () => {}),
      record(EMR.FILLPATH, (w) => w.i32(10).i32(10).i32(90).i32(90)),
      record(EMR.EOF, () => {}),
    );
    const m = makeRecordingCtx();

    expect(playEmf(file, m.ctx, 100, 100)).toBe(true);
    expect(m.calls.filter((call) => call.op === 'lineTo')).toHaveLength(2);
    expect(m.calls.filter((call) => call.op === 'stroke')).toHaveLength(0);
    expect(m.calls.filter((call) => call.op === 'fill')).toHaveLength(1);
  });

  it('discards a path bracket that exceeds the cumulative geometry budget', () => {
    const file = concat(
      emfHeader(0, 0, 100, 100),
      record(EMR.CREATEBRUSHINDIRECT, (w) => w.u32(1).u32(0).u32(0x0000ff00).u32(0)),
      record(EMR.SELECTOBJECT, (w) => w.u32(1)),
      record(EMR.BEGINPATH, () => {}),
      record(EMR.MOVETOEX, (w) => w.i32(0).i32(0)),
      record(EMR.POLYLINE16, (w) =>
        w.i32(0).i32(0).i32(30).i32(30).u32(3)
          .i16(0).i16(0).i16(20).i16(20).i16(30).i16(30),
      ),
      record(EMR.LINETO, (w) => w.i32(40).i32(40)),
      record(EMR.ENDPATH, () => {}),
      record(EMR.FILLPATH, (w) => w.i32(0).i32(0).i32(40).i32(40)),
      // Rendering must recover after the rejected bracket.
      record(EMR.CREATEPEN, (w) => w.u32(2).u32(0).i32(1).i32(0).u32(0)),
      record(EMR.SELECTOBJECT, (w) => w.u32(2)),
      record(EMR.POLYLINE16, (w) =>
        w.i32(0).i32(0).i32(10).i32(10).u32(2).i16(0).i16(0).i16(10).i16(10),
      ),
      record(EMR.EOF, () => {}),
    );
    const m = makeRecordingCtx();

    const playWithLimits = playEmf as unknown as (
      bytes: Uint8Array,
      ctx: CanvasRenderingContext2D,
      width: number,
      height: number,
      limits: { maxPathCommands: number },
    ) => boolean;
    expect(playWithLimits(file, m.ctx, 100, 100, { maxPathCommands: 4 })).toBe(true);
    expect(m.calls.filter((call) => call.op === 'fill')).toHaveLength(0);
    expect(m.calls.filter((call) => call.op === 'stroke')).toHaveLength(1);
  });

  it('rejects a poly-polygon whose per-polygon counts exceed its declared total', () => {
    const file = concat(
      emfHeader(0, 0, 100, 100),
      record(EMR.CREATEBRUSHINDIRECT, (w) => w.u32(1).u32(0).u32(0x0000ff00).u32(0)),
      record(EMR.SELECTOBJECT, (w) => w.u32(1)),
      record(EMR.POLYPOLYGON16, (w) =>
        w.i32(0).i32(0).i32(50).i32(50)
          .u32(1).u32(1) // one polygon, falsely declares one total point
          .u32(3)
          .i16(0).i16(0).i16(50).i16(0).i16(50).i16(50),
      ),
      record(EMR.EOF, () => {}),
    );
    const m = makeRecordingCtx();

    playEmf(file, m.ctx, 100, 100);
    expect(m.calls.filter((call) => call.op === 'fill')).toHaveLength(0);
  });

  it('decodes a 4bpp DIB pattern brush (MATLAB bar-chart fill colour)', () => {
    // A 2×2 4bpp BI_RGB DIB, 1-entry palette = blue, all pixels index 0. The
    // pattern brush averages to that blue, so a polygon fills blue — exercising
    // the 4bpp decode path that paints sample-13 Fig.3's bars.
    const file = concat(
      emfHeader(0, 0, 100, 100),
      record(EMR.CREATEDIBPATTERNBRUSHPT, (w) => {
        w.u32(1).u32(0).u32(32).u32(44).u32(76).u32(8); // ih,iUsage,offBmi,cbBmi,offBits,cbBits
        // BITMAPINFOHEADER (40 bytes) @ record offset 32:
        w.u32(40).i32(2).i32(2).u16(1).u16(4).u32(0).u32(0).i32(0).i32(0).u32(1).u32(0);
        w.raw(0xff, 0x00, 0x00, 0x00); // palette[0] = blue (B,G,R,reserved)
        w.raw(0, 0, 0, 0, 0, 0, 0, 0); // 2 rows × 4-byte stride, all index 0
      }),
      record(EMR.SELECTOBJECT, (w) => w.u32(1)),
      record(EMR.POLYGON16, (w) =>
        w.i32(0).i32(0).i32(40).i32(40).u32(3).i16(0).i16(0).i16(40).i16(0).i16(40).i16(40),
      ),
      record(EMR.EOF, () => {}),
    );
    const m = makeRecordingCtx();
    expect(playEmf(file, m.ctx, 100, 100)).toBe(true);
    expect(m.styles.fill.at(-1)?.toLowerCase()).toBe('#0000ff');
  });
});

// ── playEmf: robustness ──────────────────────────────────────────────────────

describe('playEmf — robustness', () => {
  it('returns false for non-EMF bytes', () => {
    const m = makeRecordingCtx();
    expect(playEmf(new Uint8Array([0xde, 0xad, 0xbe, 0xef, 0, 0, 0, 0]), m.ctx, 10, 10)).toBe(
      false,
    );
  });

  it('stops gracefully on a record with a bogus (misaligned/too-small) size', () => {
    const bad = concat(
      emfHeader(0, 0, 10, 10),
      record(EMR.CREATEPEN, (w) => w.u32(1).u32(0).i32(1).i32(0).u32(0)),
      record(EMR.SELECTOBJECT, (w) => w.u32(1)),
      record(EMR.POLYLINE16, (w) =>
        w.i32(0).i32(0).i32(10).i32(10).u32(2).i16(0).i16(0).i16(5).i16(5),
      ),
      // a corrupt record: nSize = 4 (< 8) → must stop the loop, not throw.
      new Writer().u32(EMR.POLYLINE16).u32(4).build(),
      record(EMR.EOF, () => {}),
    );
    const m = makeRecordingCtx();
    expect(() => playEmf(bad, m.ctx, 10, 10)).not.toThrow();
    expect(m.calls.some((c) => c.op === 'stroke')).toBe(true);
  });

  it('skips unrecognized records by nSize without throwing', () => {
    const file = concat(
      emfHeader(0, 0, 10, 10),
      // an unknown record type with arbitrary payload — must be skipped by nSize.
      record(9999, (w) => w.u32(1).u32(2).u32(3)),
      record(EMR.CREATEPEN, (w) => w.u32(1).u32(0).i32(1).i32(0).u32(0)),
      record(EMR.SELECTOBJECT, (w) => w.u32(1)),
      record(EMR.POLYLINE16, (w) =>
        w.i32(0).i32(0).i32(10).i32(10).u32(2).i16(0).i16(0).i16(5).i16(5),
      ),
      record(EMR.EOF, () => {}),
    );
    const m = makeRecordingCtx();
    expect(() => playEmf(file, m.ctx, 10, 10)).not.toThrow();
    expect(m.calls.some((c) => c.op === 'stroke')).toBe(true);
  });
});

// ── renderEmfToBitmap: OffscreenCanvas wrapper (browser/worker only) ──────────

describe('renderEmfToBitmap', () => {
  beforeEach(() => {
    // OffscreenCanvas + createImageBitmap don't exist in the node test env.
    vi.stubGlobal(
      'OffscreenCanvas',
      class {
        width: number;
        height: number;
        constructor(w: number, h: number) {
          this.width = w;
          this.height = h;
        }
        getContext() {
          return {
            fillStyle: '#000',
            strokeStyle: '#000',
            lineWidth: 1,
            font: '10px sans-serif',
            textAlign: 'left',
            textBaseline: 'top',
            lineJoin: 'miter',
            lineCap: 'butt',
            save() {},
            restore() {},
            beginPath() {},
            closePath() {},
            moveTo() {},
            lineTo() {},
            bezierCurveTo() {},
            ellipse() {},
            rect() {},
            stroke() {},
            fill() {},
            fillText() {},
          };
        }
      },
    );
    vi.stubGlobal(
      'createImageBitmap',
      vi.fn(
        async (src: { width: number; height: number }) =>
          ({ width: src.width, height: src.height, close() {} }) as unknown as ImageBitmap,
      ),
    );
  });
  afterEach(() => vi.unstubAllGlobals());

  it('rasterizes a minimal EMF to an ImageBitmap of the requested size', async () => {
    const file = concat(
      emfHeader(0, 0, 100, 100),
      record(EMR.CREATEPEN, (w) => w.u32(1).u32(0).i32(1).i32(0).u32(0)),
      record(EMR.SELECTOBJECT, (w) => w.u32(1)),
      record(EMR.POLYLINE16, (w) =>
        w.i32(0).i32(0).i32(100).i32(100).u32(2).i16(0).i16(0).i16(50).i16(50),
      ),
      record(EMR.EOF, () => {}),
    );
    const { bitmap: bmp, unsupported } = await renderEmfToBitmap(file, 64, 48);
    expect(bmp).not.toBeNull();
    expect(bmp?.width).toBe(64);
    expect(bmp?.height).toBe(48);
    expect(unsupported).toEqual([]);
  });

  it('returns a null bitmap for non-EMF bytes', async () => {
    expect(await renderEmfToBitmap(new Uint8Array([1, 2, 3, 4]), 10, 10)).toEqual({ bitmap: null, unsupported: [] });
  });

  it('returns a null bitmap when nothing draws (header + EOF only)', async () => {
    const file = concat(emfHeader(0, 0, 10, 10), record(EMR.EOF, () => {}));
    expect(await renderEmfToBitmap(file, 16, 16)).toEqual({ bitmap: null, unsupported: [] });
  });

  it('returns a null bitmap for a non-positive target size', async () => {
    const file = concat(emfHeader(0, 0, 10, 10), record(EMR.EOF, () => {}));
    expect((await renderEmfToBitmap(file, 0, 10)).bitmap).toBeNull();
    expect((await renderEmfToBitmap(file, 10, 0)).bitmap).toBeNull();
  });

  // A drawn polyline followed by POLYTEXTOUTW, a text record the player does
  // not implement: the partial picture the player has always drawn.
  const partial = () => concat(
    emfHeader(0, 0, 100, 100),
    record(EMR.CREATEPEN, (w) => w.u32(1).u32(0).i32(1).i32(0).u32(0)),
    record(EMR.SELECTOBJECT, (w) => w.u32(1)),
    record(EMR.POLYLINE16, (w) => w.i32(0).i32(0).i32(100).i32(100).u32(2).i16(0).i16(0).i16(50).i16(50)),
    record(97, (w) => w.u32(0)),
    record(EMR.EOF, () => {}),
  );

  it('returns the unsupported content with the bitmap through the normal loading path', async () => {
    const { bitmap, unsupported } = await renderEmfToBitmap(partial(), 64, 48);
    expect(bitmap).not.toBeNull();
    expect(unsupported).toEqual(['EMR_POLYTEXTOUTW']);
  });

  it('decodeRasterOrMetafile keeps the partial drawing with its report by default, and rejects it on request', async () => {
    const blob = () => new Blob([partial() as Uint8Array<ArrayBuffer>], { type: 'image/x-emf' });
    const drawn = await decodeRasterOrMetafile(blob(), { widthPt: 48, heightPt: 36 });
    expect(drawn).not.toBeNull();
    expect(getIncompleteMetafileReport(drawn)).toEqual({ format: 'emf', unsupported: ['EMR_POLYTEXTOUTW'] });

    const rejected = await decodeRasterOrMetafile(blob(), { widthPt: 48, heightPt: 36, incompleteMetafile: 'reject' })
      .then(() => undefined, (error: unknown) => error);
    expect(isOoxmlIncompleteMetafileError(rejected)).toBe(true);
    expect(rejected).toMatchObject({ code: 'ooxml-incomplete-metafile', format: 'emf', unsupported: ['EMR_POLYTEXTOUTW'] });

    // A complete metafile carries no report under either policy.
    const complete = concat(
      emfHeader(0, 0, 100, 100),
      record(EMR.CREATEPEN, (w) => w.u32(1).u32(0).i32(1).i32(0).u32(0)),
      record(EMR.SELECTOBJECT, (w) => w.u32(1)),
      record(EMR.POLYLINE16, (w) => w.i32(0).i32(0).i32(100).i32(100).u32(2).i16(0).i16(0).i16(50).i16(50)),
      record(EMR.EOF, () => {}),
    );
    const strict = await decodeRasterOrMetafile(new Blob([complete as Uint8Array<ArrayBuffer>]), { widthPt: 48, heightPt: 36, incompleteMetafile: 'reject' });
    expect(strict).not.toBeNull();
    expect(getIncompleteMetafileReport(strict)).toBeUndefined();
  });

  it('keeps the incomplete report on a bitmap derived by a pixel effect', async () => {
    const fetchImage = async () => new Blob([partial() as Uint8Array<ArrayBuffer>], { type: 'image/x-emf' });
    const offscreenFactory = (w: number, h: number) => ({
      width: w,
      height: h,
      getContext: () => ({
        drawImage() {},
        getImageData: () => ({ data: new Uint8ClampedArray(w * h * 4) }),
        putImageData() {},
      }),
    }) as unknown as OffscreenCanvas;
    const opts = { widthPt: 48, heightPt: 36, offscreenFactory };
    const base = await getCachedBitmapByPath('p.emf', 'image/x-emf', fetchImage, opts);
    const effects = { effects: [{ type: 'grayscale' as const }] };
    const derived = await getCachedDuotoneBitmapByPath('p.emf', 'image/x-emf', effects, fetchImage, opts);
    expect(derived).not.toBe(base);
    expect(getIncompleteMetafileReport(derived)).toEqual({ format: 'emf', unsupported: ['EMR_POLYTEXTOUTW'] });
    // A later hit on the derived cache entry carries it as well.
    const again = await getCachedDuotoneBitmapByPath('p.emf', 'image/x-emf', effects, fetchImage, opts);
    expect(getIncompleteMetafileReport(again)?.unsupported).toEqual(['EMR_POLYTEXTOUTW']);
    dropBitmapCacheByPath(fetchImage);
  });

  it('never serves a cached partial picture to a strict request', async () => {
    const fetchImage = vi.fn(async () => new Blob([partial() as Uint8Array<ArrayBuffer>], { type: 'image/x-emf' }));
    const opts = { widthPt: 48, heightPt: 36 };
    const drawn = await getCachedBitmapByPath('word/media/partial.emf', 'image/x-emf', fetchImage, opts);
    expect(getIncompleteMetafileReport(drawn)?.unsupported).toEqual(['EMR_POLYTEXTOUTW']);
    await expect(getCachedBitmapByPath('word/media/partial.emf', 'image/x-emf', fetchImage, {
      ...opts,
      incompleteMetafile: 'reject',
    })).rejects.toMatchObject({ code: 'ooxml-incomplete-metafile' });
    dropBitmapCacheByPath(fetchImage);
  });
});

// ── arcs, pies, chords, rounded rectangles and clip records ────────────────

describe('playEmf — elliptical arc records ([MS-EMF] ARC/ARCTO/CHORD/PIE/ANGLEARC/ROUNDRECT)', () => {
  const ARC = 45;
  const CHORD = 46;
  const PIE = 47;
  const ARCTO = 55;
  const ANGLEARC = 41;
  const ROUNDRECT = 44;
  const SETARCDIRECTION = 57;
  const MOVETOEX = 27;
  const LINETO = 54;
  const stock = (id: number) => record(EMR.SELECTOBJECT, (w) => w.u32(0x80000000 + id));
  const radial = (type: number, box: number[], start: number[], end: number[]) =>
    record(type, (w) => {
      for (const v of [...box, ...start, ...end]) w.i32(v);
    });
  function run(records: Uint8Array[], onUnsupported?: (r: readonly string[]) => void) {
    const m = makeRecordingCtx();
    const drew = playEmf(concat(emfHeader(), ...records, record(EMR.EOF, () => {})), m.ctx, 100, 100, {
      onUnsupported: onUnsupported ?? (() => {}),
    });
    return { ...m, drew };
  }
  const ops = (m: MockCtx) => m.calls.map((c) => c.op);
  const last = (m: MockCtx, op: string) => m.calls.filter((c) => c.op === op).at(-1)?.args;
  const close = (actual: (number | string)[] | undefined, expected: number[]) => {
    expect(actual).toHaveLength(expected.length);
    expected.forEach((v, i) => expect(actual?.[i] as number).toBeCloseTo(v, 6));
  };

  it('PIE draws centre → start radial → counterclockwise arc → closed, filled and stroked', () => {
    // Start radial points right, end radial up: counterclockwise (the default)
    // is the short quarter from (100,50) up to (50,0) on a y-down surface.
    const m = run([stock(4), stock(7), radial(PIE, [0, 0, 100, 100], [100, 50], [50, 0])]);
    expect(ops(m)).toEqual(['save', 'beginPath', 'moveTo', 'lineTo', 'bezierCurveTo', 'closePath', 'fill', 'stroke', 'restore']);
    close(last(m, 'moveTo'), [50, 50]);
    close(last(m, 'lineTo'), [100, 50]);
    const k = (4 / 3) * Math.tan(Math.PI / 8) * 50;
    close(last(m, 'bezierCurveTo'), [100, 50 - k, 50 + k, 0, 50, 0]);
    expect(m.styles.fill).toEqual(['#000000']);
    expect(m.drew).toBe(true);
  });

  it('SETARCDIRECTION AD_CLOCKWISE takes the long way round, and SAVEDC scopes it', () => {
    const m = run([
      stock(4),
      record(EMR.SAVEDC, () => {}),
      record(SETARCDIRECTION, (w) => w.u32(2)),
      radial(PIE, [0, 0, 100, 100], [100, 50], [50, 0]),
      record(EMR.RESTOREDC, (w) => w.i32(-1)),
      radial(PIE, [0, 0, 100, 100], [100, 50], [50, 0]),
    ]);
    const curves = m.calls.filter((c) => c.op === 'bezierCurveTo');
    expect(curves).toHaveLength(4); // three quarters clockwise, then one counterclockwise
    close(curves[2].args.slice(4), [50, 0]);
    close(curves[3].args.slice(4), [50, 0]);
  });

  it('ARC strokes only and leaves the current position; ARCTO lines from it and moves it', () => {
    const m = run([
      stock(7),
      record(MOVETOEX, (w) => w.i32(10).i32(90)),
      radial(ARC, [0, 0, 100, 100], [100, 50], [50, 0]),
      record(LINETO, (w) => w.i32(20).i32(90)),
      radial(ARCTO, [0, 0, 100, 100], [100, 50], [50, 0]),
      record(LINETO, (w) => w.i32(0).i32(0)),
    ]);
    expect(m.styles.fill).toEqual([]);
    const moves = m.calls.filter((c) => c.op === 'moveTo').map((c) => c.args);
    // ARC starts its own figure; LINETO still starts at (10,90); ARCTO starts at
    // the current position (20,90) and the final LINETO at the arc end (50,0).
    expect(moves).toEqual([[100, 50], [10, 90], [20, 90], [50, 0]]);
  });

  it('CHORD closes the arc with a straight chord and fills it', () => {
    const m = run([stock(4), radial(CHORD, [0, 0, 100, 100], [100, 50], [50, 0])]);
    expect(ops(m)).toEqual(['save', 'beginPath', 'moveTo', 'bezierCurveTo', 'closePath', 'fill', 'restore']);
  });

  it('equal radials draw the complete ellipse', () => {
    const m = run([stock(7), radial(ARC, [0, 0, 100, 50], [100, 25], [100, 25])]);
    const curves = m.calls.filter((c) => c.op === 'bezierCurveTo');
    expect(curves).toHaveLength(4);
    close(curves[3].args.slice(4), [100, 25]);
  });

  it('maps arc control points through the world transform (rotated, not axis-aligned)', () => {
    // 90° rotation about the origin, then translate back on-canvas.
    const m = run([
      stock(4),
      record(EMR.SETWORLDTRANSFORM, (w) => w.f32(0).f32(1).f32(-1).f32(0).f32(100).f32(0)),
      radial(PIE, [0, 0, 100, 100], [100, 50], [50, 0]),
    ]);
    // Logical (100,50) → page (50,100); logical (50,0) → page (100,50).
    close(last(m, 'lineTo'), [50, 100]);
    close(last(m, 'bezierCurveTo')?.slice(4), [100, 50]);
  });

  it('ANGLEARC lines from the current position and sweeps counterclockwise from the x-axis', () => {
    const m = run([
      stock(7),
      record(MOVETOEX, (w) => w.i32(0).i32(50)),
      record(ANGLEARC, (w) => w.i32(50).i32(50).u32(50).f32(0).f32(90)),
      record(LINETO, (w) => w.i32(0).i32(0)),
    ]);
    close(m.calls.find((c) => c.op === 'lineTo')?.args, [100, 50]);
    close(m.calls.find((c) => c.op === 'bezierCurveTo')?.args.slice(4), [50, 0]);
    // The current position moved to the arc end.
    close(m.calls.filter((c) => c.op === 'moveTo')[1].args, [50, 0]);
  });

  it('ROUNDRECT fills a closed outline with four elliptical corners', () => {
    const m = run([stock(4), record(ROUNDRECT, (w) => w.i32(0).i32(0).i32(100).i32(60).i32(20).i32(10))]);
    expect(m.calls.filter((c) => c.op === 'bezierCurveTo')).toHaveLength(4);
    expect(m.calls.filter((c) => c.op === 'lineTo')).toHaveLength(3);
    expect(m.styles.fill).toEqual(['#000000']);
  });

  it('builds pies and ellipses inside path brackets for a later FILLPATH', () => {
    const m = run([
      stock(4),
      record(59, () => {}),
      radial(PIE, [0, 0, 100, 100], [100, 50], [50, 0]),
      record(42, (w) => w.i32(0).i32(0).i32(20).i32(20)),
      record(60, () => {}),
      record(62, (w) => w.i32(0).i32(0).i32(100).i32(100)),
    ]);
    expect(m.calls.filter((c) => c.op === 'bezierCurveTo')).toHaveLength(5);
    expect(m.styles.fill).toEqual(['#000000']);
  });

  it('reports instead of guessing the direction when the mapping reflects an axis', () => {
    const reported: string[] = [];
    const m = run(
      [
        stock(4),
        record(EMR.SETWORLDTRANSFORM, (w) => w.f32(1).f32(0).f32(0).f32(-1).f32(0).f32(100)),
        radial(PIE, [0, 0, 100, 100], [100, 50], [50, 0]),
      ],
      (records) => reported.push(...records),
    );
    expect(m.calls.filter((c) => c.op === 'fill')).toHaveLength(0);
    expect(reported).toEqual(['EMR_PIE (reflected mapping)']);
  });
});

describe('playEmf — clip rectangles and regions', () => {
  const INTERSECTCLIPRECT = 30;
  const EXCLUDECLIPRECT = 29;
  const EXTSELECTCLIPRGN = 75;
  const rect = (type: number, l: number, t: number, r: number, b: number) =>
    record(type, (w) => w.i32(l).i32(t).i32(r).i32(b));
  function run(records: Uint8Array[], onUnsupported: (r: readonly string[]) => void = () => {}) {
    const m = makeRecordingCtx();
    playEmf(concat(emfHeader(), ...records, record(EMR.EOF, () => {})), m.ctx, 100, 100, { onUnsupported });
    return m;
  }

  it('INTERSECTCLIPRECT clips to the rectangle; EXCLUDECLIPRECT clips to its complement', () => {
    const m = run([rect(INTERSECTCLIPRECT, 10, 10, 50, 50), rect(EXCLUDECLIPRECT, 20, 20, 30, 30)]);
    expect(m.calls.filter((c) => c.op === 'clip').map((c) => c.args)).toEqual([['nonzero'], ['evenodd']]);
  });

  it('scopes clips to SAVEDC/RESTOREDC and balances every canvas save', () => {
    const m = run([record(EMR.SAVEDC, () => {}), rect(INTERSECTCLIPRECT, 10, 10, 50, 50), record(EMR.RESTOREDC, (w) => w.i32(-1))]);
    const saves = m.calls.filter((c) => c.op === 'save').length;
    expect(saves).toBe(2);
    expect(m.calls.filter((c) => c.op === 'restore').length).toBe(saves);
    // An unmatched SAVEDC is also unwound at the end of playback.
    const open = run([record(EMR.SAVEDC, () => {}), rect(INTERSECTCLIPRECT, 10, 10, 50, 50)]);
    expect(open.calls.filter((c) => c.op === 'restore').length).toBe(2);
  });

  it('EXTSELECTCLIPRGN RGN_COPY with no region resets a clip set at the same level', () => {
    const m = run([rect(INTERSECTCLIPRECT, 10, 10, 50, 50), record(EXTSELECTCLIPRGN, (w) => w.u32(0).u32(5))]);
    expect(m.calls.map((c) => c.op).slice(-4)).toEqual(['clip', 'restore', 'save', 'restore']);
  });

  it('reports a reset that would have to drop a clip inherited from an outer level', () => {
    const reported: string[] = [];
    run(
      [rect(INTERSECTCLIPRECT, 10, 10, 50, 50), record(EMR.SAVEDC, () => {}), record(EXTSELECTCLIPRGN, (w) => w.u32(0).u32(5))],
      (r) => reported.push(...r),
    );
    expect(reported).toEqual(['EMR_EXTSELECTCLIPRGN (reset of an inherited clip)']);
  });

  it('EXTSELECTCLIPRGN intersects device-unit region rectangles, and subtracts them for RGN_DIFF', () => {
    const region = (mode: number, rects: number[][]) =>
      record(EXTSELECTCLIPRGN, (w) => {
        w.u32(32 + rects.length * 16).u32(mode);
        w.u32(32).u32(1).u32(rects.length).u32(rects.length * 16).i32(0).i32(0).i32(100).i32(100);
        for (const r of rects) for (const v of r) w.i32(v);
      });
    const m = run([region(1, [[0, 0, 10, 10], [20, 0, 30, 10]]), region(4, [[0, 0, 5, 5], [2, 2, 8, 8]])]);
    expect(m.calls.filter((c) => c.op === 'clip').map((c) => c.args)).toEqual([['nonzero'], ['evenodd'], ['evenodd']]);
    const reported: string[] = [];
    run([region(2, [[0, 0, 10, 10]])], (r) => reported.push(...r));
    expect(reported).toEqual(['EMR_EXTSELECTCLIPRGN (mode 2)']);
  });
});

describe('playEmf — clip records inside an open path bracket ([MS-EMF] 2.3.2, 2.3.10)', () => {
  const INTERSECTCLIPRECT = 30;
  const EXTSELECTCLIPRGN = 75;
  afterEach(() => vi.unstubAllGlobals());
  const bracket = (clip: Uint8Array) => [
    record(EMR.SELECTOBJECT, (w) => w.u32(0x80000004)),
    record(EMR.BEGINPATH, () => {}),
    record(EMR.POLYGON16, (w) => w.i32(0).i32(0).i32(10).i32(10).u32(3).i16(0).i16(0).i16(40).i16(0).i16(0).i16(40)),
    clip,
    record(EMR.ENDPATH, () => {}),
    record(EMR.FILLPATH, (w) => w.i32(0).i32(0).i32(100).i32(100)),
  ];
  function run(records: Uint8Array[]) {
    const m = makeRecordingCtx();
    const reported: string[] = [];
    playEmf(concat(emfHeader(), ...records, record(EMR.EOF, () => {})), m.ctx, 100, 100, {
      onUnsupported: (r) => reported.push(...r),
    });
    return { ...m, reported };
  }
  /** The ops from the first clip on: the clip region, then the DC path traced
   *  for FILLPATH. */
  const fromClip = (m: ReturnType<typeof run>) => {
    const ops = m.calls.map((c) => c.op);
    return ops.slice(ops.indexOf('clip') - 5, ops.indexOf('fill') + 1);
  };

  it('clips at once and still fills the bracket figures under that clip', () => {
    const m = run(bracket(record(INTERSECTCLIPRECT, (w) => w.i32(10).i32(10).i32(50).i32(50))));
    expect(fromClip(m)).toEqual([
      'moveTo', 'lineTo', 'lineTo', 'lineTo', 'closePath', 'clip', // the clip rectangle
      'beginPath', 'moveTo', 'lineTo', 'lineTo', 'closePath', 'fill', // the bracket's polygon
    ]);
    expect(m.calls.find((c) => c.op === 'clip')?.args).toEqual(['nonzero']);
    expect(m.styles.fill).toEqual(['#000000']);
    expect(m.reported).toEqual([]);
  });

  it('applies region data (EXTSELECTCLIPRGN) inside a bracket the same way', () => {
    const m = run(bracket(record(EXTSELECTCLIPRGN, (w) => {
      w.u32(48).u32(1);
      w.u32(32).u32(1).u32(1).u32(16).i32(0).i32(0).i32(100).i32(100);
      w.i32(0).i32(0).i32(20).i32(20);
    })));
    expect(fromClip(m).slice(-6)).toEqual(['beginPath', 'moveTo', 'lineTo', 'lineTo', 'closePath', 'fill']);
    expect(m.styles.fill).toEqual(['#000000']);
  });

  it('keeps a closed path across a clip record until FILLPATH consumes it', () => {
    // BEGINPATH, RECTANGLE, ENDPATH, then INTERSECTCLIPRECT before FILLPATH.
    const m = run([
      record(EMR.SELECTOBJECT, (w) => w.u32(0x80000004)),
      record(EMR.BEGINPATH, () => {}),
      record(43, (w) => w.i32(10).i32(10).i32(20).i32(20)),
      record(EMR.ENDPATH, () => {}),
      record(INTERSECTCLIPRECT, (w) => w.i32(0).i32(0).i32(90).i32(90)),
      record(EMR.FILLPATH, (w) => w.i32(0).i32(0).i32(100).i32(100)),
    ]);
    const moves = m.calls.filter((c) => c.op === 'moveTo').map((c) => c.args);
    expect(moves).toEqual([[0, 0], [10, 10]]); // clip region, then the held rectangle
    expect(fromClip(m).slice(-7)).toEqual(['beginPath', 'moveTo', 'lineTo', 'lineTo', 'lineTo', 'closePath', 'fill']);
    expect(m.reported).toEqual([]);
  });
});

describe('playEmf — the device context path ([MS-EMF] 2.3.10, 2.3.11)', () => {
  const RECTANGLE = 43;
  const ABORTPATH = 68;
  const rect = (l: number, t: number, r: number, b: number) => record(RECTANGLE, (w) => w.i32(l).i32(t).i32(r).i32(b));
  const fillPath = record(EMR.FILLPATH, (w) => w.i32(0).i32(0).i32(100).i32(100));
  const saveDc = record(EMR.SAVEDC, () => {});
  const restoreDc = record(EMR.RESTOREDC, (w) => w.i32(-1));
  /** BEGINPATH, rectangle A (10..20), ENDPATH: A is held in the DC. */
  const holdA = [record(EMR.BEGINPATH, () => {}), rect(10, 10, 20, 20), record(EMR.ENDPATH, () => {})];
  const A = [['M', 10, 10], ['L', 20, 10], ['L', 20, 20], ['L', 10, 20], ['Z']];
  const B = [['M', 50, 50], ['L', 60, 50], ['L', 60, 60], ['L', 50, 60], ['Z']];

  /** Each fill's geometry, and each clip's. `maxPathCommands` lowers the
   *  path budget; `reported` then collects what playback left out. */
  function play(records: Uint8Array[], maxPathCommands?: number) {
    const m = makeRecordingCtx();
    let current: unknown[] = [];
    const fills: unknown[][] = [];
    const clips: unknown[][] = [];
    m.ctx.beginPath = () => { current = []; };
    m.ctx.moveTo = (...a) => { current.push(['M', ...a]); };
    m.ctx.lineTo = (...a) => { current.push(['L', ...a]); };
    m.ctx.closePath = () => { current.push(['Z']); };
    m.ctx.fill = () => { fills.push([...current]); };
    m.ctx.clip = () => { clips.push([...current]); };
    const unsupported: string[] = [];
    playEmf(concat(
      emfHeader(),
      record(EMR.SELECTOBJECT, (w) => w.u32(0x80000004)), // BLACK_BRUSH
      record(EMR.SELECTOBJECT, (w) => w.u32(0x80000007)), // BLACK_PEN
      ...records,
      record(EMR.EOF, () => {}),
    ), m.ctx, 100, 100, {
      onUnsupported: (r) => unsupported.push(...r),
      ...(maxPathCommands ? { maxPathCommands } : {}),
    } as Parameters<typeof playEmf>[4]);
    if (!maxPathCommands) expect(unsupported).toEqual([]);
    return { fills: fills.filter((p) => p.length), clips, reported: unsupported };
  }

  it('keeps the held path across ordinary drawing outside the bracket', () => {
    // RECTANGLE and LINETO after ENDPATH paint at once and leave A held.
    expect(play([...holdA, rect(50, 50, 60, 60), fillPath]).fills).toEqual([B, A]);
    expect(play([...holdA, record(EMR.LINETO, (w) => w.i32(90).i32(90)), fillPath]).fills).toEqual([A]);
  });

  it('consumes the path once, and paints nothing after ABORTPATH', () => {
    expect(play([...holdA, fillPath, fillPath]).fills).toEqual([A]);
    expect(play([...holdA, record(ABORTPATH, () => {}), rect(50, 50, 60, 60), fillPath]).fills).toEqual([B]);
  });

  it('fails FILLPATH and SELECTCLIPPATH while the bracket is open, leaving it open', () => {
    const r = play([
      record(EMR.BEGINPATH, () => {}), rect(10, 10, 20, 20),
      fillPath, record(EMR.SELECTCLIPPATH, (w) => w.u32(1)),
      rect(10, 10, 20, 20), record(EMR.ENDPATH, () => {}), fillPath,
    ]);
    expect(r.clips).toEqual([]);
    expect(r.fills).toEqual([[...A, ...A]]);
    // With no path at all SELECTCLIPPATH changes no clip either.
    expect(play([rect(50, 50, 60, 60), record(EMR.SELECTCLIPPATH, (w) => w.u32(1))]).clips).toEqual([]);
  });

  it('saves and restores the path with the DC', () => {
    // A held, saved twice, aborted, restored twice: A is back and fills.
    expect(play([...holdA, saveDc, saveDc, record(ABORTPATH, () => {}), restoreDc, restoreDc, fillPath]).fills).toEqual([A]);
    // A bracket saved while open is open again after RESTOREDC (Wine gdi32
    // path tests), so the figures drawn before ENDPATH join the path.
    expect(play([
      record(EMR.BEGINPATH, () => {}), rect(10, 10, 20, 20), saveDc,
      record(EMR.ENDPATH, () => {}), fillPath, restoreDc,
      fillPath, rect(50, 50, 60, 60), record(EMR.ENDPATH, () => {}), fillPath,
    ]).fills).toEqual([A, [...A, ...B]]);
    // A path begun after SAVEDC goes away with RESTOREDC.
    expect(play([saveDc, ...holdA, restoreDc, fillPath]).fills).toEqual([]);
  });

  // A RECTANGLE is 5 path commands; the budget covers the current path plus
  // every prefix a SAVEDC snapshot keeps, counting a shared buffer once.
  const hold = (l: number) => [record(EMR.BEGINPATH, () => {}), rect(l, l, l + 10, l + 10), record(EMR.ENDPATH, () => {})];
  const BUDGET = ['EMF path (path command budget)'];

  it('bounds the paths kept by repeated SAVEDC and reports the one past the budget', () => {
    // Two saved 5-command paths retain 10 of 12; a third bracket would hold 15.
    const r = play([
      ...hold(10), saveDc, ...hold(50), saveDc, ...hold(70),
      fillPath, restoreDc, fillPath, restoreDc, fillPath,
    ], 12);
    expect(r.reported).toEqual(BUDGET);
    expect(r.fills).toEqual([B, A]); // the over-budget bracket paints nothing
    // Snapshots of one buffer are one retained copy: A saved three times
    // still leaves room for a second 5-command path.
    const shared = play([...holdA, saveDc, saveDc, saveDc, ...hold(50), fillPath, restoreDc, fillPath], 12);
    expect(shared.reported).toEqual([]);
    expect(shared.fills).toEqual([B, A]);
  });

  it('keeps only the saved prefix, so an empty snapshot retains nothing', () => {
    // Each round saves an empty path, then appends 10 commands and aborts.
    // Those tails are unreachable; were they kept, the last path (10 of 12)
    // would exceed the budget.
    const round = [record(EMR.BEGINPATH, () => {}), saveDc, rect(0, 0, 5, 5), rect(0, 0, 6, 6), record(ABORTPATH, () => {})];
    const r = play([
      ...round, ...round, ...round,
      record(EMR.BEGINPATH, () => {}), rect(10, 10, 20, 20), rect(50, 50, 60, 60), record(EMR.ENDPATH, () => {}), fillPath,
    ], 12);
    expect(r.reported).toEqual([]);
    expect(r.fills).toEqual([[...A, ...B]]);
    // A snapshot of an open bracket keeps its prefix: restoring it drops the
    // commands appended after SAVEDC, and they no longer count.
    const prefix = play([
      record(EMR.BEGINPATH, () => {}), rect(10, 10, 20, 20), saveDc, rect(50, 50, 60, 60), restoreDc,
      rect(50, 50, 60, 60), record(EMR.ENDPATH, () => {}), fillPath,
    ], 10);
    expect(prefix.reported).toEqual([]);
    expect(prefix.fills).toEqual([[...A, ...B]]);
  });
});

describe('playEmf — explicit report of records it cannot draw', () => {
  function run(records: Uint8Array[]) {
    const m = makeRecordingCtx();
    const reported: string[] = [];
    playEmf(concat(emfHeader(), ...records, record(EMR.EOF, () => {})), m.ctx, 100, 100, {
      onUnsupported: (r) => reported.push(...r),
    });
    return { ...m, reported };
  }

  it('reports each unsupported drawing record once per playback, and nothing for state records', () => {
    const polyText = record(97, (w) => w.u32(0));
    const m = run([polyText, polyText, record(21 /* SETROP2-class state */, (w) => w.u32(13))]);
    expect(m.reported).toEqual(['EMR_POLYTEXTOUTW']);
  });

  it('has no default report: without a callback nothing reaches the console', () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    try {
      const file = concat(emfHeader(), record(118, (w) => w.u32(0)), record(EMR.EOF, () => {}));
      playEmf(file, makeRecordingCtx().ctx, 10, 10);
      expect(warn).not.toHaveBeenCalled();
    } finally {
      warn.mockRestore();
    }
  });

  it('keeps the other figures of a path bracket around an unsupported record, as before it was reported', () => {
    const polygon = record(EMR.POLYGON16, (w) => w.i32(0).i32(0).i32(10).i32(10).u32(3).i16(0).i16(0).i16(40).i16(0).i16(0).i16(40));
    const m = run([
      record(EMR.SELECTOBJECT, (w) => w.u32(0x80000004)),
      record(EMR.BEGINPATH, () => {}),
      polygon,
      record(97, (w) => w.u32(0)), // POLYTEXTOUTW
      record(EMR.ENDPATH, () => {}),
      record(66 /* WIDENPATH */, () => {}),
      record(EMR.FILLPATH, (w) => w.i32(0).i32(0).i32(100).i32(100)),
    ]);
    expect(m.styles.fill).toEqual(['#000000']);
    expect(m.calls.filter((c) => c.op === 'lineTo')).toHaveLength(2);
    expect(m.reported).toEqual(['EMR_POLYTEXTOUTW', 'EMR_WIDENPATH']);
  });

  it('draws EXTTEXTOUTW inside a path bracket directly and reports the missing glyph path', () => {
    const font = record(EMR.EXTCREATEFONTINDIRECTW, (w) => {
      w.u32(1).i32(-12).i32(0).i32(0).i32(0).i32(400).raw(0, 0, 0, 0).raw(0, 0, 0, 0);
      w.utf16('Arial');
      for (let i = 5; i < 32; i++) w.u16(0);
    });
    const text = record(EMR.EXTTEXTOUTW, (w) => {
      w.i32(0).i32(0).i32(100).i32(100).u32(1).f32(1).f32(1); // rclBounds, iGraphicsMode, exScale, eyScale
      w.i32(10).i32(20).u32(1).u32(76).u32(0).i32(0).i32(0).i32(0).i32(0).u32(0); // EMRTEXT
      w.utf16('A');
    });
    const m = run([font, record(EMR.SELECTOBJECT, (w) => w.u32(1)), record(EMR.BEGINPATH, () => {}), text, record(EMR.ENDPATH, () => {})]);
    expect(m.styles.text).toHaveLength(1);
    expect(m.reported).toEqual(['EMR_EXTTEXTOUTW (glyph path)']);
  });

  it('reports a truncated record stream and a malformed record, keeping what drew', () => {
    const polygon = record(EMR.POLYGON16, (w) => w.i32(0).i32(0).i32(10).i32(10).u32(3).i16(0).i16(0).i16(40).i16(0).i16(0).i16(40));
    const truncated = concat(emfHeader(), record(EMR.SELECTOBJECT, (w) => w.u32(0x80000004)), polygon, new Uint8Array([86, 0, 0, 0, 64, 0, 0, 0]));
    const m = makeRecordingCtx();
    const reported: string[] = [];
    expect(playEmf(truncated, m.ctx, 100, 100, { onUnsupported: (r) => reported.push(...r) })).toBe(true);
    expect(reported).toEqual(['EMF record stream (truncated or invalid record size)']);
    // SETWORLDTRANSFORM without its XFORM is malformed.
    expect(run([record(EMR.SETWORLDTRANSFORM, (w) => w.u32(0))]).reported).toEqual(['EMF record 35 (malformed)']);
  });

  it('paints PATCOPY/BLACKNESS pattern blits and reports other brush-only raster operations', () => {
    const bitblt = (rop: number) =>
      record(76, (w) => {
        w.i32(0).i32(0).i32(10).i32(10); // rclBounds
        w.i32(10).i32(20).i32(30).i32(40).u32(rop).i32(0).i32(0);
        for (let i = 0; i < 6; i++) w.f32(i === 0 || i === 3 ? 1 : 0);
        w.u32(0).u32(0).u32(0).u32(0).u32(0).u32(0);
      });
    const m = run([record(EMR.SELECTOBJECT, (w) => w.u32(0x80000000 + 2)), bitblt(0x00f00021), bitblt(0x00000042), bitblt(0x005a0049)]);
    expect(m.styles.fill).toEqual(['#808080', '#000000']);
    expect(m.reported).toEqual(['EMR_BITBLT (raster operation 0x5a0049)']);
  });

});

// ── EMF+ ([MS-EMFPLUS]) bitmap playback ─────────────────────────────────────

describe('playEmf — EMF+ bitmap records', () => {
  /** One EMF+ record: Type, Flags, Size, DataSize, data (4-byte aligned). */
  const plusRecord = (type: number, flags: number, data: number[] = []) => {
    const body = [...data];
    while (body.length % 4) body.push(0);
    const w = new Writer().u16(type).u16(flags).u32(12 + body.length).u32(data.length);
    for (const byte of body) w.raw(byte);
    return w.build();
  };
  const f32 = (v: number) => [...new Uint8Array(new Float32Array([v]).buffer)];
  const u32 = (v: number) => [v & 255, (v >>> 8) & 255, (v >>> 16) & 255, (v >>> 24) & 255];
  /** EMR_COMMENT carrying EMF+ records. */
  const comment = (...records: Uint8Array[]) => {
    const payload = concat(...records);
    return record(70, (w) => {
      w.u32(4 + payload.length).u32(0x2b464d45);
      for (const byte of payload) w.raw(byte);
    });
  };
  const header = (dual: boolean) => plusRecord(0x4001, dual ? 1 : 0, [...u32(0xdbc01002), ...u32(1), ...u32(96), ...u32(96)]);
  // A 2×1 premultiplied-ARGB bitmap: opaque red, half-transparent white.
  const bitmap = (flags = 0x0501, pixelFormat = 0x000e200b, bitmapType = 0) =>
    plusRecord(0x4008, flags, [
      ...u32(0xdbc01002), ...u32(1), ...u32(2), ...u32(1), ...u32(8), ...u32(pixelFormat), ...u32(bitmapType),
      0, 0, 255, 255, 128, 128, 128, 128,
    ]);
  /** EmfPlusImageAttributes object 0: Version, Reserved1, WrapMode (Tile, as
   *  Excel writes), ClampColor, ObjectClamp, Reserved2. */
  const attributes = (wrapMode = 0, objectClamp = 0) =>
    plusRecord(0x4008, 0x0800, [...u32(0xdbc01002), ...u32(0), ...u32(wrapMode), ...u32(0xffffffff), ...u32(objectClamp), ...u32(0)]);
  const drawImage = (dest: number[], src = [0, 0, 2, 1]) =>
    plusRecord(0x401a, 0x0001, [...u32(0), ...u32(2), ...src.flatMap(f32), ...dest.flatMap(f32)]);

  function run(records: Uint8Array[], gdi: Uint8Array[] = []) {
    const draws: { data: number[]; rect: number[] }[] = [];
    vi.stubGlobal('OffscreenCanvas', class {
      width: number;
      height: number;
      constructor(w: number, h: number) {
        this.width = w;
        this.height = h;
      }
      getContext() {
        const owner = this;
        return {
          createImageData: (w: number, h: number) => ({ data: new Uint8ClampedArray(w * h * 4), width: w, height: h }),
          putImageData(img: { data: Uint8ClampedArray }) {
            (owner as unknown as { pixels: number[] }).pixels = [...img.data];
          },
        };
      }
    });
    const m = makeRecordingCtx();
    (m.ctx as unknown as { drawImage: unknown }).drawImage = (img: { pixels: number[] }, ...rect: number[]) => {
      draws.push({ data: img.pixels, rect });
    };
    (m.ctx as unknown as { clearRect: unknown }).clearRect = (...a: number[]) => m.calls.push({ op: 'clearRect', args: a });
    (m.ctx as unknown as { fillRect: unknown }).fillRect = (...a: number[]) => m.calls.push({ op: 'fillRect', args: a });
    const reported: string[] = [];
    const drew = playEmf(concat(emfHeader(0, 0, 100, 100), ...gdi, ...records, record(EMR.EOF, () => {})), m.ctx, 100, 100, {
      onUnsupported: (r) => reported.push(...r),
    });
    vi.unstubAllGlobals();
    return { ...m, draws, reported, drew };
  }

  it('draws an uncompressed 32bpp bitmap object through EmfPlusDrawImage', () => {
    const result = run([
      comment(
        header(true),
        plusRecord(0x4030, 0x0002, f32(1)), // SetPageTransform: UnitPixel, scale 1
        plusRecord(0x4009, 0, u32(0x00ffffff)), // Clear to transparent white
        plusRecord(0x4023, 0), // SourceOver
        attributes(),
        bitmap(),
        drawImage([10, 20, 50, 40]),
      ),
      comment(plusRecord(0x4002, 0)),
    ]);
    expect(result.drew).toBe(true);
    expect(result.reported).toEqual([]);
    expect(result.draws).toHaveLength(1);
    expect(result.draws[0].rect).toEqual([10, 20, 50, 40]);
    // BGRA → RGBA, and premultiplied 128/128 un-premultiplied to white.
    expect(result.draws[0].data).toEqual([255, 0, 0, 255, 255, 255, 255, 128]);
    expect(result.calls.some((c) => c.op === 'clearRect')).toBe(true);
  });

  it('assembles continued objects, applies world transforms and crops the source', () => {
    const whole = bitmap();
    const data = whole.slice(12); // object data after the record header
    const first = plusRecord(0x4008, 0x8501, [...u32(data.length), ...data.slice(0, 10)]);
    const second = plusRecord(0x4008, 0x8501, [...u32(data.length), ...data.slice(10)]);
    const result = run([
      comment(
        header(true),
        attributes(),
        first,
        second,
        plusRecord(0x402d, 0, [...f32(5), ...f32(5)]), // translate
        plusRecord(0x402e, 0x2000, [...f32(2), ...f32(2)]), // then scale (append)
        drawImage([0, 0, 10, 10], [1, 0, 1, 1]),
      ),
    ]);
    expect(result.reported).toEqual([]);
    expect(result.draws[0].rect).toEqual([10, 10, 20, 20]);
    expect(result.draws[0].data).toEqual([255, 255, 255, 128]);
  });

  it('keeps the GDI rendering of a dual file whose EMF+ part is not implemented', () => {
    const gdiPolygon = [
      record(EMR.SELECTOBJECT, (w) => w.u32(0x80000004)),
      record(EMR.POLYGON16, (w) => w.i32(0).i32(0).i32(10).i32(10).u32(3).i16(0).i16(0).i16(10).i16(0).i16(0).i16(10)),
    ];
    // FillRects (0x400A) is not implemented: the GDI alternative is played.
    const dual = run([comment(header(true), plusRecord(0x400a, 0, u32(0)))], gdiPolygon);
    expect(dual.styles.fill).toEqual(['#000000']);
    expect(dual.draws).toHaveLength(0);
    // A dual file whose EMF+ part has no drawing also keeps its GDI drawing.
    expect(run([comment(header(true), plusRecord(0x401e, 0))], gdiPolygon).styles.fill).toEqual(['#000000']);
    // Played EMF+ skips GDI records outside an EmfPlusGetDC scope.
    const plus = run([comment(header(true), attributes(), bitmap(), drawImage([0, 0, 2, 1]))], gdiPolygon);
    expect(plus.styles.fill).toEqual([]);
    expect(plus.draws).toHaveLength(1);
  });

  it('keeps the GDI rendering when implemented EMF+ records would be rejected in playback', () => {
    const gdiPolygon = [
      record(EMR.SELECTOBJECT, (w) => w.u32(0x80000004)),
      record(EMR.POLYGON16, (w) => w.i32(0).i32(0).i32(10).i32(10).u32(3).i16(0).i16(0).i16(10).i16(0).i16(0).i16(10)),
    ];
    const dual = (...records: Uint8Array[]) => run([comment(header(true), ...records)], gdiPolygon);
    const expectGdi = (result: ReturnType<typeof run>) => {
      expect(result.styles.fill).toEqual(['#000000']);
      expect(result.draws).toHaveLength(0);
    };
    // Unsupported page unit (UnitInch).
    expectGdi(dual(plusRecord(0x4030, 0x0004, f32(1)), attributes(), bitmap(), drawImage([0, 0, 2, 1])));
    // Rotating and mirroring world transforms.
    expectGdi(dual(attributes(), bitmap(), plusRecord(0x402f, 0, f32(30)), drawImage([0, 0, 2, 1])));
    expectGdi(dual(attributes(), bitmap(), plusRecord(0x402e, 0, [...f32(-1), ...f32(1)]), drawImage([0, 0, 2, 1])));
    // An empty destination and a source reaching outside the image.
    expectGdi(dual(attributes(), bitmap(), drawImage([0, 0, 0, 1])));
    expectGdi(dual(attributes(), bitmap(), drawImage([0, 0, 2, 1], [0, 0, 3, 1])));
    // Missing or invalid image attributes.
    expectGdi(dual(bitmap(), drawImage([0, 0, 2, 1])));
    expectGdi(dual(attributes(9), bitmap(), drawImage([0, 0, 2, 1])));
    expectGdi(dual(attributes(0, 2), bitmap(), drawImage([0, 0, 2, 1])));
    // An object table entry replaced by another kind is no longer usable.
    expectGdi(dual(attributes(), plusRecord(0x4008, 0x0400, u32(0)), bitmap(), drawImage([0, 0, 2, 1])));
    // Restoring a state that was never saved.
    expectGdi(dual(attributes(), bitmap(), plusRecord(0x4026, 0, u32(7)), drawImage([0, 0, 2, 1])));
  });

  it('saves and restores the compositing mode with the graphics state', () => {
    const played = (...records: Uint8Array[]) => run([comment(header(false), attributes(), bitmap(), ...records)]);
    const clears = (result: ReturnType<typeof run>) => result.calls.filter((c) => c.op === 'clearRect').length;
    // SourceCopy set after Save is undone by Restore: no clearing blit.
    expect(clears(played(
      plusRecord(0x4025, 0, u32(1)),
      plusRecord(0x4023, 1),
      plusRecord(0x4026, 0, u32(1)),
      drawImage([0, 0, 2, 1]),
    ))).toBe(0);
    // SourceCopy saved with the state survives a later SourceOver.
    expect(clears(played(
      plusRecord(0x4023, 1),
      plusRecord(0x4025, 0, u32(2)),
      plusRecord(0x4023, 0),
      plusRecord(0x4026, 0, u32(2)),
      drawImage([0, 0, 2, 1]),
    ))).toBe(1);
  });

  const gdiPolygon = () => [
    record(EMR.SELECTOBJECT, (w) => w.u32(0x80000004)),
    record(EMR.POLYGON16, (w) => w.i32(0).i32(0).i32(10).i32(10).u32(3).i16(0).i16(0).i16(10).i16(0).i16(0).i16(10)),
  ];
  const file = (records: Uint8Array[], gdi: Uint8Array[] = gdiPolygon()) =>
    concat(emfHeader(0, 0, 100, 100), ...gdi, ...records, record(EMR.EOF, () => {}));
  /** An EMR_COMMENT whose EMF+ payload is written byte for byte. */
  const rawComment = (dataSize: number, payload: Uint8Array) => record(70, (w) => {
    w.u32(dataSize).u32(0x2b464d45);
    for (const byte of payload) w.raw(byte);
  });

  it('does not play EMF+ over a complete GDI rendering when a later record is truncated', () => {
    // A valid DrawImage, then 8 bytes: less than one EMF+ record header.
    const payload = concat(header(true), attributes(), bitmap(), drawImage([0, 0, 2, 1]), new Uint8Array([0x02, 0x40, 0, 0, 12, 0, 0, 0]));
    const records = [rawComment(4 + payload.length, payload)];
    expect(scanEmfPlus(file(records))).toEqual({
      present: true, dual: true, play: false, failures: ['EMF+ record (truncated header)'],
    });
    const played = run(records, gdiPolygon());
    expect(played.draws).toHaveLength(0);
    expect(played.styles.fill).toEqual(['#000000']);
    // The dual file's GDI rendering is complete: nothing is reported.
    expect(played.reported).toEqual([]);
  });

  it('treats invalid record and comment sizes as validation failures', () => {
    const valid = concat(header(true), attributes(), bitmap(), drawImage([0, 0, 2, 1]));
    // A record whose Size runs past its comment.
    const overrun = concat(valid, new Writer().u16(0x4002).u16(0).u32(64).u32(0).build());
    expect(scanEmfPlus(file([rawComment(4 + overrun.length, overrun)])).failures)
      .toEqual(['EMF+ record 0x4002 (invalid Size or DataSize)']);
    // DataSize larger than Size − 12.
    const badData = concat(valid, new Writer().u16(0x4002).u16(0).u32(12).u32(4).build());
    expect(scanEmfPlus(file([rawComment(4 + badData.length, badData)])).failures)
      .toEqual(['EMF+ record 0x4002 (invalid Size or DataSize)']);
    // An EMR_COMMENT DataSize that leaves the comment record.
    expect(scanEmfPlus(file([rawComment(4 + valid.length + 16, valid)]))).toMatchObject({
      play: false, failures: ['EMF+ comment (invalid DataSize)'],
    });
  });

  it('treats an unfinished or interleaved continued object as a validation failure', () => {
    const data = bitmap().slice(12);
    const first = plusRecord(0x4008, 0x8501, [...u32(data.length), ...data.slice(0, 10)]);
    const second = plusRecord(0x4008, 0x8501, [...u32(data.length), ...data.slice(10)]);
    const unfinished = scanEmfPlus(file([comment(header(true), attributes(), first, drawImage([0, 0, 2, 1]))]));
    expect(unfinished.play).toBe(false);
    expect(unfinished.failures).toContain('EMF+ continued object (unfinished)');
    const interleaved = scanEmfPlus(file([comment(header(true), attributes(), first, plusRecord(0x401e, 0), second, drawImage([0, 0, 2, 1]))]));
    expect(interleaved.play).toBe(false);
    expect(interleaved.failures).toContain('EMF+ continued object (unfinished)');
    // Consecutive fragments remain a valid object.
    expect(scanEmfPlus(file([comment(header(true), attributes(), first, second, drawImage([0, 0, 2, 1]))])))
      .toEqual({ present: true, dual: true, play: true, failures: [] });
  });

  it('reports the failures of an EMF+-only file, which has no GDI alternative, and plays its GDI records as before', () => {
    const result = run([
      comment(header(false), plusRecord(0x400a, 0, u32(0)), bitmap(0x0501, 0x00022009), drawImage([0, 0, 2, 1])),
    ], gdiPolygon());
    expect(result.draws).toHaveLength(0);
    expect(result.styles.fill).toEqual(['#000000']);
    expect(result.reported).toEqual(expect.arrayContaining([
      'EMF+ record 0x400a',
      'EMF+ image other than an uncompressed 32-bit bitmap',
      'EMF+ DrawImage of an unavailable image',
    ]));
    // A truncated EMF+-only stream is reported the same way.
    const payload = concat(header(false), attributes(), bitmap(), drawImage([0, 0, 2, 1]), new Uint8Array([0, 0, 0, 0]));
    expect(run([rawComment(4 + payload.length, payload)]).reported).toEqual(['EMF+ record (truncated header)']);
  });
  it('never admits a non-finite placement, so the GDI alternative is kept', () => {
    const draw = (...state: Uint8Array[]) => [comment(header(true), attributes(), bitmap(), ...state, drawImage([0, 0, 2, 1]))];
    const cases = [
      draw(plusRecord(0x4030, 0x0002, f32(NaN))), // PageScale NaN
      draw(plusRecord(0x4030, 0x0002, f32(Infinity))), // PageScale infinite
      draw(plusRecord(0x402a, 0, [1, NaN, 0, 1, 0, 0].flatMap(f32))), // world matrix
      draw(plusRecord(0x402d, 0, [...f32(Infinity), ...f32(0)])), // translate
      [comment(header(true), attributes(), bitmap(), drawImage([Infinity, 0, 2, 1]))],
      [comment(header(true), attributes(), bitmap(), drawImage([0, 0, 2, 1], [NaN, 0, 2, 1]))],
    ];
    for (const records of cases) {
      const scan = scanEmfPlus(file(records));
      expect(scan.play).toBe(false);
      expect(scan.failures).toEqual(['EMF+ DrawImage placement (non-finite value)']);
      const played = run(records, gdiPolygon());
      expect(played.draws).toHaveLength(0);
      expect(played.styles.fill).toEqual(['#000000']);
      expect(played.reported).toEqual([]);
    }
  });

  it('applies the shared image budget before a bitmap or a continued object is allocated', () => {
    // 40,000 × 1 exceeds the shared per-axis ceiling (MAX_RASTER_DIMENSION).
    const wide = new Uint8Array(28 + 40000 * 4);
    const v = new DataView(wide.buffer);
    [0xdbc01002, 1, 40000, 1, 160000, 0x0026200a, 0].forEach((x, i) => v.setUint32(i * 4, x, true));
    const oversize = [comment(header(true), attributes(), plusRecord(0x4008, 0x0501, [...wide]), drawImage([0, 0, 2, 1]))];
    expect(scanEmfPlus(file(oversize))).toMatchObject({ play: false, failures: expect.arrayContaining(['EMF+ bitmap (decoded-image budget)']) });
    const played = run(oversize, gdiPolygon());
    expect(played.draws).toHaveLength(0);
    expect(played.styles.fill).toEqual(['#000000']);
    // A continued object declaring more than the player's decoded-byte
    // ceiling is refused at its first fragment, before any assembly buffer.
    const huge = plusRecord(0x4008, 0x8501, [...u32(0xffffff00), 1, 2, 3, 4]);
    expect(scanEmfPlus(file([comment(header(true), attributes(), huge, drawImage([0, 0, 2, 1]))])).failures)
      .toContain('EMF+ bitmap (decoded-image budget)');
  });

  it('reports a bitmap blit that fails at run time instead of dropping it silently', () => {
    const noHelper = () => vi.stubGlobal('OffscreenCanvas', class { getContext() { return null; } });
    const records = [comment(header(true), attributes(), bitmap(), drawImage([0, 0, 2, 1]))];
    // Dual file without GDI drawing (as Excel writes): nothing replaces the image.
    noHelper();
    const m = makeRecordingCtx();
    const reported: string[] = [];
    const drew = playEmf(file(records, []), m.ctx, 100, 100, { onUnsupported: (r) => reported.push(...r) });
    vi.unstubAllGlobals();
    expect(drew).toBe(false);
    expect(reported).toEqual(['EMF+ DrawImage (the bitmap could not be drawn)']);
    // Dual file with a GDI rendering: the partial EMF+ drawing is cleared and
    // the complete GDI alternative plays; the EMF+ failure is still reported.
    noHelper();
    const g = makeRecordingCtx();
    const clears: number[][] = [];
    (g.ctx as unknown as { clearRect: unknown }).clearRect = (...a: number[]) => clears.push(a);
    const gdiReported: string[] = [];
    expect(playEmf(file(records), g.ctx, 100, 100, { onUnsupported: (r) => gdiReported.push(...r) })).toBe(true);
    vi.unstubAllGlobals();
    expect(clears).toContainEqual([0, 0, 100, 100]);
    expect(g.styles.fill).toEqual(['#000000']);
    expect(gdiReported).toEqual(['EMF+ DrawImage (the bitmap could not be drawn)']);
    // EMF+-only file: no alternative, the failure is reported.
    noHelper();
    const only: string[] = [];
    playEmf(file([comment(header(false), attributes(), bitmap(), drawImage([0, 0, 2, 1]))], []), makeRecordingCtx().ctx, 100, 100, {
      onUnsupported: (r) => only.push(...r),
    });
    vi.unstubAllGlobals();
    expect(only).toEqual(['EMF+ DrawImage (the bitmap could not be drawn)']);
  });

  it('rejects a strict decode whose EMF+ blit fails instead of returning null', async () => {
    // A dual file: even its GDI rendering does not stand in silently.
    const bytes = file([comment(header(true), attributes(), bitmap(), drawImage([0, 0, 2, 1]))]);
    vi.stubGlobal('OffscreenCanvas', class {
      constructor(public width: number, public height: number) {}
      // The 2×1 helper surface of the blit is refused; the target is not.
      getContext() { return this.width === 2 ? null : makeRecordingCtx().ctx; }
    });
    vi.stubGlobal('createImageBitmap', async (src: { width: number; height: number }) => ({ width: src.width, height: src.height, close() {} }));
    try {
      const outcome = await decodeRasterOrMetafile(new Blob([bytes as Uint8Array<ArrayBuffer>]), {
        widthPt: 48, heightPt: 36, incompleteMetafile: 'reject',
      }).then(() => undefined, (error: unknown) => error);
      expect(isOoxmlIncompleteMetafileError(outcome)).toBe(true);
      expect(outcome).toMatchObject({ unsupported: ['EMF+ DrawImage (the bitmap could not be drawn)'] });
    } finally {
      vi.unstubAllGlobals();
    }
  });
});
