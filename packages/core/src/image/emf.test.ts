import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { isEmf } from './wmf.js';
import { playEmf, renderEmfToBitmap } from './emf.js';

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
    clip(rule?: string) {
      calls.push({ op: 'clip', args: rule ? [rule] : [] });
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
    const bmp = await renderEmfToBitmap(file, 64, 48);
    expect(bmp).not.toBeNull();
    expect(bmp?.width).toBe(64);
    expect(bmp?.height).toBe(48);
  });

  it('returns null for non-EMF bytes', async () => {
    const bmp = await renderEmfToBitmap(new Uint8Array([1, 2, 3, 4]), 10, 10);
    expect(bmp).toBeNull();
  });

  it('returns null when nothing draws (header + EOF only)', async () => {
    const file = concat(emfHeader(0, 0, 10, 10), record(EMR.EOF, () => {}));
    const bmp = await renderEmfToBitmap(file, 16, 16);
    expect(bmp).toBeNull();
  });

  it('returns null for a non-positive target size', async () => {
    const file = concat(emfHeader(0, 0, 10, 10), record(EMR.EOF, () => {}));
    expect(await renderEmfToBitmap(file, 0, 10)).toBeNull();
    expect(await renderEmfToBitmap(file, 10, 0)).toBeNull();
  });
});
