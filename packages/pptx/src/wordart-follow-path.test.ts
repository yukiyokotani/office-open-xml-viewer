import { describe, it, expect } from 'vitest';
import { renderTextBody } from './renderer.js';
import { buildWarpEnvelope, warpGlyphTransform } from '@silurus/ooxml-core';
import type { TextBody, Paragraph } from './types';
import type { TextRunData } from '@silurus/ooxml-core';

/**
 * WordArt "Follow Path" semantics for single-edge warps (ECMA-376 §20.1.9.19,
 * issue #846). PowerPoint lays text along an arch/circle baseline at its NATURAL
 * width — the word follows the arc for only its own ink length from the start
 * (stAng), it is NOT scattered around the whole ellipse. Paired-edge presets
 * (waves, inflate/deflate, …) DO stretch the flat ink box to fill the envelope.
 *
 * These tests drive `renderTextBody` against a mock 2D context that tracks the
 * current transform matrix (CTM) through save/restore/translate/rotate/scale, so
 * the FINAL device-space position of each warped glyph is recoverable. The
 * horizontal ink SPAN and centre of the placed glyphs are the observables that
 * matter: the natural-width arc segment must also honour paragraph alignment.
 */

// 2×3 affine matrix [a,b,c,d,e,f] mapping (x,y) → (a·x+c·y+e, b·x+d·y+f).
type M = [number, number, number, number, number, number];
const I: M = [1, 0, 0, 1, 0, 0];
function mul(m: M, n: M): M {
  // m ∘ n  (apply n first, then m) — matches ctx.transform composition.
  return [
    m[0] * n[0] + m[2] * n[1],
    m[1] * n[0] + m[3] * n[1],
    m[0] * n[2] + m[2] * n[3],
    m[1] * n[2] + m[3] * n[3],
    m[0] * n[4] + m[2] * n[5] + m[4],
    m[1] * n[4] + m[3] * n[5] + m[5],
  ];
}
function apply(m: M, x: number, y: number): { x: number; y: number } {
  return { x: m[0] * x + m[2] * y + m[4], y: m[1] * x + m[3] * y + m[5] };
}

/** Mock ctx that records the device-space origin (0,0 in local frame) of every
 *  fillText — i.e. each warped glyph's baseline point after all transforms. */
function trackingCtx() {
  const glyphs: Array<{ ch: string; x: number; y: number }> = [];
  let ctm: M = I;
  const stack: M[] = [];
  let fillStyle = '';
  let font = '';
  const ctx = {
    get fillStyle() {
      return fillStyle;
    },
    set fillStyle(v: string) {
      fillStyle = v;
    },
    get font() {
      return font;
    },
    set font(v: string) {
      font = v;
    },
    direction: 'ltr' as CanvasDirection,
    textAlign: 'left' as CanvasTextAlign,
    textBaseline: 'alphabetic' as CanvasTextBaseline,
    // Fixed 10px/char advance and ink metrics → predictable natural width.
    measureText: (s: string) => ({
      width: [...s].length * 10,
      actualBoundingBoxAscent: 7,
      actualBoundingBoxDescent: 2,
    }),
    fillText: (t: string, x: number, y: number) => {
      const p = apply(ctm, x, y);
      glyphs.push({ ch: t, x: p.x, y: p.y });
    },
    save: () => {
      stack.push(ctm);
    },
    restore: () => {
      ctm = stack.pop() ?? I;
    },
    translate: (x: number, y: number) => {
      ctm = mul(ctm, [1, 0, 0, 1, x, y]);
    },
    rotate: (a: number) => {
      ctm = mul(ctm, [Math.cos(a), Math.sin(a), -Math.sin(a), Math.cos(a), 0, 0]);
    },
    scale: (sx: number, sy: number) => {
      ctm = mul(ctm, [sx, 0, 0, sy, 0, 0]);
    },
    fillRect: () => {},
    beginPath: () => {},
    moveTo: () => {},
    lineTo: () => {},
    stroke: () => {},
    clip: () => {},
    rect: () => {},
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, glyphs };
}

function run(text: string, over: Partial<TextRunData> = {}): TextRunData {
  return {
    type: 'text',
    text,
    bold: null,
    italic: null,
    underline: false,
    strikethrough: false,
    fontSize: 40,
    color: '000000',
    fontFamily: 'Arial',
    ...over,
  };
}

function warpBody(
  preset: string,
  text: string,
  alignment: Paragraph['alignment'] = 'ctr',
): TextBody {
  const para: Paragraph = {
    alignment,
    marL: 0,
    marR: 0,
    indent: 0,
    spaceBefore: null,
    spaceAfter: null,
    spaceLine: null,
    lvl: 0,
    bullet: { type: 'none' } as Paragraph['bullet'],
    defFontSize: null,
    defColor: null,
    defBold: null,
    defItalic: null,
    defFontFamily: null,
    tabStops: [],
    eaLnBrk: true,
    runs: [run(text)],
  } as Paragraph;
  return {
    verticalAnchor: 'ctr',
    paragraphs: [para],
    defaultFontSize: 40,
    defaultBold: null,
    defaultItalic: null,
    lIns: 91440,
    rIns: 91440,
    tIns: 45720,
    bIns: 45720,
    wrap: 'square',
    vert: 'horz',
    autoFit: 'none',
    textWarp: { preset, adj: [] },
  } as TextBody;
}

function warpBodyWithBreak(preset: string, first: string, second: string): TextBody {
  const body = warpBody(preset, first);
  body.paragraphs[0]!.runs = [
    run(first),
    { type: 'break' },
    run(second),
  ];
  return body;
}

// The sample-16 WordArt boxes are 6.2in × 1.5in. At SCALE below, that box is
// BOX_W × BOX_H px. A short word ("Arch Up") is far narrower than the arch, so
// Follow Path should visibly compress its span.
const BOX_W = 620; // 6.2in → 620px
const BOX_H = 150; // 1.5in → 150px
const SCALE = 1 / 12700; // 40 pt becomes 40 px; match the mocked pixel geometry.

/** Horizontal device-space span of all placed glyph origins. */
function span(glyphs: Array<{ x: number }>): number {
  if (glyphs.length === 0) return 0;
  const xs = glyphs.map((g) => g.x);
  return Math.max(...xs) - Math.min(...xs);
}

describe('WordArt Follow Path — single-edge span (issue #846)', () => {
  it('textArchUp centres a centred word within its natural-width arc segment', () => {
    const { ctx, glyphs } = trackingCtx();
    renderTextBody(ctx, warpBody('textArchUp', 'Arch Up'), 0, 0, BOX_W, BOX_H, SCALE);
    expect(glyphs.length).toBeGreaterThan(0);
    // Natural ink width of "Arch Up" = 7 chars × 10px = 70px. The arch baseline
    // arc-length for a 620×150 box is several hundred px, so the ink span must
    // stay far below the box width — the word does NOT wrap around the ellipse.
    const s = span(glyphs);
    expect(s).toBeLessThan(BOX_W * 0.5);
    const xs = glyphs.map(({ x }) => x);
    expect((Math.min(...xs) + Math.max(...xs)) / 2).toBeCloseTo(BOX_W / 2, -1);
  });

  it('places left- and right-aligned text at the corresponding path ends', () => {
    const left = trackingCtx();
    renderTextBody(left.ctx, warpBody('textArchUp', 'Arch Up', 'l'), 0, 0, BOX_W, BOX_H, SCALE);
    const right = trackingCtx();
    renderTextBody(right.ctx, warpBody('textArchUp', 'Arch Up', 'r'), 0, 0, BOX_W, BOX_H, SCALE);

    const leftXs = left.glyphs.map(({ x }) => x);
    const rightXs = right.glyphs.map(({ x }) => x);
    expect(Math.max(...leftXs)).toBeLessThan(BOX_W / 2);
    expect(Math.min(...rightXs)).toBeGreaterThan(BOX_W / 2);
  });

  it('offsets an arch baseline by the glyph ink box, not the shape height', () => {
    const { ctx, glyphs } = trackingCtx();
    renderTextBody(ctx, warpBody('textArchUp', 'A'), 0, 0, BOX_W, BOX_H, SCALE);

    const env = buildWarpEnvelope('textArchUp', [], BOX_W, BOX_H)!;
    const inkHeight = 7 + 2;
    const expected = warpGlyphTransform(env, 0.5, inkHeight, 0.8);
    expect(glyphs).toHaveLength(1);
    // fillText is offset by half the glyph advance; the flattened arc's centre
    // tangent is within a small sampling angle of horizontal.
    expect(glyphs[0]!.y).toBeCloseTo(expected.y, 1);
  });

  it('keeps the last clockwise line on the authored arch and expands preceding lines', () => {
    const { ctx, glyphs } = trackingCtx();
    renderTextBody(ctx, warpBodyWithBreak('textArchUp', 'Top', 'Bottom'), 0, 0, BOX_W, BOX_H, SCALE);

    const topY = glyphs.slice(0, 3).reduce((sum, glyph) => sum + glyph.y, 0) / 3;
    const bottomY = glyphs.slice(3).reduce((sum, glyph) => sum + glyph.y, 0) / 6;
    expect(bottomY - topY).toBeGreaterThan(20);
    expect(topY).toBeLessThan(0);
    expect(bottomY).toBeLessThan(10);
  });

  it('expands later counterclockwise lines outside the lower arch', () => {
    const { ctx, glyphs } = trackingCtx();
    renderTextBody(ctx, warpBodyWithBreak('textArchDown', 'A', 'A'), 0, 0, BOX_W, BOX_H, SCALE);
    expect(glyphs[1]!.y - glyphs[0]!.y).toBeCloseTo(48, 0);
    expect(glyphs[0]!.y).toBeGreaterThan(BOX_H);
  });

  it('honours fixed line spacing independently of the font size', () => {
    const body = warpBodyWithBreak('textArchUp', 'A', 'A');
    body.paragraphs[0]!.spaceLine = { type: 'pts', val: 18 };
    const { ctx, glyphs } = trackingCtx();
    renderTextBody(ctx, body, 0, 0, BOX_W, BOX_H, SCALE);
    expect(glyphs[1]!.y - glyphs[0]!.y).toBeCloseTo(18, 0);
  });

  it('retains paragraph spacing only between paragraphs', () => {
    const body = warpBody('textArchUp', 'A');
    body.paragraphs[0]!.spaceBefore = 1200;
    body.paragraphs[0]!.spaceAfter = 600;
    body.paragraphs.push({ ...body.paragraphs[0]! });
    const { ctx, glyphs } = trackingCtx();
    renderTextBody(ctx, body, 0, 0, BOX_W, BOX_H, SCALE);
    expect(glyphs[1]!.y - glyphs[0]!.y).toBeCloseTo(66, 0);
    expect(glyphs[1]!.y).toBeLessThan(10);
  });

  it.each([50, 100, 150, 200])('uses %i%% authored line spacing', (percent) => {
    const body = warpBodyWithBreak('textArchUp', 'A', 'A');
    body.paragraphs[0]!.spaceLine = { type: 'pct', val: percent * 1000 };
    const { ctx, glyphs } = trackingCtx();
    renderTextBody(ctx, body, 0, 0, BOX_W, BOX_H, SCALE);
    expect(glyphs[1]!.y - glyphs[0]!.y).toBeCloseTo(48 * percent / 100, 0);
  });

  it('accumulates mixed-size line boxes before anchoring the last baseline', () => {
    const body = warpBody('textArchUp', 'A');
    body.paragraphs = [16, 32, 48].map(size => ({ ...body.paragraphs[0]!, runs: [run('A', { fontSize: size })] }));
    const { ctx, glyphs } = trackingCtx();
    renderTextBody(ctx, body, 0, 0, BOX_W, BOX_H, SCALE);
    expect(glyphs[1]!.y - glyphs[0]!.y).toBeCloseTo(34.56, 0);
    expect(glyphs[2]!.y - glyphs[1]!.y).toBeCloseTo(53.76, 0);
    expect(glyphs[2]!.y).toBeLessThan(10);
  });

  it.each([135, 225])('uses path winding for a %i degree downward arch', (angle) => {
    const body = warpBodyWithBreak('textArchDown', 'A', 'A');
    body.textWarp!.adj = [angle * 60000];
    const { ctx, glyphs } = trackingCtx();
    renderTextBody(ctx, body, 0, 0, BOX_W, BOX_H, SCALE);
    // These authored counterclockwise arcs are on the TOP half; following
    // screen-down instead of path winding would reverse the outward order.
    expect(glyphs[1]!.y).toBeLessThan(glyphs[0]!.y - 40);
  });

  it('anchors the final circle line and expands preceding lines radially', () => {
    const { ctx, glyphs } = trackingCtx();
    renderTextBody(ctx, warpBodyWithBreak('textCircle', 'A', 'A'), 0, 0, BOX_W, BOX_H, SCALE);
    expect(glyphs[0]!.x - glyphs[1]!.x).toBeCloseTo(48, 0);
    expect(glyphs[1]!.x).toBeGreaterThan(BOX_W - 10);
  });

  it('recomputes each line on its expanded ellipse, including arc-length alignment', () => {
    const body = warpBodyWithBreak('textArchUp', 'ABCDEFG', 'ABCDEFG');
    const multi = trackingCtx();
    renderTextBody(multi.ctx, body, 0, 0, 200, 180, SCALE);
    const outer = trackingCtx();
    renderTextBody(outer.ctx, warpBody('textArchUp', 'ABCDEFG'), -48, -48, 296, 276, SCALE);
    for (let i = 0; i < 7; i++) {
      expect(multi.glyphs[i]!.x).toBeCloseTo(outer.glyphs[i]!.x, 6);
      expect(multi.glyphs[i]!.y).toBeCloseTo(outer.glyphs[i]!.y, 6);
    }
  });

  it('textCircle keeps the word compact rather than scattering around the ellipse', () => {
    const { ctx, glyphs } = trackingCtx();
    renderTextBody(ctx, warpBody('textCircle', 'Circle'), 0, 0, BOX_W, BOX_H, SCALE);
    const s = span(glyphs);
    // "Circle" = 6 chars × 10px = 60px natural. Full-circle distribution would
    // spread glyphs across the whole ellipse width (≈ box width); Follow Path
    // keeps them in a compact arc segment.
    expect(s).toBeLessThan(BOX_W * 0.6);
  });

  it('a word wider than the arch still fills (clamps to) the full path', () => {
    const { ctx, glyphs } = trackingCtx();
    // 40 chars × 10px = 400px natural — comparable to the arc length, so it
    // spans (nearly) the whole arch and the span is large.
    const long = 'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA';
    renderTextBody(ctx, warpBody('textArchUp', long), 0, 0, BOX_W, BOX_H, SCALE);
    const s = span(glyphs);
    expect(s).toBeGreaterThan(BOX_W * 0.5);
  });
});
