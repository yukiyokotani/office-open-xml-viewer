import { layoutDocument } from './document-layout.js';
import { normalizeInternalDocumentModel } from './parser-model.js';
import { createLayoutServices } from './layout-runtime.js';
import { paintTextBoxLayout } from './paint/canvas-text.js';
import { afterEach, describe, it, expect, vi } from 'vitest';
import { rasterizeMathSvg } from '@silurus/ooxml-core';
import { createPaintResourceRegistry } from './layout/paint-resources.js';
import { createPaintResourceSession, unavailablePaintResourceHandle } from './paint/resource-session.js';
import { createCanvasPaintResourcePainter } from './paint/canvas-page.js';
import { canonicalCanvasPaintResourceHandlers } from './paint/canonical-resource-handlers.js';
import { prepareMathResources } from './paint/math-resources.js';

vi.mock('@silurus/ooxml-core', async (load) => ({
  ...await load<typeof import('@silurus/ooxml-core')>(),
  rasterizeMathSvg: vi.fn(async () => ({ source: {} })),
}));
afterEach(() => vi.restoreAllMocks());
import {
  acquireAndPaintShapeTextBox,
  acquireShapeTextBoxForTest,
} from './retained-shape-textbox.test-support.js';
import { createPageLayers } from './layout/page-graph.js';
import { textRunGeometryForPage } from './layout/text-index.js';
import type { DocumentLayout } from './layout/types.js';
import type { ShapeRun, ShapeText, ShapeTextRun } from './types';

// ECMA-376 §20.1.10.83 ST_TextVerticalType — a DrawingML text-box body direction
// (`<wps:bodyPr vert>`). Word writes it on a Word text box (`<wps:txbx>` shape):
//   - `vert`    : ALL glyphs rotated 90° CW  (chars T→B, lines R→L).
//   - `vert270` : ALL glyphs rotated 270° CW (= 90° CCW; chars B→T, lines L→R).
//   - `eaVert`  : East-Asian upright vertical — CJK stands UPRIGHT, non-EA glyphs
//                 rotated 90° (chars T→B, lines R→L). Mirrors the section-level
//                 tbRl per-glyph path (UAX#50 vo) and pptx's verified eaVert.
//   - `horz` / absent : horizontal (unchanged legacy path).
//
// The renderer laies the box out with the SAME horizontal engine, rotated ±90°
// about the box centre with width/height swapped, so the layout/kinsoku/bidi are
// reused. Because the text box does NOT emit `onTextRun`, we verify by recording
// the Canvas transform at every glyph draw and asserting each glyph's NET
// rotation (the angle its local +x axis makes in device space).

interface GlyphCall {
  text: string;
  /** Net rotation of the drawn glyph in device space, degrees (atan2(b,a)). */
  angleDeg: number;
  /** Device-space position of the draw origin (local (x,y) mapped by the CTM). */
  devX: number;
  devY: number;
}

interface ImageCall {
  /** Net rotation of the drawn image in device space, degrees (atan2(b,a)). */
  angleDeg: number;
  /** Device-space position of the draw origin (local (x,y) mapped by the CTM). */
  devX: number;
  devY: number;
  /** Local draw width/height (before the CTM). */
  w: number;
  h: number;
}

/** Recording 2D context that tracks the full affine CTM (a,b,c,d,e,f) across
 *  save/restore/translate/rotate/scale, so a glyph's net orientation is
 *  recoverable at fillText time. measureText gives every code point the current
 *  font px as advance (1 em / CJK), with symmetric font/ink boxes. */
function makeMatrixCtx(): {
  ctx: CanvasRenderingContext2D;
  glyphs: GlyphCall[];
  images: ImageCall[];
  clips: boolean[];
} {
  let m = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };
  const stack: (typeof m)[] = [];
  let font = '10px serif';
  let textAlign = 'start';
  let textBaseline = 'alphabetic';
  let letterSpacing = '0px';
  let fillStyle = '#000';
  let direction = 'ltr';
  let fontKerning = 'auto';
  const glyphs: GlyphCall[] = [];
  const images: ImageCall[] = [];
  const clips: boolean[] = [];
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    get textAlign() { return textAlign; },
    set textAlign(v: string) { textAlign = v; },
    get textBaseline() { return textBaseline; },
    set textBaseline(v: string) { textBaseline = v; },
    get letterSpacing() { return letterSpacing; },
    set letterSpacing(v: string) { letterSpacing = v; },
    get fillStyle() { return fillStyle; },
    set fillStyle(v: string) { fillStyle = v; },
    get direction() { return direction; },
    set direction(v: string) { direction = v; },
    get fontKerning() { return fontKerning; },
    set fontKerning(v: string) { fontKerning = v; },
    strokeStyle: '#000',
    lineWidth: 1,
    globalAlpha: 1,
    save() { stack.push({ ...m }); },
    restore() { const s = stack.pop(); if (s) m = s; },
    translate(tx: number, ty: number) {
      m = { ...m, e: m.e + m.a * tx + m.c * ty, f: m.f + m.b * tx + m.d * ty };
    },
    rotate(t: number) {
      const cos = Math.cos(t), sin = Math.sin(t);
      m = {
        a: m.a * cos + m.c * sin,
        b: m.b * cos + m.d * sin,
        c: -m.a * sin + m.c * cos,
        d: -m.b * sin + m.d * cos,
        e: m.e,
        f: m.f,
      };
    },
    transform(a: number, b: number, c: number, d: number, e: number, f: number) {
      m = { a: m.a * a + m.c * b, b: m.b * a + m.d * b,
        c: m.a * c + m.c * d, d: m.b * c + m.d * d,
        e: m.a * e + m.c * f + m.e, f: m.b * e + m.d * f + m.f };
    },
    scale(sx: number, sy: number) {
      m = { ...m, a: m.a * sx, b: m.b * sx, c: m.c * sy, d: m.d * sy };
    },
    beginPath() {}, closePath() {}, moveTo() {}, lineTo() {}, rect() {},
    fill() {}, stroke() {}, clip() { clips.push(true); }, fillRect() {}, strokeRect() {}, clearRect() {},
    setTransform() {}, resetTransform() {},
    measureText(s: string) {
      const p = px();
      return {
        width: [...s].length * p,
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    fillText(text: string, x: number, y: number) {
      const angleDeg = (Math.atan2(m.b, m.a) * 180) / Math.PI;
      const devX = m.a * x + m.c * y + m.e;
      const devY = m.b * x + m.d * y + m.f;
      glyphs.push({ text, angleDeg, devX, devY });
    },
    strokeText() {},
    drawImage(_bmp: unknown, x: number, y: number, w: number, h: number) {
      const angleDeg = (Math.atan2(m.b, m.a) * 180) / Math.PI;
      const devX = m.a * x + m.c * y + m.e;
      const devY = m.b * x + m.d * y + m.f;
      images.push({ angleDeg, devX, devY, w, h });
    },
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, glyphs, images, clips };
}

function richTextbox(
  runs: ShapeTextRun[],
  textVert?: string | null,
  alignment = 'left',
): ShapeRun {
  const block: ShapeText = {
    text: runs.map((r) => r.text).join(''),
    fontSizePt: runs[0]?.fontSizePt ?? 10,
    alignment,
    runs,
  };
  return {
    type: 'shape', zOrder: 0, subpaths: [], presetGeometry: 'rect', fill: null, stroke: null,
    textBlocks: [block], textAnchor: 't',
    textInsetL: 0, textInsetT: 0, textInsetR: 0, textInsetB: 0,
    textVert: textVert ?? null,
  } as unknown as ShapeRun;
}

const CJK = '経'; // UAX#50 vo=U (upright)
const LAT = 'A'; // vo=R (sideways)
const NEAR = (a: number, b: number, tol = 1e-6) => Math.abs(a - b) <= tol;
// Normalise an angle to (-180, 180].
const norm = (deg: number) => ((((deg + 180) % 360) + 360) % 360) - 180;

describe('§20.1.10.83 textbox <wps:bodyPr vert> — vertical text-box rendering', () => {
  const run = (text: string): ShapeTextRun => ({ text, fontSizePt: 10, fontFamily: 'NotInMetrics' });

  it('horz (absent vert): no rotation — glyphs upright, CTM identity (legacy path)', () => {
    const { ctx, glyphs } = makeMatrixCtx();
    acquireAndPaintShapeTextBox(richTextbox([run(CJK + LAT)]), 0, 0, 200, 100, ctx, 1, {});
    const drawn = glyphs.filter((g) => g.text.includes(CJK) || g.text.includes(LAT) || /経|A/.test(g.text));
    expect(drawn.length).toBeGreaterThan(0);
    for (const g of drawn) expect(NEAR(norm(g.angleDeg), 0)).toBe(true);
  });

  it.each([
    [-90, false, null, -90], [30, false, null, 30],
    [-90, true, null, 0], [-90, false, 'vert', 0],
    [90, false, 'vert270', 0],
  ])('composes shape rotation=%s upright=%s vert=%s once', (rotation, textUpright, textVert, angle) => {
    const { ctx, glyphs } = makeMatrixCtx();
    const shape = { ...richTextbox([run('A')], textVert), rotation, textUpright };
    acquireAndPaintShapeTextBox(shape, 10, 20, 200, 100, ctx, 1, {});
    expect(norm(glyphs.find((g) => g.text === 'A')!.angleDeg)).toBeCloseTo(angle);
  });

  it('rotates ordinary text around its shape center without mirroring readable glyphs', () => {
    const original = makeMatrixCtx();
    const rotated = makeMatrixCtx();
    const shape = richTextbox([run('A')]);
    acquireAndPaintShapeTextBox(shape, 10, 20, 200, 100, original.ctx, 1, {});
    acquireAndPaintShapeTextBox({ ...shape, rotation: -90, flipH: true, flipV: true },
      10, 20, 200, 100, rotated.ctx, 1, {});
    const before = original.glyphs.find((g) => g.text === 'A')!;
    const after = rotated.glyphs.find((g) => g.text === 'A')!;
    expect(after.devX).toBeCloseTo(110 + before.devY - 70);
    expect(after.devY).toBeCloseTo(70 - before.devX + 110);
    expect(norm(after.angleDeg)).toBeCloseTo(-90);
  });

  it('keeps upright stacked text independent of the shape rotation and vertical flip', () => {
    const reference = makeMatrixCtx();
    const transformed = makeMatrixCtx();
    const shape = { ...richTextbox([run('A')], 'wordArtVert'), textUpright: true };
    acquireAndPaintShapeTextBox(shape, 10, 20, 200, 100, reference.ctx, 1, {});
    acquireAndPaintShapeTextBox({ ...shape, rotation: 30, flipV: true },
      10, 20, 200, 100, transformed.ctx, 1, {});
    expect(transformed.glyphs).toEqual(reference.glyphs);
  });

  it('projects selectable text through the retained shape rotation', () => {
    const { ctx } = makeMatrixCtx();
    const shape = { ...richTextbox([run('A')]), rotation: -90 };
    const textBox = acquireShapeTextBoxForTest(shape, 10, 20, 200, 100, ctx, 1, {});
    expect(textBox).toBeDefined();
    const layers = createPageLayers([{
      layer: 'body', node: textBox!, coordinateSpace: 'upright-physical',
    }]);
    const layout = { pages: [{
      pageIndex: 0, flowDomains: [], readingOrder: [textBox!.id], sectionRegions: [], layers,
    }] } as unknown as DocumentLayout;
    const [geometry] = textRunGeometryForPage(layout, 0);

    expect(geometry?.placement.text).toBe('A');
    expect(geometry?.pointToPage).toMatchObject({
      a: 0, b: -1, c: 1, d: 0, e: 40, f: 180,
    });
  });

  it('vert: every glyph rotated +90° CW (all-rotate; CJK included)', () => {
    const { ctx, glyphs } = makeMatrixCtx();
    acquireAndPaintShapeTextBox(richTextbox([run(CJK + LAT)], 'vert'), 0, 0, 200, 100, ctx, 1, {});
    expect(glyphs.length).toBeGreaterThan(0);
    for (const g of glyphs) expect(NEAR(norm(g.angleDeg), 90), `${g.text}@${g.angleDeg}`).toBe(true);
  });

  it('vert270: every glyph rotated −90° (270° CW)', () => {
    const { ctx, glyphs } = makeMatrixCtx();
    acquireAndPaintShapeTextBox(richTextbox([run(CJK + LAT)], 'vert270'), 0, 0, 200, 100, ctx, 1, {});
    expect(glyphs.length).toBeGreaterThan(0);
    for (const g of glyphs) expect(NEAR(norm(g.angleDeg), -90), `${g.text}@${g.angleDeg}`).toBe(true);
  });

  it('eaVert: CJK stands UPRIGHT (net 0°) while Latin stays sideways (net +90°)', () => {
    const { ctx, glyphs } = makeMatrixCtx();
    acquireAndPaintShapeTextBox(richTextbox([run(CJK), run(LAT)], 'eaVert'), 0, 0, 200, 100, ctx, 1, {});
    const cjk = glyphs.find((g) => g.text.includes(CJK));
    const lat = glyphs.find((g) => g.text.includes(LAT));
    expect(cjk, 'CJK glyph drawn').toBeDefined();
    expect(lat, 'Latin glyph drawn').toBeDefined();
    expect(NEAR(norm(cjk!.angleDeg), 0), `CJK @${cjk!.angleDeg}`).toBe(true);
    expect(NEAR(norm(lat!.angleDeg), 90), `Latin @${lat!.angleDeg}`).toBe(true);
  });

  it.each(['wordArtVert', 'wordArtVertRtl'])(
    '%s: Word keeps Latin sideways, CJK upright and columns left-to-right',
    (mode) => {
      const { ctx, glyphs } = makeMatrixCtx();
      const shape = richTextbox([run('AB'), run(CJK)], mode);
      shape.textBlocks = [...shape.textBlocks!, { text: 'CD', fontSizePt: 10, alignment: 'left', runs: [run('CD')] }];
      acquireAndPaintShapeTextBox(shape, 0, 0, 200, 100, ctx, 1, {});
      const a = glyphs.find((g) => g.text === 'A')!;
      const b = glyphs.find((g) => g.text === 'B')!;
      const c = glyphs.find((g) => g.text === 'C')!;
      const cjk = glyphs.find((g) => g.text === CJK)!;
      expect(norm(a.angleDeg)).toBeCloseTo(90);
      expect(norm(cjk.angleDeg)).toBeCloseTo(0);
      expect(b.devY - a.devY).toBeCloseTo(10); // ordinary horizontal advance
      expect(b.devX).toBeCloseTo(a.devX);
      expect(c.devX).toBeGreaterThan(a.devX);
    },
  );

  it('WordArt wrap=none retains one continuous overflowing column', () => {
    const { ctx, glyphs, clips } = makeMatrixCtx();
    const shape = Object.assign(richTextbox([run('ABCDE')], 'wordArtVert'), { textWrap: 'none', textAutofit: 'none' });
    acquireAndPaintShapeTextBox(shape, 0, 0, 100, 30, ctx, 1, {});
    expect(glyphs.map((g) => g.text).join('')).toBe('ABCDE');
    expect(glyphs.every((g) => Math.abs(g.devX - glyphs[0].devX) < 0.001)).toBe(true);
    expect(glyphs.at(-1)!.devY).toBeGreaterThan(30);
    expect(clips).toEqual([]);
  });

  it('WordArt keeps physical inset axes and moves anchors along LTR columns', () => {
    const draw = (extra: Partial<ShapeRun>) => {
      const { ctx, glyphs } = makeMatrixCtx();
      const shape = Object.assign(richTextbox([run('AB')], 'wordArtVert'), extra);
      acquireAndPaintShapeTextBox(shape, 0, 0, 100, 100, ctx, 1, {});
      return glyphs[0];
    };
    const start = draw({});
    const inset = draw({ textInsetL: 10, textInsetT: 20 });
    expect(inset.devX - start.devX).toBeCloseTo(10);
    expect(inset.devY - start.devY).toBeCloseTo(20);
    const center = draw({ textAnchor: 'ctr' });
    const end = draw({ textAnchor: 'b' });
    expect(center.devX).toBeGreaterThan(start.devX);
    expect(end.devX).toBeGreaterThan(center.devX);
  });

  it('WordArt keeps emoji presentation clusters upright alongside sideways Latin', () => {
    const { ctx, glyphs } = makeMatrixCtx();
    acquireAndPaintShapeTextBox(richTextbox([run('A👩‍💻🇯🇵')], 'wordArtVert'), 0, 0, 200, 100, ctx, 1, {});
    expect(norm(glyphs.find((g) => g.text === 'A')!.angleDeg)).toBeCloseTo(90);
    const emoji = glyphs.filter((g) => /\p{Emoji_Presentation}/u.test(g.text));
    expect(emoji.length).toBeGreaterThan(0);
    for (const glyph of emoji) expect(norm(glyph.angleDeg)).toBeCloseTo(0);
  });

  it.each([
    [30, false, false, 120], [90, false, false, 180],
    [0, true, false, 90], [0, false, true, -90],
  ])('WordArt text frame rotation=%s flipH=%s flipV=%s', (rotation, flipH, flipV, angle) => {
    const { ctx, glyphs } = makeMatrixCtx();
    const shape = { ...richTextbox([run('AB')], 'wordArtVert'), rotation, flipH, flipV };
    acquireAndPaintShapeTextBox(shape, 0, 0, 200, 100, ctx, 1, {});
    expect(norm(glyphs.find((g) => g.text === 'A')!.angleDeg)).toBeCloseTo(norm(angle));
  });

  it('rotated glyphs land INSIDE the physical box (transform pivots on box centre)', () => {
    // A vert box of physical 200×100: after the +90° rotation about the centre,
    // every drawn glyph's device origin must still fall within the physical box.
    const { ctx, glyphs } = makeMatrixCtx();
    acquireAndPaintShapeTextBox(richTextbox([run(CJK + CJK + LAT)], 'vert'), 0, 0, 200, 100, ctx, 1, {});
    expect(glyphs.length).toBeGreaterThan(0);
    for (const g of glyphs) {
      expect(g.devX, `${g.text} devX in [0,200]`).toBeGreaterThanOrEqual(-1);
      expect(g.devX).toBeLessThanOrEqual(201);
      expect(g.devY, `${g.text} devY in [0,100]`).toBeGreaterThanOrEqual(-1);
      expect(g.devY).toBeLessThanOrEqual(101);
    }
  });

  it('eaVert justified (both) column advances by NATURAL width — no §17.18.44 stretch drift', () => {
    // A justified (`both`) eaVert paragraph that WRAPS: the first column has slack
    // (a long Latin word can't break mid-word, so it wraps whole), which the
    // horizontal justify pass would distribute into inter-segment gaps. Inside an
    // eaVert cell that distribution must NOT be painted — the column flows
    // start-aligned by its natural measured widths — so the four CJK cells on the
    // first column stay UNIFORMLY spaced across the two-run boundary. (The prior
    // bug advanced by the justify-expanded width, opening a gap at the run seam.)
    const { ctx, glyphs } = makeMatrixCtx();
    const two: ShapeTextRun[] = [run('経経'), run('済済'), run('ABCDEFGHIJ')];
    const block: ShapeText = {
      text: '経経済済ABCDEFGHIJ',
      fontSizePt: 10,
      alignment: 'both',
      runs: two,
    };
    const shape = {
      type: 'shape', zOrder: 0, subpaths: [], presetGeometry: 'rect', fill: null, stroke: null,
      textBlocks: [block], textAnchor: 't',
      textInsetL: 0, textInsetT: 0, textInsetR: 0, textInsetB: 0,
      textVert: 'eaVert',
    } as unknown as ShapeRun;
    // Box 200×100 → logical column length 100 → 10 cells of the 10px font. The
    // first column holds 経経済済 (4 cells); ABCDEFGHIJ (10 cells) wraps whole.
    acquireAndPaintShapeTextBox(shape, 0, 0, 200, 100, ctx, 1, {});
    const cjk = glyphs.filter((g) => /[経済]/.test(g.text) && [...g.text].length === 1);
    expect(cjk.length, 'four upright CJK cells on the justified first column').toBe(4);
    // The along-column axis is device +y (the +90° frame maps local +x → +y).
    const dys = cjk.map((g) => g.devY).sort((a, b) => a - b);
    const gaps = dys.slice(1).map((v, i) => v - dys[i]);
    const minGap = Math.min(...gaps);
    const maxGap = Math.max(...gaps);
    // All three inter-cell gaps equal (uniform natural pitch); the bug widened the
    // run1→run2 gap by the distributed slack.
    expect(maxGap - minGap, `uniform cell pitch, gaps=${gaps}`).toBeLessThan(0.5);
  });

  it('vert vs vert270 stack lines to OPPOSITE sides of the box centre', () => {
    // Two lines (a hard wrap via two blocks) → the second line sits on the
    // opposite cross-side for vert (R→L, leftwards) vs vert270 (L→R, rightwards).
    const two = (v: string) => {
      const { ctx, glyphs } = makeMatrixCtx();
      const block2: ShapeText = { text: CJK, fontSizePt: 10, alignment: 'left', runs: [run(CJK)] };
      const shape = richTextbox([run(CJK)], v);
      (shape as unknown as { textBlocks: ShapeText[] }).textBlocks.push(block2);
      acquireAndPaintShapeTextBox(shape, 0, 0, 200, 100, ctx, 1, {});
      return glyphs;
    };
    const gv = two('vert');
    const g270 = two('vert270');
    // Box centre is at device x = 100 (w=200). vert: lines go R→L, so line 1 sits
    // right of centre, line 2 left. vert270: L→R, mirrored.
    expect(gv[0].devX).toBeGreaterThan(gv[gv.length - 1].devX);
    expect(g270[0].devX).toBeLessThan(g270[g270.length - 1].devX);
  });

  // ── (a) mongolianVert (§20.1.10.83) ──────────────────────────────────────
  // GT (batch-3 adjudication): identical per-glyph orientation to eaVert (CJK
  // UPRIGHT, Latin sideways 90° CW), but the line/column progression is the
  // MIRROR of eaVert — columns advance LEFT→RIGHT instead of right→left.
  it('mongolianVert: CJK upright (0°) while Latin stays sideways (+90°) — same as eaVert', () => {
    const { ctx, glyphs } = makeMatrixCtx();
    acquireAndPaintShapeTextBox(richTextbox([run(CJK), run(LAT)], 'mongolianVert'), 0, 0, 200, 100, ctx, 1, {});
    const cjk = glyphs.find((g) => g.text.includes(CJK));
    const lat = glyphs.find((g) => g.text.includes(LAT));
    expect(cjk, 'CJK glyph drawn').toBeDefined();
    expect(lat, 'Latin glyph drawn').toBeDefined();
    expect(NEAR(norm(cjk!.angleDeg), 0), `CJK @${cjk!.angleDeg}`).toBe(true);
    expect(NEAR(norm(lat!.angleDeg), 90), `Latin @${lat!.angleDeg}`).toBe(true);
  });

  it('mongolianVert stacks columns LEFT→RIGHT (mirror of eaVert R→L)', () => {
    // Two blocks (two columns). eaVert puts the first column on the RIGHT and the
    // second to its left; mongolianVert mirrors it — first column on the LEFT,
    // second to its right.
    const twoCols = (v: string) => {
      const { ctx, glyphs } = makeMatrixCtx();
      const block2: ShapeText = { text: CJK, fontSizePt: 10, alignment: 'left', runs: [run(CJK)] };
      const shape = richTextbox([run(CJK)], v);
      (shape as unknown as { textBlocks: ShapeText[] }).textBlocks.push(block2);
      acquireAndPaintShapeTextBox(shape, 0, 0, 200, 100, ctx, 1, {});
      return glyphs;
    };
    const ea = twoCols('eaVert');
    const mn = twoCols('mongolianVert');
    // eaVert: first column right of the last. mongolianVert: first column LEFT.
    expect(ea[0].devX).toBeGreaterThan(ea[ea.length - 1].devX);
    expect(mn[0].devX, `mongolianVert first col @${mn[0].devX} < last @${mn[mn.length - 1].devX}`)
      .toBeLessThan(mn[mn.length - 1].devX);
  });

  // ── (b) eaVert + ruby (§17.3.3.25) ───────────────────────────────────────
  // GT: furigana sits on the RIGHT side of the vertical base column, upright,
  // running top→bottom. In the +90° CW frame the physical RIGHT is device +x.
  it('eaVert ruby draws furigana upright on the device-RIGHT of the base column', () => {
    const { ctx, glyphs } = makeMatrixCtx();
    const baseRun: ShapeTextRun = {
      text: '漢字',
      fontSizePt: 10,
      fontFamily: 'NotInMetrics',
      ruby: { text: 'かんじ', fontSizePt: 5 },
    };
    acquireAndPaintShapeTextBox(richTextbox([baseRun], 'eaVert'), 0, 0, 200, 100, ctx, 1, {});
    const base = glyphs.filter((g) => /[漢字]/.test(g.text));
    const ruby = glyphs.filter((g) => /[かんじ]/.test(g.text));
    expect(base.length, 'base glyphs drawn').toBeGreaterThan(0);
    expect(ruby.length, 'ruby glyphs drawn').toBe(3);
    // Ruby stands upright (net 0°), like the upright CJK base cells.
    for (const r of ruby) expect(NEAR(norm(r.angleDeg), 0), `ruby ${r.text}@${r.angleDeg}`).toBe(true);
    // Physical RIGHT = device +x: every ruby glyph is right of every base glyph.
    const baseMaxX = Math.max(...base.map((g) => g.devX));
    const rubyMinX = Math.min(...ruby.map((g) => g.devX));
    expect(rubyMinX, `ruby minX ${rubyMinX} > base maxX ${baseMaxX}`).toBeGreaterThan(baseMaxX);
    // Exact cross offset comes from retained selected-face ink: base ascent 8pt
    // plus guide descent 1pt = 9pt. No em-ratio placement is reconstructed.
    const meanBaseX = base.reduce((s, g) => s + g.devX, 0) / base.length;
    const meanRubyX = ruby.reduce((s, g) => s + g.devX, 0) / ruby.length;
    expect(meanRubyX - meanBaseX, `cross offset ${meanRubyX - meanBaseX} = base ascent + guide descent`)
      .toBeCloseTo(9, 5);
  });

  // ── (c) vert + inline image: image stays UPRIGHT ─────────────────────────
  // GT: the inline raster keeps its physical orientation (a graphic is not text),
  // even though the surrounding `vert` glyphs rotate 90° CW.
  it('vert inline image is drawn UPRIGHT (net 0°), not rotated with the text frame', () => {
    const { ctx, images } = makeMatrixCtx();
    const imgBlock: ShapeText = {
      text: '', fontSizePt: 10, alignment: 'left',
      imagePath: 'word/media/image1.png', imageWidthPt: 24, imageHeightPt: 36,
    } as unknown as ShapeText;
    const shape = richTextbox([{ text: CJK, fontSizePt: 10, fontFamily: 'NotInMetrics' }], 'vert');
    (shape as unknown as { textBlocks: ShapeText[] }).textBlocks.push(imgBlock);
    const fakeBmp = { width: 24, height: 36 } as unknown as ImageBitmap;
    const imgs = new Map<string, ImageBitmap>([['word/media/image1.png', fakeBmp]]);
    acquireAndPaintShapeTextBox(shape, 0, 0, 200, 100, ctx, 1, {}, imgs as never);
    expect(images.length, 'one image drawn').toBe(1);
    // Upright: the net rotation cancels the +90° page frame → 0°.
    expect(NEAR(norm(images[0].angleDeg), 0), `image @${images[0].angleDeg}`).toBe(true);
    // Portrait aspect preserved (physical 24×36, not swapped to 36×24). The draw
    // dimensions in the upright local frame are width=physical-width,
    // height=physical-height (the callback receives dw=cross, dh=along).
    expect(Math.abs(images[0].w)).toBeCloseTo(24, 3);
    expect(Math.abs(images[0].h)).toBeCloseTo(36, 3);
  });

  it('vert270 inline image is ALSO drawn upright (net 0°), not −180° (counter-rotation sign)', () => {
    // vert270's page frame is −90°, so the image counter-rotation must be +90°
    // (not the −90° the +90° modes use) — otherwise the raster is flipped 180°.
    const { ctx, images } = makeMatrixCtx();
    const imgBlock: ShapeText = {
      text: '', fontSizePt: 10, alignment: 'left',
      imagePath: 'word/media/image1.png', imageWidthPt: 24, imageHeightPt: 36,
    } as unknown as ShapeText;
    const shape = richTextbox([{ text: CJK, fontSizePt: 10, fontFamily: 'NotInMetrics' }], 'vert270');
    (shape as unknown as { textBlocks: ShapeText[] }).textBlocks.push(imgBlock);
    const fakeBmp = { width: 24, height: 36 } as unknown as ImageBitmap;
    const imgs = new Map<string, ImageBitmap>([['word/media/image1.png', fakeBmp]]);
    acquireAndPaintShapeTextBox(shape, 0, 0, 200, 100, ctx, 1, {}, imgs as never);
    expect(images.length).toBe(1);
    expect(NEAR(norm(images[0].angleDeg), 0), `vert270 image @${images[0].angleDeg}`).toBe(true);
  });

  it('vertical inline image reserves its physical WIDTH (crossExtent), not its height, along the column stack', () => {
    // Two CJK text columns sandwich the image column; their cross (device-x)
    // separation = text-line-box + the image's reserved cross extent. A tall-thin
    // image (10 wide × 90 tall) reserves its physical WIDTH 10 (crossExtent) — the
    // fitH-vs-crossExtent bug would instead reserve its 90-tall height, pushing the
    // trailing column ~80px further, so the separation cleanly distinguishes them.
    const { ctx, glyphs } = makeMatrixCtx();
    const imgBlock: ShapeText = {
      text: '', fontSizePt: 10, alignment: 'left',
      imagePath: 'word/media/image1.png', imageWidthPt: 10, imageHeightPt: 90,
    } as unknown as ShapeText;
    const shape = richTextbox([{ text: CJK, fontSizePt: 10, fontFamily: 'NotInMetrics' }], 'vert');
    const tb = shape as unknown as { textBlocks: ShapeText[] };
    tb.textBlocks.push(imgBlock);
    tb.textBlocks.push({ text: CJK, fontSizePt: 10, alignment: 'left', runs: [run(CJK)] } as ShapeText);
    const fakeBmp = { width: 10, height: 90 } as unknown as ImageBitmap;
    const imgs = new Map<string, ImageBitmap>([['word/media/image1.png', fakeBmp]]);
    acquireAndPaintShapeTextBox(shape, 0, 0, 200, 100, ctx, 1, {}, imgs as never);
    const cjk = glyphs.filter((g) => g.text.includes(CJK));
    const firstCol = cjk[0].devX;
    const lastCol = cjk[cjk.length - 1].devX;
    const sep = Math.abs(firstCol - lastCol);
    // Separation ≈ one text line box (~10-14) + image cross 10 ≈ 20-30. With the
    // bug (reserving the 90 height) it would exceed 90. Assert well below 90.
    expect(sep, `two-column separation ${sep} reflects image cross 10, not height 90`).toBeLessThan(45);
    expect(sep, `columns are actually separated ${sep}`).toBeGreaterThan(12);
  });

  // ── cross-axis spacing (sample-53 class regressions) ─────────────────────
  // GT (Word PDF): a ruby-bearing line's COLUMN is widened by the ruby
  // selected-face ink reservation acquired by layoutLines(), so the base column clears the
  // previous column instead of the furigana overprinting it.
  it('eaVert ruby line widens its column advance by the ruby reservation (§17.3.3.25)', () => {
    const advance = (withRuby: boolean): number => {
      const { ctx, glyphs } = makeMatrixCtx();
      const labelBlock: ShapeText = {
        text: CJK, fontSizePt: 10, alignment: 'left',
        runs: [{ text: CJK, fontSizePt: 10, fontFamily: 'NotInMetrics' }],
      };
      const baseRun: ShapeTextRun = {
        text: '漢', fontSizePt: 10, fontFamily: 'NotInMetrics',
        ...(withRuby ? { ruby: { text: 'かん', fontSizePt: 5 } } : {}),
      };
      const baseBlock: ShapeText = { text: '漢', fontSizePt: 10, alignment: 'left', runs: [baseRun] };
      const shape = richTextbox([{ text: CJK, fontSizePt: 10, fontFamily: 'NotInMetrics' }], 'eaVert');
      (shape as unknown as { textBlocks: ShapeText[] }).textBlocks = [labelBlock, baseBlock];
      acquireAndPaintShapeTextBox(shape, 0, 0, 200, 100, ctx, 1, {});
      const label = glyphs.find((g) => g.text.includes(CJK));
      const base = glyphs.find((g) => g.text.includes('漢'));
      expect(label, 'label glyph drawn').toBeDefined();
      expect(base, 'base glyph drawn').toBeDefined();
      // eaVert stacks columns R→L: the base column sits LEFT of the label column.
      return label!.devX - base!.devX;
    };
    const plain = advance(false);
    const withRuby = advance(true);
    // Base ascent 8pt + guide descent 1pt = 9pt of exact retained reserve.
    expect(withRuby - plain, `ruby advance ${withRuby} vs plain ${plain}`).toBeCloseTo(9, 5);
  });

  // GT (Word PDF): mongolianVert's first column is the exact MIRROR of the
  // eaVert first column (sample-53: flow-start-edge → label glyph-box gap is
  // preserved within 0.5pt). Reflecting only the band ORIGIN keeps the baseline
  // measured from the band's physical RIGHT edge, so asymmetric leading (line
  // box taller than the glyph em box) shoves the mongolian line toward the
  // physical left — the centerline inside the band must be reflected too.
  it('mongolianVert mirrors the line centerline within its cross-axis band', () => {
    const firstColumnCenterX = (mode: 'eaVert' | 'mongolianVert'): number => {
      const { ctx, glyphs } = makeMatrixCtx();
      const shape = richTextbox([{ text: CJK, fontSizePt: 11, fontFamily: 'NotInMetrics' }], mode);
      const [block] = (shape as unknown as { textBlocks: ShapeText[] }).textBlocks;
      // Exact 20pt line box > the 11pt glyph box → asymmetric leading exposes a
      // non-mirrored centerline (natural == lineH would accidentally pass).
      (block as unknown as { lineSpacingRule: string; lineSpacingVal: number }).lineSpacingRule = 'exact';
      (block as unknown as { lineSpacingRule: string; lineSpacingVal: number }).lineSpacingVal = 20;
      acquireAndPaintShapeTextBox(shape, 0, 0, 200, 100, ctx, 1, {});
      const glyph = glyphs.find((g) => g.text.includes(CJK));
      expect(glyph, 'first-column glyph drawn').toBeDefined();
      // Upright CJK cells draw centred on the column centerline, so devX IS the
      // centerline's cross position.
      return glyph!.devX;
    };
    const eaX = firstColumnCenterX('eaVert');
    const mongolianX = firstColumnCenterX('mongolianVert');
    // Equal glyphs have equal cross extents, so mirrored centerlines imply equal
    // flow-start-edge → glyph-box gaps (box spans x ∈ [0, 200], insets 0).
    expect(mongolianX, `mongolian centerline ${mongolianX} mirrors eaVert ${eaX}`).toBeCloseTo(200 - eaX, 5);
  });

  // mongolianVert + ruby: the furigana keeps its orientation side (physical
  // RIGHT of the base column, like eaVert) — so the mirror must keep the BASE
  // cell at its non-ruby mirrored position and spend the ruby reservation on
  // the band's interior (right) side. Mirroring the ruby-inflated baseline
  // offset naively would shove the base toward the band's right edge and paint
  // the furigana on top of the NEXT column.
  it('mongolianVert ruby line keeps its base column mirrored and reserves toward the next column', () => {
    const render = (withRuby: boolean) => {
      const { ctx, glyphs } = makeMatrixCtx();
      const rubyBlock: ShapeText = {
        text: '漢', fontSizePt: 10, alignment: 'left',
        runs: [{
          text: '漢', fontSizePt: 10, fontFamily: 'NotInMetrics',
          ...(withRuby ? { ruby: { text: 'かん', fontSizePt: 5 } } : {}),
        }],
      };
      const nextBlock: ShapeText = {
        text: CJK, fontSizePt: 10, alignment: 'left',
        runs: [{ text: CJK, fontSizePt: 10, fontFamily: 'NotInMetrics' }],
      };
      const shape = richTextbox([{ text: '漢', fontSizePt: 10, fontFamily: 'NotInMetrics' }], 'mongolianVert');
      (shape as unknown as { textBlocks: ShapeText[] }).textBlocks = [rubyBlock, nextBlock];
      acquireAndPaintShapeTextBox(shape, 0, 0, 200, 100, ctx, 1, {});
      const base = glyphs.find((g) => g.text.includes('漢'));
      const next = glyphs.find((g) => g.text.includes(CJK));
      const ruby = glyphs.filter((g) => /[かん]/.test(g.text));
      expect(base, 'base glyph drawn').toBeDefined();
      expect(next, 'next-column glyph drawn').toBeDefined();
      if (withRuby) expect(ruby.length, 'ruby glyphs drawn').toBe(2);
      return { baseX: base!.devX, nextX: next!.devX, rubyXs: ruby.map((g) => g.devX) };
    };
    const plain = render(false);
    const withRuby = render(true);
    // The base column does NOT move: the reservation goes to the ruby side
    // (band interior), not the flow-start edge.
    expect(withRuby.baseX, `base ${withRuby.baseX} unmoved from ${plain.baseX}`).toBeCloseTo(plain.baseX, 5);
    // The NEXT column (to the right, L→R) advances by the ruby reservation.
    expect(withRuby.nextX - plain.nextX, 'next column pushed by the reservation').toBeCloseTo(9, 5);
    // The furigana sits BETWEEN its base column and the next column — physical
    // right of the base (orientation preserved), clear of the next column.
    for (const rx of withRuby.rubyXs) {
      expect(rx, `ruby ${rx} right of base ${withRuby.baseX}`).toBeGreaterThan(withRuby.baseX);
      expect(rx, `ruby ${rx} clears next column ${withRuby.nextX}`).toBeLessThan(withRuby.nextX - 5);
    }
  });

  it('mongolianVert hpsRaise increases the ruby-bearing column advance by the reservation delta', () => {
    const render = (ruby: false | { hpsRaisePt?: number }) => {
      const { ctx, glyphs } = makeMatrixCtx();
      const baseSizePt = 15;
      const rubySizePt = 7.5;
      const baseRun: ShapeTextRun = {
        text: '漢', fontSizePt: baseSizePt, fontFamily: 'NotInMetrics',
        ...(ruby ? { ruby: { text: 'かん', fontSizePt: rubySizePt, ...ruby } } : {}),
      };
      const rubyBlock: ShapeText = {
        text: '漢', fontSizePt: baseSizePt, alignment: 'left',
        runs: [baseRun],
      };
      const nextBlock: ShapeText = {
        text: CJK, fontSizePt: baseSizePt, alignment: 'left',
        runs: [{ text: CJK, fontSizePt: baseSizePt, fontFamily: 'NotInMetrics' }],
      };
      const shape = richTextbox([baseRun], 'mongolianVert');
      (shape as unknown as { textBlocks: ShapeText[] }).textBlocks = [rubyBlock, nextBlock];
      acquireAndPaintShapeTextBox(shape, 0, 0, 200, 100, ctx, 1, {});
      const base = glyphs.find((g) => g.text.includes('漢'));
      const next = glyphs.find((g) => g.text.includes(CJK));
      const rubyGlyphs = glyphs.filter((g) => /[かん]/.test(g.text));
      expect(base, 'base glyph drawn').toBeDefined();
      expect(next, 'next-column glyph drawn').toBeDefined();
      if (ruby) expect(rubyGlyphs.length, 'ruby glyphs drawn').toBe(2);
      return {
        columnAdvance: next!.devX - base!.devX,
        rubyCrossOffset: rubyGlyphs.length > 0
          ? Math.max(...rubyGlyphs.map((g) => g.devX)) - base!.devX
          : NaN,
      };
    };

    const plain = render(false);
    const fallback = render({});
    const zero = render({ hpsRaisePt: 0 });
    const raised = render({ hpsRaisePt: 15 });
    expect(
      fallback.columnAdvance - plain.columnAdvance,
      'absent hpsRaise reserves selected-face base ascent + guide descent',
    ).toBeCloseTo(13.5, 5);
    expect(
      raised.columnAdvance - fallback.columnAdvance,
      'the authored 15pt raise grows from the 13.5pt ink-touching fallback',
    ).toBeCloseTo(1.5, 5);
    expect(zero.columnAdvance, 'zero hpsRaise adds no column reservation')
      .toBeCloseTo(plain.columnAdvance, 5);
    expect(zero.rubyCrossOffset, 'zero hpsRaise draws ruby on the base centerline').toBeCloseTo(0, 5);
    expect(
      raised.rubyCrossOffset,
      'authored hpsRaise remains the exact cross-axis displacement',
    ).toBeCloseTo(15, 5);
    expect(fallback.rubyCrossOffset, 'fallback touches retained base and guide ink')
      .toBeCloseTo(13.5, 5);
  });

  it('mongolianVert without hpsRaise keeps the selected-face ink reservation', () => {
    const columnAdvance = (withRuby: boolean): number => {
      const { ctx, glyphs } = makeMatrixCtx();
      const baseRun: ShapeTextRun = {
        text: '漢', fontSizePt: 15, fontFamily: 'NotInMetrics',
        ...(withRuby ? { ruby: { text: 'かん', fontSizePt: 7.5 } } : {}),
      };
      const rubyBlock: ShapeText = {
        text: '漢', fontSizePt: 15, alignment: 'left',
        runs: [baseRun],
      };
      const nextBlock: ShapeText = {
        text: CJK, fontSizePt: 15, alignment: 'left',
        runs: [{ text: CJK, fontSizePt: 15, fontFamily: 'NotInMetrics' }],
      };
      const shape = richTextbox([baseRun], 'mongolianVert');
      (shape as unknown as { textBlocks: ShapeText[] }).textBlocks = [rubyBlock, nextBlock];
      acquireAndPaintShapeTextBox(shape, 0, 0, 200, 100, ctx, 1, {});
      const base = glyphs.find((g) => g.text.includes('漢'));
      const next = glyphs.find((g) => g.text.includes(CJK));
      expect(base, 'base glyph drawn').toBeDefined();
      expect(next, 'next-column glyph drawn').toBeDefined();
      return next!.devX - base!.devX;
    };

    expect(columnAdvance(true) - columnAdvance(false)).toBeCloseTo(13.5, 5);
  });

  // Autofit (spAutoFit) shares the same line resolver: a ruby-bearing vertical
  // line must widen the fitted cross extent by the same reservation, or the
  // grown box would clip the furigana (sample-53 box (d) class).
  it('eaVert spAutoFit measurement grows by the ruby reservation', () => {
    const measure = (withRuby: boolean): number => {
      const { ctx } = makeMatrixCtx();
      const shape = richTextbox([{
        text: '漢', fontSizePt: 10, fontFamily: 'NotInMetrics',
        ...(withRuby ? { ruby: { text: 'かん', fontSizePt: 5 } } : {}),
      }], 'eaVert');
      shape.textAutofit = 'sp';
      return acquireShapeTextBoxForTest(shape, 0, 0, 200, 200, ctx, 1)
        ?.flowBounds.widthPt ?? 0;
    };
    expect(measure(true) - measure(false), 'fitted extent grows by retained base/guide ink').toBeCloseTo(9, 5);
  });

  // GT (Word PDF): §21.1.2.1.1 names the insets by PHYSICAL bounding-box edge
  // (lIns = left inset, default 91440 EMU = 7.2pt; bIns = bottom, 45720 EMU) —
  // reflecting the column stack for mongolianVert's L→R flow must not hand the
  // physical LEFT edge to bIns.
  it('mongolianVert first column origin honors lIns (physical left), not bIns (§21.1.2.1.1)', () => {
    const firstColX = (lIns: number, bIns: number): number => {
      const { ctx, glyphs } = makeMatrixCtx();
      const shape = richTextbox([{ text: CJK, fontSizePt: 10, fontFamily: 'NotInMetrics' }], 'mongolianVert');
      const s = shape as unknown as {
        textInsetL: number; textInsetT: number; textInsetR: number; textInsetB: number;
      };
      s.textInsetL = lIns;
      s.textInsetT = 3;
      s.textInsetR = 7;
      s.textInsetB = bIns;
      acquireAndPaintShapeTextBox(shape, 0, 0, 200, 100, ctx, 1, {});
      const g = glyphs.find((gl) => gl.text.includes(CJK));
      expect(g, 'glyph drawn').toBeDefined();
      return g!.devX;
    };
    // +12pt of lIns moves the first (leftmost) column +12pt right.
    expect(firstColX(17, 3) - firstColX(5, 3), 'lIns owns the physical-left origin').toBeCloseTo(12, 5);
    // bIns must NOT move the physical-left first column.
    expect(firstColX(7, 20), 'bIns does not own the physical-left origin').toBeCloseTo(firstColX(7, 3), 5);
  });
});

// Compare cached and unavailable OMML geometry with the horizontal host rule.
it.each(['cached', 'no engine', 'conversion', 'rasterization'] as const)('%s preserves horizontal math geometry and diagnostics in both stacked modes', async (failure) => {
  const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
  const error = vi.spyOn(console, 'error').mockImplementation(() => {});
  if (failure === 'rasterization') vi.mocked(rasterizeMathSvg).mockRejectedValue(new Error('rasterization failed'));
  try {
    for (const display of [false, true]) {
      const outcomes = [];
      let horizontalGeometry: { advance: number; image?: ImageCall } | undefined;
      for (const textVert of ['horz', 'wordArtVert', 'wordArtVertRtl']) {
        warn.mockClear(); error.mockClear();
        const { ctx, glyphs, images } = makeMatrixCtx();
        const text = (value: string) => ({ type: 'text', text: value, fontSize: 10, fontFamily: 'NotInMetrics',
          bold: false, italic: false, underline: false, strikethrough: false });
        const paraProps = { alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
          spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [] };
        const shape = { ...richTextbox([], textVert), widthPt: 200, heightPt: 100,
          textBoxContent: [{ type: 'paragraph', ...paraProps, runs: [text('A'),
            { type: 'math', nodes: [{ kind: 'run', text: 'x+y', style: 'italic' }], display, fontSize: 10 }, text('B')] }] };
        const normalized = normalizeInternalDocumentModel({
          section: { pageWidth: 300, pageHeight: 200, marginTop: 0, marginRight: 0, marginBottom: 0,
            marginLeft: 0, headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false },
          body: [{ type: 'paragraph', ...paraProps, runs: [shape] }],
          headers: { default: null, first: null, even: null }, footers: { default: null, first: null, even: null },
          fontFamilyClasses: {},
        } as unknown as import('./types.js').DocxDocumentModel);
        const prepared = failure === 'no engine' ? undefined : await prepareMathResources(normalized.mathOccurrences, {
          loadMathJax: async () => {},
          mathMLToSvg: async () => {
            if (failure === 'conversion') throw new Error('conversion failed');
            return { svg: '<svg/>', widthEm: 3, ascentEm: 1.5, descentEm: .5 };
          },
        });
        const services = createLayoutServices(normalized.document, { measureContext: ctx,
          mathResources: prepared?.records, mathDrawables: prepared?.drawables });
        const layout = layoutDocument(normalized.document, services, { currentDateMs: 0 });
        const paragraph = layout.pages[0].layers.body.find((node) => node.kind === 'paragraph');
        if (!paragraph || paragraph.kind !== 'paragraph') throw new Error('expected body paragraph');
        const registry = createPaintResourceRegistry(normalized.mathOccurrences.map(({ resourceKey }) => ({ kind: 'math', resourceKey })));
        const session = createPaintResourceSession(registry, normalized.mathOccurrences.map(({ resourceKey }) => ({
          kind: 'math', resourceKey, handle: unavailablePaintResourceHandle('optional or failed math'),
        })));
        paintTextBoxLayout(paragraph.textBoxes[0]!, { ctx, scale: 1, dpr: 1, resources:
          failure === 'cached' ? { paint(_key, _kind, bounds, target) {
            target.drawImage({} as CanvasImageSource, bounds.xPt, bounds.yPt, bounds.widthPt, bounds.heightPt);
          } } : createCanvasPaintResourcePainter(session, canonicalCanvasPaintResourceHandlers) });
        const diagnostics = services.math.resolve(normalized.mathOccurrences[0].resourceKey).diagnostics;
        const outcome = { text: glyphs.map((g) => g.text).join(''), images: images.length,
          diagnostics, warnings: [...warn.mock.calls], errors: [...error.mock.calls] };
        expect(outcome).toEqual({ text: 'AB', images: failure === 'cached' ? 1 : 0, diagnostics: failure === 'cached' ? [] : [{ code: 'UNSUPPORTED_FEATURE', severity: 'warning',
          message: failure === 'no engine'
            ? 'The optional math renderer is unavailable; using the deterministic text fallback'
            : 'Math conversion failed; using the deterministic text fallback' }], warnings: [], errors: [] });
        const [a, b] = glyphs;
        if (textVert === 'horz') {
          horizontalGeometry = { advance: b.devX - a.devX, image: images[0] };
          if (failure === 'cached') {
            expect(images[0]).toMatchObject({ w: 30, h: 20 });
            expect(images[0].devX - a.devX).toBeCloseTo(10);
          }
          // A non-square fallback detects confusing horizontal advance with
          // height: three 10 pt ems still reserve their inline extent.
          const inlineAdvance = failure === 'cached' ? 30 : 3 * 10;
          expect(b.devX - a.devX).toBeCloseTo(10 + inlineAdvance);
          expect(b.devY).toBeCloseTo(a.devY);
        } else {
          const inlineAdvance = failure === 'cached' ? 10 + horizontalGeometry!.image!.h
            : horizontalGeometry!.advance;
          expect(b.devY - a.devY).toBeCloseTo(inlineAdvance);
          expect(b.devX).toBeCloseTo(a.devX);
          if (failure === 'cached') {
            expect(images[0]).toMatchObject({ w: horizontalGeometry!.image!.w, h: horizontalGeometry!.image!.h });
            expect(images[0].devY - a.devY).toBeCloseTo(10);
            expect(norm(images[0].angleDeg)).toBeCloseTo(0);
          } else {
            expect(b.devY - a.devY).toBeCloseTo(horizontalGeometry!.advance);
          }
        }
        outcomes.push(outcome);
      }
      expect(outcomes.slice(1)).toEqual([outcomes[0], outcomes[0]]);
    }
  } finally {
    vi.mocked(rasterizeMathSvg).mockResolvedValue({ source: {} as CanvasImageSource, widthPx: 3, heightPx: 2 });
  }
});
