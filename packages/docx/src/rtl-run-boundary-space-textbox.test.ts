import { describe, expect, it } from 'vitest';
import { acquireAndPaintShapeTextBox } from './retained-shape-textbox.test-support.js';
import type { ShapeRun, ShapeText, ShapeTextRun } from './types.js';
import { shapeRunToDocRun } from './layout/text.js';

// ─────────────────────────────────────────────────────────────────────────────
// ECMA-376 §17.3.1.6 `<w:bidi>` — inter-word space at a run/segment boundary in
// Word shape/text-box text (issue #929, text-box mirror of PR #949).
//
// PR #949 fixed the BODY paragraph draw loop, but explicitly scoped out the
// retained text-box visual-order projection. That path
// still drew an RTL segment's WHOLE logical string, including its trailing
// inter-word space, with one `fillText`. Chrome moves that edge whitespace to
// the segment's physical LEFT under `ctx.direction='rtl'`; skia-canvas does not,
// leaving the space on the physical RIGHT and collapsing the gap to the next
// reading word on the left.
//
// These tests acquire and paint retained text-box geometry with fixed-width mock glyphs.
// Since shape text has no `onTextRun` callback, all geometry comes from recorded
// `fillText` device positions. The recording context tracks translate/scale and
// honors `letterSpacing` in `measureText`, matching the BODY regression test's
// backend-independent measurement model.
// ─────────────────────────────────────────────────────────────────────────────

interface FillCall {
  text: string;
  /** Device x after the active translate/scale transform. */
  x: number;
  direction: CanvasDirection;
  letterSpacingPx: number;
}

function makeRecordingCanvas(): { ctx: CanvasRenderingContext2D; fills: FillCall[] } {
  let font = '10px serif';
  let letterSpacing = '0px';
  let direction: CanvasDirection = 'ltr';
  let tx = 0;
  let sx = 1;
  const stack: { tx: number; sx: number }[] = [];
  const fills: FillCall[] = [];
  const ctx = {
    get font() { return font; },
    set font(value: string) { font = value; },
    get letterSpacing() { return letterSpacing; },
    set letterSpacing(value: string) { letterSpacing = value; },
    get direction() { return direction; },
    set direction(value: CanvasDirection) { direction = value; },
    fontKerning: 'auto',
    measureText(text: string) {
      const px = Number(/([\d.]+)px/.exec(font)?.[1] ?? 10);
      const spacing = Number.parseFloat(letterSpacing) || 0;
      const cps = [...text].length;
      return {
        width: cps * px + cps * spacing,
        fontBoundingBoxAscent: px * 0.8,
        fontBoundingBoxDescent: px * 0.2,
        actualBoundingBoxAscent: px * 0.8,
        actualBoundingBoxDescent: px * 0.2,
      } as TextMetrics;
    },
    save() { stack.push({ tx, sx }); },
    restore() {
      const frame = stack.pop();
      if (frame) ({ tx, sx } = frame);
    },
    scale(fx: number) { sx *= fx; },
    translate(dx: number) { tx += dx * sx; },
    fillText(text: string, x: number) {
      fills.push({
        text,
        x: tx + sx * x,
        direction,
        letterSpacingPx: Number.parseFloat(letterSpacing) || 0,
      });
    },
    beginPath() {}, closePath() {}, moveTo() {}, lineTo() {}, stroke() {}, fill() {},
    fillRect() {}, strokeRect() {}, clip() {}, rect() {}, setLineDash() {},
    drawImage() {}, clearRect() {}, arc() {}, quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    rotate() {}, strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, fills };
}

const SHAPE_X = 100;
const SHAPE_Y = 50;
const SHAPE_W = 200;
const SHAPE_H = 100;
const FONT_SIZE = 10;
const W1 = 'ابج';
const W2 = 'دهو';
const SPACE_WIDTH = FONT_SIZE;

function shapeWith(blocks: ShapeText[]): ShapeRun {
  return {
    type: 'shape',
    presetGeometry: 'rect',
    wrapMode: 'none',
    textAnchor: 't',
    textInsetL: 0,
    textInsetT: 0,
    textInsetR: 0,
    textInsetB: 0,
    textBlocks: blocks,
  } as unknown as ShapeRun;
}

function textRun(text: string): ShapeTextRun {
  return { text, fontSizePt: FONT_SIZE, fontFamily: 'serif' };
}

function block(text: string, runs: ShapeTextRun[], bidi: boolean): ShapeText {
  return {
    text,
    runs,
    fontSizePt: FONT_SIZE,
    fontFamily: 'serif',
    alignment: 'left',
    bidi,
  } as unknown as ShapeText;
}

function render(blocks: ShapeText[]): FillCall[] {
  const { ctx, fills } = makeRecordingCanvas();
  acquireAndPaintShapeTextBox(
    shapeWith(blocks),
    SHAPE_X,
    SHAPE_Y,
    SHAPE_W,
    SHAPE_H,
    ctx,
    1,
  );
  return fills;
}

function wordFill(fills: FillCall[], word: string): FillCall {
  const fill = fills.find((event) => event.text.trim() === word);
  expect(fill).toBeDefined();
  return fill as FillCall;
}

function expectFullRtlGap(fills: FillCall[]): void {
  const w1 = wordFill(fills, W1);
  const w2 = wordFill(fills, W2);

  expect(w1.direction).toBe('rtl');
  expect(w2.direction).toBe('rtl');
  expect(/\s$/u.test(w1.text)).toBe(false);

  const w2Right = w2.x + [...W2].length * FONT_SIZE;
  expect(w1.x - w2Right).toBeCloseTo(SPACE_WIDTH, 6);
}

describe('ECMA-376 §17.3.1.6 RTL inter-word space in shape text (issue #929 / PR #949 scope-out)', () => {
  it('keeps one full visual gap between adjacent RTL runs', () => {
    const fills = render([
      block(`${W1} ${W2}`, [textRun(`${W1} `), textRun(W2)], true),
    ]);

    expectFullRtlGap(fills);
  });

  it('keeps one full visual gap when the main line engine splits a single RTL run', () => {
    const fills = render([
      block(`${W1} ${W2}`, [textRun(`${W1} ${W2}`)], true),
    ]);

    expectFullRtlGap(fills);
  });

  it('leaves the LTR trailing-space fast path byte-identical', () => {
    const fills = render([
      block('ABC DEF', [textRun('ABC DEF')], false),
    ]);
    const abc = wordFill(fills, 'ABC');

    expect(abc.text).toBe('ABC ');
    expect(abc.direction).toBe('ltr');
    expect(abc.x).toBeCloseTo(SHAPE_X, 6);
    expect(abc.letterSpacingPx).toBe(0);
  });
});

// Issue #2078 — §17.3.2.30 `<w:rtl>` on a text-box run (`ShapeTextRun.rtl`)
// reaches the same run-direction rule the body uses: `shapeRunToDocRun` maps it
// onto the body `DocRun.rtl`, so the shared segment builder / UAX#9 pass marks
// the run RTL even in an LTR paragraph, while unmarked runs stay unchanged.
describe('ECMA-376 §17.3.2.30 w:rtl on text-box runs (issue #2078)', () => {
  it('maps ShapeTextRun.rtl onto the body DocRun contract', () => {
    expect(shapeRunToDocRun({ ...textRun('A'), rtl: true }).rtl).toBe(true);
    expect(shapeRunToDocRun({ ...textRun('A'), rtl: false }).rtl).toBe(false);
    expect('rtl' in shapeRunToDocRun(textRun('A'))).toBe(false);
  });

  // §17.3.2.30: rtl-marked punctuation/symbols are RTL-strong (Word HL1), the
  // same rule `computeLineVisualOrder` applies to body runs; unmarked ones are not.
  it('draws rtl-marked punctuation right-to-left inside an LTR text-box paragraph', () => {
    const marked = render([
      block('AB !?', [textRun('AB '), { ...textRun('!?'), rtl: true }], false),
    ]);
    expect(wordFill(marked, 'AB').direction).toBe('ltr');
    expect(wordFill(marked, '!?').direction).toBe('rtl');

    const unmarked = render([block('AB !?', [textRun('AB '), textRun('!?')], false)]);
    expect(unmarked.every((fill) => fill.direction === 'ltr')).toBe(true);
  });

  it('reorders rtl-marked runs like the body UAX#9 pass', () => {
    const fills = render([
      block("!? #", [{ ...textRun('!?'), rtl: true }, textRun(' '), { ...textRun('#'), rtl: true }], false),
    ]);
    // Both marked runs are R; the space between them resolves to R (N1), so
    // the logical sequence `!?` ` ` `#` is one RTL run and `#` paints leftmost.
    expect(wordFill(fills, '#').x).toBeLessThan(wordFill(fills, '!?').x);
  });
});
