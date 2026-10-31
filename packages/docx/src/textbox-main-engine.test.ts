import { describe, it, expect } from 'vitest';
import { acquireAndPaintShapeTextBox } from './retained-shape-textbox.test-support.js';
import type { ShapeRun, ShapeText, ShapeTextRun, TabStop } from './types';

// B2/B5 — text-box (shape) text is now laid out by the MAIN line-layout engine
// (buildSegments → layoutLines), so a text box gets the SAME kinsoku
// (§17.15.1.58–.60), UAX#9 bidi (§17.3.1.6), §17.18.44 justification and
// §17.3.1.37 tab stops the body does. The old simplified wrapper applied NONE of
// these. These are characterization tests over the draw pass: a recording canvas
// captures every fillText/measureText so the four features can be asserted from
// the drawn glyph positions, plus a guard that a plain LTR box (using none of the
// four) is unchanged.

interface FillTextEvent { text: string; x: number; y: number }

/** Recording 2D context. `measureText` width is code-point count × the current
 *  font's px (so widths are deterministic and independent of any real font).
 *  `letterSpacing` widens each measured advance, matching the browser, so the
 *  justify letterSpacing path is measured faithfully. */
function makeRecordingCanvas(): { ctx: CanvasRenderingContext2D; fillTexts: FillTextEvent[] } {
  let font = '10px serif';
  let letterSpacing = '0px';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ls = () => parseFloat(/(-?\d+(?:\.\d+)?)px/.exec(letterSpacing)?.[1] ?? '0');
  const fillTexts: FillTextEvent[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    get letterSpacing() { return letterSpacing; },
    set letterSpacing(v: string) { letterSpacing = v; },
    measureText: (s: string) => {
      const p = px();
      const n = [...s].length;
      // width = n glyphs × p, plus letterSpacing between/around glyphs (the
      // browser adds it after every glyph incl. the last).
      return {
        width: n * p + n * ls(),
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {}, rotate() {},
    setLineDash() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {},
    fillText(s: string, x: number, y: number) { fillTexts.push({ text: s, x, y }); },
    strokeText(s: string, x: number, y: number) { fillTexts.push({ text: s, x, y }); },
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, fillTexts };
}

const SHAPE_X = 100, SHAPE_Y = 50;
const INNER_X = SHAPE_X; // no insets
const SCALE = 1;

function shapeWith(blocks: ShapeText[], w: number, h: number): ShapeRun {
  return {
    type: 'shape',
    presetGeometry: 'rect', wrapMode: 'none', textAnchor: 't',
    textInsetL: 0, textInsetT: 0, textInsetR: 0, textInsetB: 0,
    textBlocks: blocks,
  } as unknown as ShapeRun;
}

function render(blocks: ShapeText[], w: number, h: number): FillTextEvent[] {
  const { ctx, fillTexts } = makeRecordingCanvas();
  acquireAndPaintShapeTextBox(shapeWith(blocks, w, h), SHAPE_X, SHAPE_Y, w, h, ctx, SCALE);
  return fillTexts;
}

/** Group fillText events into lines by their y (baseline), top to bottom. */
function lines(events: FillTextEvent[]): FillTextEvent[][] {
  const byY = new Map<number, FillTextEvent[]>();
  for (const e of events) {
    const arr = byY.get(e.y) ?? [];
    arr.push(e);
    byY.set(e.y, arr);
  }
  return [...byY.entries()].sort((a, b) => a[0] - b[0]).map(([, arr]) => arr);
}

/** The concatenated text drawn on a given line, in draw (visual) order. */
function lineText(lineEvents: FillTextEvent[]): string {
  return [...lineEvents].sort((a, b) => a.x - b.x).map((e) => e.text).join('');
}

function block(text: string, extra: Partial<ShapeText> = {}): ShapeText {
  return {
    text,
    fontSizePt: 10,
    fontFamily: 'serif',
    alignment: 'left',
    runs: [{ text, fontSizePt: 10, fontFamily: 'serif' } as ShapeTextRun],
    ...extra,
  } as unknown as ShapeText;
}

describe('text-box text on the main line engine (B2/B5)', () => {
  // (a) KINSOKU (§17.15.1.59 行頭禁則) — a line-start-forbidden mark (、 U+3001)
  // must never BEGIN a wrapped line. 10px/char, box inner width 60px ⇒ 6 CJK
  // cells/line. Because omitted w:overflowPunct defaults to true (§17.3.1.21),
  // the 7th glyph (、) hangs past the first line instead of opening line 2.
  it('(a) keeps a line-start-forbidden mark off the head of a wrapped line', () => {
    // "一二三四五六、七八九". Inner box 60px wide (shape width 60, no insets) ⇒ 6 CJK
    // cells fit per line. A naive greedy wrap fills line 1 with six glyphs
    // (一二三四五六) and pushes 、 to the head of line 2. The default-true
    // overflowPunct rule instead lets 、 extend one character beyond the line;
    // [MS-OE376] §2.1.56(b) says that rule wins when it conflicts with kinsoku.
    const evs = render([block('一二三四五六、七八九')], 60, 400);
    const ls = lines(evs);
    expect(ls.length).toBeGreaterThanOrEqual(2);
    // No wrapped line may START with the forbidden mark.
    for (let i = 1; i < ls.length; i++) {
      expect(lineText(ls[i]).startsWith('、')).toBe(false);
    }
    // Concretely the mark hangs from line 1 and line 2 starts with ordinary text.
    expect(lineText(ls[0]).endsWith('六、')).toBe(true);
    expect(lineText(ls[1]).startsWith('七')).toBe(true);
  });

  // (b) BIDI (§17.3.1.6 <w:bidi>) — the flag seeds the paragraph base direction,
  // which drives BOTH the UAX#9 segment reorder AND right-aligned line layout. The
  // old shape wrapper honoured neither. Two observable consequences:
  //   1. A pure-Latin RTL box lays its line from the RIGHT edge (LTR sits left).
  //   2. Mixed Hebrew+Latin reorders visually: the Hebrew (strong-RTL) run is
  //      drawn to the RIGHT of the Latin run (they swap relative to logical order).
  it('(b) right-aligns and UAX#9-reorders an RTL (bidi) paragraph', () => {
    const W = 2000;
    const ltr = render([block('AAA', { alignment: 'left' })], W, 400);
    const rtl = render([block('AAA', { alignment: 'left', bidi: true })], W, 400);
    const ltrX = ltr.find((e) => e.text === 'AAA')!.x;
    const rtlX = rtl.find((e) => e.text === 'AAA')!.x;
    // LTR sits at the inner-left; the RTL flag pushes the same word to the right
    // edge of the box (region right − word width).
    expect(ltrX).toBeCloseTo(INNER_X, 5);
    expect(rtlX).toBeGreaterThan(INNER_X + W / 2);
    expect(rtlX).toBeCloseTo(INNER_X + W - 3 * 10, 0); // "AAA" = 30px wide

    // Mixed Hebrew + Latin under an RTL base: the Hebrew run is visually to the
    // RIGHT of the Latin run (larger x), i.e. reading order runs right→left.
    const mixed = render([block('שלום AB', { bidi: true })], W, 400);
    const heb = mixed.find((e) => /[֐-׿]/.test(e.text))!;
    const lat = mixed.find((e) => e.text.includes('A'))!;
    expect(heb).toBeDefined();
    expect(lat).toBeDefined();
    expect(heb.x).toBeGreaterThan(lat.x);
  });

  // (c) JUSTIFY (§17.18.44 both) — a non-final justified line stretches its
  // inter-word gaps so its last glyph reaches the region's right edge. 10px/char,
  // inner width chosen so two words wrap (a 3rd forces line 1 to be non-final),
  // then line 1's right edge must sit at innerX + innerW.
  it('(c) stretches a justified (both) non-final line to the region edge', () => {
    // "AAAA BBBB CCCC" — inner width 110px. "AAAA " (5) + "BBBB " (5) = 100px fits;
    // "CCCC" (4×10=40) overflows ⇒ line 1 = "AAAA BBBB" (non-final), line 2 = "CCCC".
    const W = 110;
    const justified = render([block('AAAA BBBB CCCC', { alignment: 'both' })], W, 400);
    const ls = lines(justified);
    expect(ls.length).toBeGreaterThanOrEqual(2);
    const line1 = [...ls[0]].sort((a, b) => a.x - b.x);
    // Right edge of the last drawn token on line 1 = its x + its measured width.
    // Under justification it must reach the inner-right edge (INNER_X + W). The
    // last token "BBBB" is 40px wide.
    const last = line1[line1.length - 1];
    // The line-end space hangs past the edge (WORD_LINE_END_SPACE_HANGS).
    const lastRight = last.x + [...last.text.trimEnd()].length * 10; // letterSpacing 0 on this LTR path
    expect(lastRight).toBeCloseTo(INNER_X + W, 0);
    // The last line ("CCCC") is NOT stretched — it starts at the inner-left.
    expect(ls[1][0].x).toBeCloseTo(INNER_X, 5);
  });

  // (d) TAB STOP (§17.3.1.37) — content after a \t advances to the paragraph's
  // explicit tab stop. A left tab at 100pt places the post-tab text at
  // innerX + 100 (scale 1), regardless of the pre-tab label width.
  it('(d) advances post-tab text to an explicit tab stop', () => {
    const tabStops: TabStop[] = [{ pos: 100, alignment: 'left', leader: 'none' }];
    const events = render(
      [block('A\tB', { tabStops, runs: [{ text: 'A\tB', fontSizePt: 10, fontFamily: 'serif' } as ShapeTextRun] })],
      2000,
      400,
    );
    const a = events.find((e) => e.text === 'A');
    const b = events.find((e) => e.text === 'B');
    expect(a).toBeDefined();
    expect(b).toBeDefined();
    // "A" at the inner-left; "B" at the 100pt tab stop (innerX + 100).
    expect((a as FillTextEvent).x).toBeCloseTo(INNER_X, 5);
    expect((b as FillTextEvent).x).toBeCloseTo(INNER_X + 100, 5);
  });

  // (e) GUARD — a plain LTR box that uses NONE of the four features wraps Latin
  // words at spaces exactly as a greedy wrapper would, and the continuation line
  // sits at the inner-left. This pins that the engine swap did not disturb the
  // ordinary path.
  it('(e) wraps a plain LTR paragraph at spaces with an unshifted continuation line', () => {
    // 10px/char, inner width 60px ⇒ 6 chars/line. "aa bb cc dd" → "aa bb " (6)
    // then "cc dd" wraps.
    const evs = render([block('aa bb cc dd', { alignment: 'left' })], 60, 400);
    const ls = lines(evs);
    expect(ls.length).toBeGreaterThanOrEqual(2);
    // Every line starts flush at the inner-left (no indent, no justify shift).
    for (const ln of ls) {
      const leftmost = [...ln].sort((a, b) => a.x - b.x)[0];
      expect(leftmost.x).toBeCloseTo(INNER_X, 5);
    }
    // The wrapped word order is preserved top-to-bottom.
    expect(lineText(ls[0]).replace(/\s+$/, '')).toBe('aa bb');
    expect(lineText(ls[1]).replace(/\s+$/, '')).toBe('cc dd');
  });
});
