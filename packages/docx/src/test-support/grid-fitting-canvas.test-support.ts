/**
 * Canvas double for a host that grid-fits small sizes the way Chromium on
 * Windows measures a face with embedded bitmap strikes (core
 * reference-advances.ts): below {@link STRIKE_LIMIT_PX} every half-width cell
 * is rounded to a whole pixel and a full-width glyph spans two of them, so
 * U+3000 at 14.667px measures 16 and at 13.333px 14; at and above the limit
 * advances are linear (half-width 0.5 em, full-width 1 em). `fillText` is
 * shaped from the same advances, as on that host.
 *
 * The double keeps a real affine transform and the font/letter-spacing drawing
 * state across save/restore, and records every `fillText` with the state in
 * effect, so a test can reconstruct where each glyph lands on the page.
 */
export const STRIKE_LIMIT_PX = 64;

const FULL_WIDTH = /[ᄀ-ᇿ⺀-鿿가-힯豈-﫿！-｠￠-￦]/u;

/** Advance of one character at `px` on the grid-fitting host. */
export function gridFittedAdvancePx(character: string, px: number): number {
  const cells = FULL_WIDTH.test(character) ? 2 : 1;
  const half = px < STRIKE_LIMIT_PX ? Math.floor(px / 2 + 0.5) : px / 2;
  return cells * half;
}

/** Linear (design) advance of one character at `px`. */
export function linearAdvancePx(character: string, px: number): number {
  return (FULL_WIDTH.test(character) ? 1 : 0.5) * px;
}

export interface Matrix {
  readonly a: number;
  readonly b: number;
  readonly c: number;
  readonly d: number;
  readonly e: number;
  readonly f: number;
}

export interface PaintedText {
  readonly text: string;
  readonly x: number;
  readonly y: number;
  readonly font: string;
  readonly letterSpacing: string;
  readonly matrix: Matrix;
}

function fontPx(font: string): number {
  const match = /(?:^|\s)((?:\d+(?:\.\d*)?|\.\d+))px/.exec(font);
  return match ? Number(match[1]) : 10;
}

function spacingPx(spacing: string): number {
  const match = /^(-?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?)px$/.exec(spacing);
  return match ? Number(match[1]) : 0;
}

/** Width the host reports for `text` (Canvas adds letter spacing after
 * every character). */
export function gridFittedWidth(text: string, font: string, letterSpacing = '0px'): number {
  const px = fontPx(font);
  const characters = [...text];
  return characters.reduce((sum, character) => sum + gridFittedAdvancePx(character, px), 0)
    + characters.length * spacingPx(letterSpacing);
}

export function makeGridFittingCanvas(): {
  canvas: HTMLCanvasElement;
  ctx: CanvasRenderingContext2D;
  fills: PaintedText[];
  measured: string[];
} {
  let state = {
    font: '10px serif',
    letterSpacing: '0px',
    fontKerning: 'auto' as CanvasFontKerning,
    textAlign: 'start' as CanvasTextAlign,
    textBaseline: 'alphabetic' as CanvasTextBaseline,
    direction: 'ltr' as CanvasDirection,
    matrix: { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 } as Matrix,
  };
  const stack: (typeof state)[] = [];
  const fills: PaintedText[] = [];
  const measured: string[] = [];
  const multiply = (m: Matrix, n: Matrix): Matrix => ({
    a: m.a * n.a + m.c * n.b,
    b: m.b * n.a + m.d * n.b,
    c: m.a * n.c + m.c * n.d,
    d: m.b * n.c + m.d * n.d,
    e: m.a * n.e + m.c * n.f + m.e,
    f: m.b * n.e + m.d * n.f + m.f,
  });
  const apply = (n: Matrix) => { state = { ...state, matrix: multiply(state.matrix, n) }; };
  const ctx = {
    get font() { return state.font; },
    set font(value: string) { state = { ...state, font: value }; },
    get letterSpacing() { return state.letterSpacing; },
    set letterSpacing(value: string) { state = { ...state, letterSpacing: value }; },
    get fontKerning() { return state.fontKerning; },
    set fontKerning(value: CanvasFontKerning) { state = { ...state, fontKerning: value }; },
    get textAlign() { return state.textAlign; },
    set textAlign(value: CanvasTextAlign) { state = { ...state, textAlign: value }; },
    get textBaseline() { return state.textBaseline; },
    set textBaseline(value: CanvasTextBaseline) { state = { ...state, textBaseline: value }; },
    get direction() { return state.direction; },
    set direction(value: CanvasDirection) { state = { ...state, direction: value }; },
    measureText(text: string): TextMetrics {
      measured.push(text);
      const px = fontPx(state.font);
      const width = gridFittedWidth(text, state.font, state.letterSpacing);
      return {
        width,
        actualBoundingBoxLeft: 0,
        actualBoundingBoxRight: width,
        actualBoundingBoxAscent: px * 0.8,
        actualBoundingBoxDescent: px * 0.2,
        fontBoundingBoxAscent: px * 0.8,
        fontBoundingBoxDescent: px * 0.2,
      } as TextMetrics;
    },
    save() { stack.push(state); },
    restore() { state = stack.pop() ?? state; },
    translate(x: number, y: number) { apply({ a: 1, b: 0, c: 0, d: 1, e: x, f: y }); },
    scale(x: number, y: number) { apply({ a: x, b: 0, c: 0, d: y, e: 0, f: 0 }); },
    rotate(angle: number) {
      const cos = Math.cos(angle), sin = Math.sin(angle);
      apply({ a: cos, b: sin, c: -sin, d: cos, e: 0, f: 0 });
    },
    transform(a: number, b: number, c: number, d: number, e: number, f: number) { apply({ a, b, c, d, e, f }); },
    setTransform(a = 1, b = 0, c = 0, d = 1, e = 0, f = 0) { state = { ...state, matrix: { a, b, c, d, e, f } }; },
    getTransform() { return { ...state.matrix }; },
    fillText(text: string, x: number, y: number) {
      fills.push({ text, x, y, font: state.font, letterSpacing: state.letterSpacing, matrix: state.matrix });
    },
    strokeText() {},
    beginPath() {}, closePath() {}, moveTo() {}, lineTo() {}, stroke() {}, fill() {},
    fillRect() {}, strokeRect() {}, clearRect() {}, clip() {}, rect() {}, setLineDash() {},
    drawImage() {}, arc() {}, quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1, globalAlpha: 1,
    lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return {
    canvas: canvas as unknown as HTMLCanvasElement,
    ctx: ctx as unknown as CanvasRenderingContext2D,
    fills,
    measured,
  };
}
