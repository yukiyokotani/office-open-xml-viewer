/**
 * Headless probe for the xlsx shape-text EMPTY-paragraph line-height fix.
 *
 * Background: in `renderer.ts` `drawShapeText`, the per-paragraph line builder
 * starts `lineHeight = 0` and only raises it from text/math runs
 * (`lineHeight = Math.max(lineHeight, pxSize * 1.2)`). An EMPTY paragraph (a
 * `<a:p>` with no run and no `<a:br>`) therefore flushed a line of `height: 0`,
 * so a blank line inside a textbox collapsed to nothing and pulled the
 * following paragraphs up. ECMA-376 §21.1.2.2.6 / §21.1.2.3 — the paragraph
 * mark still reserves one line at its run size. The fix seeds a blank line's
 * height from the paragraph's effective size (nearest text size → the
 * 11pt DEFAULT_FONT_SIZE) using xlsx's synthetic `1.2` leading (xlsx has no
 * real font metrics, unlike docx PR #582).
 *
 * Why this probe and not the VRT: a full scan of the xlsx sample corpus shows
 * NO textbox mixes a blank line with following text — every empty shape
 * paragraph lives in an entirely-empty textbox that draws no glyphs either way,
 * so the VRT cannot observe this fix (a coverage gap, like docx's list gap).
 * This probe injects the missing case directly.
 *
 * It MEASURES device pixels (not eyeballs): it renders a control textbox
 * [text, text] and a test textbox [text, EMPTY, text] through the real
 * `drawShapeText`, locates the two glyph bands in each, and asserts the gap
 * between the two text lines grows by exactly ONE line height when an empty
 * paragraph is interposed. Against the pre-fix renderer the empty line adds 0,
 * so the delta is ~0 and this test fails — a genuine regression guard, not a
 * tautology.
 *
 * CI-safe: skia-canvas ships a native binding CI omits, so the suite is gated
 * with `describe.skipIf(!skia)` exactly like render.test.ts / the border probe.
 */
import { describe, it, expect } from 'vitest';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';
import type { ShapeText, ShapeParagraph } from '@silurus/ooxml-xlsx';

const skia = await import('skia-canvas').catch(() => null);
type Skia = typeof import('skia-canvas');
const { Canvas } = (skia ?? {}) as Skia;

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, '../../..');
// Import the renderer's exported text typesetter directly by source path
// (renderer.ts has no static WASM import, so this needs no parser build).
const RENDERER_PATH = resolve(ROOT, 'packages/xlsx/src/renderer.ts');

// Mirror of core's PT_TO_PX (96/72) and renderer.ts's DEFAULT_FONT_SIZE / the
// synthetic 1.2 single-line leading. One blank line at 11pt is:
const PT_TO_PX = 96 / 72;
const DEFAULT_FONT_SIZE = 11;
const EXPECTED_BLANK_LINE_PX = DEFAULT_FONT_SIZE * PT_TO_PX * 1.2;

const W = 220;
const H = 160;

type DrawShapeText = (
  ctx: CanvasRenderingContext2D,
  txt: ShapeText,
  sw: number,
  sh: number,
  cs: number,
) => void;

function textPara(text: string): ShapeParagraph {
  return {
    align: 'l',
    runs: [
      { type: 'text', text, bold: false, italic: false, size: 11, color: '#000000' },
    ],
  };
}

function emptyPara(): ShapeParagraph {
  // No run, no break — the empty paragraph mark whose line collapsed pre-fix.
  return { align: 'l', runs: [] };
}

/** Render a top-anchored, non-wrapping textbox and return its RGBA pixels. */
async function render(paragraphs: ShapeParagraph[]): Promise<{
  data: Uint8ClampedArray;
  w: number;
  h: number;
}> {
  const { drawShapeText } = (await import(RENDERER_PATH)) as {
    drawShapeText: DrawShapeText;
  };
  const canvas = new Canvas(W, H);
  const ctx = canvas.getContext('2d') as unknown as CanvasRenderingContext2D;
  const txt: ShapeText = { anchor: 't', wrap: 'none', paragraphs };
  drawShapeText(ctx, txt, W, H, 1);
  const img = ctx.getImageData(0, 0, W, H);
  return { data: img.data, w: W, h: H };
}

/** Ink-weighted vertical centroids of each contiguous band of text rows. Only
 *  black glyphs are drawn (no fill/border), so painted alpha == ink. */
function bandCentroids(data: Uint8ClampedArray, w: number, h: number): number[] {
  const INK_ALPHA = 40;
  const MIN_ROW_INK = 3; // reject stray AA specks
  const rowInk: number[] = [];
  for (let y = 0; y < h; y++) {
    let n = 0;
    for (let x = 0; x < w; x++) {
      if (data[(y * w + x) * 4 + 3] > INK_ALPHA) n++;
    }
    rowInk.push(n >= MIN_ROW_INK ? n : 0);
  }
  const bands: number[] = [];
  let y = 0;
  while (y < h) {
    if (rowInk[y] === 0) {
      y++;
      continue;
    }
    let wsum = 0;
    let isum = 0;
    while (y < h && rowInk[y] > 0) {
      wsum += y * rowInk[y];
      isum += rowInk[y];
      y++;
    }
    bands.push(wsum / isum);
  }
  return bands;
}

describe.skipIf(!skia)('xlsx shape-text empty paragraph reserves a line', () => {
  it('an interposed empty paragraph pushes the next line down by one blank line height', async () => {
    const control = await render([textPara('Above'), textPara('Below')]);
    const test = await render([textPara('Above'), emptyPara(), textPara('Below')]);

    const controlBands = bandCentroids(control.data, control.w, control.h);
    const testBands = bandCentroids(test.data, test.w, test.h);

    // eslint-disable-next-line no-console
    console.log(
      `\n[PROBE] control bands=${controlBands.map((b) => b.toFixed(1)).join(', ')}` +
        `  test bands=${testBands.map((b) => b.toFixed(1)).join(', ')}` +
        `\n  expected blank-line px=${EXPECTED_BLANK_LINE_PX.toFixed(2)}`,
    );

    // Both render exactly two glyph bands ("Above" + "Below"); the empty
    // paragraph draws nothing, so it adds no band — only vertical space.
    expect(controlBands.length).toBe(2);
    expect(testBands.length).toBe(2);

    const controlGap = controlBands[1] - controlBands[0];
    const testGap = testBands[1] - testBands[0];
    const delta = testGap - controlGap;

    // eslint-disable-next-line no-console
    console.log(
      `  controlGap=${controlGap.toFixed(2)}  testGap=${testGap.toFixed(2)}  delta=${delta.toFixed(2)}`,
    );

    // "Above" sits at the same top in both, so the delta isolates exactly the
    // empty paragraph's reserved height — font-metric offsets cancel. With the
    // fix this equals one blank line (~17.6px); pre-fix the empty line added 0
    // (delta ~0), which fails this assertion.
    expect(Math.abs(delta - EXPECTED_BLANK_LINE_PX)).toBeLessThan(2);
  });
});
