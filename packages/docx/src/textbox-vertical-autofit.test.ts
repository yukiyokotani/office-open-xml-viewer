import { describe, it, expect } from 'vitest';
import { acquireShapeTextBoxForTest } from './retained-shape-textbox.test-support.js';
import type { ShapeRun, ShapeText, ShapeTextRun } from './types';

// ECMA-376 §20.1.10.83 (`<wps:bodyPr vert>`) + §21.1.2.1.1 (`<a:spAutoFit/>`) —
// cross-axis auto-grow/shrink for a VERTICAL text box (issue #1000, final
// remainder of #988).
//
// Word ground truth (the autofit adjudication fixture, measured from the Word
// PDF at 300 dpi):
//   • OVERFLOW  eaVert spAutoFit box, authored 0.70in wide → rendered 2.070in
//     (GREW +1.37in) with the authored LEFT edge fixed (grew RIGHTWARD).
//   • UNDERFLOW eaVert spAutoFit box, authored 2.40in wide → rendered 0.953in
//     (SHRANK −1.45in) with the authored LEFT edge fixed (shrank from the RIGHT).
//   • CONTROL   identical content, `noAutofit` → stayed at authored 0.70in and
//     CLIPPED the overflow.
// So a vertical spAutoFit box resizes its CROSS axis (physical WIDTH = the
// column-stacking axis after the ±90° rotate-layout) to the content column
// stack, keeping the authored TOP-LEFT anchor (`off.x` = LEFT edge) fixed and
// moving the far (RIGHT) edge — grow on overflow, shrink on underflow. This is
// the same top-left bbox anchor the horizontal branch keeps (it fixes the TOP
// edge and moves the BOTTOM).

/** Minimal measuring 2D context for retained acquisition. Each code point
 * advances by the current font px; symmetric-ish font box (0.8 / 0.2 em). */
function makeMeasureCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const stack: string[] = [];
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    textAlign: 'start',
    textBaseline: 'alphabetic',
    direction: 'ltr',
    fontKerning: 'auto',
    save() { stack.push(font); },
    restore() { const s = stack.pop(); if (s !== undefined) font = s; },
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
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

const run = (text: string): ShapeTextRun => ({ text, fontSizePt: 10, fontFamily: 'NotInMetrics' });

/** A single-paragraph text box. Insets 0 keep the arithmetic clean: the vertical
 *  wrap width (column length) equals the physical HEIGHT and the returned cross
 *  extent equals the physical WIDTH. */
function verticalTextbox(
  text: string,
  textVert: string | null,
  autofit: string | null,
): ShapeRun {
  const block: ShapeText = { text, fontSizePt: 10, alignment: 'left', runs: [run(text)] };
  return {
    type: 'shape', zOrder: 0, subpaths: [], presetGeometry: 'rect', fill: null, stroke: null,
    textBlocks: [block], textAnchor: 't',
    textInsetL: 0, textInsetT: 0, textInsetR: 0, textInsetB: 0,
    textVert,
    textAutofit: autofit,
  } as unknown as ShapeRun;
}

const CJK = '経';

function retainedAutofitBox(
  shape: ShapeRun,
  box: { x: number; y: number; w: number; h: number },
  ctx: CanvasRenderingContext2D,
): { x: number; y: number; w: number; h: number } {
  const layout = acquireShapeTextBoxForTest(
    shape, box.x, box.y, box.w, box.h, ctx, 1,
  );
  return layout ? {
    x: layout.flowBounds.xPt,
    y: layout.flowBounds.yPt,
    w: layout.flowBounds.widthPt,
    h: layout.flowBounds.heightPt,
  } : box;
}

describe('§20.1.10.83 + §21.1.2.1.1 — vertical spAutoFit cross-axis grow/shrink (#1000)', () => {
  it('GROW: an overflowing eaVert spAutoFit box widens, LEFT edge fixed, RIGHT edge moves', () => {
    const ctx = makeMeasureCtx();
    // Box physical 30w × 100h, insets 0 → logical column length 100 → 10 CJK per
    // column. 60 CJK ⇒ ~6 columns ⇒ required cross width ≫ authored 30.
    const shape = verticalTextbox(CJK.repeat(60), 'eaVert', 'sp');
    const box = { x: 100, y: 50, w: 30, h: 100 };
    const fit = retainedAutofitBox(shape, box, ctx);
    expect(fit.w, `grew past authored 30 (fit.w=${fit.w})`).toBeGreaterThan(45);
    expect(fit.x, 'authored LEFT edge fixed').toBe(100);
    expect(fit.y, 'top unchanged').toBe(50);
    expect(fit.h, 'flow axis (column length = physical height) unchanged').toBe(100);
    // Right edge moved outward; left edge did not.
    expect(fit.x + fit.w, 'RIGHT edge moved right').toBeGreaterThan(box.x + box.w);
  });

  it('SHRINK: an underflowing eaVert spAutoFit box narrows, LEFT edge fixed, RIGHT edge moves in', () => {
    const ctx = makeMeasureCtx();
    // Box physical 200w × 100h. 5 CJK ⇒ 1 column ⇒ required cross width ≪ 200.
    const shape = verticalTextbox(CJK.repeat(5), 'eaVert', 'sp');
    const box = { x: 100, y: 50, w: 200, h: 100 };
    const fit = retainedAutofitBox(shape, box, ctx);
    expect(fit.w, `shrank below authored 200 (fit.w=${fit.w})`).toBeLessThan(40);
    expect(fit.w, 'still positive').toBeGreaterThan(0);
    expect(fit.x, 'authored LEFT edge fixed').toBe(100);
    expect(fit.h, 'flow axis unchanged').toBe(100);
    expect(fit.x + fit.w, 'RIGHT edge moved inward').toBeLessThan(box.x + box.w);
  });

  it('noAutofit vertical box keeps its authored extent (clips instead of growing)', () => {
    const ctx = makeMeasureCtx();
    const shape = verticalTextbox(CJK.repeat(60), 'eaVert', 'none');
    const box = { x: 100, y: 50, w: 30, h: 100 };
    const fit = retainedAutofitBox(shape, box, ctx);
    expect(fit).toEqual(box);
  });

  it('all four vert modes grow the cross (width) axis, never the flow (height) axis', () => {
    const ctx = makeMeasureCtx();
    for (const mode of ['vert', 'vert270', 'eaVert', 'mongolianVert']) {
      const shape = verticalTextbox(CJK.repeat(60), mode, 'sp');
      const box = { x: 0, y: 0, w: 30, h: 100 };
      const fit = retainedAutofitBox(shape, box, ctx);
      expect(fit.w, `${mode} grew width`).toBeGreaterThan(45);
      expect(fit.h, `${mode} kept height`).toBe(100);
    }
  });

  it('HORIZONTAL spAutoFit still grows the HEIGHT (flow axis), width untouched — regression', () => {
    const ctx = makeMeasureCtx();
    // Horizontal (vert absent): 60 CJK wraps within width 200 → several lines →
    // height grows well past 20; width stays 200.
    const shape = verticalTextbox(CJK.repeat(60), null, 'sp');
    const box = { x: 100, y: 50, w: 200, h: 20 };
    const fit = retainedAutofitBox(shape, box, ctx);
    expect(fit.w, 'width unchanged for a horizontal box').toBe(200);
    expect(fit.h, 'height grew to the line stack').toBeGreaterThan(20);
    expect(fit.x, 'x unchanged').toBe(100);
    expect(fit.y, 'top edge fixed').toBe(50);
  });

  it('a vertical box carrying an inline image keeps its authored extent (image cross extent not line-measurable)', () => {
    const ctx = makeMeasureCtx();
    const imgBlock: ShapeText = {
      text: '', fontSizePt: 10, alignment: 'left',
      imagePath: 'word/media/image1.png', imageWidthPt: 24, imageHeightPt: 36,
    } as unknown as ShapeText;
    const shape = verticalTextbox(CJK.repeat(60), 'eaVert', 'sp');
    (shape as unknown as { textBlocks: ShapeText[] }).textBlocks.push(imgBlock);
    const box = { x: 10, y: 10, w: 30, h: 100 };
    const fit = retainedAutofitBox(shape, box, ctx);
    expect(fit).toEqual(box);
  });

  it('inset-axis mapping: lIns/rIns add to the CROSS (width) extent (not the column length)', () => {
    // WORD_TEXT_BOX_PHYSICAL_INSETS (Windows Word 16.0.20527 vertical-inset
    // controls): insets stay on their physical edges, so the left/right insets
    // bound the column stack and a fitted box is columns + lIns + rIns wide.
    // Two boxes differing only by lIns/rIns differ by exactly 2·lr.
    const ctx = makeMeasureCtx();
    const mk = (lr: number): ShapeRun => {
      const block: ShapeText = { text: CJK.repeat(30), fontSizePt: 10, alignment: 'left', runs: [run(CJK.repeat(30))] };
      return {
        type: 'shape', zOrder: 0, subpaths: [], presetGeometry: 'rect', fill: null, stroke: null,
        textBlocks: [block], textAnchor: 't',
        textInsetL: lr, textInsetT: 0, textInsetR: lr, textInsetB: 0,
        textVert: 'eaVert', textAutofit: 'sp',
      } as unknown as ShapeRun;
    };
    const box = { x: 0, y: 0, w: 30, h: 100 };
    const w0 = retainedAutofitBox(mk(0), box, ctx).w;
    const w7 = retainedAutofitBox(mk(7), box, ctx).w;
    expect(w7 - w0, 'lIns+rIns=2·7 added to the cross width').toBeCloseTo(14, 5);
  });

  it('inset-axis mapping: tIns/bIns govern the COLUMN LENGTH (wrap) axis — more inset ⇒ shorter columns ⇒ wider box', () => {
    // tIns/bIns reduce the column length (= physical HEIGHT), so larger
    // top/bottom insets fit fewer chars per column ⇒ more columns ⇒ a WIDER
    // cross extent. This distinguishes the wrap axis from the cross axis.
    const ctx = makeMeasureCtx();
    const mk = (tb: number): ShapeRun => {
      const block: ShapeText = { text: CJK.repeat(20), fontSizePt: 10, alignment: 'left', runs: [run(CJK.repeat(20))] };
      return {
        type: 'shape', zOrder: 0, subpaths: [], presetGeometry: 'rect', fill: null, stroke: null,
        textBlocks: [block], textAnchor: 't',
        textInsetL: 0, textInsetT: tb, textInsetR: 0, textInsetB: tb,
        textVert: 'eaVert', textAutofit: 'sp',
      } as unknown as ShapeRun;
    };
    const box = { x: 0, y: 0, w: 30, h: 100 };
    const wNarrowInset = retainedAutofitBox(mk(0), box, ctx).w;
    const wWideInset = retainedAutofitBox(mk(25), box, ctx).w;
    // 20 chars of 10pt: 10 per 100pt column (2 columns) vs 5 per 50pt
    // column (4 columns), so the column stack doubles.
    expect(wWideInset, 'shorter columns from bigger tIns/bIns pack into more columns ⇒ wider')
      .toBeCloseTo(2 * wNarrowInset, 5);
  });

  it('a shape without spAutoFit or without text blocks is returned unchanged', () => {
    const ctx = makeMeasureCtx();
    const noText = { type: 'shape', zOrder: 0, subpaths: [], presetGeometry: 'rect', fill: null, stroke: null, textVert: 'eaVert', textAutofit: 'sp' } as unknown as ShapeRun;
    const box = { x: 1, y: 2, w: 3, h: 4 };
    expect(retainedAutofitBox(noText, box, ctx)).toEqual(box);
    const notSp = verticalTextbox(CJK.repeat(60), 'eaVert', null);
    expect(retainedAutofitBox(notSp, { ...box }, ctx)).toEqual(box);
  });
});
