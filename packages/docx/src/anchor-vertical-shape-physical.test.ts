import { describe, expect, it } from 'vitest';
import { __test_resolveAnchorBox, __test_resolveShapeBox } from './test-support/renderer-internals.test-support.js';
import type { AnchorFloatRegistrationState } from './layout/acquisition-context.js';
import { layoutBodyModel } from './test-support/document-layout.test-support.js';
import type { BodyElement, DocParagraph, ImageRun, SectionProps, ShapeRun } from './types.js';

// ECMA-376 §17.6.20 (tbRl) + §20.4.3.x — issue #988 batch-3 adjudication ②:
// `positionH`/`positionV` anchors in a VERTICAL section resolve against the
// PHYSICAL (un-rotated) page, on the physical vertical axis, increasing
// downward — independent of the section text direction. `paragraph`-relative
// vertical anchors resolve from the PHYSICAL TOP of the anchor paragraph's
// column (the top content margin for a single-column body), NOT from the
// paragraph's logical flow position.
//
// Ground truth = the batch-3 positionV fixture's Word PDF (Letter portrait
// 612×792 pt, 1 in margins, three identical anchored rectangles):
//
//   | relativeFrom | posOffset       | measured physical box top |
//   |--------------|-----------------|---------------------------|
//   | page         | 3.0 in (216 pt) | y = 216                   |
//   | margin       | 1.5 in (108 pt) | y = 72 + 108 = 180        |
//   | paragraph    | 0.3 in (21.6pt) | y = 72 + 21.6 = 93.6      |
//
// positionH is `page`-relative in all three (72 / 230.4 / 388.8 pt), and the
// physical x equals the raw offset. The shapes are 100.8 × 86.4 pt.

const PHYS_W = 612;
const PHYS_H = 792;
const MARGIN = 72;

// Vertical anchor capability: LOGICAL (swapped) geometry + `verticalPhys`. The
// logical frame of a portrait Letter tbRl page: logical width = physical
// height, margins rotated one quarter-turn (verticalLayoutSection). contentX
// is the current column band start in LOGICAL x — its physical image is the
// column's physical TOP (72 = the top content margin), which is what a
// paragraph-relative positionV anchors from.
const verticalState = {
  pageWidth: PHYS_H,
  marginLeft: MARGIN, // logical left = physical top
  marginRight: MARGIN, // logical right = physical bottom
  marginTop: MARGIN, // logical top = physical right
  marginBottom: MARGIN, // logical bottom = physical left
  pageH: PHYS_W, // logical height in points = physical width
  contentX: MARGIN, // column band start (logical x) ⇒ physical column top
  contentW: PHYS_H - 2 * MARGIN,
  verticalCJK: true,
  verticalPhys: {
    pageWidth: PHYS_W,
    pageHeight: PHYS_H,
    marginLeft: MARGIN,
    marginRight: MARGIN,
    marginTop: MARGIN,
    marginBottom: MARGIN,
    writingMode: 'vertical-rl',
  },
} as unknown as AnchorFloatRegistrationState;

const SHAPE_W = 100.8;
const SHAPE_H = 86.4;

function rectShape(
  xPt: number,
  yPt: number,
  yRelativeFrom: string,
): ShapeRun {
  return {
    widthPt: SHAPE_W,
    heightPt: SHAPE_H,
    anchorXPt: xPt,
    anchorYPt: yPt,
    anchorXFromMargin: false,
    anchorYFromPara: yRelativeFrom === 'paragraph',
    anchorXRelativeFrom: 'page',
    anchorYRelativeFrom: yRelativeFrom,
    zOrder: 0,
    subpaths: [],
    presetGeometry: 'rect',
    fill: null,
    stroke: null,
    wrapMode: 'none',
  } as unknown as ShapeRun;
}

/** Invert the logical projection back to the physical frame: the page paint
 *  transform is `physical = (cssW − logical.y, logical.x)`, so a logical box
 *  `{x, y, w, h}` images the physical box
 *  `{x: cssW − (y + h), y: x, w: h, h: w}`. */
function toPhysical(
  box: { x: number; y: number; w: number; h: number },
  cssW: number,
): { x: number; y: number; w: number; h: number } {
  return { x: cssW - (box.y + box.h), y: box.x, w: box.h, h: box.w };
}

// A deliberately WRONG logical flow-Y for the anchor paragraph: the physical
// resolution must anchor paragraph-relative offsets from the column's PHYSICAL
// top (state.contentX), never from this logical coordinate.
const LOGICAL_PARA_TOP = 500;

describe('resolveShapeBox — vertical (tbRl) physical anchor mapping (§20.4.3.x, #988 ②)', () => {
  it.each([
    ['page', 388.8, 216, 216],
    ['margin', 230.4, 108, MARGIN + 108],
    ['paragraph', 72, 21.6, MARGIN + 21.6],
  ] as const)(
    'positionV relativeFrom=%s resolves on the physical vertical axis',
    (rf, xPt, yPt, expectedPhysTop) => {
      const box = __test_resolveShapeBox(
        rectShape(xPt, yPt, rf),
        verticalState,
        LOGICAL_PARA_TOP,
      );
      const phys = toPhysical(box, PHYS_W);
      expect(phys.x).toBeCloseTo(xPt, 1);
      expect(phys.y).toBeCloseTo(expectedPhysTop, 1);
      expect(phys.w).toBeCloseTo(SHAPE_W, 4);
      expect(phys.h).toBeCloseTo(SHAPE_H, 4);
    },
  );

  it('returns the logical-projected box (w↔h swap) so the flow band matches', () => {
    const box = __test_resolveShapeBox(
      rectShape(388.8, 216, 'page'),
      verticalState,
      LOGICAL_PARA_TOP,
    );
    // logical x = physical y; logical y = cssW − (px + w); w/h swapped.
    expect(box.x).toBeCloseTo(216, 4);
    expect(box.y).toBeCloseTo(PHYS_W - (388.8 + SHAPE_W), 4);
    expect(box.w).toBeCloseTo(SHAPE_H, 4);
    expect(box.h).toBeCloseTo(SHAPE_W, 4);
  });
});

// ── Pagination/paint agreement (Codex review F1) ───────────────────────────
// The PAGINATOR's float registration must resolve a wrapped anchored shape
// through the SAME physical projection the paint pass uses — its measure state
// carries `verticalPhys` — otherwise a `topAndBottom` band is reserved at the
// raw logical rectangle and page assignment diverges from the painted layout.

function testParagraph(text: string, runs?: unknown[]): DocParagraph {
  return {
    alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: runs ?? [{
      type: 'text', text,
      bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: 'Times New Roman', fontFamilyEastAsia: 'Times New Roman',
      isLink: false, background: null, vertAlign: null, hyperlink: null,
    }],
    defaultFontSize: 10, defaultFontFamily: 'Times New Roman',
    widowControl: false,
  } as unknown as DocParagraph;
}

function makeMeasureCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx: Record<string, unknown> = {
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {}, moveTo() {}, lineTo() {},
    stroke() {}, fill() {}, fillRect() {}, strokeRect() {}, rect() {}, clip() {},
    scale() {}, translate() {}, rotate() {}, setTransform() {},
    setLineDash() {}, clearRect() {}, arc() {}, quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {}, fillText() {}, strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left', direction: 'ltr', globalAlpha: 1, lineCap: 'butt', lineJoin: 'miter',
  };
  Object.defineProperty(ctx, 'font', { get: () => font, set: (v: string) => { font = v; } });
  return ctx as unknown as CanvasRenderingContext2D;
}

describe('canonical body layout — vertical wrapped-shape physical projection (#988 ②, F1)', () => {
  it('a topAndBottom shape displaces the flow by its LOGICAL projection (physical width)', () => {
    // Physical portrait 200×300, margins T20 R30 B40 L24 ⇒ logical flow starts
    // at y=30 and ends at 30+146=176 (insets 30/24 on the 200 pt flow axis).
    const phys = {
      pageWidth: 200, pageHeight: 300,
      marginTop: 20, marginRight: 30, marginBottom: 40, marginLeft: 24,
      headerDistance: 0, footerDistance: 0,
      titlePage: false, evenAndOddHeaders: false,
      textDirection: 'tbRl',
    } as unknown as SectionProps;
    // Physical box (40, 40)–(155, 60): logical projection y ∈ [45, 160]
    // (cssW − px − w = 45), a 115 pt-tall topAndBottom band. The RAW logical
    // rectangle would be y ∈ [40, 60] — only 20 pt tall.
    const shape = {
      type: 'shape',
      widthPt: 115,
      heightPt: 20,
      anchorXPt: 40,
      anchorYPt: 40,
      anchorXFromMargin: false,
      anchorYFromPara: false,
      anchorXRelativeFrom: 'page',
      anchorYRelativeFrom: 'page',
      wrapMode: 'topAndBottom',
      zOrder: 0,
      subpaths: [],
      presetGeometry: 'rect',
      fill: null,
      stroke: null,
    };
    const body = [
      { type: 'paragraph', ...testParagraph('', [
        { type: 'text', text: 'zz', bold: false, italic: false, underline: false, strikethrough: false, fontSize: 10, color: null, fontFamily: 'Times New Roman', fontFamilyEastAsia: 'Times New Roman', isLink: false, background: null, vertAlign: null, hyperlink: null },
        shape,
      ]) },
      { type: 'paragraph', ...testParagraph('qq') },
      { type: 'paragraph', ...testParagraph('ww') },
    ] as unknown as BodyElement[];
    const layout = layoutBodyModel(
      body, phys, makeMeasureCtx(),
      { 'Times New Roman': 'roman' },
    );
    // With the PHYSICAL projection the band ends at logical y=160: the two
    // following one-line (~11.5 pt) paragraphs are pushed to ~160–171.5 and
    // ~171.5–183 — the second crosses the 176 flow bottom onto page 2. With
    // the raw logical band ([40, 60]) they sit at ~60–83 and fit page 1.
    expect(layout.pages.length).toBe(2);
  });
});

describe('resolveAnchorBox — paragraph-relative positionV under tbRl (§20.4.3.5, #988 ②)', () => {
  it('anchors from the physical top of the paragraph column, not the logical flow y', () => {
    const img: ImageRun = {
      imagePath: 'word/media/image1.png',
      mimeType: 'image/png',
      widthPt: SHAPE_W,
      heightPt: SHAPE_H,
      anchor: true,
      wrapMode: 'none',
      anchorXPt: 72,
      anchorYPt: 21.6,
      anchorXRelativeFrom: 'page',
      anchorYRelativeFrom: 'paragraph',
      anchorXFromMargin: false,
      anchorYFromPara: true,
    } as unknown as ImageRun;
    const box = __test_resolveAnchorBox(img, verticalState, LOGICAL_PARA_TOP);
    const phys = toPhysical(box, PHYS_W);
    expect(phys.x).toBeCloseTo(72, 1);
    expect(phys.y).toBeCloseTo(MARGIN + 21.6, 1); // 93.6 — Word GT
  });
});
