import { describe, it, expect } from 'vitest';
import { createLayoutServices } from './layout-runtime.js';
import { renderDocumentToCanvas } from './renderer.js';
import { testFontSnapshot } from './layout/test-font-snapshot.js';
import type {
  BodyElement,
  DocParagraph,
  DocxDocumentModel,
  DocxTextRun,
  LineSpacing,
  SectionProps,
  ShapeRun,
  ShapeText,
} from './types';

// ECMA-376 §17.3.1.33 `<w:spacing w:lineRule>` — baseline placement inside the
// line box.
//
//   auto  → MULTIPLE spacing: `w:line` is a 240ths-of-a-line multiplier. Word
//           pins the baseline at the NATURAL ascent from the box top and places
//           the multiplier's extra leading ENTIRELY BELOW the glyphs. It does
//           NOT centre the natural line in the enlarged box. (Measured against
//           Word's PDF export during the #981 / #990 follow-up: a 2.0× 48pt
//           title was displaced ~22pt when centred.)
//   exact / atLeast → the extra space (value − natural) stays split half above /
//           half below the glyphs (centred). Unchanged by #990.
//
// Draw-only: the line BOX height (lineBoxHeight, §17.3.1.33) is identical either
// way, so pagination is unaffected — this pins only WHERE the glyphs sit inside
// the box. The line pitch (baseline-to-baseline) is therefore invariant.

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Any = any;

// ---------- recording 2D context ----------
interface FillTextCall { text: string; x: number; y: number; font: string; }

function makeRecordingCanvas(): {
  canvas: HTMLCanvasElement;
  fillTextCalls: FillTextCall[];
} {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const fillTextCalls: FillTextCall[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p,
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {},
    setLineDash() {}, drawImage() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    fillText(text: string, x: number, y: number) {
      fillTextCalls.push({ text, x, y, font });
    },
    strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = {
    width: 0, height: 0,
    style: {} as Record<string, string>,
    getContext: () => ctx,
  };
  return { canvas: canvas as unknown as HTMLCanvasElement, fillTextCalls };
}

// A SYNTHETIC, untabled font: the mock canvas reports a clean 1.0 em box
// (ascent 0.8 / descent 0.2) for it, so these tests isolate the line-spacing
// MULTIPLIER from the substituted-font single-line FLOOR (intendedSingleLinePx).
const TEST_FONT = 'Synthetic Untabled Serif';

function textRun(text: string): DocxTextRun {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 10, color: null, fontFamily: TEST_FONT, fontFamilyEastAsia: '',
    isLink: false, background: null, vertAlign: null, hyperlink: null,
  } as DocxTextRun;
}

function paragraph(text: string, lineSpacing: LineSpacing | null): BodyElement {
  return {
    type: 'paragraph',
    alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing,
    numbering: null, tabStops: [],
    runs: [{ type: 'text', ...textRun(text) } as DocParagraph['runs'][number]],
    defaultFontSize: 10, defaultFontFamily: TEST_FONT,
    widowControl: false,
  } as unknown as BodyElement;
}

function docWith(...body: BodyElement[]): DocxDocumentModel {
  return {
    section: {
      pageWidth: 400, pageHeight: 4000,
      marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps,
    body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { [TEST_FONT]: 'roman' },
  } as unknown as DocxDocumentModel;
}

async function renderAndRead(...body: BodyElement[]) {
  const { canvas, fillTextCalls } = makeRecordingCanvas();
  const model = docWith(...body);
  await renderDocumentToCanvas(model, canvas, 0, {
    dpr: 1,
    width: 400, // scale = 400/400 = 1 (px per pt) ⇒ asserts are in pt-equivalent units
    layoutServices: createLayoutServices(model, {
      localMetrics: testFontSnapshot([{ family: TEST_FONT }]), measureContext: canvas.getContext('2d'),
    }),
  });
  return fillTextCalls;
}

const auto = (v: number): LineSpacing => ({ value: v, rule: 'auto', explicit: true });
const exact = (pt: number): LineSpacing => ({ value: pt, rule: 'exact', explicit: true });
const atLeast = (pt: number): LineSpacing => ({ value: pt, rule: 'atLeast', explicit: true });

// Synthetic 10pt untabled metrics at scale 1: ascent 8, descent 2, natural 10.
const ASCENT = 8;
const NATURAL = 10;
const TOL = 0.05;

describe('lineRule=auto (multiple spacing) pins the baseline at natural ascent — extra leading below (§17.3.1.33, #990)', () => {
  it('1.0× places the baseline at top + ascent (no extra)', async () => {
    const calls = await renderAndRead(paragraph('T', auto(1.0)));
    const t = calls.find((c) => c.text === 'T');
    expect(t).toBeDefined();
    expect(t!.y).toBeCloseTo(ASCENT, TOL); // 8
  });

  it('1.5× keeps the baseline at top + ascent (extra 0.5× leading below, NOT centred)', async () => {
    const calls = await renderAndRead(paragraph('T', auto(1.5)));
    const t = calls.find((c) => c.text === 'T');
    expect(t).toBeDefined();
    // Centring would give top + (15 − 10)/2 + 8 = 10.5; Word pins at 8.
    expect(t!.y).toBeCloseTo(ASCENT, TOL); // 8, not 10.5
  });

  it('2.0× keeps the baseline at top + ascent (extra 1.0× leading below, NOT centred)', async () => {
    const calls = await renderAndRead(paragraph('T', auto(2.0)));
    const t = calls.find((c) => c.text === 'T');
    expect(t).toBeDefined();
    // Centring would give top + (20 − 10)/2 + 8 = 13; Word pins at 8.
    expect(t!.y).toBeCloseTo(ASCENT, TOL); // 8, not 13
  });

  it('line PITCH (baseline-to-baseline) equals the full box height — pagination invariant', async () => {
    // Two words, each 300px wide at 10px/char in a 400px column ⇒ two lines.
    const w = 'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAA BBBBBBBBBBBBBBBBBBBBBBBBBBBBBB';
    const calls = await renderAndRead(paragraph(w, auto(2.0)));
    const line1 = calls.find((c) => c.text.startsWith('A'));
    const line2 = calls.find((c) => c.text.startsWith('B'));
    expect(line1).toBeDefined();
    expect(line2).toBeDefined();
    // First line baseline pinned at ascent.
    expect(line1!.y).toBeCloseTo(ASCENT, TOL); // 8
    // Box height = natural × 2 = 20; line pitch is that regardless of placement.
    expect(line2!.y - line1!.y).toBeCloseTo(NATURAL * 2, TOL); // 20
  });
});

describe('lineRule=exact / atLeast keep the centred placement (regression guard, unchanged by #990)', () => {
  it('exact 20pt centres the glyph box (half-leading above and below)', async () => {
    const calls = await renderAndRead(paragraph('T', exact(20)));
    const t = calls.find((c) => c.text === 'T');
    expect(t).toBeDefined();
    // Centred: top + (20 − 10)/2 + 8 = 13.
    expect(t!.y).toBeCloseTo(13, TOL);
  });

  it('atLeast 20pt centres the glyph box (half-leading above and below)', async () => {
    const calls = await renderAndRead(paragraph('T', atLeast(20)));
    const t = calls.find((c) => c.text === 'T');
    expect(t).toBeDefined();
    // atLeast → max(natural 10, 20) = 20; centred: top + (20 − 10)/2 + 8 = 13.
    expect(t!.y).toBeCloseTo(13, TOL);
  });

  it('atLeast below natural collapses to natural (no extra, baseline at ascent)', async () => {
    const calls = await renderAndRead(paragraph('T', atLeast(5)));
    const t = calls.find((c) => c.text === 'T');
    expect(t).toBeDefined();
    // max(natural 10, 5) = 10 ⇒ no extra ⇒ baseline at ascent 8.
    expect(t!.y).toBeCloseTo(ASCENT, TOL);
  });

  it('sub-single auto multiplier (0.5×) centres the glyph in the compressed line box', async () => {
    const calls = await renderAndRead(paragraph('T', auto(0.5)));
    const t = calls.find((c) => c.text === 'T');
    expect(t).toBeDefined();
    // lineH = 10 × 0.5 = 5 < glyph box 10. Word centres the glyph box in the
    // compressed advance: top + (5 − 10)/2 + ascent 8 = 5.5. The glyph may
    // extend beyond both line-box edges, but it does not shift wholly downward
    // into the following row/content.
    expect(t!.y).toBeCloseTo(5.5, TOL);
  });
});

describe('lineRule=auto — the substituted-font single-line FLOOR is centred; only the multiplier extra goes below', () => {
  // Times New Roman is tabled (font-metrics.ts design floor 2355/2048 ≈ 1.1499 em
  // ⇒ 11.499px for 10pt), while the mock canvas reports a naive 1.0 em (10px)
  // glyph box. So intendedSingle 11.499 > glyphNatural 10: the floor half-leading
  // ((11.499−10)/2 = 0.75) is centred, and the multiplier's extra falls BELOW.
  const timesPara = (ls: LineSpacing): BodyElement => {
    const p = paragraph('T', ls) as Any;
    p.defaultFontFamily = 'Times New Roman';
    p.runs[0].fontFamily = 'Times New Roman';
    p.runs[0].fontFamilyEastAsia = '';
    return p as BodyElement;
  };
  const timesDoc = (ls: LineSpacing): DocxDocumentModel => {
    const d = docWith(timesPara(ls)) as Any;
    d.fontFamilyClasses = { 'Times New Roman': 'roman' };
    return d as DocxDocumentModel;
  };
  const readTimes = async (ls: LineSpacing): Promise<number> => {
    const { canvas, fillTextCalls } = makeRecordingCanvas();
    const model = timesDoc(ls);
    await renderDocumentToCanvas(model, canvas, 0, {
      dpr: 1, width: 400,
      layoutServices: createLayoutServices(model, {
        localMetrics: testFontSnapshot([{ family: 'Times New Roman', lineHeightRatio: 2355 / 2048 }]), measureContext: canvas.getContext('2d'),
      }),
    });
    const t = fillTextCalls.find((c) => c.text === 'T');
    expect(t).toBeDefined();
    return t!.y;
  };

  it('the multiplier does NOT move the baseline over the floor (auto 1.0× ≡ auto 2.0× baseline)', async () => {
    const y1 = await readTimes(auto(1.0));
    const y2 = await readTimes(auto(2.0));
    // Floor centred (0.75 below box top over the ascent) is present in both; the
    // 2.0× multiplier's extra (11.499) all falls below, so the baseline is equal.
    expect(y2).toBeCloseTo(y1, TOL);
    // And it is the floor-centred position (≈ (11.499−10)/2 + 8 = 8.75), not the
    // old full-box centring (which for 2.0× would be (22.998−10)/2 + 8 = 14.5).
    expect(y1).toBeCloseTo((11.499 - 10) / 2 + 8, 0.1);
  });
});

describe('lineRule=auto on an ACTIVE docGrid keeps the full-box centring (grid gate)', () => {
  const gridDoc = (ls: LineSpacing): DocxDocumentModel => {
    const d = docWith(paragraph('T', ls)) as Any;
    d.section.docGridType = 'lines';
    d.section.docGridLinePitch = 18; // pt
    return d as DocxDocumentModel;
  };
  it('a gridded auto 2.0× line stays centred in the grid-snapped box (NOT pinned)', async () => {
    const { canvas, fillTextCalls } = makeRecordingCanvas();
    const model = gridDoc(auto(2.0));
    await renderDocumentToCanvas(model, canvas, 0, {
      dpr: 1, width: 400,
      layoutServices: createLayoutServices(model, {
        localMetrics: testFontSnapshot([{ family: TEST_FONT }]), measureContext: canvas.getContext('2d'),
      }),
    });
    const t = fillTextCalls.find((c) => c.text === 'T');
    expect(t).toBeDefined();
    // Gridded auto: lineH = max(glyphNatural 10, pitch 18 × 2 = 36) = 36. The grid
    // gate excludes it from the pin, so the glyph is CENTRED: (36 − 10)/2 + 8 = 21.
    // A pinned baseline would sit at 8 — the assertion distinguishes the two.
    expect(t!.y).toBeCloseTo(21, 0.2);
  });
});

// ---------------------------------------------------------------------------
// Text-box (shape) path — shapeLineMetrics mirrors the body baseline rule.
// ---------------------------------------------------------------------------

function textboxDoc(rule: 'auto' | 'exact' | 'atLeast' | null, val: number): DocxDocumentModel {
  const block = {
    text: 'T', fontSizePt: 10, fontFamily: TEST_FONT, fontFamilyEastAsia: TEST_FONT,
    alignment: 'left', spaceBefore: 0, spaceAfter: 0,
    ...(rule ? { lineSpacingRule: rule, lineSpacingVal: val } : {}),
  } as unknown as ShapeText;
  const shape = {
    type: 'shape',
    widthPt: 300, heightPt: 200,
    anchorXPt: 0, anchorYPt: 0,
    anchorXFromMargin: false, anchorYFromPara: true,
    anchorXRelativeFrom: 'column', anchorYRelativeFrom: 'paragraph',
    presetGeometry: 'rect', wrapMode: 'none', textAnchor: 't',
    textInsetL: 0, textInsetT: 0, textInsetR: 0, textInsetB: 0,
    textBlocks: [block],
  } as unknown as ShapeRun;
  const para = {
    type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: [shape as unknown as DocParagraph['runs'][number]],
    defaultFontSize: 10, defaultFontFamily: TEST_FONT, widowControl: false,
  } as unknown as DocParagraph;
  return {
    section: {
      pageWidth: 400, pageHeight: 600,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 4, footerDistance: 4, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps,
    body: [para as unknown as BodyElement],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { [TEST_FONT]: 'roman' },
  } as unknown as DocxDocumentModel;
}

async function renderTextbox(rule: 'auto' | 'exact' | 'atLeast' | null, val: number): Promise<number> {
  const { canvas, fillTextCalls } = makeRecordingCanvas();
  const model = textboxDoc(rule, val);
  await renderDocumentToCanvas(model, canvas, 0, {
    dpr: 1, width: 400,
    layoutServices: createLayoutServices(model, {
      localMetrics: testFontSnapshot([{ family: TEST_FONT }]), measureContext: canvas.getContext('2d'),
    }),
  });
  const t = fillTextCalls.find((c) => c.text === 'T');
  expect(t).toBeDefined();
  return t!.y;
}

describe('text-box (shape) lineRule=auto pins the baseline — extra leading below (mirrors the body, #990)', () => {
  it('the auto multiplier does NOT move the first-line baseline (top-anchored box, extra below)', async () => {
    // A top-anchored text box: the first line's box top is fixed by the anchor +
    // inset, so a pinned baseline = boxTop + ascent is invariant to the auto
    // multiplier. Centring would push the 2.0× baseline down by the extra/2.
    const yAuto1 = await renderTextbox('auto', 1.0);
    const yAuto2 = await renderTextbox('auto', 2.0);
    const yNone = await renderTextbox(null, 0);
    expect(yAuto2).toBeCloseTo(yAuto1, TOL);
    expect(yAuto1).toBeCloseTo(yNone, TOL); // 1.0× auto ≡ single spacing
  });

  it('exact spacing still centres the glyph box (regression guard, unchanged by #990)', async () => {
    // exact 20pt expands the box; the glyph stays centred, so its baseline sits
    // BELOW the pinned auto baseline by the half-leading (extra/2 = 5).
    const yAuto1 = await renderTextbox('auto', 1.0);
    const yExact = await renderTextbox('exact', 20);
    expect(yExact - yAuto1).toBeCloseTo((20 - NATURAL) / 2, TOL); // +5
  });
});
