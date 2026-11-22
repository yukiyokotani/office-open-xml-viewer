import { describe, it, expect } from 'vitest';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { renderDocumentToCanvas, type DocxTextRunInfo } from './renderer.js';
import type {
  BodyElement, DocParagraph, DocxTextRun, SectionProps, DocxDocumentModel,
  SectionGeom,
} from './types';

// ECMA-376 §17.6.20 — text direction is PER-SECTION (issue #1000, #988 batch-3
// adjudication ①): a vertical (tbRl, or btLr which shares the tbRl page FRAME
// per the #988 re-adjudication — its GLYPHS all ride the page rotation, see
// btlr-rotated-glyphs.test.ts) NON-FINAL section can
// coexist with a horizontal FINAL section in one document. The parser carries
// the ending section's `<w:textDirection>` on its SectionBreak marker; the
// paginator lays a vertical section out in its SWAPPED LOGICAL geometry
// (`verticalLayoutSection`'s quarter-turn: logical width = physical height,
// margins rotated) and stamps `sectionGeom` (logical) + `sectionTextDirection`
// in lockstep; the renderer rotates each page +90° per ITS OWN direction.

// Deterministic stub canvas: glyph advance = charCount × fontPx, font box =
// 0.8/0.2 em (a single line is exactly fontPx tall). Copied from
// per-section-page-geometry.test.ts.
function makeCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, fillText() {}, strokeText() {}, beginPath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fillRect() {}, drawImage() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    letterSpacing: '0px',
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

// Canonical layout builds its default measure ctx from `new OffscreenCanvas(...)`,
// which the node test env lacks — polyfill with the same deterministic stub.
(globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = class {
  getContext() { return makeCtx(); }
};

type DocRun = DocParagraph['runs'][number];
function textRun(text: string, fontSize: number): DocRun {
  const run: DocxTextRun = {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize, color: null, fontFamily: 'NotInMetrics', isLink: false, background: null,
    vertAlign: null, hyperlink: null,
  } as unknown as DocxTextRun;
  return { type: 'text', ...run } as DocRun;
}
function para(text: string, fontSize = 20): BodyElement {
  const p: DocParagraph = {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: [textRun(text, fontSize)],
    defaultFontSize: fontSize, defaultFontFamily: 'NotInMetrics', widowControl: false,
  } as unknown as DocParagraph;
  return { type: 'paragraph', ...p } as BodyElement;
}

const EMPTY_HF = { default: null, first: null, even: null };

// Horizontal FINAL (body) section with fully asymmetric margins so any margin
// rotation slip is caught.
function horizontalBodySection(): SectionProps {
  return {
    pageWidth: 400, pageHeight: 500,
    marginTop: 12, marginRight: 24, marginBottom: 36, marginLeft: 48,
    headerDistance: 5, footerDistance: 6, titlePage: false, evenAndOddHeaders: false,
  } as SectionProps;
}

// A vertical section's PHYSICAL page geometry (as the parser emits it — the
// sectPr's verbatim pgSz/pgMar), asymmetric margins.
const VERT_PHYS: SectionGeom = {
  pageWidth: 612, pageHeight: 792,
  marginTop: 10, marginRight: 20, marginBottom: 30, marginLeft: 40,
  headerDistance: 7, footerDistance: 8,
};

// body: [V1] <break: vertical btLr, explicit geom> [H1] (horizontal body section).
function mixedDoc(): DocxDocumentModel {
  const body: BodyElement[] = [
    para('V1'),
    { type: 'sectionBreak', kind: 'nextPage', geom: VERT_PHYS, textDirection: 'btLr' } as BodyElement,
    para('H1'),
  ];
  return {
    section: horizontalBodySection(), body,
    headers: EMPTY_HF, footers: EMPTY_HF,
    fontFamilyClasses: {},
  } as unknown as DocxDocumentModel;
}

function canonicalLayout(doc: DocxDocumentModel) {
  return layoutDocument(doc, createLayoutServices(doc, { measureContext: makeCtx() }), { currentDateMs: 0 });
}

const retainedText = (node: import('./layout/types.js').PaintNode): string => node.kind === 'paragraph'
  ? node.lines.flatMap((line) => line.placements)
      .filter((placement) => placement.kind === 'text')
      .map((placement) => placement.text).join('')
  : '';

describe('per-section text direction (§17.6.20, issue #1000) — paginator stamps', () => {
  it('stamps a vertical mid-section with its SWAPPED logical geometry + direction', () => {
    const doc = mixedDoc();
    const pages = canonicalLayout(doc).pages;
    // Page 0 = the btLr section. Its stamped geometry is the LOGICAL quarter-turn
    // of the marker's physical geom: logical width = physical height, and
    // logical margin{Left,Top,Right,Bottom} = physical margin{Top,Right,Bottom,Left}
    // (verticalLayoutSection). header/footer distances are preserved.
    expect(pages[0].section.textDirection).toBe('btLr');
    expect(pages[0].section.geometry).toEqual({
      pageWidth: 792, pageHeight: 612,
      marginLeft: 10, marginTop: 20, marginRight: 30, marginBottom: 40,
      headerDistance: 7, footerDistance: 8,
    });
    // Page 1 = horizontal final section: body-level geometry, direction null.
    expect(pages[1].section.textDirection).toBe('lrTb');
    expect(pages[1].section.geometry.pageWidth).toBe(400);
    expect(pages[1].section.geometry.pageHeight).toBe(500);
    expect(pages[1].section.geometry.marginLeft).toBe(48);
  });

  it('a geom-less continuous vertical section inherits and swaps the following page box', () => {
    const doc = mixedDoc();
    doc.body[1] = { type: 'sectionBreak', kind: 'continuous', textDirection: 'tbRl' } as BodyElement;
    const pages = canonicalLayout(doc).pages;
    expect(pages[0].section.textDirection).toBe('tbRl');
    // swap of body 400×500 / margins T12 R24 B36 L48 → logical
    // 500×400 / L12 T24 R36 B48.
    expect(pages[0].section.geometry).toEqual({
      pageWidth: 500, pageHeight: 400,
      marginLeft: 12, marginTop: 24, marginRight: 36, marginBottom: 48,
      headerDistance: 5, footerDistance: 6,
    });
  });

  it('all-vertical single-section document keeps the legacy swapped stamps', () => {
    const section: SectionProps = {
      ...horizontalBodySection(),
      pageWidth: 612, pageHeight: 792,
      marginTop: 10, marginRight: 20, marginBottom: 30, marginLeft: 40,
      headerDistance: 7, footerDistance: 8,
      textDirection: 'tbRl',
    } as SectionProps;
    const doc = {
      section, body: [para('V1'), para('V2')],
      headers: EMPTY_HF, footers: EMPTY_HF, fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const pages = canonicalLayout(doc).pages;
    for (const page of pages) {
      expect(page.section.textDirection).toBe('tbRl');
      expect(page.section.geometry.pageWidth).toBe(792);
      expect(page.section.geometry.pageHeight).toBe(612);
      expect(page.section.geometry.marginLeft).toBe(10);
    }
  });

  // Height- AND width-sensitive: the vertical section must paginate against its
  // SWAPPED frame. Physical 206(w)×300(h), margins all 20 ⇒ logical page 300×206,
  // content height 166 (8 one-line paras: a 20pt vertical single line is
  // 20 + 20/30pt, WORD_VERTICAL_LINE_CENTRED_BOX), wrap width 260 (a 10-char
  // 20pt para = 200pt fits on ONE line). If the frame were left physical:
  // content height 260 and wrap width 166 (two lines per para) — a totally
  // different distribution — so this fails the moment the swap is dropped.
  it('paginates a vertical mid-section against its SWAPPED logical frame', () => {
    const vertPhys: SectionGeom = {
      pageWidth: 206, pageHeight: 300,
      marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
      headerDistance: 0, footerDistance: 0,
    };
    const body: BodyElement[] = [
      ...Array.from({ length: 10 }, (_v, i) => para(`A${i + 1}______A${i + 1}`.slice(0, 10))),
      { type: 'sectionBreak', kind: 'nextPage', geom: vertPhys, textDirection: 'tbRl' } as BodyElement,
      para('B1'),
    ];
    const doc = {
      section: horizontalBodySection(), body,
      headers: EMPTY_HF, footers: EMPTY_HF, fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const pages = canonicalLayout(doc).pages;
    const counts = pages.map((page) => page.layers.body.filter((node) => node.kind === 'paragraph').length);
    // 8 one-line paras on page 0, the remaining 2 on page 1, B1 on page 2.
    expect(counts).toEqual([8, 2, 1]);
  });

  it('promotes a CONTINUOUS boundary to a page break when the direction changes', () => {
    // The paint model rotates a whole page: one direction per physical page is a
    // documented renderer constraint, so a continuous break whose flow direction
    // changes must open a fresh page. (Word GT for this boundary is unverified —
    // this pins the model constraint, not a Word behavior claim.)
    const section: SectionProps = { ...horizontalBodySection(), sectionStart: 'continuous' } as SectionProps;
    const body: BodyElement[] = [
      para('V1'),
      { type: 'sectionBreak', kind: 'continuous', textDirection: 'btLr' } as BodyElement,
      para('H1'),
    ];
    const doc = {
      section, body, headers: EMPTY_HF, footers: EMPTY_HF, fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const pages = canonicalLayout(doc).pages;
    expect(pages.length).toBe(2);

    // Control: the same continuous boundary WITHOUT a direction change stays on
    // one page (the existing §17.18.77 continuous behavior is untouched).
    const flatBody: BodyElement[] = [
      para('S1'),
      { type: 'sectionBreak', kind: 'continuous' } as BodyElement,
      para('S2'),
    ];
    const flatDoc = {
      section, body: flatBody, headers: EMPTY_HF, footers: EMPTY_HF, fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const flatPages = canonicalLayout(flatDoc).pages;
    expect(flatPages.length).toBe(1);
  });
});

// A recording canvas that captures ctx.rotate + canvas dims + text-run overlay
// payloads, so a test can assert the per-page +90° rotation.
function makeRecordingCanvas(): { canvas: HTMLCanvasElement; rotations: number[] } {
  let font = '10px serif';
  const rotations: number[] = [];
  const ctx = {
    get font() { return font; }, set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      return {
        width: [...s].length * p * 0.5,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {},
    rotate(r: number) { rotations.push(r); },
    setTransform() {}, setLineDash() {}, clearRect() {}, arc() {},
    quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {}, fillText() {}, strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, rotations };
}

function makeBorderRecordingCanvas(): {
  canvas: HTMLCanvasElement;
  strokes: Array<{ color: string; points: Array<readonly [number, number]>; deviceWidth: number }>;
} {
  type Matrix = { a: number; b: number; c: number; d: number; e: number; f: number };
  let matrix: Matrix = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };
  const stack: Matrix[] = [];
  let font = '10px serif';
  let path: Array<readonly [number, number]> = [];
  let strokeStyle = '#000';
  let lineWidth = 1;
  const strokes: Array<{ color: string; points: Array<readonly [number, number]>; deviceWidth: number }> = [];
  const point = (x: number, y: number): readonly [number, number] => [
    matrix.a * x + matrix.c * y + matrix.e,
    matrix.b * x + matrix.d * y + matrix.f,
  ];
  const ctx = {
    get font() { return font; }, set font(v: string) { font = v; },
    get strokeStyle() { return strokeStyle; }, set strokeStyle(v: string) { strokeStyle = v; },
    get lineWidth() { return lineWidth; }, set lineWidth(v: number) { lineWidth = v; },
    letterSpacing: '0px', fontKerning: 'auto',
    measureText: (s: string) => {
      const p = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      return {
        width: [...s].length * p * 0.5,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() { stack.push({ ...matrix }); },
    restore() { matrix = stack.pop()!; },
    scale(x: number, y: number) {
      matrix = { ...matrix, a: matrix.a * x, b: matrix.b * x, c: matrix.c * y, d: matrix.d * y };
    },
    translate(x: number, y: number) {
      matrix = {
        ...matrix,
        e: matrix.e + matrix.a * x + matrix.c * y,
        f: matrix.f + matrix.b * x + matrix.d * y,
      };
    },
    rotate(angle: number) {
      const cos = Math.cos(angle), sin = Math.sin(angle);
      matrix = {
        ...matrix,
        a: matrix.a * cos + matrix.c * sin,
        b: matrix.b * cos + matrix.d * sin,
        c: -matrix.a * sin + matrix.c * cos,
        d: -matrix.b * sin + matrix.d * cos,
      };
    },
    setTransform(a: number, b: number, c: number, d: number, e: number, f: number) {
      matrix = { a, b, c, d, e, f };
    },
    beginPath() { path = []; },
    moveTo(x: number, y: number) { path.push(point(x, y)); },
    lineTo(x: number, y: number) { path.push(point(x, y)); },
    stroke() {
      const normalScale = path.length >= 2 && Math.abs(path[0]![0] - path[1]![0]) < 1e-8
        ? Math.hypot(matrix.a, matrix.b)
        : Math.hypot(matrix.c, matrix.d);
      strokes.push({ color: strokeStyle, points: [...path], deviceWidth: lineWidth * normalScale });
    },
    closePath() {}, fill() {}, fillRect() {}, strokeRect() {}, clip() {}, rect() {},
    setLineDash() {}, clearRect() {}, arc() {}, quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {}, fillText() {}, strokeText() {},
    fillStyle: '#000', textAlign: 'left' as CanvasTextAlign,
    textBaseline: 'alphabetic' as CanvasTextBaseline,
    direction: 'ltr' as CanvasDirection, globalAlpha: 1,
    lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, strokes };
}

describe('per-section text direction (§17.6.20, issue #1000) — renderer', () => {
  it('rotates ONLY the vertical page +90°; the horizontal page stays unrotated', async () => {
    const doc = mixedDoc();
    const runs0: DocxTextRunInfo[] = [];
    const { canvas: c0, rotations: r0 } = makeRecordingCanvas();
    await renderDocumentToCanvas(doc, c0, 0, { dpr: 1, onTextRun: (r) => runs0.push(r) });
    const runs1: DocxTextRunInfo[] = [];
    const { canvas: c1, rotations: r1 } = makeRecordingCanvas();
    await renderDocumentToCanvas(doc, c1, 1, { dpr: 1, onTextRun: (r) => runs1.push(r) });

    // Page 0 (btLr): the whole-page +90° paint rotation fires (one +π/2 among the
    // rotations; per-glyph counter-rotations may add more entries).
    expect(r0.some((r) => Math.abs(r - Math.PI / 2) < 1e-9)).toBe(true);
    // Its canvas is the PHYSICAL page box (612×792 portrait).
    expect(c0.width / c0.height).toBeCloseTo(612 / 792, 2);
    // The overlay run is handed the vertical placement transform.
    expect(runs0.length).toBeGreaterThan(0);
    expect(runs0[0].transform).toBeDefined();

    // Page 1 (horizontal body): no page rotation, no overlay transform, its own
    // 400×500 box.
    expect(r1.some((r) => Math.abs(r - Math.PI / 2) < 1e-9)).toBe(false);
    expect(c1.width / c1.height).toBeCloseTo(400 / 500, 2);
    expect(runs1.length).toBeGreaterThan(0);
    expect(runs1[0].transform).toBeUndefined();
  });

  it.each([
    { width: 137.3, dpr: 1 },
    { width: 205.1, dpr: 2 },
  ])('snaps retained borders after the production vertical-page transform at width=$width dpr=$dpr', async ({ width, dpr }) => {
    const section = {
      pageWidth: 100.3, pageHeight: 200.7,
      marginTop: 10.2, marginRight: 11.3, marginBottom: 12.4, marginLeft: 13.1,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
      textDirection: 'tbRl',
    } as SectionProps;
    const bordered = para('B', 10) as DocParagraph & BodyElement;
    bordered.borders = {
      top: { style: 'single', color: '123456', width: 1, space: 0 },
      bottom: null, left: null, right: null, between: null,
    };
    const doc = {
      section, body: [bordered], headers: EMPTY_HF, footers: EMPTY_HF, fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const { canvas, strokes } = makeBorderRecordingCanvas();

    await renderDocumentToCanvas(doc, canvas, 0, { width, dpr });

    const border = strokes.find((stroke) => stroke.color.toLowerCase() === '#123456');
    expect(border?.points).toHaveLength(2);
    const xDevice = border!.points[0]![0];
    expect(border!.points[1]![0]).toBeCloseTo(xDevice, 8);
    const target = Math.round(border!.deviceWidth) % 2 === 1 ? .5 : 0;
    expect(xDevice - target).toBeCloseTo(Math.round(xDevice - target), 8);
  });
});

describe('per-section text direction (§17.6.20, issue #1000) — physical page size', () => {
  it('un-swaps each page by ITS OWN stamped direction', () => {
    const doc = mixedDoc();
    const pages = canonicalLayout(doc).pages;
    // Page 0 (vertical): stamped LOGICAL 792×612 → physical 612×792.
    expect(pages[0].geometry).toMatchObject({ widthPt: 612, heightPt: 792 });
    // Page 1 (horizontal): stamped physical box verbatim.
    expect(pages[1].geometry).toMatchObject({ widthPt: 400, heightPt: 500 });
  });

  it('resolves an EMPTY parity page from its page metadata: the blank belongs to the OUTGOING section', () => {
    // A VERTICAL first section followed by an oddPage-starting horizontal
    // section: the parity padding leaves a BLANK page (pages: [V1][blank][H1]
    // [H2]). §17.18.77 — the new section "begins on the next odd-numbered page,
    // LEAVING the next even page blank": the blank precedes the incoming
    // section's start, so it carries the OUTGOING (vertical) section's frame —
    // physical 612×792 — not the body-level 400×500 the element-stamp fallback
    // would report (the blank has no elements).
    const body: BodyElement[] = [
      para('V1'),
      // marker ending section 1 (vertical). The UPCOMING section's start type
      // is the NEXT marker's kind (§17.6.22): oddPage.
      { type: 'sectionBreak', kind: 'nextPage', geom: VERT_PHYS, textDirection: 'tbRl' } as BodyElement,
      para('H1'),
      {
        type: 'sectionBreak', kind: 'oddPage', geom: horizontalBodySection(),
      } as BodyElement,
      para('H2'),
    ];
    const doc = {
      section: horizontalBodySection(), body,
      headers: EMPTY_HF, footers: EMPTY_HF, fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const pages = canonicalLayout(doc).pages;
    // Page 0: V1 (vertical). Page 1: parity blank (the oddPage section pads
    // past the even slot). Page 2: H1. Page 3: H2 (nextPage default into the
    // final section).
    expect(pages.length).toBe(4);
    expect(pages[1].layers.body).toHaveLength(0);
    expect(pages[0].geometry).toMatchObject({ widthPt: 612, heightPt: 792 });
    // The blank parity page belongs to the OUTGOING vertical section
    // (§17.18.77): physical 612×792, not the body 400×500.
    expect(pages[1].geometry).toMatchObject({ widthPt: 612, heightPt: 792 });
    // The incoming oddPage section itself is horizontal at the body box.
    expect(pages[2].geometry).toMatchObject({ widthPt: 400, heightPt: 500 });
  });
});

describe('per-section text direction (§17.6.20, issue #1000) — vertical BODY + horizontal mid', () => {
  // The reverse mix: the FINAL (body) section is vertical, a NON-final section
  // horizontal. The body-level swap (verticalLayoutDoc) still runs, so the
  // horizontal marker explicitly carries the body's PHYSICAL geometry, so the
  // coordinate projection must not double-swap it. The horizontal page paints
  // unrotated while the vertical final page rotates.
  it('an authored horizontal mid-section keeps its PHYSICAL page geometry', () => {
    const section: SectionProps = {
      ...horizontalBodySection(),
      pageWidth: 612, pageHeight: 792,
      marginTop: 10, marginRight: 20, marginBottom: 30, marginLeft: 40,
      headerDistance: 7, footerDistance: 8,
      textDirection: 'tbRl',
    } as SectionProps;
    const body: BodyElement[] = [
      para('H1'),
      { type: 'sectionBreak', kind: 'nextPage', geom: VERT_PHYS } as BodyElement,
      para('V1'),
    ];
    const doc = {
      section, body, headers: EMPTY_HF, footers: EMPTY_HF, fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const pages = canonicalLayout(doc).pages;
    // The horizontal mid-section's frame is the body's verbatim PHYSICAL box —
    // NOT the swapped logical one, and NOT a double-swap of it.
    expect(pages[0].section.textDirection).toBe('lrTb');
    expect(pages[0].section.geometry).toEqual({
      pageWidth: 612, pageHeight: 792,
      marginTop: 10, marginRight: 20, marginBottom: 30, marginLeft: 40,
      headerDistance: 7, footerDistance: 8,
    });
    // The vertical final section keeps the swapped logical frame.
    expect(pages[1].section.textDirection).toBe('tbRl');
    expect(pages[1].section.geometry).toEqual({
      pageWidth: 792, pageHeight: 612,
      marginLeft: 10, marginTop: 20, marginRight: 30, marginBottom: 40,
      headerDistance: 7, footerDistance: 8,
    });
    // Physical page boxes agree per page.
    expect(pages[0].geometry).toMatchObject({ widthPt: 612, heightPt: 792 });
    expect(pages[1].geometry).toMatchObject({ widthPt: 612, heightPt: 792 });
  });

  it('paints the horizontal mid page unrotated and the vertical final page rotated', async () => {
    const section: SectionProps = {
      ...horizontalBodySection(),
      pageWidth: 612, pageHeight: 792,
      marginTop: 10, marginRight: 20, marginBottom: 30, marginLeft: 40,
      textDirection: 'tbRl',
    } as SectionProps;
    const body: BodyElement[] = [
      para('H1'),
      { type: 'sectionBreak', kind: 'nextPage' } as BodyElement,
      para('V1'),
    ];
    const doc = {
      section, body, headers: EMPTY_HF, footers: EMPTY_HF, fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const { canvas: c0, rotations: r0 } = makeRecordingCanvas();
    await renderDocumentToCanvas(doc, c0, 0, { dpr: 1 });
    const { canvas: c1, rotations: r1 } = makeRecordingCanvas();
    await renderDocumentToCanvas(doc, c1, 1, { dpr: 1 });
    expect(r0.some((r) => Math.abs(r - Math.PI / 2) < 1e-9)).toBe(false);
    expect(r1.some((r) => Math.abs(r - Math.PI / 2) < 1e-9)).toBe(true);
    // Both pages share the same physical portrait box here.
    expect(c0.width / c0.height).toBeCloseTo(612 / 792, 2);
    expect(c1.width / c1.height).toBeCloseTo(612 / 792, 2);
  });
});

describe('per-section text direction (§17.6.20, issue #1000) — split slices and layoutDocument', () => {
  it('stamps the direction on every slice of a split paragraph in a vertical mid-section', () => {
    // Physical 200×300, margins 20 ⇒ logical frame 300×200: wrap width 260
    // (13 chars × 20 pt), content height 160 (8 lines/page). One 40-char
    // paragraph wraps to 4 lines; 26 of them (104 lines) force paragraph
    // SPLITS across pages, so the split-slice stamping path (tagSection…
    // thunks) is exercised, not just pushTagged.
    const vertPhys: SectionGeom = {
      pageWidth: 200, pageHeight: 300,
      marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
      headerDistance: 0, footerDistance: 0,
    };
    const body: BodyElement[] = [
      para('X'.repeat(13 * 12)), // 12 lines > 8 per page ⇒ must split
      { type: 'sectionBreak', kind: 'nextPage', geom: vertPhys, textDirection: 'tbRl' } as BodyElement,
      para('B1'),
    ];
    const doc = {
      section: horizontalBodySection(), body,
      headers: EMPTY_HF, footers: EMPTY_HF, fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const pages = canonicalLayout(doc).pages;
    // The paragraph split across ≥ 2 pages inside the vertical section.
    expect(pages.length).toBeGreaterThanOrEqual(3);
    for (let i = 0; i < pages.length - 1; i++) {
      const slices = pages[i].layers.body.filter((node) => node.kind === 'paragraph');
      expect(slices.length).toBeGreaterThan(0);
      for (const s of slices) {
        expect(pages[i].section.textDirection).toBe('tbRl');
        expect(pages[i].section.geometry.pageWidth).toBe(300);
        expect(pages[i].section.geometry.pageHeight).toBe(200);
      }
    }
  });

  it('layoutDocument reports each page own textDirection', async () => {
    const { layoutDocument } = await import('./document-layout.js');
    const layout = layoutDocument(mixedDoc());
    expect(layout.pages.length).toBe(2);
    expect(layout.pages[0].section.textDirection).toBe('btLr');
    // Page geometry remains the physical page box; section regions own the
    // logical-to-physical writing-mode transform.
    expect(layout.pages[0].geometry.widthPt).toBe(612);
    expect(layout.pages[1].section.textDirection).toBe('lrTb');
    expect(layout.pages[1].geometry.widthPt).toBe(400);
  });
});

describe('per-section text direction (§17.6.20, issue #1000) — HF reserve', () => {
  it('reserves NOTHING on a vertical page (paint draws its HF physically, reserve-free)', () => {
    // Vertical mid-section: physical 206×300, margins 20 ⇒ logical content height
    // 166 ⇒ exactly 8 one-line 20pt paras (20 + 20/30pt vertical single lines,
    // WORD_VERTICAL_LINE_CENTRED_BOX). A 40pt footer whose extent (10+40)
    // overflows the 20pt logical bottom margin would phantom-reserve 30pt and
    // push A7/A8 off the page — the paint pass never reserves on vertical pages
    // (issue #988: vertical HF draw horizontally at the physical margins), so
    // pagination must not either.
    const vertPhys: SectionGeom = {
      pageWidth: 206, pageHeight: 300,
      marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
      headerDistance: 0, footerDistance: 10,
    };
    const footer40 = { body: [para('F', 40)] };
    const body: BodyElement[] = [
      ...Array.from({ length: 8 }, (_v, i) => para(`A${i + 1}`)),
      {
        type: 'sectionBreak', kind: 'nextPage', geom: vertPhys, textDirection: 'tbRl',
        headers: EMPTY_HF, footers: { default: footer40, first: null, even: null },
        titlePage: false,
      } as BodyElement,
      para('B1'),
    ];
    const doc = {
      section: { ...horizontalBodySection(), footerDistance: 10 },
      body,
      headers: EMPTY_HF, footers: { default: footer40, first: null, even: null },
      fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const pages = canonicalLayout(doc).pages;
    const textsOn = (i: number) =>
      (pages[i]?.layers.body ?? []).filter((node) => node.kind === 'paragraph').map(retainedText);
    expect(textsOn(0)).toEqual(['A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7', 'A8']);
    expect(textsOn(1)).toEqual(['B1']);
  });
});
