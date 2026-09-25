import { describe, it, expect } from 'vitest';
import { layoutDocument } from './document-layout.js';
import { renderDocumentToCanvas } from './renderer.js';
import type {
  BodyElement, DocParagraph, DocxTextRun, FieldRun, SectionProps, DocxDocumentModel,
  SectionGeom, PageNumType, HeaderFooter,
} from './types';

// ECMA-376 §17.6.12 `<w:pgNumType>` — end-to-end coverage of per-section page-number
// RESTART (`w:start`) + FORMAT (`w:fmt`) and the §17.16.4.3.1 field `\*` switch,
// from the parsed model through pagination (which stamps `sectionPageNumType`) to
// the PAGE field text a footer paints. Mirrors per-section-page-geometry.test.ts's
// deterministic stub-canvas approach (glyph advance = charCount × fontPx, line
// height = fontPx) so pagination is exact and headless.

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

type TextDraw = Readonly<{ text: string; x: number; y: number }>;

function makeRecordingCanvas(): { canvas: HTMLCanvasElement; texts: string[]; draws: TextDraw[] } {
  const texts: string[] = [];
  const draws: TextDraw[] = [];
  const ctx = makeCtx() as CanvasRenderingContext2D & Record<string, unknown>;
  ctx.fillText = (text: string, x: number, y: number) => {
    texts.push(text);
    draws.push({ text, x, y });
  };
  Object.assign(ctx, {
    setTransform() {}, clearRect() {}, closePath() {}, rect() {}, clip() {},
    translate() {}, rotate() {}, scale() {}, setLineDash() {}, arc() {},
    quadraticCurveTo() {}, bezierCurveTo() {}, strokeRect() {},
    createLinearGradient() { return { addColorStop() {} }; },
    globalAlpha: 1, textBaseline: 'alphabetic', lineCap: 'butt', lineJoin: 'miter',
  });
  const canvas = {
    width: 0,
    height: 0,
    getContext: () => ctx,
  } as unknown as HTMLCanvasElement;
  return { canvas, texts, draws };
}

// OffscreenCanvas polyfill (paginateWithHeaderFooterReserve builds its measure ctx
// from `new OffscreenCanvas`, absent in node). Same deterministic stub.
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
function pageFieldRun(instruction = 'PAGE', fontSize = 20): DocRun {
  const f: FieldRun = {
    fieldType: 'page', instruction, fallbackText: '?',
    bold: false, italic: false, underline: false, strikethrough: false,
    fontSize, color: null, fontFamily: 'NotInMetrics', background: null, vertAlign: null,
  } as unknown as FieldRun;
  return { type: 'field', ...f } as DocRun;
}

function numPagesFieldRun(fontSize = 10): DocRun {
  const f: FieldRun = {
    fieldType: 'numPages', instruction: 'NUMPAGES', fallbackText: '8',
    bold: true, italic: false, underline: false, strikethrough: false,
    fontSize, color: null, fontFamily: 'NotInMetrics', background: null, vertAlign: null,
  } as unknown as FieldRun;
  return { type: 'field', ...f } as DocRun;
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
function footerWithPageField(instruction = 'PAGE'): HeaderFooter {
  const p: DocParagraph = {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: [pageFieldRun(instruction, 20)],
    defaultFontSize: 20, defaultFontFamily: 'NotInMetrics', widowControl: false,
  } as unknown as DocParagraph;
  return { body: [{ type: 'paragraph', ...p } as BodyElement] };
}

const GEOM = (): SectionGeom => ({
  pageWidth: 200, pageHeight: 140,
  marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
  headerDistance: 0, footerDistance: 0,
});
// content height 100 ⇒ five 20pt lines per page.

/** A 2-section doc: front matter (fmt/start on the mid-body break) + body (final
 *  section, its pgNumType on doc.section). `bodyLines` per section controls how
 *  many physical pages each spans. */
function twoSectionDoc(
  frontPgNum: PageNumType | null,
  bodyPgNum: PageNumType | null,
  frontLines = 6,
  bodyLines = 6,
  footerInstr = 'PAGE',
): DocxDocumentModel {
  const front: BodyElement[] = [];
  for (let i = 0; i < frontLines; i++) front.push(para(`F${i}`));
  const body: BodyElement[] = [];
  for (let i = 0; i < bodyLines; i++) body.push(para(`B${i}`));
  const footer = footerWithPageField(footerInstr);
  const bodySection: SectionProps = {
    ...GEOM(), titlePage: false, evenAndOddHeaders: false,
    pageNumType: bodyPgNum,
  } as unknown as SectionProps;
  return {
    section: bodySection,
    body: [
      ...front,
      {
        type: 'sectionBreak', kind: 'nextPage', geom: GEOM(),
        headers: { default: null, first: null, even: null },
        footers: { default: footer, first: null, even: null },
        titlePage: false,
        pageNumType: frontPgNum,
      } as BodyElement,
      ...body,
    ],
    headers: { default: null, first: null, even: null },
    footers: { default: footer, first: null, even: null },
    fontFamilyClasses: {},
  } as unknown as DocxDocumentModel;
}

/** A 2-section doc joined by a CONTINUOUS section break (§17.18.77): the second
 *  section begins MID-PAGE below the first (no page break) and its content SPILLS
 *  onto the following physical pages. The second section is the FINAL (body-level)
 *  section, so its `w:type="continuous"` lives on `section.sectionStart` and its
 *  `<w:pgNumType>` on `section.pageNumType` — exactly how the parser models
 *  sample-27 (§17.6.12). `firstLines` keeps section 1 short so section 2 shares its
 *  page; `secondLines` controls how far section 2 spills. */
function continuousSpilloverDoc(
  secondPgNum: PageNumType | null,
  firstLines: number,
  secondLines: number,
): DocxDocumentModel {
  const front: BodyElement[] = [];
  for (let i = 0; i < firstLines; i++) front.push(para(`S1-${i}`));
  const body: BodyElement[] = [];
  for (let i = 0; i < secondLines; i++) body.push(para(`S2-${i}`));
  const footer = footerWithPageField('PAGE');
  const bodySection: SectionProps = {
    ...GEOM(), titlePage: false, evenAndOddHeaders: false,
    // §17.18.77 — the body-level (final) section starts CONTINUOUS, so it shares the
    // page where section 1 ends. §17.6.12 — and it carries the restart.
    sectionStart: 'continuous',
    pageNumType: secondPgNum,
  } as unknown as SectionProps;
  return {
    section: bodySection,
    body: [
      ...front,
      {
        // The mid-body marker ENDS section 1 (which carries no restart). The break's
        // effective kind is read from the UPCOMING section (the body section's
        // `sectionStart: 'continuous'`), so this marker's own `kind` is irrelevant.
        type: 'sectionBreak', kind: 'nextPage', geom: GEOM(),
        headers: { default: null, first: null, even: null },
        footers: { default: footer, first: null, even: null },
        titlePage: false,
        pageNumType: null,
      } as BodyElement,
      ...body,
    ],
    headers: { default: null, first: null, even: null },
    footers: { default: footer, first: null, even: null },
    fontFamilyClasses: {},
  } as unknown as DocxDocumentModel;
}

describe('PAGE field renders the per-section displayed number (footer)', () => {
  async function footerTexts(doc: DocxDocumentModel, pageIndex: number): Promise<string[]> {
    const { canvas, texts } = makeRecordingCanvas();
    await renderDocumentToCanvas(doc, canvas, pageIndex, { dpr: 1 });
    return texts;
  }

  async function continuousDecimalPageFieldTexts(
    doc: DocxDocumentModel,
    pageIndex: number,
  ): Promise<string[]> {
    return (await footerTexts(doc, pageIndex)).filter((text) => /^\d+$/.test(text));
  }

  it('paints i, ii, then restarts to 1, 2 across the two sections', async () => {
    const doc = twoSectionDoc(
      { fmt: 'lowerRoman', start: 1 },
      { fmt: 'decimal', start: 1 },
    );
    // The footer's PAGE field is the only glyph on each page besides body text; the
    // roman/decimal string is unique enough to assert via `.includes`.
    expect(await footerTexts(doc, 0)).toContain('i');
    expect(await footerTexts(doc, 1)).toContain('ii');
    expect(await footerTexts(doc, 2)).toContain('1');
    expect(await footerTexts(doc, 3)).toContain('2');
  });

  it('start=0 offsets the first page number to 0 (Word writes start="0")', async () => {
    const doc = twoSectionDoc({ start: 0 }, null, 6, 1);
    // page 0 shows "0", page 1 shows "1" (decimal — no fmt on the front section).
    expect(await footerTexts(doc, 0)).toContain('0');
    expect(await footerTexts(doc, 1)).toContain('1');
  });

  it('restarts a spilling continuous section after its shared first page', async () => {
    const doc = continuousSpilloverDoc({ start: 50 }, 3, 9);
    const layout = layoutDocument(doc);
    expect(layout.pages).toHaveLength(3);
    const sharedRegions = layout.pages[0]!.sectionRegions;

    expect(sharedRegions).toHaveLength(2);
    const outgoingOwner = sharedRegions[0]!.sectionOccurrenceId;
    const incomingOwner = sharedRegions[1]!.sectionOccurrenceId;

    expect(sharedRegions.map((region) => region.sectionOccurrenceId)).toEqual([
      outgoingOwner,
      incomingOwner,
    ]);
    expect(layout.pages.map((page) => page.sectionOccurrenceId)).toEqual([
      outgoingOwner,
      incomingOwner,
      incomingOwner,
    ]);
    expect(layout.pages.map((page) => page.pageNumber.displayNumber)).toEqual([1, 51, 52]);
    expect(await Promise.all(layout.pages.map((_, pageIndex) =>
      continuousDecimalPageFieldTexts(doc, pageIndex))))
      .toEqual([['1'], ['51'], ['52']]);
  });

  it('starts at the authored number when a continuous section first paints on the next page', async () => {
    const doc = continuousSpilloverDoc({ start: 2 }, 5, 6);
    const layout = layoutDocument(doc);

    expect(layout.pages).toHaveLength(3);
    const sharedRegions = layout.pages[0]!.sectionRegions;
    expect(sharedRegions).toHaveLength(2);
    expect(sharedRegions[0]!.sectionOccurrenceId).not.toBe(sharedRegions[1]!.sectionOccurrenceId);
    expect(sharedRegions.map((region) => region.blockEndPt - region.blockStartPt))
      .toEqual([100, 0]);
    expect(layout.pages.map((page) => page.pageNumber.displayNumber)).toEqual([1, 2, 3]);
    expect(await Promise.all(layout.pages.map((_, pageIndex) =>
      continuousDecimalPageFieldTexts(doc, pageIndex))))
      .toEqual([['1'], ['2'], ['3']]);
  });

  it('keeps a continuous restart page-wide when both sections fit the shared page', async () => {
    const doc = continuousSpilloverDoc({ start: 99 }, 2, 2);
    const layout = layoutDocument(doc);
    expect(layout.pages).toHaveLength(1);
    const sharedRegions = layout.pages[0]!.sectionRegions;

    expect(sharedRegions).toHaveLength(2);
    expect(layout.pages[0]!.sectionOccurrenceId).toBe(sharedRegions[0]!.sectionOccurrenceId);
    expect(layout.pages[0]!.sectionOccurrenceId).not.toBe(sharedRegions[1]!.sectionOccurrenceId);
    expect(layout.pages[0]!.pageNumber).toEqual({
      displayNumber: 1,
      format: 'decimal',
      sectionOccurrenceId: sharedRegions[0]!.sectionOccurrenceId,
    });
    expect(await continuousDecimalPageFieldTexts(doc, 0)).toEqual(['1']);
  });

  it('field \\* switch overrides the section fmt (PAGE \\* Roman on a decimal section)', async () => {
    // Body section decimal; the footer field carries `\* Roman` ⇒ uppercase roman.
    const doc = twoSectionDoc(null, { fmt: 'decimal', start: 3 }, 1, 1, 'PAGE \\* Roman');
    // front page 0 = 1 (decimal, no fmt) but its footer also carries \* Roman ⇒ "I".
    expect(await footerTexts(doc, 0)).toContain('I');
    // body page 1 restarts to 3; \* Roman ⇒ "III".
    expect(await footerTexts(doc, 1)).toContain('III');
  });

  it('routes an international section fmt end-to-end (chineseCounting §17.6.12/§17.18.59)', async () => {
    // §17.18.59 chineseCounting: 1 → 一, 2 → 二. The core converter extension flows
    // straight through the PAGE-field path (resolveFieldText → formatOrdinalNumber)
    // with no renderer change — this pins that wiring for the CJK systems.
    const doc = twoSectionDoc(null, { fmt: 'chineseCounting', start: 1 }, 1, 6);
    // front page 0 = 1 (decimal, no fmt); body pages restart to 一, 二 (chineseCounting).
    expect(await footerTexts(doc, 0)).toContain('1');
    expect(await footerTexts(doc, 1)).toContain('一');
    expect(await footerTexts(doc, 2)).toContain('二');
  });

  it('routes an international field \\* switch end-to-end (PAGE \\* HEBREW1 §17.16.4.3.1)', async () => {
    // §17.16.4.3.1 HEBREW1 → hebrew1 gematria; §17.18.59: 1 → א. The field switch
    // overrides the (absent) section fmt on the single body page.
    const doc = twoSectionDoc(null, { start: 1 }, 1, 1, 'PAGE \\* HEBREW1');
    expect(await footerTexts(doc, 0)).toContain('א'); // page 1 → א
  });

  it('single-section document without pgNumType is unchanged (decimal 1..N)', async () => {
    const footer = footerWithPageField('PAGE');
    const section: SectionProps = {
      ...GEOM(), titlePage: false, evenAndOddHeaders: false,
    } as unknown as SectionProps;
    const doc = {
      section,
      body: [para('A0'), para('A1'), para('A2'), para('A3'), para('A4'), para('A5')],
      headers: { default: null, first: null, even: null },
      footers: { default: footer, first: null, even: null },
      fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    // 6 lines, 5 per page ⇒ 2 pages: footers "1" and "2".
    expect(await footerTexts(doc, 0)).toContain('1');
    expect(await footerTexts(doc, 1)).toContain('2');
  });

  it('keeps recomputed PAGE and NUMPAGES on the authored run-position baseline', async () => {
    // §17.16.18 complex-field delimiters split one displayed result across
    // physical runs. The instruction run inherits w:position=6 (3pt) through
    // the style hierarchy; the recomputed field must retain that effective
    // §17.3.2.24 position instead of dropping to the unshifted baseline.
    const positionedTypography = {
      strike: false, doubleStrike: false, caps: false, smallCaps: false, colorAuto: false,
      verticalAlign: { status: 'missing', raw: null, value: null },
      positionPt: { status: 'valid', raw: '6', value: 3 },
      snapToGrid: null, characterSpacingPt: null, characterScale: null,
      kerningThresholdPt: null,
      emphasis: { status: 'missing', raw: null, value: null },
      languages: { eastAsia: null, bidi: null },
      eastAsianLayout: {
        vert: null, vertCompress: null, combine: null,
        combineBrackets: { status: 'missing', raw: null, value: null },
      },
    } as const;
    const positionedText = (text: string): DocRun => ({
      ...textRun(text, 10), position: 3,
    } as DocRun);
    const positionedField = (field: DocRun): DocRun => ({
      ...field, __typographyAcquisition: positionedTypography,
    } as unknown as DocRun);
    const footer: HeaderFooter = { body: [{
      type: 'paragraph', alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
      spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
      runs: [
        positionedText('Page '),
        positionedField(pageFieldRun('PAGE', 10)),
        positionedText(' of '),
        positionedField(numPagesFieldRun()),
      ],
      defaultFontSize: 10, defaultFontFamily: 'NotInMetrics', widowControl: false,
    } as unknown as BodyElement] };
    const section: SectionProps = {
      ...GEOM(), titlePage: false, evenAndOddHeaders: false,
    } as unknown as SectionProps;
    const doc = {
      section,
      // 56 × 20pt lines at five lines/page ⇒ 12 pages, so the runtime NUMPAGES
      // value is two digits even though its cached fallback is one digit.
      body: Array.from({ length: 56 }, (_, index) => para(`L${index}`)),
      headers: { default: null, first: null, even: null },
      footers: { default: footer, first: null, even: null },
      fontFamilyClasses: {},
    } as unknown as DocxDocumentModel;
    const { canvas, draws } = makeRecordingCanvas();

    await renderDocumentToCanvas(doc, canvas, 0, { dpr: 1 });

    const footerDraws = draws.filter(({ text }) => ['Page ', '1', '12'].includes(text));
    expect(footerDraws.map(({ text }) => text)).toEqual(['Page ', '1', '12']);
    expect(new Set(footerDraws.map(({ y }) => y)).size).toBe(1);
  });
});
