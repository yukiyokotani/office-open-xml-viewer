import { describe, it, expect } from 'vitest';
import { createLayoutServices } from './layout-runtime.js';
import { renderDocumentToCanvas } from './renderer.js';
import { testFontSnapshot } from './layout/test-font-snapshot.js';
import type { BodyElement, DocNote, DocParagraph, DocxDocumentModel, HeaderFooter, SectionProps } from './types';

// ECMA-376 §17.6.11 (pgMar/@bottom) — the main-document text bottom is placed at the
// GREATER of the bottom margin and the footer's extent, so a footer taller than the
// bottom-margin allowance (marginBottom − footerDistance) rises into the content area
// and content (body AND footnotes) must clear it: Word never lays main text over a
// footer. The paginator measures each page's footer and re-paginates with that
// reservation (paginateWithHeaderFooterReserve), and the footnote block is raised by the
// same overflow. A NEGATIVE bottom margin is the spec's explicit exception — text is
// then measured from the page bottom regardless of the footer and overlaps it, so
// nothing is reserved. These tests pin those rules with a synthetic doc whose footer
// is far taller than its bottom margin (reconstructed from sample-13's masthead
// footer: a DOI / corresponding-author block ~53pt tall over a ~49pt margin).

interface Call { text: string; y: number; }

function makeRecordingCanvas(): { canvas: HTMLCanvasElement; calls: Call[] } {
  let font = '10px serif';
  const calls: Call[] = [];
  let transform = { scaleX: 1, scaleY: 1, translateX: 0, translateY: 0 };
  const transformStack: typeof transform[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      return {
        width: [...s].length * p * 0.5,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() { transformStack.push({ ...transform }); },
    restore() { transform = transformStack.pop() ?? transform; },
    beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {},
    scale(x: number, y: number) {
      transform.scaleX *= x;
      transform.scaleY *= y;
    },
    translate(x: number, y: number) {
      transform.translateX += transform.scaleX * x;
      transform.translateY += transform.scaleY * y;
    },
    rotate() {},
    setLineDash() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {},
    fillText(s: string, _x: number, y: number) {
      calls.push({ text: s, y: transform.translateY + transform.scaleY * y });
    },
    strokeText(s: string, _x: number, y: number) {
      calls.push({ text: s, y: transform.translateY + transform.scaleY * y });
    },
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, calls };
}

function para(text: string): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: text
      ? [{
          type: 'text', text, bold: false, italic: false, underline: false,
          strikethrough: false, fontSize: 10, color: null, fontFamily: 'Times New Roman',
          fontFamilyEastAsia: '', isLink: false, background: null, vertAlign: null, hyperlink: null,
        } as DocParagraph['runs'][number]]
      : [],
    defaultFontSize: 10, defaultFontFamily: 'Times New Roman', widowControl: false,
  } as unknown as DocParagraph;
}

// A footnote reference run (kind 'footnote') appended to a body paragraph, so the
// page that holds it retains the note block in the canonical notes layer.
function paraWithFootnoteRef(text: string, noteId: string): DocParagraph {
  const p = para(text);
  (p.runs as DocParagraph['runs']).push({
    type: 'text', text: '', bold: false, italic: false, underline: false,
    strikethrough: false, fontSize: 10, color: null, fontFamily: 'Times New Roman',
    fontFamilyEastAsia: '', isLink: false, background: null, vertAlign: null, hyperlink: null,
    noteRef: { kind: 'footnote', id: noteId },
  } as unknown as DocParagraph['runs'][number]);
  return p;
}

// pageHeight 600, margins 10, footerDistance 4 → bottom-margin allowance for the
// footer is marginBottom − footerDistance = 6pt. A footer taller than 6pt overflows.
function docWithFooter(
  body: BodyElement[],
  footer: HeaderFooter | null,
  opts: { footnotes?: DocNote[]; marginBottom?: number; vAlign?: SectionProps['vAlign'] } = {},
): DocxDocumentModel {
  const section: SectionProps = {
    pageWidth: 400, pageHeight: 600,
    marginTop: 10, marginRight: 10, marginBottom: opts.marginBottom ?? 10, marginLeft: 10,
    headerDistance: 4, footerDistance: 4, titlePage: false, evenAndOddHeaders: false,
    sectionStart: 'nextPage',
    vAlign: opts.vAlign ?? null,
  } as SectionProps;
  return {
    section,
    body,
    headers: { default: null, first: null, even: null },
    footers: { default: footer, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
    footnotes: opts.footnotes ?? [],
  } as unknown as DocxDocumentModel;
}

async function renderPage0(doc: DocxDocumentModel): Promise<Call[]> {
  const { canvas, calls } = makeRecordingCanvas();
  await renderDocumentToCanvas(doc, canvas, 0, {
    dpr: 1, width: 400,
    layoutServices: createLayoutServices(doc, {
      localMetrics: testFontSnapshot([{ family: 'Times New Roman', lineHeightRatio: 2355 / 2048 }]), measureContext: canvas.getContext('2d'),
    }),
  });
  return calls;
}

async function renderPage0Default(doc: DocxDocumentModel): Promise<Call[]> {
  const { canvas, calls } = makeRecordingCanvas();
  await renderDocumentToCanvas(doc, canvas, 0, { dpr: 1, width: 400 });
  return calls;
}

describe('footer reserve — content never overlaps a tall footer (ECMA-376 §17.6.11)', () => {
  // Enough single-line body paragraphs to overflow page 0 (content area is 580pt;
  // ~50 lines is well over a page) so, absent any reservation, body text packs all
  // the way down to the content bottom (600 − marginBottom = 590pt).
  const body = (): BodyElement[] =>
    Array.from({ length: 50 }, () => para('BODY') as unknown as BodyElement);
  // A footer far taller than the 6pt allowance (6 lines ≈ ~70pt), so its top edge
  // rises well above the content bottom and into the body's region.
  const tallFooter: HeaderFooter = {
    body: Array.from({ length: 6 }, () => para('FTR') as unknown as BodyElement),
  };

  it('breaks body to the next page so no line is painted into the tall footer band', async () => {
    const withFooter = await renderPage0(docWithFooter(body(), tallFooter));
    const noFooter = await renderPage0(docWithFooter(body(), null));

    const bodyY = (calls: Call[]) => calls.filter((c) => c.text === 'BODY').map((c) => c.y);
    const footerY = (calls: Call[]) => calls.filter((c) => c.text === 'FTR').map((c) => c.y);

    const maxBodyTall = Math.max(...bodyY(withFooter));
    const minFooter = Math.min(...footerY(withFooter));
    const maxBodyNone = Math.max(...bodyY(noFooter));

    // Sanity: the tall footer is actually painted on page 0.
    expect(footerY(withFooter).length).toBeGreaterThan(0);

    // INVARIANT: with the tall footer reserved, the lowest body line on page 0 sits
    // ABOVE the topmost footer line — body never overlaps the footer.
    expect(maxBodyTall).toBeLessThan(minFooter);

    // NON-TRIVIALITY: the SAME body, with no footer to reserve against, packs down
    // PAST where the tall footer's top sits — proving the invariant above is not
    // satisfied trivially (the body genuinely reaches into that band) and that the
    // reservation, not a short body, is what clears the footer.
    expect(maxBodyNone).toBeGreaterThan(minFooter);
  });

  it('bottom-aligns body ink to the effective top edge of a tall footer', async () => {
    const calls = await renderPage0Default(docWithFooter(
      [para('BODY') as unknown as BodyElement],
      tallFooter,
      { vAlign: 'bottom' },
    ));
    const topAligned = await renderPage0Default(docWithFooter(
      [para('BODY') as unknown as BodyElement],
      null,
      { vAlign: 'top' },
    ));
    const footerBaselines = calls.filter((call) => call.text === 'FTR').map((call) => call.y);
    const lineAdvance = footerBaselines[1]! - footerBaselines[0]!;
    const baselineFromFlowTop = topAligned.find((call) => call.text === 'BODY')!.y - 10;
    const bodyFlowBottom = calls.find((call) => call.text === 'BODY')!.y
      + lineAdvance - baselineFromFlowTop;
    const footerFlowTop = footerBaselines[0]! - baselineFromFlowTop;

    // Both stories use the same retained line metrics. Deriving the flow-box edges
    // from their baselines avoids hard-coded ascent/descent guesses (+2/-8).
    expect(bodyFlowBottom).toBeCloseTo(footerFlowTop, 8);
  });

  it('raises the footnote block above a tall footer so notes do not overlap it', async () => {
    // A short body (fits page 0) whose paragraph references a footnote, plus the tall
    // footer. The footnote block is anchored at the bottom margin; the tall footer
    // overflows above the margin, so without the reserve the note prints over the
    // footer. The same overflow raises the note block to clear it (§17.6.11).
    const footnotes: DocNote[] = [{ id: 'fn1', content: [para('NOTE') as unknown as BodyElement] }];
    const docBody: BodyElement[] = [
      para('BODY') as unknown as BodyElement,
      paraWithFootnoteRef('BODY', 'fn1') as unknown as BodyElement,
    ];
    const calls = await renderPage0(docWithFooter(docBody, tallFooter, { footnotes }));

    const noteY = calls.filter((c) => c.text === 'NOTE').map((c) => c.y);
    const footerY = calls.filter((c) => c.text === 'FTR').map((c) => c.y);

    // Sanity: both the footnote and the tall footer are painted on page 0.
    expect(noteY.length).toBeGreaterThan(0);
    expect(footerY.length).toBeGreaterThan(0);

    // INVARIANT: the lowest footnote line sits ABOVE the topmost footer line — the
    // footnote block clears the footer just like the body does.
    expect(Math.max(...noteY)).toBeLessThan(Math.min(...footerY));
  });

  it('places the body bottom at |bottom| above the page bottom for a negative bottom margin (§17.6.11)', async () => {
    // §17.6.11 (pgMar/@bottom), the SYMMETRIC twin of the @top rule: a negative bottom
    // margin measures the main text from the bottom of the page extent by |bottom|
    // "regardless of the footer ... and therefore shall overlap the footer text". The
    // spec's own example (w:bottom="-720") keeps the body ½ inch (|bottom|) ABOVE the
    // page bottom. So with marginBottom = −10 the body's lowest line must clear the
    // bottom of the page by 10pt — the SAME extent a +10 bottom margin (no footer to
    // reserve against) gives it — NOT packed down to the raw −10 offset (which sits
    // BELOW the page bottom, off-canvas). Pre-fix the content height used the raw signed
    // marginBottom (pageHeight − marginTop − (−10) ⇒ a TALLER page), so the body packed
    // past the page bottom and diverged from the +10 placement.
    // A body long enough that page 0 fills to its content bottom in BOTH cases, so the
    // page-0 fill genuinely depends on the content height. A negative bottom margin must
    // NOT enlarge that height (pageHeight − marginTop − |bottom|); pre-fix it did
    // (… − (−10) ⇒ +10 taller), letting ~2 extra lines pack onto page 0 past the bottom.
    const longBody = (): BodyElement[] =>
      Array.from({ length: 70 }, () => para('BODY') as unknown as BodyElement);
    const bodyY = (calls: Call[]) => calls.filter((c) => c.text === 'BODY').map((c) => c.y);
    const negTall = await renderPage0(docWithFooter(longBody(), tallFooter, { marginBottom: -10 }));
    const negNone = await renderPage0(docWithFooter(longBody(), null, { marginBottom: -10 }));
    const posNone = await renderPage0(docWithFooter(longBody(), null, { marginBottom: 10 }));

    // Sanity: the tall footer is still painted (the negative margin only stops the
    // RESERVE, it does not suppress the footer).
    expect(negTall.filter((c) => c.text === 'FTR').length).toBeGreaterThan(0);

    // EXCEPTION: nothing is reserved — the body reaches the same depth WITH the tall
    // footer as WITHOUT one.
    expect(Math.max(...bodyY(negTall))).toBeCloseTo(Math.max(...bodyY(negNone)), 1);

    // MAGNITUDE: |−10| keeps the body bottom exactly where a +10 bottom margin (no
    // reserve) does. This is the spec's "measured from the page bottom by |bottom|"
    // rule, and it is what fixes the content-height computation under a negative margin.
    expect(Math.max(...bodyY(negNone))).toBeCloseTo(Math.max(...bodyY(posNone)), 1);

    // DIRECTION: the body's lowest line on page 0 stays within the page (above the page
    // bottom), not pushed off-canvas below it as the raw negative offset would.
    expect(Math.max(...bodyY(negNone))).toBeLessThanOrEqual(600);
  });
});
