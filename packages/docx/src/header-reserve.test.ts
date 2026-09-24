import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, HeaderFooter, SectionProps } from './types';

// ECMA-376 §17.6.11 (pgMar/@top) — the SYMMETRIC twin of the footer rule in
// footer-reserve.test.ts. The main-document text TOP is placed at the GREATER of the
// top margin and the header's extent ("The value of top / The extent of the header
// text"), so a header taller than its top-margin allowance (marginTop − headerDistance)
// rises into the content area and the body must start BELOW it: Word never lays main
// text over a header ("the main text extent ends at the bottom of the header region").
// The paginator reserves that overflow at the TOP of every page's content area and the
// body's first line is pushed down by the same amount. A NEGATIVE top margin is the
// spec's explicit exception — the main text is then "measured from the top of the page
// extent regardless of the header ... and therefore shall overlap the header text", so
// nothing is reserved. These tests pin those rules with a synthetic doc whose header is
// far taller than its top margin (the header-side mirror of sample-13's masthead).

interface Call { text: string; x: number; y: number; }

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
    fillText(s: string, x: number, y: number) {
      calls.push({
        text: s,
        x: transform.translateX + transform.scaleX * x,
        y: transform.translateY + transform.scaleY * y,
      });
    },
    strokeText(s: string, x: number, y: number) {
      calls.push({
        text: s,
        x: transform.translateX + transform.scaleX * x,
        y: transform.translateY + transform.scaleY * y,
      });
    },
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  (ctx as unknown as { canvas: { width: number; height: number } }).canvas = {
    width: 400,
    height: 600,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, calls };
}

// Keep this geometry-only fixture outside the reference-font catalog. These
// tests exercise header reservation and vertical alignment; binding them to a
// real catalogued face would make their fixed synthetic Canvas metrics compete
// with that face's real design metrics.
const TEST_FONT_FAMILY = 'Header Reserve Test Serif';

function para(text: string): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: text
      ? [{
          type: 'text', text, bold: false, italic: false, underline: false,
          strikethrough: false, fontSize: 10, color: null, fontFamily: TEST_FONT_FAMILY,
          fontFamilyEastAsia: '', isLink: false, background: null, vertAlign: null, hyperlink: null,
        } as DocParagraph['runs'][number]]
      : [],
    defaultFontSize: 10, defaultFontFamily: TEST_FONT_FAMILY, widowControl: false,
  } as unknown as DocParagraph;
}

// pageHeight 600, margins 10, headerDistance 4 → top-margin allowance for the header
// is marginTop − headerDistance = 6pt. A header taller than 6pt overflows the top.
function docWithHeader(
  body: BodyElement[],
  header: HeaderFooter | null,
  opts: {
    marginTop?: number;
    pageHeight?: number;
    columns?: SectionProps['columns'];
    vAlign?: SectionProps['vAlign'];
  } = {},
): DocxDocumentModel {
  const section: SectionProps = {
    pageWidth: 400, pageHeight: opts.pageHeight ?? 600,
    marginTop: opts.marginTop ?? 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
    headerDistance: 4, footerDistance: 4, titlePage: false, evenAndOddHeaders: false,
    sectionStart: 'nextPage', columns: opts.columns ?? null,
    vAlign: opts.vAlign ?? null,
  } as SectionProps;
  return {
    section,
    body,
    headers: { default: header, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { [TEST_FONT_FAMILY]: 'roman' },
    footnotes: [],
  } as unknown as DocxDocumentModel;
}

async function renderPage(doc: DocxDocumentModel, pageIndex = 0): Promise<Call[]> {
  const { canvas, calls } = makeRecordingCanvas();
  await renderDocumentToCanvas(doc, canvas, pageIndex, { dpr: 1, width: 400 });
  return calls;
}

const renderPage0 = (doc: DocxDocumentModel): Promise<Call[]> => renderPage(doc, 0);

describe('header reserve — content never overlaps a tall header (ECMA-376 §17.6.11)', () => {
  // A few single-line body paragraphs (page 0 holds them comfortably) — enough to read
  // the topmost body line. The reserve, not a long body, is what places that line.
  const body = (): BodyElement[] =>
    Array.from({ length: 6 }, () => para('BODY') as unknown as BodyElement);
  // A header far taller than the 6pt allowance (6 lines ≈ ~70pt), so its bottom edge
  // sinks well below the top margin and into the body's region.
  const tallHeader: HeaderFooter = {
    body: Array.from({ length: 6 }, () => para('HDR') as unknown as BodyElement),
  };

  it('starts the body below a tall header so no line is painted into the header band', async () => {
    const withHeader = await renderPage0(docWithHeader(body(), tallHeader));
    const noHeader = await renderPage0(docWithHeader(body(), null));

    const bodyY = (calls: Call[]) => calls.filter((c) => c.text === 'BODY').map((c) => c.y);
    const headerY = (calls: Call[]) => calls.filter((c) => c.text === 'HDR').map((c) => c.y);

    const minBodyTall = Math.min(...bodyY(withHeader));
    const maxHeader = Math.max(...headerY(withHeader));
    const minBodyNone = Math.min(...bodyY(noHeader));

    // Sanity: the tall header is actually painted on page 0.
    expect(headerY(withHeader).length).toBeGreaterThan(0);

    // INVARIANT: with the tall header reserved, the topmost body line on page 0 sits
    // BELOW the bottom-most header line — body never overlaps the header.
    expect(minBodyTall).toBeGreaterThan(maxHeader);

    // NON-TRIVIALITY: the SAME body, with no header to reserve against, starts at the
    // top margin — ABOVE where the tall header's bottom sits — proving the invariant
    // above is not satisfied trivially and that the reservation, not a short header, is
    // what clears the header band.
    expect(minBodyNone).toBeLessThan(maxHeader);
  });

  it('starts EVERY newspaper column below a tall header (ECMA-376 §17.6.4 + §17.6.11)', async () => {
    // The reserve shrinks the content area from the top for the WHOLE page, not just
    // column 0. A short page so column 0 overflows into column 1 on page 0; a tall
    // header that overflows the top margin. Each newspaper column (§17.6.4) restarts
    // at the section's region top, which must be the RESERVED top, not the bare margin.
    const cols = { count: 2, spacePt: 20, equalWidth: true, sep: false, cols: [] };
    const manyBody = Array.from({ length: 40 }, () => para('BODY') as unknown as BodyElement);
    const calls = await renderPage0(docWithHeader(manyBody, tallHeader, { columns: cols, pageHeight: 300 }));

    const bodyCalls = calls.filter((c) => c.text === 'BODY');
    const headerY = calls.filter((c) => c.text === 'HDR').map((c) => c.y);

    // Preconditions: the header is painted, and the body genuinely reached the SECOND
    // column (≥2 distinct x positions) — otherwise the invariant below is vacuous.
    expect(headerY.length).toBeGreaterThan(0);
    const distinctX = new Set(bodyCalls.map((c) => Math.round(c.x)));
    expect(distinctX.size).toBeGreaterThanOrEqual(2);

    // INVARIANT: the topmost body line across BOTH columns sits below the bottom-most
    // header line. Pre-fix, column 1 reset to the bare top margin and painted into the
    // header band (its first line is the global minimum), so this caught the overlap.
    expect(Math.min(...bodyCalls.map((c) => c.y))).toBeGreaterThan(Math.max(...headerY));
  });

  it('centres body ink in the effective post-header-reserve text band', async () => {
    const calls = await renderPage0(docWithHeader(
      [para('BODY') as unknown as BodyElement],
      tallHeader,
      { vAlign: 'center' },
    ));
    const bodyBaseline = calls.find((call) => call.text === 'BODY')!.y;
    const headerBottom = Math.max(...calls.filter((call) => call.text === 'HDR').map((call) => call.y)) + 2;
    const effectiveBottom = 590;
    const expectedInkMidpoint = (headerBottom + effectiveBottom) / 2;
    const bodyInkMidpoint = bodyBaseline - 3;
    expect(bodyInkMidpoint).toBeCloseTo(expectedInkMidpoint, 0);
  });

  it('centres table-only body flow in the effective reserved band', async () => {
    const bodyTable = {
      type: 'table', colWidths: [380],
      rows: [{
        cells: [{
          content: [{ type: 'paragraph', ...para('TABLE') }], colSpan: 1, vMerge: null,
          borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
          background: null, vAlign: 'top', widthPt: 380,
        }],
        rowHeight: 20, rowHeightRule: 'exact', isHeader: false,
      }],
      borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
      cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
      jc: 'left', layout: 'fixed',
    } as unknown as BodyElement;
    const baseline = (calls: Call[]) => calls.find((call) => call.text === 'TABLE')!.y;
    const centredWithHeader = baseline(await renderPage0(
      docWithHeader([bodyTable], tallHeader, { vAlign: 'center' }),
    ));
    const centredWithoutHeader = baseline(await renderPage0(
      docWithHeader([bodyTable], null, { vAlign: 'center' }),
    ));
    const topWithHeader = baseline(await renderPage0(
      docWithHeader([bodyTable], tallHeader, { vAlign: 'top' }),
    ));
    const topWithoutHeader = baseline(await renderPage0(
      docWithHeader([bodyTable], null, { vAlign: 'top' }),
    ));
    const effectiveTopReserve = topWithHeader - topWithoutHeader;

    expect(effectiveTopReserve).toBeGreaterThan(0);
    // Raising the effective band top by R translates a centred retained table by
    // exactly R/2. Comparing the same fragment removes any glyph-baseline estimate.
    expect(centredWithHeader - centredWithoutHeader).toBeCloseTo(effectiveTopReserve / 2, 8);
  });

  it('centres a continuation page in the same effective reserved band', async () => {
    const calls = await renderPage(
      docWithHeader(
        Array.from({ length: 80 }, () => para('BODY') as unknown as BodyElement),
        tallHeader,
        { vAlign: 'center', pageHeight: 300 },
      ),
      1,
    );
    const bodyY = calls.filter((call) => call.text === 'BODY').map((call) => call.y);
    const headerBottom = Math.max(...calls.filter((call) => call.text === 'HDR').map((call) => call.y)) + 2;
    expect(bodyY.length).toBeGreaterThan(0);
    const bodyInkMidpoint = (Math.min(...bodyY) - 8 + Math.max(...bodyY) + 2) / 2;
    expect(bodyInkMidpoint).toBeCloseTo((headerBottom + 290) / 2, 0);
  });

  it('places the body at |top| below the page top for a negative top margin (§17.6.11)', async () => {
    // §17.6.11 (pgMar/@top): a negative top margin measures the main text from the top
    // of the page extent by |top| "regardless of the header ... and therefore shall
    // overlap the header text". The spec's own example (w:top="-720") puts the body
    // ½ inch (|top|) BELOW the page top. So with marginTop = −10 the body's first line
    // must land 10pt below the page top — the SAME place a +10 top margin (no header to
    // reserve against) puts it — NOT at the raw −10pt offset (which sits ABOVE the page
    // top, off-canvas). Pre-fix the render path used the raw signed marginTop, so the
    // body was painted above the page top and diverged from the +10 placement.
    const bodyY = (calls: Call[]) => calls.filter((c) => c.text === 'BODY').map((c) => c.y);
    const negTall = await renderPage0(docWithHeader(body(), tallHeader, { marginTop: -10 }));
    const negNone = await renderPage0(docWithHeader(body(), null, { marginTop: -10 }));
    const posNone = await renderPage0(docWithHeader(body(), null, { marginTop: 10 }));

    // Sanity: the tall header is still painted (the negative margin only stops the
    // RESERVE, it does not suppress the header).
    expect(negTall.filter((c) => c.text === 'HDR').length).toBeGreaterThan(0);

    // EXCEPTION: nothing is reserved — the body starts at the same y WITH the tall
    // header as WITHOUT one (a naive max(0, header extent − top) would wrongly push
    // it down here).
    expect(Math.min(...bodyY(negTall))).toBeCloseTo(Math.min(...bodyY(negNone)), 1);

    // MAGNITUDE: |−10| places the body exactly where a +10 top margin (no reserve)
    // does. This is the spec's "measured from the page top by |top|" rule.
    expect(Math.min(...bodyY(negNone))).toBeCloseTo(Math.min(...bodyY(posNone)), 1);

    // DIRECTION: the body's first line sits BELOW the page top (positive y), not
    // off-canvas above it as the raw negative offset would place it.
    expect(Math.min(...bodyY(negNone))).toBeGreaterThan(0);
  });
});
