import { describe, expect, it } from 'vitest';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import { transformRect } from './layout/coordinate-space.js';
import type { LayoutPage, LayoutRect, ParagraphLayout } from './layout/types.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, SectionProps } from './types.js';

// ECMA-376 §17.6.11 w:pgMar names physical page edges and §17.6.20 /
// §17.18.93 (Part 4 §14.11.7) choose the line progression: tbRl lines flow
// right to left, tbLrV (≡ lrV) lines flow left to right. The first line of
// the body therefore sits against the physical right margin in tbRl and
// against the physical left margin in tbLrV, and every line starts at the
// physical top margin. Word 16 (Windows) exports of A4 sections with distinct
// margins on all four edges confirm both placements, with btLr on the tbRl
// frame. The margins here are distinct on every edge so a mirrored or
// rotated mapping cannot pass.

function makeCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    measureText: (text: string) => {
      const size = px();
      return {
        width: [...text].length * size,
        fontBoundingBoxAscent: size * 0.8, fontBoundingBoxDescent: size * 0.2,
        actualBoundingBoxAscent: size * 0.8, actualBoundingBoxDescent: size * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, fillText() {}, strokeText() {}, beginPath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fillRect() {}, drawImage() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1, letterSpacing: '0px',
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
  } as unknown as CanvasRenderingContext2D;
}

(globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = class {
  getContext() { return makeCtx(); }
};

function para(text: string): BodyElement {
  return {
    type: 'paragraph',
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: [{
      type: 'text', text, bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: 'NotInMetrics', isLink: false, background: null,
      vertAlign: null, hyperlink: null,
    }],
    defaultFontSize: 10, defaultFontFamily: 'NotInMetrics', widowControl: false,
  } as unknown as DocParagraph & BodyElement;
}

const PAGE = { width: 600, height: 800 };
const MARGINS = { top: 60, right: 90, bottom: 80, left: 40 };

function layoutPage(textDirection: string): LayoutPage {
  const section = {
    pageWidth: PAGE.width, pageHeight: PAGE.height,
    marginTop: MARGINS.top, marginRight: MARGINS.right,
    marginBottom: MARGINS.bottom, marginLeft: MARGINS.left,
    headerDistance: 36, footerDistance: 36, titlePage: false, evenAndOddHeaders: false,
    textDirection,
  } as SectionProps;
  const document = {
    section, body: [para('AB'), para('CD')],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: {},
  } as unknown as DocxDocumentModel;
  const layout = layoutDocument(
    document,
    createLayoutServices(document, { measureContext: makeCtx() }),
    { currentDateMs: 0 },
  );
  return layout.pages[0]!;
}

/** Physical rectangles of the body paragraphs' first lines, through the
 * page's own paint matrix. */
function physicalParagraphs(page: LayoutPage): LayoutRect[] {
  const matrix = page.sectionRegions[0]!.coordinateSpace.logicalToPhysical;
  return page.layers.body
    .filter((node): node is ParagraphLayout => node.kind === 'paragraph')
    .map((node) => transformRect(matrix, node.lines[0]!.bounds));
}

describe('vertical section margins on the physical page (§17.6.11, §17.6.20)', () => {
  it('starts a tbLrV body at the physical left and top margins and advances rightward', () => {
    const page = layoutPage('tbLrV');
    const [first, second] = physicalParagraphs(page);

    expect(page.section.geometry).toMatchObject({
      pageWidth: PAGE.height, pageHeight: PAGE.width,
      marginTop: MARGINS.left, marginBottom: MARGINS.right,
      marginLeft: MARGINS.top, marginRight: MARGINS.bottom,
    });
    expect(first!.xPt).toBeCloseTo(MARGINS.left, 6);
    expect(first!.yPt).toBeCloseTo(MARGINS.top, 6);
    expect(second!.xPt).toBeCloseTo(first!.xPt + first!.widthPt, 6);
  });

  it.each(['tbRl', 'btLr'])(
    'starts a %s body at the physical right and top margins and advances leftward',
    (textDirection) => {
      const [first, second] = physicalParagraphs(layoutPage(textDirection));

      expect(first!.xPt + first!.widthPt).toBeCloseTo(PAGE.width - MARGINS.right, 6);
      expect(first!.yPt).toBeCloseTo(MARGINS.top, 6);
      expect(second!.xPt + second!.widthPt).toBeCloseTo(first!.xPt, 6);
    },
  );
});
