import { describe, expect, it } from 'vitest';
import { layoutBodyModel } from './test-support/document-layout.test-support.js';
import type { BodyElement, DocParagraph, DocxTextRun, FramePr, SectionProps } from './types';

// ECMA-376 §17.3.1.11: a drop cap's height is `lines` lines of its anchor
// paragraph. The anchor is a NON-empty paragraph, so the empty-mark face rule
// (WORD_EMPTY_MARK_ASCII_FACE) must not size it: with East Asian anchor text the
// established East Asian line keeps the three-line exclusion.

function context(): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    measureText: (text: string) => {
      const size = px();
      // East Asian face: a 1.6 em box (16pt at 10pt); ascii face: 1.0 em.
      const eastAsian = font.includes('Anchor EA');
      return {
        width: [...text].length * size,
        fontBoundingBoxAscent: size * (eastAsian ? 1.28 : 0.8),
        fontBoundingBoxDescent: size * (eastAsian ? 0.32 : 0.2),
        actualBoundingBoxAscent: size * (eastAsian ? 1.28 : 0.8),
        actualBoundingBoxDescent: size * (eastAsian ? 0.32 : 0.2),
      } as TextMetrics;
    },
    save() {}, restore() {}, fillText() {}, strokeText() {}, beginPath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fillRect() {}, drawImage() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
  } as unknown as CanvasRenderingContext2D;
}

const section: SectionProps = {
  pageWidth: 220, pageHeight: 400,
  marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
  headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
} as SectionProps;

function run(text: string, fontSize: number): DocParagraph['runs'][number] {
  const value: DocxTextRun = {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize, color: null, fontFamily: 'Anchor Latin', fontFamilyEastAsia: 'Anchor EA',
    isLink: false, background: null, vertAlign: null, hyperlink: null,
  } as DocxTextRun;
  return { type: 'text', ...value } as DocParagraph['runs'][number];
}

function paragraph(text: string, fontSize: number, framePr?: FramePr): BodyElement {
  const value: DocParagraph = {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: [run(text, fontSize)],
    defaultFontSize: fontSize, defaultFontFamily: 'Anchor Latin', defaultFontFamilyEastAsia: 'Anchor EA',
    ...(framePr ? { framePr } : {}),
  } as DocParagraph;
  return { type: 'paragraph', ...value } as BodyElement;
}

describe('drop-cap anchor line (§17.3.1.11)', () => {
  it('sizes a three-line drop cap from the East Asian anchor line, not the empty-mark ascii face', () => {
    const frame: FramePr = {
      dropCap: 'drop', lines: 3, wrap: 'around', hAnchor: 'text', vAnchor: 'text',
      hRule: 'auto', hSpace: 0, vSpace: 0,
    };
    const layout = layoutBodyModel([
      paragraph('W', 10, frame),
      paragraph('あ'.repeat(60), 10),
    ], section, context());
    const anchor = layout.pages[0]!.layers.roots.map((root) => root.node)
      .find((node) => node.kind === 'paragraph' && node.ordinaryFlow);
    const lineStarts = anchor?.kind === 'paragraph' ? anchor.lines.map((line) => line.bounds.xPt) : [];
    // Three 16pt East Asian anchor lines (48pt) wrap beside the 10pt cap; the
    // ascii mark face (3 x 10pt = 30pt) would release the third line.
    expect(lineStarts.slice(0, 4)).toEqual([30, 30, 30, 20]);
  });
});
