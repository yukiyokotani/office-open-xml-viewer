import { describe, it, expect } from 'vitest';
import { layoutBodyModel } from './test-support/document-layout.test-support.js';
import type { LineLayout, TextPlacement } from './layout/types.js';
import type { BodyElement, DocParagraph, DocSettings, DocxTextRun, SectionProps } from './types';

// A run's explicit U+0020 advance (`__spaceAdvancePt`, set by producers
// other than DOCX such as the OpenDocument reader's balanced spaces) replaces
// the space glyph's advance in measurement, line breaking and placement.

/** Stub glyph advance: every character 0.5 em (10 pt at 20 pt). */
function makeCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p * 0.5,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

// A 360 pt band from x = 20 to x = 380.
const section: SectionProps = {
  pageWidth: 400, pageHeight: 400,
  marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
  headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
};

function textRun(text: string, extra: Partial<DocxTextRun> = {}): DocParagraph['runs'][number] {
  const run: DocxTextRun = {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 20, color: null, fontFamily: 'NotInMetrics', isLink: false, background: null,
    vertAlign: null, hyperlink: null, ...extra,
  };
  return { type: 'text', ...run } as DocParagraph['runs'][number];
}

function para(runs: DocParagraph['runs'], alignment: DocParagraph['alignment']): BodyElement {
  const p: DocParagraph = {
    alignment, indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs, defaultFontSize: 20, defaultFontFamily: 'NotInMetrics',
  };
  return { type: 'paragraph', ...p } as BodyElement;
}

function lines(body: BodyElement, settings?: DocSettings): LineLayout[] {
  const node = layoutBodyModel([body], section, makeCtx(), {}, settings).pages[0]!.layers.body[0]!;
  if (node.kind !== 'paragraph') throw new Error('expected a paragraph');
  return [...node.lines];
}

function texts(line: LineLayout): TextPlacement[] {
  return line.placements.filter((p): p is TextPlacement => p.kind === 'text')
    .sort((a, b) => a.bounds.xPt - b.bounds.xPt);
}

const space = (advance: number, extra: Partial<DocxTextRun> = {}) =>
  textRun(' ', { ...extra, __spaceAdvancePt: advance } as Partial<DocxTextRun>);

const xOf = (line: LineLayout, text: string): number =>
  texts(line).find((p) => p.text.startsWith(text))!.bounds.xPt;

describe('explicit space advance', () => {
  it('replaces the space glyph advance', () => {
    const [line] = lines(para([textRun('aaaa'), space(3), textRun('bbbb cccc')], 'left'));
    // aaaa = 40 pt, then a 3 pt space instead of the 10 pt glyph.
    expect(xOf(line!, 'bbbb')).toBeCloseTo(20 + 40 + 3, 6);
    // Other spaces keep their glyph advance.
    expect(xOf(line!, 'bbbb')).toBeCloseTo(63, 6);
    const [wide] = lines(para([textRun('aaaa'), space(25), textRun('bbbb')], 'left'));
    expect(xOf(wide!, 'bbbb')).toBeCloseTo(20 + 40 + 25, 6);
  });

  it('applies to every U+0020 in the run, before horizontal scaling', () => {
    const [line] = lines(para([textRun('aaaa'), textRun('  ', { __spaceAdvancePt: 4, charScale: 2 } as Partial<DocxTextRun>), textRun('bbbb')], 'left'));
    expect(xOf(line!, 'bbbb')).toBeCloseTo(20 + 40 + 8, 6);
  });

  it('takes part in line breaking', () => {
    // Seven words fit 360 pt with 10 pt spaces (340 pt); widening one space
    // by 25 pt (365 pt) pushes the seventh word to the next line.
    const words = (gap: number) => [
      textRun('aaaa bbbb'), space(gap), textRun('cccc dddd eeee ffff gggg'),
    ];
    const [fits] = lines(para(words(10), 'left'));
    expect(texts(fits!).map((p) => p.text).join('').trim()).toBe('aaaa bbbb cccc dddd eeee ffff gggg');
    const [wraps] = lines(para(words(35), 'left'));
    expect(texts(wraps!).map((p) => p.text).join('').trim()).toBe('aaaa bbbb cccc dddd eeee ffff');
  });
});
