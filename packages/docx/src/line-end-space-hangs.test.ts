import { describe, it, expect } from 'vitest';
import { layoutBodyModel } from './test-support/document-layout.test-support.js';
import type { LineLayout, TextPlacement } from './layout/types.js';
import type { BodyElement, DocParagraph, DocxTextRun, SectionProps } from './types';

// WORD_LINE_END_SPACE_HANGS: the U+0020 that ends a soft-wrapped LTR line
// stands outside justification and centre/right alignment. Windows Word ends
// every justified line flush at the margin in compatibility modes unset, 11,
// 12 and 14 (mode 15 already fitted the visible width), and places
// centre/right-aligned wrapped lines by their visible text.

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

function textRun(text: string): DocParagraph['runs'][number] {
  const run: DocxTextRun = {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 20, color: null, fontFamily: 'NotInMetrics', isLink: false, background: null,
    vertAlign: null, hyperlink: null,
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

function lines(body: BodyElement): LineLayout[] {
  const node = layoutBodyModel([body], section, makeCtx()).pages[0]!.layers.body[0]!;
  if (node.kind !== 'paragraph') throw new Error('expected a paragraph');
  return [...node.lines];
}

function texts(line: LineLayout): TextPlacement[] {
  return line.placements.filter((p): p is TextPlacement => p.kind === 'text')
    .sort((a, b) => a.bounds.xPt - b.bounds.xPt);
}

const visibleStart = (line: LineLayout): number =>
  texts(line).find((p) => p.text.trim() !== '')!.bounds.xPt;

/** Right edge of the line's visible text (each character 10 pt). */
function visibleEnd(line: LineLayout): number {
  const last = texts(line).filter((p) => p.text.trim() !== '').at(-1)!;
  const trailing = last.text.length - last.text.trimEnd().length;
  return last.bounds.xPt + last.advancePt - trailing * 10;
}

// Seven 4-letter words (340 pt visible, 350 pt with the line-end space) fit
// the 360 pt band; the eighth wraps.
const WORDS = 'aaaa bbbb cccc dddd eeee ffff gggg hhhh';

describe('WORD_LINE_END_SPACE_HANGS', () => {
  it('justifies a wrapped line flush to the band edge, the line-end space hanging', () => {
    const [first, last] = lines(para([textRun(WORDS)], 'both'));
    expect(texts(first!).map((p) => p.text).join('')).toBe('aaaa bbbb cccc dddd eeee ffff gggg ');
    expect(visibleStart(first!)).toBeCloseTo(20, 6);
    expect(visibleEnd(first!)).toBeCloseTo(380, 6);
    // The final line is not justified.
    expect(visibleStart(last!)).toBeCloseTo(20, 6);
  });

  it('centres and right-aligns a wrapped line by its visible text', () => {
    const [centred, centredLast] = lines(para([textRun(WORDS)], 'center'));
    expect(visibleStart(centred!)).toBeCloseTo(20 + (360 - 340) / 2, 6);
    expect(visibleStart(centredLast!)).toBeCloseTo(20 + (360 - 40) / 2, 6);
    const [right, rightLast] = lines(para([textRun(WORDS)], 'right'));
    expect(visibleEnd(right!)).toBeCloseTo(380, 6);
    expect(visibleEnd(rightLast!)).toBeCloseTo(380, 6);
  });

  it('keeps the space before a manual break inside the aligned width', () => {
    // Outside the observation: a line ended by a break keeps its geometry.
    const [line] = lines(para([textRun('aaaa '), { type: 'break', breakType: 'line' } as DocParagraph['runs'][number], textRun('bbbb')], 'right'));
    expect(visibleEnd(line!)).toBeCloseTo(370, 6);
  });
});
