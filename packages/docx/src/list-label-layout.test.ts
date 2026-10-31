import { describe, it, expect } from 'vitest';
import { layoutBodyModel } from './test-support/document-layout.test-support.js';
import type { LineLayout, TextPlacement } from './layout/types.js';
import type { BodyElement, DocParagraph, DocSettings, DocxTextRun, SectionProps } from './types';

// List labels emitted as leading runs (`__listLabel`, set by producers other
// than DOCX such as the OpenDocument reader) follow LibreOffice's label
// geometry: label lines keep the first-line indent through the body's first
// line, the separator tab breaks before a stop at the line end and has no
// width at a line start, a label-only line fits the next word as an empty
// line would, label spaces take no justification, and end/centre labels end
// or centre at the label position.

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

function para(runs: DocParagraph['runs'], alignment: DocParagraph['alignment'], extra: Partial<DocParagraph> = {}): BodyElement {
  const p: DocParagraph = {
    alignment, indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs, defaultFontSize: 20, defaultFontFamily: 'NotInMetrics', ...extra,
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

type Align = 'start' | 'end' | 'center';
const label = (text: string, align: Align = 'start') =>
  textRun(text, { __listLabel: align } as Partial<DocxTextRun>);

// M = 40 pt, I = -20 pt: the label starts at x = 20 + 20 = 40 (the band
// starts at x = 20), later lines at x = 60; the list tab stop T = 80 pt
// (x = 100).
const list = (runs: DocParagraph['runs'], alignment: DocParagraph['alignment'] = 'left', tab = 80) =>
  para(runs, alignment, {
    indentLeft: 40, indentFirst: -20,
    tabStops: [{ pos: tab, alignment: 'left', leader: 'none' }],
    __noHangingIndentTab: true,
  } as Partial<DocParagraph>);

const xOf = (line: LineLayout, text: string): number =>
  texts(line).find((p) => p.text.startsWith(text))!.bounds.xPt;

describe('list labels emitted as leading runs', () => {
  it('starts the body at the tab stop after the label', () => {
    const [line] = lines(list([label('1.'), label('\t'), textRun('body')]));
    expect(xOf(line!, '1.')).toBeCloseTo(40, 6);
    expect(xOf(line!, 'body')).toBeCloseTo(100, 6);
  });

  it('keeps the first-line indent on wrapped label lines and the body\'s first line', () => {
    // A 40-character label (400 pt) wraps in the 340 pt first-line band.
    const [first, second, third] = lines(list([
      label('PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP PPPPP'), label('\t'),
      textRun('aaaa bbbb cccc dddd eeee ffff gggg hhhh iiii jjjj kkkk'),
    ]));
    expect(texts(first!)[0]!.bounds.xPt).toBeCloseTo(40, 6);
    expect(texts(second!)[0]!.bounds.xPt).toBeCloseTo(40, 6);
    expect(texts(third!)[0]!.bounds.xPt).toBeCloseTo(60, 6);
  });

  it('collapses a separator tab that starts a line and breaks before a stop at the line end', () => {
    // T at the band end (x = 380): the body starts the next line at the
    // first-line indent.
    const [first, second] = lines(list([label('1.'), label('\t'), textRun('body')], 'left', 360));
    expect(texts(first!).map((p) => p.text).join('').trim()).toBe('1.');
    expect(xOf(second!, 'body')).toBeCloseTo(40, 6);
  });

  it('breaks the first body word inside the word on a label-only line', () => {
    // 30 characters (300 pt) do not fit the 280 pt after the tab stop.
    const [first, second] = lines(list([label('1.'), label('\t'), textRun('UUUUUUUUUUUUUUUUUUUUUUUUUUUUUU')]));
    expect(xOf(first!, 'U')).toBeCloseTo(100, 6);
    expect(texts(first!).at(-1)!.text.length).toBeGreaterThan(0);
    expect(texts(second!)[0]!.bounds.xPt).toBeCloseTo(60, 6);
  });

  it('leaves label spaces out of justification', () => {
    const [line] = lines(list([
      label('A B'), textRun(' aaaa bbbb cccc dddd eeee ffff gggg hhhh'),
    ], 'both'));
    expect(xOf(line!, 'B')).toBeCloseTo(60, 6);
  });

  it('ends an end-aligned label and centres a centre-aligned one at the label position', () => {
    const [end] = lines(list([label('iii.', 'end'), label('\t', 'end'), textRun('body')]));
    expect(xOf(end!, 'iii.')).toBeCloseTo(0, 6);
    expect(xOf(end!, 'body')).toBeCloseTo(100, 6);
    const [centre] = lines(list([label('ii.', 'center'), label('\t', 'center'), textRun('body')]));
    expect(xOf(centre!, 'ii.')).toBeCloseTo(25, 6);
    expect(xOf(centre!, 'body')).toBeCloseTo(100, 6);
  });
});
