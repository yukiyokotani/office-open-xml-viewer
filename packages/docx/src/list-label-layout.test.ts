import { describe, it, expect } from 'vitest';
import { layoutBodyModel } from './test-support/document-layout.test-support.js';
import type { LineLayout, ResourcePlacement, TextPlacement } from './layout/types.js';
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

  it('adds no metrics from a collapsed separator tab to the body line', () => {
    const big = { fontSize: 40 } as Partial<DocxTextRun>;
    const runs = [textRun('1.', { ...big, __listLabel: 'start' } as Partial<DocxTextRun>),
      textRun('\t', { ...big, __listLabel: 'start' } as Partial<DocxTextRun>), textRun('body')];
    const [, body] = lines(list(runs, 'left', 360));
    const [plain] = lines(para([textRun('body')], 'left'));
    expect(body!.advancePt).toBeCloseTo(plain!.advancePt, 6);
  });

  it('breaks the first body word inside the word on a label-only line', () => {
    // 30 characters (300 pt) do not fit the 280 pt after the tab stop.
    const [first, second] = lines(list([label('1.'), label('\t'), textRun('UUUUUUUUUUUUUUUUUUUUUUUUUUUUUU')]));
    expect(xOf(first!, 'U')).toBeCloseTo(100, 6);
    expect(texts(first!).at(-1)!.text.length).toBeGreaterThan(0);
    expect(texts(second!)[0]!.bounds.xPt).toBeCloseTo(60, 6);
  });

  it('breaks a first body word across formatting seams the same way', () => {
    const [first, second] = lines(list([
      label('1.'), label('\t'),
      textRun('UUUUUUUUUUUUUUU'), textRun('UUUUUUUUUUUUUUU', { color: 'ff0000' }),
    ]));
    expect(xOf(first!, 'U')).toBeCloseTo(100, 6);
    // 28 characters fit the 280 pt after the tab stop.
    expect(texts(first!).map((p) => p.text).join('').replace(/[^U]/g, '').length).toBe(28);
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

// A picture label is the label font's zero-advance metric, the picture and
// its separator, all flagged. LibreOffice gives the picture a box 0.5 pt
// larger on every side, sizes the line with the box top at the label font's
// ascent (so the box extends the descent at most), and centres the box in
// the natural line; proportional spacing adds its extra from the text.
describe('picture list labels', () => {
  const metric = (align: Align = 'start') => ({
    type: 'anchorHost', fontSize: 20, fontFamily: 'NotInMetrics', __listLabel: align,
  }) as unknown as DocParagraph['runs'][number];
  const picture = (widthPt: number, heightPt: number, align: Align = 'start') => ({
    type: 'image', imagePath: 'Pictures/a.png', mimeType: 'image/png',
    widthPt, heightPt, rotation: 0, flipH: false, flipV: false, anchor: false,
    anchorXPt: 0, anchorYPt: 0, anchorXFromMargin: false, anchorYFromPara: false,
    allowOverlap: false, __listLabel: align,
  }) as unknown as DocParagraph['runs'][number];
  const pictureList = (widthPt: number, heightPt: number, body: DocParagraph['runs'], extra: Partial<DocParagraph> = {}, align: Align = 'start', tab = true) =>
    ({ ...list([metric(align), picture(widthPt, heightPt, align), ...(tab ? [label('\t', align)] : []), ...body]), ...extra }) as BodyElement;
  const image = (line: LineLayout): ResourcePlacement =>
    line.placements.find((p): p is ResourcePlacement => p.kind === 'resource')!;
  const [plain] = lines(list([label('1.'), label('\t'), textRun('body')]));
  const natural = plain!.advancePt;
  const textTop = (line: LineLayout) => texts(line).find((p) => p.text.startsWith('body'))!.bounds.yPt - line.bounds.yPt;

  it('advances by the box and draws the picture inset within it', () => {
    const [line] = lines(pictureList(10, 8, [textRun('body')], {}, 'start', false));
    expect(image(line!).bounds.xPt).toBeCloseTo(40.5, 6);
    expect(image(line!).bounds.widthPt).toBeCloseTo(10, 6);
    expect(xOf(line!, 'body')).toBeCloseTo(51, 6);
    const [tabbed] = lines(pictureList(10, 8, [textRun('body')]));
    expect(xOf(tabbed!, 'body')).toBeCloseTo(100, 6);
  });

  it('centres a box shorter than the line without changing the line', () => {
    const [line] = lines(pictureList(10, 8, [textRun('body')]));
    expect(line!.advancePt).toBeCloseTo(natural, 6);
    expect(image(line!).bounds.yPt - line!.bounds.yPt).toBeCloseTo((natural - 9) / 2 + 0.5, 6);
    expect(textTop(line!)).toBeCloseTo(textTop(plain!), 6);
  });

  it('extends the line below the label font\'s ascent by a taller box, text kept at the top', () => {
    // Label font ascent 16 pt: a 31 pt box makes a 16 + 15 pt line.
    const [line] = lines(pictureList(10, 30, [textRun('body')]));
    expect(line!.advancePt).toBeCloseTo(31, 6);
    expect(image(line!).bounds.yPt - line!.bounds.yPt).toBeCloseTo(0.5, 6);
    expect(textTop(line!)).toBeCloseTo(textTop(plain!), 6);
  });

  it('sizes from the label font, not a taller run on the line', () => {
    // A 40 pt word (ascent 32, descent 8) and a 31 pt box below the 20 pt
    // label font's 16 pt ascent: descent 15, a 47 pt line, the box centred.
    const [line] = lines(pictureList(10, 30, [textRun('body', { fontSize: 40 })]));
    expect(line!.advancePt).toBeCloseTo(47, 6);
    expect(image(line!).bounds.yPt - line!.bounds.yPt).toBeCloseTo((47 - 31) / 2 + 0.5, 6);
  });

  it('adds proportional spacing from the text line, below the centred box', () => {
    const spacing = { lineSpacing: { value: 1.5, rule: 'auto', explicit: true } } as Partial<DocParagraph>;
    const [plainSpaced] = lines({ ...list([label('1.'), label('\t'), textRun('body')]), ...spacing } as BodyElement);
    const [small] = lines(pictureList(10, 8, [textRun('body')], spacing));
    expect(small!.advancePt).toBeCloseTo(plainSpaced!.advancePt, 6);
    expect(image(small!).bounds.yPt - small!.bounds.yPt).toBeCloseTo((natural - 9) / 2 + 0.5, 6);
    const [tall] = lines(pictureList(10, 30, [textRun('body')], spacing));
    expect(tall!.advancePt).toBeCloseTo(31 + plainSpaced!.advancePt - natural, 6);
    expect(image(tall!).bounds.yPt - tall!.bounds.yPt).toBeCloseTo(0.5, 6);
  });

  it('ends an end-aligned box and centres a centre-aligned one at the label position', () => {
    const [end] = lines(pictureList(10, 8, [textRun('body')], {}, 'end'));
    expect(image(end!).bounds.xPt).toBeCloseTo(40 - 11 + 0.5, 6);
    expect(xOf(end!, 'body')).toBeCloseTo(100, 6);
    const [centre] = lines(pictureList(10, 8, [textRun('body')], {}, 'center'));
    expect(image(centre!).bounds.xPt).toBeCloseTo(40 - 5.5 + 0.5, 6);
  });
});
