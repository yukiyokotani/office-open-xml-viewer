import { describe, expect, it } from 'vitest';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import type { LineLayout, ParagraphLayout, TextPlacement } from './layout/types.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, SectionProps } from './types.js';

// WORD_VERTICAL_LINE_CENTRED_BOX, WORD_VERTICAL_RUBY_UPRIGHT_GUIDE and
// WORD_RUBY_ALIGN_PLACEMENT over a stub face (0.8/0.2 em box, one em per
// character) that no metric table describes.

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

type Ruby = Readonly<{ text: string; fontSizePt: number; hpsRaisePt: number; align?: string }>;
type Run = Readonly<{ text: string; ruby?: Ruby; position?: number }>;

function typographyInput(ruby: Ruby) {
  const value = <T>(v: T | undefined) => v === undefined
    ? { status: 'missing', raw: null, value: null }
    : { status: 'valid', raw: String(v), value: v };
  return {
    ruby: {
      align: value(ruby.align), baseFontSizePt: value(undefined),
      raisePt: value(ruby.hpsRaisePt), language: value(undefined), guideRuns: [],
    },
  };
}

function layoutParagraph(
  runs: readonly Run[],
  textDirection?: string,
  lineSpacing: DocParagraph['lineSpacing'] = null,
): Readonly<{ node: ParagraphLayout; lines: readonly LineLayout[] }> {
  const paragraph = {
    type: 'paragraph',
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing, numbering: null, tabStops: [],
    runs: runs.map((run) => ({
      type: 'text', text: run.text, bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: 'NotInMetrics', isLink: false, background: null,
      vertAlign: null, hyperlink: null,
      ...(run.position !== undefined ? { position: run.position } : {}),
      ...(run.ruby ? {
        ruby: { text: run.ruby.text, fontSizePt: run.ruby.fontSizePt, hpsRaisePt: run.ruby.hpsRaisePt },
        typographyInput: typographyInput(run.ruby),
      } : {}),
    })),
    defaultFontSize: 10, defaultFontFamily: 'NotInMetrics', widowControl: false,
  } as unknown as DocParagraph & BodyElement;
  const section = {
    pageWidth: 600, pageHeight: 800,
    marginTop: 60, marginRight: 90, marginBottom: 80, marginLeft: 40,
    headerDistance: 36, footerDistance: 36, titlePage: false, evenAndOddHeaders: false,
    ...(textDirection ? { textDirection } : {}),
  } as SectionProps;
  const document = {
    section, body: [paragraph],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: {},
  } as unknown as DocxDocumentModel;
  const page = layoutDocument(
    document,
    createLayoutServices(document, { measureContext: makeCtx() }),
    { currentDateMs: 0 },
  ).pages[0]!;
  const node = page.layers.body.find((entry): entry is ParagraphLayout => entry.kind === 'paragraph')!;
  return { node, lines: node.lines };
}

function textPlacements(line: LineLayout): TextPlacement[] {
  return line.placements.filter((entry): entry is TextPlacement => entry.kind === 'text');
}

describe('vertical line box (WORD_VERTICAL_LINE_CENTRED_BOX)', () => {
  it.each([
    // The stub face has no admitted metrics: East Asian text takes the Far
    // East single-line class, Latin text its own box.
    { text: 'あいう', extra: 10 / 30 * 1.3 },
    { text: 'Abc', extra: 10 / 30 },
  ])('adds the measured fraction of the em to the single line of $text', ({ text, extra }) => {
    const horizontal = layoutParagraph([{ text }]).lines[0]!;
    for (const direction of ['tbRl', 'tbLrV']) {
      const vertical = layoutParagraph([{ text }], direction).lines[0]!;
      expect(vertical.advancePt).toBeCloseTo(horizontal.advancePt + extra, 6);
      // The axis the glyphs paint on lies halfway across the line.
      expect(vertical.baselinePt - vertical.bounds.yPt).toBeCloseTo(vertical.advancePt / 2, 6);
    }
    // Horizontal lines keep their ascent/descent partition.
    expect(horizontal.baselinePt - horizontal.bounds.yPt).not.toBeCloseTo(horizontal.advancePt / 2, 3);
  });

  it('keeps the extra space of an automatic multiple after the line in strip order', () => {
    const single = layoutParagraph([{ text: 'あ' }], 'tbRl').lines[0]!.advancePt;
    const auto2 = { rule: 'auto', value: 2, explicit: true } as unknown as DocParagraph['lineSpacing'];
    const tbRl = layoutParagraph([{ text: 'あ' }], 'tbRl', auto2).lines[0]!;
    expect(tbRl.advancePt).toBeCloseTo(2 * single, 6);
    expect(tbRl.baselinePt - tbRl.bounds.yPt).toBeCloseTo(single / 2, 6);
    // A tbLrV line is reflected across its band at paint: the extra space
    // moves to the other logical side so it still follows the line.
    const tbLrV = layoutParagraph([{ text: 'あ' }], 'tbLrV', auto2).lines[0]!;
    expect(tbLrV.bounds.yPt + tbLrV.advancePt - tbLrV.baselinePt).toBeCloseTo(single / 2, 6);
  });

  it('puts an atLeast minimum before the line in strip order', () => {
    const single = layoutParagraph([{ text: 'あ' }], 'tbRl').lines[0]!.advancePt;
    const atLeast = { rule: 'atLeast', value: 30, explicit: true } as unknown as DocParagraph['lineSpacing'];
    const tbRl = layoutParagraph([{ text: 'あ' }], 'tbRl', atLeast).lines[0]!;
    expect(tbRl.advancePt).toBeCloseTo(30, 6);
    expect(tbRl.bounds.yPt + tbRl.advancePt - tbRl.baselinePt).toBeCloseTo(single / 2, 6);
    const tbLrV = layoutParagraph([{ text: 'あ' }], 'tbLrV', atLeast).lines[0]!;
    expect(tbLrV.baselinePt - tbLrV.bounds.yPt).toBeCloseTo(single / 2, 6);
  });
});

describe('w:position and empty marks in vertical lines (WORD_VERTICAL_LINE_CENTRED_BOX)', () => {
  const single = () => layoutParagraph([{ text: '\u3042' }], 'tbRl').lines[0]!.advancePt;

  it('moves a positioned run\'s whole box, lengthening the line by the raise', () => {
    const line = layoutParagraph([{ text: '\u3042' }, { text: '\u3044', position: 3 }], 'tbRl').lines[0]!;
    expect(line.advancePt).toBeCloseTo(single() + 3, 6);
    expect(line.baselinePt - line.bounds.yPt).toBeCloseTo(single() / 2 + 3, 6);
  });

  it('keeps a uniformly positioned line on its axis, growing only by a raise past the tail half', () => {
    const raised = layoutParagraph([{ text: '\u3042', position: 3 }], 'tbRl').lines[0]!;
    expect(raised.advancePt).toBeCloseTo(single(), 6);
    const far = layoutParagraph([{ text: '\u3042', position: 12 }], 'tbRl').lines[0]!;
    expect(far.advancePt).toBeCloseTo(single() / 2 + 12, 6);
    expect(far.baselinePt - far.bounds.yPt).toBeCloseTo(single() / 2, 6);
  });

  it('gives an empty paragraph the vertical single line of its mark', () => {
    const horizontal = layoutParagraph([]).node;
    const vertical = layoutParagraph([], 'tbRl').node;
    expect(vertical.advancePt - horizontal.advancePt).toBeCloseTo(10 / 30, 6);
  });
});

describe('vertical ruby guide (WORD_VERTICAL_RUBY_UPRIGHT_GUIDE)', () => {
  it.each(['tbRl', 'tbLrV'])('plans a %s guide as upright and turned cells on its raised axis', (direction) => {
    const ruby = { text: 'かんAB', fontSizePt: 5, hpsRaisePt: 9 };
    const line = layoutParagraph([{ text: '漢字', ruby }], direction).lines[0]!;
    const placement = textPlacements(line).find((entry) => entry.ruby)!;
    const ops = placement.ruby!.paintOps;
    expect(ops.map((op) => op.glyphOrientation)).toEqual(['upright', 'upright', 'sideways']);
    for (const op of ops) expect(op.origin.yPt).toBeCloseTo(line.baselinePt - 9, 6);
    // Upright cells are centred on their one-em advance; the guide is
    // centred along the two-em base.
    expect(ops[0]!.origin.xPt - placement.origin.xPt).toBeCloseTo((20 - 20) / 2 + 2.5, 6);
    // The guide reserves half of its own vertical single line (the stub box
    // at 5pt plus its Far East addition) beyond its axis.
    expect(line.baselinePt - line.bounds.yPt).toBeCloseTo(9 + (5 + 5 / 30 * 1.3) / 2, 6);
  });

  it('stands a rightVertical guide half its size nearer the base', () => {
    const ruby = { text: 'か', fontSizePt: 5, hpsRaisePt: 9, align: 'rightVertical' };
    const line = layoutParagraph([{ text: '漢', ruby }], 'tbRl').lines[0]!;
    const op = textPlacements(line).find((entry) => entry.ruby)!.ruby!.paintOps[0]!;
    expect(op.origin.yPt).toBeCloseTo(line.baselinePt - 9 + 2.5, 6);
  });
});

describe('ruby alignment (WORD_RUBY_ALIGN_PLACEMENT)', () => {
  const guideStarts = (align: string, base: string, guide: string): number[] => {
    const ruby = { text: guide, fontSizePt: 5, hpsRaisePt: 9, align };
    const line = layoutParagraph([{ text: base, ruby }]).lines[0]!;
    const placement = textPlacements(line).find((entry) => entry.ruby)!;
    return placement.ruby!.paintOps.map((op) => op.origin.xPt - placement.origin.xPt);
  };

  it.each([
    // A 30pt base over a 10pt two-letter guide: 20pt of slack.
    { align: 'center', starts: [10, 15] },
    { align: 'left', starts: [0, 5] },
    { align: 'right', starts: [20, 25] },
    { align: 'distributeLetter', starts: [0, 25] },
    { align: 'distributeSpace', starts: [5, 20] },
  ])('places a shorter guide by $align', ({ align, starts }) => {
    const actual = guideStarts(align, '東京都', 'とう');
    actual.forEach((value, index) => expect(value).toBeCloseTo(starts[index]!, 6));
  });

  it.each([
    { align: 'center', baseStart: 7.5 },
    { align: 'left', baseStart: 0 },
    { align: 'right', baseStart: 15 },
  ])('widens the run to a longer guide and places the base by $align', ({ align, baseStart }) => {
    const ruby = { text: 'ひがしがわ', fontSizePt: 5, hpsRaisePt: 9, align };
    const line = layoutParagraph([{ text: '東', ruby }, { text: '後' }]).lines[0]!;
    const [annotated, next] = textPlacements(line);
    expect(next!.origin.xPt - annotated!.origin.xPt).toBeCloseTo(25, 6);
    expect(annotated!.ruby!.paintOps[0]!.origin.xPt - annotated!.origin.xPt).toBeCloseTo(0, 6);
    expect(annotated!.paintOps[0]!.offset.xPt).toBeCloseTo(baseStart, 6);
  });
});
