import { describe, expect, it } from 'vitest';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import type { LineLayout, ParagraphLayout, TextPlacement } from './layout/types.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, SectionProps } from './types.js';
import { createFontResolver } from './layout/font-service.js';
import { createTextLayoutService } from './layout/text.js';
import { paragraphMarkLineMetrics } from './line-layout.js';
import { wordVerticalSingleLinePx } from './layout/line-compatibility.js';

// WORD_VERTICAL_LINE_CENTRED_BOX, WORD_VERTICAL_RUBY_UPRIGHT_GUIDE and
// WORD_RUBY_ALIGN_PLACEMENT over a stub face (0.8/0.2 em box, one em per
// character) that no metric table describes.

/** Characters every stub context has measured (a work counter). */
let measuredCharacters = 0;

function makeCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    measureText: (text: string) => {
      measuredCharacters += text.length;
      const size = px();
      return {
        // One em per character; the pair "AV" kerns 0.2 em tighter.
        width: [...text].length * size - (text.match(/AV/g)?.length ?? 0) * size * 0.2,
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
type Run = Readonly<{ text: string; ruby?: Ruby; position?: number; fontFamily?: string; eastAsianVert?: boolean; fontSize?: number; extra?: Record<string, unknown> }>;

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

function paragraphOf(
  runs: readonly Run[],
  lineSpacing: DocParagraph['lineSpacing'] = null,
  markFontFamily = 'NotInMetrics',
): DocParagraph & BodyElement {
  return {
    type: 'paragraph',
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing, numbering: null, tabStops: [],
    runs: runs.map((run) => ({
      type: 'text', text: run.text, bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: run.fontSize ?? 10, color: null, fontFamily: run.fontFamily ?? 'NotInMetrics', isLink: false, background: null,
      ...(run.eastAsianVert ? { eastAsianVert: true } : {}),
      vertAlign: null, hyperlink: null,
      ...(run.position !== undefined ? { position: run.position } : {}),
      ...run.extra,
      ...(run.ruby ? {
        ruby: { text: run.ruby.text, fontSizePt: run.ruby.fontSizePt, hpsRaisePt: run.ruby.hpsRaisePt },
        typographyInput: typographyInput(run.ruby),
      } : {}),
    })),
    defaultFontSize: 10, defaultFontFamily: markFontFamily, widowControl: false,
  } as unknown as DocParagraph & BodyElement;
}

function layoutBody(
  body: readonly BodyElement[],
  textDirection?: string,
  sectionExtra: Partial<SectionProps> = {},
) {
  const section = {
    pageWidth: 600, pageHeight: 800,
    marginTop: 60, marginRight: 90, marginBottom: 80, marginLeft: 40,
    headerDistance: 36, footerDistance: 36, titlePage: false, evenAndOddHeaders: false,
    ...(textDirection ? { textDirection } : {}),
    ...sectionExtra,
  } as SectionProps;
  const document = {
    section, body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: {},
  } as unknown as DocxDocumentModel;
  return layoutDocument(
    document,
    createLayoutServices(document, { measureContext: makeCtx() }),
    { currentDateMs: 0 },
  ).pages[0]!;
}

function layoutParagraph(
  runs: readonly Run[],
  textDirection?: string,
  lineSpacing: DocParagraph['lineSpacing'] = null,
  sectionExtra: Partial<SectionProps> = {},
  markFontFamily = 'NotInMetrics',
): Readonly<{ node: ParagraphLayout; lines: readonly LineLayout[] }> {
  const page = layoutBody([paragraphOf(runs, lineSpacing, markFontFamily)], textDirection, sectionExtra);
  const node = page.layers.body.find((entry): entry is ParagraphLayout => entry.kind === 'paragraph')!;
  return { node, lines: node.lines };
}

/** The lines of a one-cell table whose cell has `textDirection`. */
function layoutCellParagraph(
  runs: readonly Run[],
  textDirection: string,
  sectionExtra: Partial<SectionProps> = {},
): readonly LineLayout[] {
  const cell = {
    content: [paragraphOf(runs)], colSpan: 1, vMerge: null,
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    background: null, vAlign: 'top', widthPt: 200, textDirection,
  };
  const table = {
    type: 'table', colWidths: [200],
    rows: [{ cells: [cell], rowHeight: 200, rowHeightRule: 'exact', isHeader: false }],
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
    jc: 'left', layout: 'fixed',
  } as unknown as BodyElement;
  const page = layoutBody([table], undefined, sectionExtra);
  const found: LineLayout[] = [];
  const visit = (node: unknown): void => {
    if (!node || typeof node !== 'object') return;
    const record = node as Record<string, unknown>;
    if (record.kind === 'paragraph' && Array.isArray(record.lines)) {
      found.push(...(record.lines as LineLayout[]));
      return;
    }
    for (const value of Object.values(record)) {
      if (Array.isArray(value)) value.forEach(visit);
      else if (value && typeof value === 'object') visit(value);
    }
  };
  page.layers.body.forEach(visit);
  return found;
}

// Vertical single lines are whole twips (wordVerticalSingleLinePx).
const twips = (pt: number): number => Math.round(pt * 20) / 20;

function textPlacements(line: LineLayout): TextPlacement[] {
  return line.placements.filter((entry): entry is TextPlacement => entry.kind === 'text');
}

describe('vertical line box (WORD_VERTICAL_LINE_CENTRED_BOX)', () => {
  it('holds the vertical single line in whole twips at any device scale', () => {
    // 36pt Far East class: 36 + 1.56 = 37.56pt (751.2 twips) holds 37.55pt.
    expect(wordVerticalSingleLinePx(36, 36, true)).toBeCloseTo(37.55, 9);
    // 10pt Latin: 10.3333pt rounds to 10.35pt; at 2x the device value doubles.
    expect(wordVerticalSingleLinePx(10, 10, false)).toBeCloseTo(10.35, 9);
    expect(wordVerticalSingleLinePx(20, 20, false, 2)).toBeCloseTo(20.7, 9);
    // 10.5pt Far East: 10.5 + 0.455 = 10.955pt rounds to 10.95pt.
    expect(wordVerticalSingleLinePx(10.5, 10.5, true)).toBeCloseTo(10.95, 9);
  });

  it.each([
    // The stub face has no admitted metrics: East Asian text takes the Far
    // East single-line class, Latin text its own box.
    { text: 'あいう', extra: 10 / 30 * 1.3 },
    { text: 'Abc', extra: 10 / 30 },
  ])('adds the measured fraction of the em to the single line of $text', ({ text, extra }) => {
    const horizontal = layoutParagraph([{ text }]).lines[0]!;
    for (const direction of ['tbRl', 'tbLrV']) {
      const vertical = layoutParagraph([{ text }], direction).lines[0]!;
      expect(vertical.advancePt).toBeCloseTo(twips(horizontal.advancePt + extra), 6);
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
    expect(vertical.advancePt).toBeCloseTo(twips(horizontal.advancePt + 10 / 30), 6);
  });
});

describe('vertical ruby guide (WORD_VERTICAL_RUBY_UPRIGHT_GUIDE)', () => {
  it.each(['tbRl', 'tbLrV'])('plans a %s guide as upright and turned cells on its raised axis', (direction) => {
    const ruby = { text: 'かんAB', fontSizePt: 5, hpsRaisePt: 9 };
    const line = layoutParagraph([{ text: '漢字', ruby }], direction).lines[0]!;
    const placement = textPlacements(line).find((entry) => entry.ruby)!;
    const ops = placement.ruby!.paintOps;
    // The turned Latin cell is painted whole, as the base run paints one.
    expect(ops.map((op) => [op.text, op.glyphOrientation])).toEqual([
      ['か', 'upright'], ['ん', 'upright'], ['AB', 'sideways'],
    ]);
    for (const op of ops) expect(op.origin.yPt).toBeCloseTo(line.baselinePt - 9, 6);
    // Upright cells are centred on their one-em advance; the guide is
    // centred along the two-em base.
    expect(ops[0]!.origin.xPt - placement.origin.xPt).toBeCloseTo((20 - 20) / 2 + 2.5, 6);
    // The guide reserves half of its own vertical single line (the stub box
    // at 5pt plus its Far East addition) beyond its axis.
    expect(line.baselinePt - line.bounds.yPt).toBeCloseTo(9 + twips(5 + 5 / 30 * 1.3) / 2, 6);
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

describe('review regressions', () => {
  it.each([undefined, 'tbRl'])('widens a split base to its longer guide (%s)', (direction) => {
    const ruby = { text: '\u3042\u3044\u3046\u3048\u304a\u304b\u304d', fontSizePt: 5, hpsRaisePt: 9, align: 'left' };
    const line = layoutParagraph([{ text: '\u6f22A', ruby }, { text: '\u5f8c' }], direction).lines[0]!;
    const placements = textPlacements(line);
    const first = placements.find((entry) => entry.ruby)!;
    const next = placements.find((entry) => entry.text === '\u5f8c')!;
    expect(next.origin.xPt - first.origin.xPt).toBeCloseTo(35, 6);
  });

  it('keeps a guide unwidened without an authored alignment', () => {
    const ruby = { text: '\u3072\u304c\u3057\u304c\u308f', fontSizePt: 5, hpsRaisePt: 9 };
    const line = layoutParagraph([{ text: '\u6771', ruby }, { text: '\u5f8c' }]).lines[0]!;
    const [annotated, next] = textPlacements(line);
    expect(next!.origin.xPt - annotated!.origin.xPt).toBeCloseTo(10, 6);
  });

  it.each(['tbRl', 'tbLrV'])('keeps a uniformly lowered %s line one single line wide', (direction) => {
    const single = layoutParagraph([{ text: '\u3042' }], direction).lines[0]!.advancePt;
    const lowered = layoutParagraph([{ text: '\u3042', position: -12 }], direction).lines[0]!;
    expect(lowered.advancePt).toBeCloseTo(single, 6);
  });

  it('distributes a mixed vertical guide letter by letter', () => {
    const ruby = { text: '\u304bAB', fontSizePt: 5, hpsRaisePt: 9, align: 'distributeLetter' };
    const line = layoutParagraph([{ text: '\u6f22\u5b57\u65e5\u672c', ruby }], 'tbRl').lines[0]!;
    const placement = textPlacements(line).find((entry) => entry.ruby)!;
    const ops = placement.ruby!.paintOps;
    expect(ops.map((op) => op.text)).toEqual(['\u304b', 'A', 'B']);
    // Letter starts: the upright kana's origin is its centre.
    const starts = [ops[0]!.origin.xPt - 2.5, ops[1]!.origin.xPt, ops[2]!.origin.xPt];
    expect(starts[1]! - starts[0]!).toBeCloseTo(starts[2]! - starts[1]!, 6);
    expect(starts[0]! - placement.origin.xPt).toBeCloseTo(0, 6);
    expect(starts[2]! + 5 - placement.origin.xPt).toBeCloseTo(40, 6);
  });

  it('keeps the guide hpsRaise from a raised base', () => {
    const ruby = { text: '\u304b', fontSizePt: 5, hpsRaisePt: 9 };
    const line = layoutParagraph([{ text: '\u524d' }, { text: '\u6f22', ruby, position: 6 }], 'tbRl').lines[0]!;
    const op = textPlacements(line).find((entry) => entry.ruby)!.ruby!.paintOps[0]!;
    expect(op.origin.yPt).toBeCloseTo(line.baselinePt - 6 - 9, 6);
  });

  it('plans the guide of a tate-chu-yoko run upright', () => {
    const ruby = { text: '\u304b', fontSizePt: 5, hpsRaisePt: 9 };
    const line = layoutParagraph([{ text: '\u524d' }, { text: '12', ruby, eastAsianVert: true }], 'tbRl').lines[0]!;
    const op = textPlacements(line).find((entry) => entry.ruby)!.ruby!.paintOps[0]!;
    expect(op.glyphOrientation).toBe('upright');
  });

  it('counts the grid cells of an empty vertical mark from its vertical single line', () => {
    // Noto Sans 10pt: a 13.62pt horizontal single line, 13.95pt vertically.
    const text = createTextLayoutService({
      fonts: createFontResolver([{ requestedFamily: 'Noto Sans', resolvedFamily: 'Noto Sans',
        source: 'local' as const, resourceIdentity: 'office-local:local("Noto Sans")' }]),
      measurer: { fingerprint: 'vertical-mark-fixture',
        measure: (request) => ({ advancePt: [...request.text].length * 5, ascentPt: 8, descentPt: 2 }) },
    });
    const paragraph = {
      type: 'paragraph', alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
      spaceBefore: 0, spaceAfter: 0, lineSpacing: { rule: 'auto', value: 1, explicit: true },
      numbering: null, tabStops: [], runs: [], widowControl: false,
      defaultFontSize: 10, defaultFontFamily: 'Noto Sans',
    } as unknown as DocParagraph;
    const advance = (vertical: boolean) => paragraphMarkLineMetrics(paragraph, 1,
      { type: 'lines', linePitchPt: 13.8 }, false, false, makeCtx(), {}, paragraph.lineSpacing, {},
      text, undefined, false, 0, true, vertical).advancePx;
    expect(advance(false)).toBeCloseTo(13.8, 6);
    expect(advance(true)).toBeCloseTo(27.6, 6);
  });
});

describe('ruby cell geometry (WORD_RUBY_ALIGN_PLACEMENT, review round 2)', () => {
  const guideOf = (line: LineLayout) => textPlacements(line).find((entry) => entry.ruby)!;
  const lineOf = (lines: readonly LineLayout[], text: string) =>
    lines.findIndex((line) => textPlacements(line).some((entry) => entry.text.includes(text)));

  it.each([undefined, 'tbRl'])('reserves the guide of a raised base inside its line (%s)', (direction) => {
    const ruby = { text: 'か', fontSizePt: 5, hpsRaisePt: 9 };
    const line = layoutParagraph([{ text: '前' }, { text: '漢', ruby, position: 6 }], direction).lines[0]!;
    const op = guideOf(line).ruby!.paintOps[0]!;
    // The guide stands hpsRaise from its raised base.
    expect(op.origin.yPt).toBeCloseTo(line.baselinePt - 6 - 9, 6);
    // Its head extent: the 5pt em box above the baseline horizontally, half
    // its vertical single line beyond the axis vertically.
    const headPt = direction ? twips(5 + 5 / 30 * 1.3) / 2 : 5;
    expect(op.origin.yPt - headPt).toBeGreaterThanOrEqual(line.bounds.yPt - 1e-6);
    expect(op.origin.yPt - headPt).toBeCloseTo(line.bounds.yPt, 6);
  });

  it.each([undefined, 'tbRl'])('wraps a widened split base as one cell (%s)', (direction) => {
    const ruby = { text: 'あいうえお', fontSizePt: 5, hpsRaisePt: 9, align: 'center' };
    const lineLength = direction ? { pageHeight: 60 + 80 + 32 } : { pageWidth: 40 + 90 + 32 };
    const { lines } = layoutParagraph(
      [{ text: '前' }, { text: '漢A', ruby }, { text: '後' }], direction, null, lineLength,
    );
    expect(lineOf(lines, '前')).toBe(0);
    // The 25pt cell does not fit after the first ideograph; it moves whole.
    expect(lineOf(lines, '漢')).toBe(1);
    expect(lineOf(lines, 'A')).toBe(1);
  });

  it('keeps the widened cell in whole snapToChars cells', () => {
    const grid = { docGridType: 'snapToChars', docGridLinePitch: 18, docGridCharSpace: 0 } as Partial<SectionProps>;
    const ruby = { text: 'あいうえお', fontSizePt: 5, hpsRaisePt: 9, align: 'left' };
    const line = layoutParagraph([{ text: '漢', ruby }, { text: '後' }], undefined, null, grid).lines[0]!;
    const [annotated, next] = textPlacements(line);
    const pitch = annotated!.advancePt;
    const advance = next!.origin.xPt - annotated!.origin.xPt;
    // The next character starts after the 25pt guide, on a cell boundary.
    expect(advance).toBeGreaterThanOrEqual(25 - 1e-6);
    const guideEnd = Math.max(...annotated!.ruby!.paintOps.map((op) => op.origin.xPt + (op.advancePt ?? 0)));
    expect(next!.origin.xPt).toBeGreaterThanOrEqual(guideEnd - 1e-6);
    expect(pitch).toBeCloseTo(advance, 6);
  });

  it('admits a fitText ruby base at its widened cell', () => {
    const ruby = { text: 'あいうえお', fontSizePt: 5, hpsRaisePt: 9, align: 'left' };
    const { lines } = layoutParagraph([
      { text: '前' },
      { text: '漢', ruby, extra: { fitTextVal: 300, fitTextId: 7 } },
      { text: '後' },
    ], undefined, null, { pageWidth: 40 + 90 + 32 });
    // ECMA-376 §17.3.2.14 fits the base run to 15pt; the authored alignment
    // widens the ruby cell to its 25pt guide, which no longer fits after the
    // first ideograph.
    expect(lineOf(lines, '前')).toBe(0);
    expect(lineOf(lines, '漢')).toBe(1);
    // One width for admission and placement: the next character follows
    // the cell.
    expect(textPlacements(lines[1]!)[0]!.advancePt).toBeCloseTo(25, 6);
    expect(lineOf(lines, '\u5f8c')).toBe(2);
  });

  it('paints a joining sideways guide whole', () => {
    const ruby = { text: 'سلام', fontSizePt: 5, hpsRaisePt: 9 };
    const line = layoutParagraph([{ text: '漢字', ruby }], 'tbRl').lines[0]!;
    const ops = guideOf(line).ruby!.paintOps;
    expect(ops.map((op) => op.text)).toEqual(['سلام']);
    expect(ops[0]!.glyphOrientation).toBe('sideways');
  });

  it('distributes turned letters from their contextual pens', () => {
    const ruby = { text: 'AVか', fontSizePt: 5, hpsRaisePt: 9, align: 'distributeLetter' };
    const line = layoutParagraph([{ text: '漢字日本', ruby }], 'tbRl').lines[0]!;
    const placement = guideOf(line);
    const ops = placement.ruby!.paintOps;
    expect(ops.map((op) => op.text)).toEqual(['A', 'V', 'か']);
    // "AV" measures 9pt together: the guide is 14pt over the 40pt base.
    const gap = (40 - 14) / 2;
    expect(ops[0]!.origin.xPt - placement.origin.xPt).toBeCloseTo(0, 6);
    expect(ops[1]!.origin.xPt - placement.origin.xPt).toBeCloseTo(5 + gap, 6);
    // The kana keeps its planned pen after the kerned pair: it ends at the
    // base end.
    expect(ops[2]!.origin.xPt + 2.5 - placement.origin.xPt).toBeCloseTo(40, 6);
  });

  it('centres a tate-chu-yoko base once inside its widened cell', () => {
    const ruby = { text: 'あいうえお', fontSizePt: 5, hpsRaisePt: 9, align: 'center' };
    const line = layoutParagraph([{ text: '12', ruby, eastAsianVert: true }], 'tbRl').lines[0]!;
    const placement = textPlacements(line).find((entry) => entry.ruby)!;
    expect(placement.advancePt).toBeCloseTo(25, 6);
    // The one-em cell's centre lies at the widened cell's centre.
    expect(placement.paintOps[0]!.offset.xPt).toBeCloseTo(12.5, 6);
  });

  it('raises a rightVertical guide in an upright cell as its line reserves', () => {
    const ruby = { text: 'か', fontSizePt: 5, hpsRaisePt: 9, align: 'rightVertical' };
    const line = layoutCellParagraph([{ text: '漢', ruby }], 'tbRlV')[0]!;
    const op = guideOf(line).ruby!.paintOps[0]!;
    expect(op.origin.yPt).toBeCloseTo(line.baselinePt - 9 + 2.5, 6);
    expect(op.origin.yPt - twips(5 + 5 / 30 * 1.3) / 2).toBeCloseTo(line.bounds.yPt, 6);
  });

  it('centres distributed upright guide letters in an upright cell on their own advances', () => {
    const ruby = { text: 'かき', fontSizePt: 5, hpsRaisePt: 9, align: 'distributeLetter' };
    const line = layoutCellParagraph([{ text: '漢字日本', ruby }], 'tbRlV')[0]!;
    const placement = guideOf(line);
    const centres = placement.ruby!.paintOps.map((op) => op.origin.xPt - placement.origin.xPt);
    expect(centres[0]).toBeCloseTo(2.5, 6);
    expect(centres[1]).toBeCloseTo(37.5, 6);
  });

  it('keeps w:position on a turned run in an upright cell', () => {
    const offsetOf = (position: number | undefined) => {
      const line = layoutCellParagraph([{ text: '漢' }, { text: 'A', ...(position ? { position } : {}) }], 'tbRlV')[0]!;
      const latin = textPlacements(line).find((entry) => entry.text === 'A')!;
      const op = latin.paintOps.find((entry) => entry.text === 'A')!;
      return latin.origin.yPt + op.offset.yPt - line.baselinePt;
    };
    expect(offsetOf(6) - offsetOf(undefined)).toBeCloseTo(-6, 6);
  });
});

describe('ruby cell geometry (WORD_RUBY_ALIGN_PLACEMENT, review round 3)', () => {
  const guides = (lines: readonly LineLayout[]) => lines.flatMap((line) =>
    textPlacements(line).filter((entry) => entry.ruby));
  const lineOf = (lines: readonly LineLayout[], text: string) =>
    lines.findIndex((line) => textPlacements(line).some((entry) => entry.text.includes(text)));
  const lengthOf = (direction: string | undefined, pt: number) =>
    direction ? { pageHeight: 60 + 80 + pt } : { pageWidth: 40 + 90 + pt };

  it.each([undefined, 'tbRl'])('wraps an unsplit ideographic ruby base as its whole cell (%s)', (direction) => {
    const ruby = { text: 'かきくけこさ', fontSizePt: 5, hpsRaisePt: 9, align: 'center' };
    const { lines } = layoutParagraph(
      [{ text: '前' }, { text: '漢字', ruby }, { text: '後' }], direction, null, lengthOf(direction, 35),
    );
    expect(lineOf(lines, '前')).toBe(0);
    // The 30pt cell does not fit after the first ideograph; it moves whole.
    expect(lineOf(lines, '漢字')).toBe(1);
    const annotated = guides(lines);
    expect(annotated).toHaveLength(1);
    expect(annotated[0]!.advancePt).toBeCloseTo(30, 6);
  });

  it.each([undefined, 'tbRl'])('keeps an overlong ruby cell whole and paints its guide once (%s)', (direction) => {
    const ruby = { text: 'かきくけこさ', fontSizePt: 5, hpsRaisePt: 9, align: 'center' };
    const { lines } = layoutParagraph(
      [{ text: '前' }, { text: '漢字', ruby }, { text: '後' }], direction, null, lengthOf(direction, 16),
    );
    const annotated = guides(lines);
    expect(annotated).toHaveLength(1);
    expect(annotated[0]!.text).toBe('漢字');
    expect(annotated[0]!.ruby!.paintOps.map((op) => op.text).join('')).toBe('かきくけこさ');
  });

  it('keeps the ruby base alignment of a snapped base in an upright cell', () => {
    const grid = { docGridType: 'snapToChars', docGridLinePitch: 18, docGridCharSpace: 0 } as Partial<SectionProps>;
    const baseAt = (align: string) => {
      const ruby = { text: 'かきくけこ', fontSizePt: 5, hpsRaisePt: 9, align };
      const line = layoutCellParagraph([{ text: '漢', ruby }], 'tbRlV', grid)[0]!;
      const placement = textPlacements(line).find((entry) => entry.ruby)!;
      const op = placement.paintOps.find((entry) => entry.text === '漢')!;
      return placement.origin.xPt + op.offset.xPt;
    };
    // The 25pt guide widens the 10pt base: rubyAlign moves the base along
    // the column by the 15pt overhang.
    expect(baseAt('right') - baseAt('left')).toBeCloseTo(15, 6);
  });

  // A centred guide over one ideograph widens its cell; a distributed one
  // over a longer base spaces every turned letter.
  it.each([['center', false], ['distributeLetter', true]] as const)('plans a long turned guide in linear work (%s)', (align, longBase) => {
    const work = (run: Run, length: number) => {
      const before = measuredCharacters;
      layoutParagraph([run], 'tbRl', null, { pageHeight: 60 + 80 + 10 * length });
      return measuredCharacters - before;
    };
    // The guide's share: the annotated run less its bare base.
    const guideWork = (length: number) => {
      const base = '漢'.repeat(longBase ? length : 1);
      const ruby = { text: 'Q'.repeat(length), fontSizePt: 5, hpsRaisePt: 9, align };
      return work({ text: base, ruby }, length) - work({ text: base }, length);
    };
    const short = guideWork(128);
    const long = guideWork(256);
    // Doubling the guide at most doubles its work, with headroom for fixed
    // per-guide cost; quadratic planning quadruples it.
    expect(long / short).toBeLessThan(2.5);
  });

  it('distributes turned Latin letters beside a joining run kept whole', () => {
    const ruby = { text: 'ABسلامか', fontSizePt: 5, hpsRaisePt: 9, align: 'distributeLetter' };
    const line = layoutParagraph([{ text: '漢字日本', ruby }], 'tbRl').lines[0]!;
    const placement = textPlacements(line).find((entry) => entry.ruby)!;
    const ops = placement.ruby!.paintOps;
    expect(ops.map((op) => op.text)).toEqual(['A', 'B', 'سلام', 'か']);
    // A 35pt guide over the 40pt base: four letters share 5pt in three gaps.
    const gap = 5 / 3;
    expect(ops[1]!.origin.xPt - placement.origin.xPt).toBeCloseTo(5 + gap, 6);
    expect(ops[2]!.origin.xPt - placement.origin.xPt).toBeCloseTo(10 + 2 * gap, 6);
  });

  it('moves a ruby cell whole ahead of a line-start-forbidden follower', () => {
    const ruby = { text: 'かき', fontSizePt: 5, hpsRaisePt: 9 };
    const { lines } = layoutParagraph(
      [{ text: '前' }, { text: '漢字', ruby }, { text: 'ゃ' }], undefined, null, lengthOf(undefined, 30),
    );
    expect(guides(lines).map((entry) => entry.text)).toEqual(['漢字']);
    expect(lineOf(lines, '漢字')).toBe(lineOf(lines, 'ゃ'));
  });
});

describe('ruby cell geometry (WORD_RUBY_ALIGN_PLACEMENT, review round 4)', () => {
  const lineOf = (lines: readonly LineLayout[], text: string) =>
    lines.findIndex((line) => textPlacements(line).some((entry) => entry.text.includes(text)));

  it('keeps an unwidened ruby base in its snapToChars script block', () => {
    const grid = { docGridType: 'snapToChars', docGridLinePitch: 18, docGridCharSpace: 0, pageWidth: 40 + 90 + 35 } as Partial<SectionProps>;
    const ruby = { text: 'か', fontSizePt: 5, hpsRaisePt: 9, align: 'center' };
    const narrow = { charScale: 0.4 };
    const { lines } = layoutParagraph([
      { text: 'I' }, { text: 'i', ruby, extra: narrow }, { text: 'i', extra: narrow }, { text: '後' },
    ], undefined, null, grid);
    // The 18pt Latin block takes two 10pt cells, leaving room for 後.
    expect(lineOf(lines, '後')).toBe(0);
  });

  it('centres an unwidened guide on its base as placed in the snapToChars block', () => {
    const grid = { docGridType: 'snapToChars', docGridLinePitch: 18, docGridCharSpace: 0, pageWidth: 40 + 90 + 35 } as Partial<SectionProps>;
    const ruby = { text: 'か', fontSizePt: 5, hpsRaisePt: 9, align: 'center' };
    const narrow = { charScale: 0.4 };
    const line = layoutParagraph([
      { text: 'I' }, { text: 'i', ruby, extra: narrow }, { text: 'i', extra: narrow }, { text: '後' },
    ], undefined, null, grid).lines[0]!;
    const annotated = textPlacements(line).find((entry) => entry.ruby)!;
    const base = annotated.paintOps.find((op) => op.text === 'i')!;
    // The 40%-wide i is 4pt; the 5pt guide centres on it.
    const baseCentre = annotated.origin.xPt + base.offset.xPt + 2;
    const guide = annotated.ruby!.paintOps[0]!;
    expect(guide.origin.xPt + 2.5).toBeCloseTo(baseCentre, 6);
  });

  it.each(['tbRl', 'tbLrV'])('retracts a character ahead of a ruby cell led by a forbidden kana (%s)', (direction) => {
    const ruby = { text: 'かきくけ', fontSizePt: 5, hpsRaisePt: 9, align: 'center' };
    const { lines } = layoutParagraph(
      [{ text: '前前' }, { text: 'ゃ漢', ruby }], direction, null, { pageHeight: 60 + 80 + 30 },
    );
    const texts = lines.map((line) => textPlacements(line).map((entry) => entry.text).join(''));
    expect(texts).toEqual(['前', '前ゃ漢']);
  });

  it('plans each font slot of a vertical guide in its own face', () => {
    const ruby = { text: 'ABか', fontSizePt: 5, hpsRaisePt: 9 };
    const line = layoutParagraph([{
      text: '漢字', ruby, fontFamily: 'AsciiFace', extra: { fontFamilyEastAsia: 'EastFace' },
    }], 'tbRl').lines[0]!;
    const ops = textPlacements(line).find((entry) => entry.ruby)!.ruby!.paintOps;
    const faces = ops.map((op) => [op.text, JSON.stringify(op.fontRoute)]);
    expect(faces.map(([text]) => text)).toEqual(['AB', 'か']);
    expect(faces[0]![1]).toContain('AsciiFace');
    expect(faces[1]![1]).toContain('EastFace');
  });

  it('paints a turned guide cell with the kerning it was planned with', () => {
    const ruby = { text: 'AVか', fontSizePt: 5, hpsRaisePt: 9, align: 'center' };
    const line = layoutParagraph([{ text: '漢字', ruby }], 'tbRl').lines[0]!;
    const ops = textPlacements(line).find((entry) => entry.ruby)!.ruby!.paintOps;
    const turned = ops.find((op) => op.glyphOrientation === 'sideways')!;
    expect(turned.text).toBe('AV');
    expect(turned.kerning).toBe('none');
  });
});
