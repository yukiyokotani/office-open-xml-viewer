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
      // A face named TallFace has a two-em box.
      const box = font.includes('TallFace') ? 2 : 1;
      return {
        // One em per character; the pair "AV" kerns 0.2 em tighter.
        width: [...text].length * size - (text.match(/AV/g)?.length ?? 0) * size * 0.2,
        fontBoundingBoxAscent: size * 0.8 * box, fontBoundingBoxDescent: size * 0.2 * box,
        actualBoundingBoxAscent: size * 0.8 * box, actualBoundingBoxDescent: size * 0.2 * box,
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

type GuideRun = Readonly<{ text: string; fonts: Readonly<Record<string, string>> }>;
type Ruby = Readonly<{ text: string; fontSizePt: number; hpsRaisePt: number; align?: string; guideRuns?: readonly GuideRun[] }>;
type Run = Readonly<{ text: string; ruby?: Ruby; position?: number; fontFamily?: string; eastAsianVert?: boolean; fontSize?: number; extra?: Record<string, unknown> }>;

function typographyInput(ruby: Ruby) {
  const value = <T>(v: T | undefined) => v === undefined
    ? { status: 'missing', raw: null, value: null }
    : { status: 'valid', raw: String(v), value: v };
  return {
    ruby: {
      align: value(ruby.align), baseFontSizePt: value(undefined),
      raisePt: value(ruby.hpsRaisePt), language: value(undefined),
      guideRuns: (ruby.guideRuns ?? []).map((run) => ({
        text: run.text, fontFamily: run.fonts.ascii ?? null, fontSizePt: ruby.fontSizePt,
        bold: false, italic: false, color: null, language: null,
        fonts: { ascii: null, highAnsi: null, eastAsia: null, complexScript: null, ...run.fonts },
      })),
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
  return layoutCell(runs, textDirection, sectionExtra).lines;
}

/** The lines of a one-cell table and the page-space origin of each line's
 * first placement, with an optional uniform rule and cell margin. */
function layoutCell(
  runs: readonly Run[],
  textDirection: string,
  sectionExtra: Partial<SectionProps> = {},
  rulePt = 0,
  marginPt = 0,
): Readonly<{ lines: readonly LineLayout[]; cellTransform?: { a: number; b: number; c: number; d: number; e: number; f: number } }> {
  const cell = {
    content: [paragraphOf(runs)], colSpan: 1, vMerge: null,
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    background: null, vAlign: 'top', widthPt: 200, textDirection,
  };
  const rule = rulePt > 0 ? { width: rulePt, color: '000000', style: 'single' } : null;
  const table = {
    type: 'table', colWidths: [200],
    rows: [{ cells: [cell], rowHeight: 200, rowHeightRule: 'exact', isHeader: false }],
    borders: { top: rule, bottom: rule, left: rule, right: rule, insideH: null, insideV: null },
    cellMarginTop: marginPt, cellMarginBottom: marginPt, cellMarginLeft: marginPt, cellMarginRight: marginPt,
    jc: 'left', layout: 'fixed',
  } as unknown as BodyElement;
  const page = layoutBody([table], undefined, sectionExtra);
  const found: LineLayout[] = [];
  let cellTransform: { a: number; b: number; c: number; d: number; e: number; f: number } | undefined;
  const visit = (node: unknown): void => {
    if (!node || typeof node !== 'object') return;
    const record = node as Record<string, unknown>;
    if (record.kind === 'table-cell' && record.verticalText) {
      cellTransform = (record.verticalText as { transform: typeof cellTransform }).transform;
    }
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
  return { lines: found, ...(cellTransform ? { cellTransform } : {}) };
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
    const ruby = { text: 'か', fontSizePt: 3.5, hpsRaisePt: 9, align: 'center' };
    const narrow = { charScale: 0.4 };
    const { lines } = layoutParagraph([
      { text: 'I' }, { text: 'i', ruby, extra: narrow }, { text: 'i', extra: narrow }, { text: '後' },
    ], undefined, null, grid);
    // The 18pt Latin block takes two 10pt cells, leaving room for 後.
    expect(lineOf(lines, '後')).toBe(0);
  });

  it('centres an unwidened guide on its base as placed in the snapToChars block', () => {
    const grid = { docGridType: 'snapToChars', docGridLinePitch: 18, docGridCharSpace: 0, pageWidth: 40 + 90 + 35 } as Partial<SectionProps>;
    const ruby = { text: 'か', fontSizePt: 3.5, hpsRaisePt: 9, align: 'center' };
    const narrow = { charScale: 0.4 };
    const line = layoutParagraph([
      { text: 'I' }, { text: 'i', ruby, extra: narrow }, { text: 'i', extra: narrow }, { text: '後' },
    ], undefined, null, grid).lines[0]!;
    const annotated = textPlacements(line).find((entry) => entry.ruby)!;
    const base = annotated.paintOps.find((op) => op.text === 'i')!;
    // The 40%-wide i is 4pt; the 3.5pt guide centres on it.
    const baseCentre = annotated.origin.xPt + base.offset.xPt + 2;
    const guide = annotated.ruby!.paintOps[0]!;
    expect(guide.origin.xPt + 1.75).toBeCloseTo(baseCentre, 6);
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

describe('rotated cell rule inset (WORD_ROTATED_CELL_RULE_INSET)', () => {
  // The physical x of the local frame's line-start corner of the cell.
  const frameX = (direction: string, rulePt: number, marginPt: number) =>
    layoutCell([{ text: '漢字' }], direction, {}, rulePt, marginPt).cellTransform!.e;

  it('starts a btLr cell\'s first column the rule width in from a margin-free edge', () => {
    expect(frameX('btLr', 1.5, 0) - frameX('btLr', 0, 0)).toBeCloseTo(1.5, 6);
    expect(frameX('tbRl', 1.5, 0) - frameX('tbRl', 0, 0)).toBeCloseTo(0, 6);
  });

  it('measures a cell margin from the rule centre', () => {
    expect(frameX('btLr', 0.5, 5.4) - frameX('btLr', 0, 0)).toBeCloseTo(5.65, 6);
    expect(frameX('tbRl', 0.5, 5.4) - frameX('tbRl', 0, 0)).toBeCloseTo(-5.15, 6);
  });
});

describe('ruby cell geometry (WORD_RUBY_ALIGN_PLACEMENT, review round 5)', () => {
  const lineTexts = (lines: readonly LineLayout[]) =>
    lines.map((line) => textPlacements(line).map((entry) => entry.text).join(''));

  it.each(['tbRl', 'tbLrV'])('admits a split ruby base by its whole cell (%s)', (direction) => {
    const ruby = { text: 'かきくけ', fontSizePt: 5, hpsRaisePt: 9, align: 'center' };
    const { lines } = layoutParagraph(
      [{ text: '前前前' }, { text: 'ゃA', ruby }], direction, null, { pageHeight: 60 + 80 + 40 },
    );
    expect(lineTexts(lines)).toEqual(['前前', '前ゃA']);
  });

  it('widens a guide longer than its Latin base glyphs on a character grid', () => {
    const grid = { docGridType: 'snapToChars', docGridLinePitch: 18, docGridCharSpace: 0 } as Partial<SectionProps>;
    const ruby = { text: 'かか', fontSizePt: 4, hpsRaisePt: 9, align: 'center' };
    const narrow = { charScale: 0.4 };
    const line = layoutParagraph([
      { text: 'i', ruby, extra: narrow }, { text: 'i', ruby, extra: narrow },
    ], undefined, null, grid).lines[0]!;
    const [first, second] = textPlacements(line).filter((entry) => entry.ruby);
    const firstEnd = Math.max(...first!.ruby!.paintOps.map((op) => op.origin.xPt + (op.advancePt ?? 0)));
    const secondStart = Math.min(...second!.ruby!.paintOps.map((op) => op.origin.xPt));
    // Each 8pt guide widens its 4pt base to one whole 10pt cell.
    expect(secondStart).toBeGreaterThanOrEqual(firstEnd - 1e-6);
    expect(second!.origin.xPt - first!.origin.xPt).toBeCloseTo(10, 6);
  });

  it('sizes the mark column after a trailing manual break as an empty paragraph\'s mark', () => {
    const breaks = paragraphOf([{ text: '漢' }]);
    const lineBreak = { type: 'break', breakType: 'line' };
    const withBreaks = { ...breaks, runs: [...breaks.runs, lineBreak, lineBreak] } as unknown as DocParagraph & BodyElement;
    const page = layoutBody([withBreaks, paragraphOf([])], 'tbRl');
    const [broken, empty] = page.layers.body.filter((entry): entry is ParagraphLayout => entry.kind === 'paragraph');
    expect(broken!.lines).toHaveLength(3);
    expect(broken!.lines.at(-1)!.advancePt).toBeCloseTo(empty!.advancePt, 6);
  });

  it.each([undefined, 'tbRl'])('shapes each guide run in its own faces (%s)', (direction) => {
    const fonts = (face: string) => ({ ascii: face, highAnsi: face, eastAsia: face });
    const ruby = {
      text: 'ABか', fontSizePt: 5, hpsRaisePt: 9,
      guideRuns: [{ text: 'AB', fonts: fonts('GuideLatin') }, { text: 'か', fonts: fonts('GuideKana') }],
    };
    const line = layoutParagraph([{ text: '漢字', ruby, fontFamily: 'BaseFace' }], direction).lines[0]!;
    const ops = textPlacements(line).find((entry) => entry.ruby)!.ruby!.paintOps;
    const faceOf = (text: string) => JSON.stringify(ops.find((op) => op.text.includes(text))!.fontRoute);
    expect(faceOf('A')).toContain('GuideLatin');
    expect(faceOf('か')).toContain('GuideKana');
  });

  it.each([undefined, 'tbRl'])('spaces a mixed ruby base inside its cell, not at its edges (%s)', (direction) => {
    const ruby = { text: 'あいうえおかきくけ', fontSizePt: 5, hpsRaisePt: 9, align: 'center' };
    const line = layoutParagraph([{ text: '前' }, { text: '東A', ruby }, { text: '後' }], direction).lines[0]!;
    const placements = textPlacements(line);
    const at = (text: string) => placements.find((entry) => entry.text === text)!;
    const glyph = (text: string) => at(text).origin.xPt + at(text).paintOps[0]!.offset.xPt;
    const cellStart = at('東').origin.xPt;
    // The quarter-em gap (WORD_EAST_ASIAN_AUTO_SPACE) lies between 東 and A,
    // inside the 45pt cell; 後 follows the cell directly.
    expect(at('後').origin.xPt + at('後').paintOps[0]!.offset.xPt - cellStart - (direction ? 5 : 0)).toBeCloseTo(45, 6);
    expect(glyph('A') - glyph('東') - (direction ? -5 : 0)).toBeCloseTo(10 + 2.5, 6);
  });
});

describe('ruby cell geometry (WORD_RUBY_ALIGN_PLACEMENT, review round 6)', () => {
  const lineTexts = (lines: readonly LineLayout[]) =>
    lines.map((line) => textPlacements(line).map((entry) => entry.text).join(''));

  it.each([undefined, 'tbRl'])('forms one cell from every authored base run (%s)', (direction) => {
    const ruby = { text: 'とうきょう', fontSizePt: 5, hpsRaisePt: 9, align: 'center' };
    const { lines } = layoutParagraph([
      { text: '前' }, { text: '東', ruby }, { text: '京', extra: { bold: true, rubyBaseContinuation: true } },
    ], direction, null, direction ? { pageHeight: 60 + 80 + 37 } : { pageWidth: 40 + 90 + 37 });
    expect(lineTexts(lines)).toEqual(['前東京']);
    const placements = textPlacements(lines[0]!);
    const east = placements.find((entry) => entry.text === '東')!;
    const capital = placements.find((entry) => entry.text === '京')!;
    // The two base runs share one 25pt cell under the 25pt guide.
    expect(capital.origin.xPt + capital.advancePt - east.origin.xPt).toBeCloseTo(25, 6);
    expect(placements.filter((entry) => entry.ruby)).toHaveLength(1);
  });

  it.each([undefined, 'tbRl'])('shapes identically formatted guide runs as one text (%s)', (direction) => {
    const fonts = { ascii: 'GuideFace', highAnsi: 'GuideFace', eastAsia: 'GuideFace', complexScript: 'GuideFace' };
    const ruby = {
      text: 'سلام', fontSizePt: 5, hpsRaisePt: 9,
      guideRuns: [{ text: 'سل', fonts }, { text: 'ام', fonts }],
    };
    const line = layoutParagraph([{ text: '漢字', ruby }], direction).lines[0]!;
    const ops = textPlacements(line).find((entry) => entry.ruby)!.ruby!.paintOps;
    expect(ops.map((op) => op.text)).toEqual(['سلام']);
  });

  it.each(['tbRl', 'tbLrV'])('retracts a whole split ruby cell ahead of a forbidden head (%s)', (direction) => {
    const ruby = { text: 'か', fontSizePt: 5, hpsRaisePt: 9 };
    const { lines } = layoutParagraph(
      [{ text: '前' }, { text: '東A', ruby }, { text: 'ゃ', ruby }], direction, null, { pageHeight: 60 + 80 + 40 },
    );
    expect(lineTexts(lines)).toEqual(['前', '東Aゃ']);
  });

  it('admits an unwidened split Latin ruby base by its shared grid block', () => {
    const grid = { docGridType: 'snapToChars', docGridLinePitch: 18, docGridCharSpace: 0, pageWidth: 40 + 90 + 20 } as Partial<SectionProps>;
    const ruby = { text: 'か', fontSizePt: 3.5, hpsRaisePt: 9, align: 'center' };
    const narrow = { smallCaps: true, charScale: 0.4 };
    const annotated = layoutParagraph([{ text: 'I' }, { text: 'aB', ruby, extra: narrow }], undefined, null, grid).lines;
    const plain = layoutParagraph([{ text: 'I' }, { text: 'aB', extra: narrow }], undefined, null, grid).lines;
    expect(plain).toHaveLength(1);
    expect(lineTexts(annotated)).toEqual(lineTexts(plain));
  });

  it('keeps the mark extent of the column after a ruby paragraph\'s trailing break', () => {
    const ruby = { text: 'か', fontSizePt: 5, hpsRaisePt: 18 };
    const annotated = paragraphOf([{ text: '漢', ruby }]);
    const lineBreak = { type: 'break', breakType: 'line' };
    const withBreak = { ...annotated, runs: [...annotated.runs, lineBreak] } as unknown as DocParagraph & BodyElement;
    const page = layoutBody([withBreak, paragraphOf([])], 'tbRl');
    const [broken, empty] = page.layers.body.filter((entry): entry is ParagraphLayout => entry.kind === 'paragraph');
    expect(broken!.lines).toHaveLength(2);
    expect(broken!.lines.at(-1)!.advancePt).toBeCloseTo(empty!.advancePt, 6);
  });

  it('reserves a vertical guide\'s head extent from the guide\'s own face', () => {
    const headOf = (face: string) => {
      const fonts = { ascii: face, highAnsi: face, eastAsia: face };
      const ruby = { text: 'か', fontSizePt: 5, hpsRaisePt: 9, guideRuns: [{ text: 'か', fonts }] };
      const line = layoutParagraph([{ text: '漢', ruby }], 'tbRl').lines[0]!;
      return line.baselinePt - line.bounds.yPt;
    };
    // A two-em guide box reserves half an em of the 5pt guide more.
    expect(headOf('TallFace') - headOf('NotInMetrics')).toBeCloseTo(2.5, 1);
  });
});

describe('ruby cell geometry (WORD_RUBY_ALIGN_PLACEMENT, review round 7)', () => {
  const lineTexts = (lines: readonly LineLayout[]) =>
    lines.map((line) => textPlacements(line).map((entry) => entry.text).join(''));
  const column = (pt: number) => ({ pageHeight: 60 + 80 + pt });
  const guide = { text: 'か', fontSizePt: 5, hpsRaisePt: 9, align: 'center' };

  it('admits a fitText ruby base piece by its whole cell', () => {
    const fitted = (piece: 0 | 1, length: number) => lineTexts(layoutParagraph([
      { text: '前' },
      { text: '東', ruby: guide, ...(piece === 0 ? { extra: { fitTextVal: 300, fitTextId: 41 } } : {}) },
      { text: '京', extra: { rubyBaseContinuation: true, ...(piece === 1 ? { fitTextVal: 300, fitTextId: 42 } : {}) } },
    ], 'tbRl', null, column(length)).lines);
    // 前 (10pt) and the 25pt cell need 35pt.
    expect(fitted(0, 32)).toEqual(['前', '東京']);
    expect(fitted(1, 22)).toEqual(['前', '東京']);
  });

  it('moves a line-end-forbidden opener with the ruby cell after it', () => {
    const { lines } = layoutParagraph(
      [{ text: '前「' }, { text: '東京', ruby: guide }], 'tbRl', null, column(35),
    );
    expect(lineTexts(lines)).toEqual(['前', '「東京']);
  });

  it('keeps a forbidden ruby leader on a column it cannot leave', () => {
    const { lines } = layoutParagraph(
      [{ text: '東京', ruby: guide }, { text: 'ゃ', ruby: guide }], 'tbRl', null, column(25),
    );
    expect(lineTexts(lines)).toEqual(['東京ゃ']);
  });

  it('restores the column extents a retracted ruby cell had added', () => {
    const raised = { text: 'か', fontSizePt: 5, hpsRaisePt: 9 };
    const retracted = layoutParagraph(
      [{ text: '前' }, { text: '東A', ruby: raised }, { text: 'ゃ' }], 'tbRl', null, column(40),
    ).lines;
    expect(lineTexts(retracted)).toEqual(['前', '東Aゃ']);
    const same = paragraphOf([{ text: '前' }, { text: '東A', ruby: raised }, { text: 'ゃ' }]);
    const broken = { ...same, runs: [same.runs[0], { type: 'break', breakType: 'line' }, ...same.runs.slice(1)] } as unknown as DocParagraph & BodyElement;
    const explicit = layoutBody([broken], 'tbRl', column(40)).layers.body
      .find((entry): entry is ParagraphLayout => entry.kind === 'paragraph')!.lines[0]!;
    const axis = (line: LineLayout) => line.baselinePt - line.bounds.yPt;
    expect(axis(retracted[0]!)).toBeCloseTo(axis(explicit), 6);
  });

  it('shapes guide runs as one text when their characters select one face', () => {
    const run = (eastAsia: string) => ({ ascii: 'Same', highAnsi: 'Same', eastAsia, complexScript: 'Same' });
    const ruby = {
      text: 'سلام', fontSizePt: 5, hpsRaisePt: 9,
      guideRuns: [{ text: 'سل', fonts: run('KanaOne') }, { text: 'ام', fonts: run('KanaTwo') }],
    };
    const line = layoutParagraph([{ text: '漢字', ruby }]).lines[0]!;
    expect(textPlacements(line).find((entry) => entry.ruby)!.ruby!.paintOps.map((op) => op.text)).toEqual(['سلام']);
  });

  it('turns a sideways guide in an upright cell about the guide face', () => {
    const offsetOf = (face: string) => {
      const fonts = { ascii: face, highAnsi: face, eastAsia: face };
      const ruby = { text: 'A', fontSizePt: 5, hpsRaisePt: 9, guideRuns: [{ text: 'A', fonts }] };
      const line = layoutCellParagraph([{ text: '漢', ruby }], 'tbRlV')[0]!;
      return textPlacements(line).find((entry) => entry.ruby)!.ruby!.paintOps[0]!.glyphOffsetPt?.yPt ?? 0;
    };
    expect(offsetOf('NotInMetrics')).toBeCloseTo(1.5, 6);
    expect(offsetOf('TallFace')).toBeCloseTo(3, 6);
  });
});

describe('ruby cells in generated paragraphs (WORD_RUBY_ALIGN_PLACEMENT)', () => {
  // Deterministic generator: ruby cells with split base runs, fitText pieces,
  // kinsoku leaders and openers, in lrTb, tbRl and tbLrV columns of varied
  // length.
  const random = (seed: number) => () => {
    seed = (seed + 0x6d2b79f5) | 0;
    let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
  const KANJI = [...'一二三四五六七八九十百千万東京都庁前後中左右上下山川田林森村町市区円年月日火水木金土'];
  const LEADERS = [...'ゃょー」'];
  const LINE_START_FORBIDDEN = new Set(LEADERS);
  const LINE_END_FORBIDDEN = new Set(['「']);
  const ALIGNS = ['center', 'left', 'right', 'distributeLetter', undefined];
  type Cell = { id: number; base: string; guide: string };

  function generate(seed: number) {
    const next = random(seed);
    const pick = <T,>(items: readonly T[]) => items[Math.floor(next() * items.length)]!;
    const pool = [...KANJI];
    const take = () => pool.splice(Math.floor(next() * pool.length), 1)[0]!;
    const runs: Run[] = [];
    const cells: Cell[] = [];
    let fitId = 100;
    const items = 3 + Math.floor(next() * 6);
    for (let index = 0; index < items; index++) {
      if (next() < 0.5) {
        let text = Array.from({ length: 1 + Math.floor(next() * 3) }, take).join('');
        if (next() < 0.3) text = pick(LEADERS) + text;
        if (next() < 0.25) text += '「';
        runs.push({ text });
        continue;
      }
      const pieces = [take() + (next() < 0.3 ? 'A' : ''), ...(next() < 0.5 ? [take()] : [])];
      const guide = 'あいうえおかきくけこ'.slice(0, 1 + Math.floor(next() * 6));
      const ruby = { text: guide, fontSizePt: 5, hpsRaisePt: next() < 0.5 ? 9 : 4, align: pick(ALIGNS) };
      const fitPiece = next() < 0.3 ? Math.floor(next() * pieces.length) : -1;
      pieces.forEach((text, piece) => runs.push({
        text,
        ...(piece === 0 ? { ruby } : {}),
        extra: {
          ...(piece > 0 ? { rubyBaseContinuation: true } : {}),
          ...(piece === fitPiece ? { fitTextVal: 200 + Math.floor(next() * 400), fitTextId: fitId++ } : {}),
        },
      }));
      cells.push({ id: cells.length, base: pieces.join(''), guide });
    }
    const direction = pick([undefined, 'tbRl', 'tbLrV']);
    const lengthPt = 25 + Math.floor(next() * 60);
    const grid = next() < 0.3
      ? { docGridType: 'snapToChars', docGridLinePitch: 18, docGridCharSpace: 0 } as Partial<SectionProps> : {};
    return { runs, cells, direction, lengthPt, grid };
  }

  const lineText = (line: LineLayout) => textPlacements(line).map((entry) => entry.text).join('');
  const sectionFor = (direction: string | undefined, lengthPt: number, grid: Partial<SectionProps>) =>
    ({ ...grid, ...(direction ? { pageHeight: 60 + 80 + lengthPt } : { pageWidth: 40 + 90 + lengthPt }) });

  /** The same runs with an explicit break wherever the layout broke. */
  function explicitlyBroken(runs: readonly Run[], lines: readonly LineLayout[]): DocParagraph & BodyElement {
    const ends = new Set<number>();
    let offset = 0;
    for (const line of lines.slice(0, -1)) ends.add(offset += [...lineText(line)].length);
    const out: unknown[] = [];
    const base = paragraphOf(runs);
    let position = 0;
    runs.forEach((run, index) => {
      const source = base.runs[index] as unknown as { text: string };
      const chars = [...run.text];
      let start = 0;
      for (let at = 1; at <= chars.length; at++) {
        if (!ends.has(position + at)) continue;
        out.push({ ...source, text: chars.slice(start, at).join('') });
        out.push({ type: 'break', breakType: 'line' });
        start = at;
      }
      if (start < chars.length) out.push({ ...source, text: chars.slice(start).join('') });
      position += chars.length;
    });
    return { ...base, runs: out } as unknown as DocParagraph & BodyElement;
  }

  it.each(Array.from({ length: 200 }, (_, seed) => seed + 1))('keeps every ruby cell whole and every line edge legal (seed %i)', (seed) => {
    const { runs, cells, direction, lengthPt, grid } = generate(seed);
    const section = sectionFor(direction, lengthPt, grid);
    const { lines } = layoutParagraph(runs, direction, null, section);
    const texts = lines.map(lineText);
    expect(texts.join('')).toBe(runs.map((run) => run.text).join(''));
    for (const cell of cells) {
      // One line holds the whole base, and its guide is painted once.
      // Each cell's kanji are unique in the paragraph ('A' may repeat).
      const own = [...cell.base].filter((char) => char !== 'A');
      expect(texts.filter((text) => own.some((char) => text.includes(char))), cell.base).toHaveLength(1);
      expect(texts.some((text) => text.includes(cell.base)), cell.base).toBe(true);
      const guides = lines.flatMap((line) => textPlacements(line))
        .filter((entry) => entry.ruby && cell.base.startsWith(entry.text))
        .map((entry) => entry.ruby!.paintOps.map((op) => op.text).join(''));
      expect(guides, cell.base).toEqual([cell.guide]);
    }
    const context = `${direction ?? "lrTb"} ${lengthPt}pt ${grid.docGridType ?? ''} ${JSON.stringify(texts)} ${JSON.stringify(runs.map((run) => [run.text, run.ruby?.text ?? "", run.extra ?? {}]))}`;
    texts.forEach((text, index) => {
      if (index > 0) expect(LINE_START_FORBIDDEN.has([...text][0]!), `${index} of ${context}`).toBe(false);
      if (index < texts.length - 1 && LINE_END_FORBIDDEN.has([...text].at(-1)!)) {
        // Unavoidable only when the opener and the cell after it exceed a
        // whole line.
        const following = textPlacements(lines[index + 1]!);
        const cell = cells.find((candidate) => following[0]?.ruby && candidate.base.startsWith(following[0].text));
        const cellAdvance = cell ? following.filter((entry) => cell.base.includes(entry.text))
          .reduce((sum, entry) => sum + entry.advancePt, 0) : 0;
        expect(cell !== undefined && cellAdvance + 10 > lengthPt, `${index} of ${context}`).toBe(true);
      }
    });
    // Each line's extents are those of its own content: breaking the same
    // runs explicitly at the same places gives the same lines.
    const explicit = layoutBody([explicitlyBroken(runs, lines)], direction, section).layers.body
      .find((entry): entry is ParagraphLayout => entry.kind === 'paragraph')!.lines;
    expect(explicit.map(lineText)).toEqual(texts);
    explicit.forEach((line, index) => {
      const laid = lines[index]!;
      expect(laid.advancePt, `${index}`).toBeCloseTo(line.advancePt, 6);
      expect(laid.baselinePt - laid.bounds.yPt, `${index}`).toBeCloseTo(line.baselinePt - line.bounds.yPt, 6);
    });
  });
});
