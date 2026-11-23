import { describe, expect, it } from 'vitest';
import { createLayoutServices } from './layout-runtime.js';
import { layoutDocument } from './document-layout.js';
import type { LineLayout, ParagraphLayout, TextPlacement } from './layout/types.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, SectionProps } from './types.js';
import { createFontResolver } from './layout/font-service.js';
import { createTextLayoutService } from './layout/text.js';
import { testFontSnapshot } from './layout/test-font-snapshot.js';
import { paragraphMarkLineMetrics } from './line-layout.js';
import { wordVerticalSingleLinePx } from './layout/line-compatibility.js';

// WORD_VERTICAL_LINE_CENTRED_BOX, WORD_VERTICAL_RUBY_UPRIGHT_GUIDE and
// WORD_RUBY_ALIGN_PLACEMENT over a stub face (0.8/0.2 em box, one em per
// spacing character) that no metric table describes.

/** What follows a cluster's first character without an advance of its own. */
const JOINED_RE = /\p{Mn}|\u200d\p{Extended_Pictographic}?/gu;

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
        // One em per character (a nonspacing mark, a variation selector and
        // the rest of an emoji joiner sequence have no advance); the pair
        // "AV" kerns 0.2 em tighter.
        width: [...text.replace(JOINED_RE, '')].length * size - (text.match(/AV/g)?.length ?? 0) * size * 0.2,
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

type GuideRun = Readonly<{
  text: string; fonts: Readonly<Record<string, string>>; rtl?: boolean; cs?: boolean; fontHint?: string;
  bold?: boolean; italic?: boolean; boldCs?: boolean; italicCs?: boolean;
}>;
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
        bold: run.bold ?? false, italic: run.italic ?? false,
        boldCs: run.boldCs ?? false, italicCs: run.italicCs ?? false, color: null, language: null,
        fonts: { ascii: null, highAnsi: null, eastAsia: null, complexScript: null, ...run.fonts },
        rtl: run.rtl ?? false, cs: run.cs ?? false, fontHint: run.fontHint ?? null,
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

describe('ruby cell geometry (WORD_RUBY_ALIGN_PLACEMENT, review round 8)', () => {
  const lineTexts = (lines: readonly LineLayout[]) =>
    lines.map((line) => textPlacements(line).map((entry) => entry.text).join(''));
  const column = (pt: number) => ({ pageHeight: 60 + 80 + pt });
  const guide = { text: 'か', fontSizePt: 5, hpsRaisePt: 9, align: 'center' };
  const axis = (line: LineLayout) => line.baselinePt - line.bounds.yPt;
  const paragraphLines = (paragraph: DocParagraph & BodyElement, lengthPt: number) =>
    layoutBody([paragraph], 'tbRl', column(lengthPt)).layers.body
      .find((entry): entry is ParagraphLayout => entry.kind === 'paragraph')!.lines;
  const withBreakAfter = (runs: readonly Run[], index: number) => {
    const same = paragraphOf(runs);
    return { ...same, runs: [...same.runs.slice(0, index), { type: 'break', breakType: 'line' }, ...same.runs.slice(index)] } as unknown as DocParagraph & BodyElement;
  };

  it('keeps each merged guide run in the face its own slots select', () => {
    const ruby = {
      text: 'Aか', fontSizePt: 5, hpsRaisePt: 9,
      guideRuns: [
        { text: 'A', fonts: { ascii: 'Same', highAnsi: 'Same', eastAsia: 'TallFace' } },
        { text: 'か', fonts: { ascii: 'Other', highAnsi: 'Other', eastAsia: 'Same' } },
      ],
    };
    for (const direction of ['tbRl', undefined]) {
      const line = layoutParagraph([{ text: '漢字', ruby }], direction).lines[0]!;
      const ops = textPlacements(line).find((entry) => entry.ruby)!.ruby!.paintOps;
      expect(ops.map((op) => op.text).join(''), direction).toBe('Aか');
      for (const op of ops) {
        expect(JSON.stringify(op.fontRoute), `${direction} ${op.text}`).toContain('Same');
        expect(JSON.stringify(op.fontRoute), `${direction} ${op.text}`).not.toContain('TallFace');
      }
    }
  });

  it('keeps a forbidden ruby leader on the column when nothing before it can move', () => {
    const { lines } = layoutParagraph(
      [{ text: '前' }, { text: 'ゃ東京', ruby: guide }], 'tbRl', null, column(35),
    );
    expect(lineTexts(lines)).toEqual(['前ゃ東京']);
  });

  it('moves consecutive openers of different runs with the ruby cell after them', () => {
    const { lines } = layoutParagraph(
      [{ text: '前前「' }, { text: '『', fontFamily: 'OtherFace' }, { text: '東京', ruby: guide }],
      'tbRl', null, column(40),
    );
    expect(lineTexts(lines)).toEqual(['前前', '「『東京']);
  });

  it('does not lead a line with the space a forbidden leader pulls down', () => {
    const { lines } = layoutParagraph(
      [{ text: '前前A ' }, { text: 'ー後', fontFamily: 'OtherFace' }], 'tbRl', null, column(40),
    );
    expect(lineTexts(lines)).toEqual(['前前', 'A ー後']);
  });

  it('marks every placed piece of a ruby base as a piece of its cell', () => {
    const line = layoutParagraph([
      { text: '前' }, { text: '東', ruby: guide }, { text: '京', extra: { rubyBaseContinuation: true } }, { text: '後' },
    ]).lines[0]!;
    expect(textPlacements(line).map((entry) => [entry.text, entry.rubyBasePiece === true]))
      .toEqual([['前', false], ['東', true], ['京', true], ['後', false]]);
  });

  it('takes the Far East single-line class of a guide from its face, not from its text', () => {
    // Two faces with admitted line data and the same box: one of the Far
    // East class, one not.
    const face = (family: string, farEastCodePage: boolean) => ({
      ...testFontSnapshot([{ family, lineHeightRatio: 1 }])[family.toLowerCase()]!, farEastCodePage,
    });
    const localMetrics = { fareast: face('FarEast', true), western: face('Western', false) };
    const headOf = (family: string, text: string) => {
      const fonts = { ascii: family, highAnsi: family, eastAsia: family };
      const ruby = { text, fontSizePt: 5, hpsRaisePt: 30, guideRuns: [{ text, fonts }] };
      const section = {
        pageWidth: 600, pageHeight: 800, marginTop: 60, marginRight: 90, marginBottom: 80, marginLeft: 40,
        headerDistance: 36, footerDistance: 36, titlePage: false, evenAndOddHeaders: false, textDirection: 'tbRl',
      } as SectionProps;
      const document = {
        section, body: [paragraphOf([{ text: '漢', ruby }])],
        headers: { default: null, first: null, even: null }, footers: { default: null, first: null, even: null },
        fontFamilyClasses: {},
      } as unknown as DocxDocumentModel;
      const line = (layoutDocument(document, createLayoutServices(document, { measureContext: makeCtx(), localMetrics }),
        { currentDateMs: 0 }).pages[0]!.layers.body
        .find((entry): entry is ParagraphLayout => entry.kind === 'paragraph')!).lines[0]!;
      return line.baselinePt - line.bounds.yPt;
    };
    // A Latin letter in the Far East face reserves what a kana does there,
    // and a kana in the other face what a Latin letter does there.
    expect(headOf('FarEast', 'A')).toBeCloseTo(headOf('FarEast', 'か'), 6);
    expect(headOf('Western', 'か')).toBeCloseTo(headOf('Western', 'A'), 6);
    expect(headOf('FarEast', 'A')).toBeGreaterThan(headOf('Western', 'A') + 0.01);
  });

  it('moves earlier openers with an opener of another run that does not fit', () => {
    const { lines } = layoutParagraph(
      [{ text: '前前前「' }, { text: '『', fontFamily: 'OtherFace' }, { text: '東京', ruby: guide }],
      'tbRl', null, column(40),
    );
    expect(lineTexts(lines)).toEqual(['前前前', '「『東京']);
  });

  it('moves an opener with ordinary text of another run that does not fit', () => {
    const { lines } = layoutParagraph(
      [{ text: '前前前「' }, { text: '東京', fontFamily: 'OtherFace' }], 'tbRl', null, column(40),
    );
    expect(lineTexts(lines)).toEqual(['前前前', '「東京']);
  });

  it('pulls an opener of another run down with the character a forbidden leader takes', () => {
    const { lines } = layoutParagraph(
      [{ text: '前前「' }, { text: '東', fontFamily: 'OtherFace' }, { text: 'ー後' }], 'tbRl', null, column(40),
    );
    expect(lineTexts(lines)).toEqual(['前前', '「東ー後']);
  });

  it('restores the column extents a retracted opener had added', () => {
    const runs: Run[] = [
      { text: '前前' }, { text: '「', fontSize: 20, position: 18 }, { text: '東京', ruby: guide },
    ];
    const wrapped = layoutParagraph(runs, 'tbRl', null, column(45)).lines;
    expect(lineTexts(wrapped)).toEqual(['前前', '「東京']);
    const explicit = paragraphLines(withBreakAfter(runs, 1), 45);
    expect(lineTexts(explicit)).toEqual(['前前', '「東京']);
    expect(axis(wrapped[0]!)).toBeCloseTo(axis(explicit[0]!), 6);
    expect(wrapped[0]!.advancePt).toBeCloseTo(explicit[0]!.advancePt, 6);
  });

  it('restores the column extents a retracted suffix of ordinary text had added', () => {
    // A raised 後 leaves the column ahead of a forbidden leader that does
    // not fit: the column keeps only the extents of 前前.
    const runs: Run[] = [
      { text: '前前' }, { text: '後', fontSize: 20, position: 18 }, { text: 'ゃ東京', ruby: guide },
    ];
    const wrapped = layoutParagraph(runs, 'tbRl', null, column(50)).lines;
    expect(lineTexts(wrapped)).toEqual(['前前', '後ゃ東京']);
    const explicit = paragraphLines(withBreakAfter(runs, 1), 50);
    expect(axis(wrapped[0]!)).toBeCloseTo(axis(explicit[0]!), 6);
    expect(wrapped[0]!.advancePt).toBeCloseTo(explicit[0]!.advancePt, 6);
  });

  it('shapes a guide of many formatted runs once, not once per run seam', () => {
    const count = 400;
    const work = (runs: number) => {
      const fonts = (index: number) => ({ ascii: 'Same', highAnsi: 'Same', eastAsia: `Unused${index}` });
      const ruby = {
        text: 'a'.repeat(count), fontSizePt: 5, hpsRaisePt: 9,
        guideRuns: Array.from({ length: runs }, (_, index) => ({ text: 'a'.repeat(count / runs), fonts: fonts(index) })),
      };
      const before = measuredCharacters;
      const line = layoutParagraph([{ text: '漢字', ruby }]).lines[0]!;
      expect(textPlacements(line).find((entry) => entry.ruby)!.ruby!.paintOps).toHaveLength(count);
      return measuredCharacters - before;
    };
    // The same letters in one run and in one run per letter.
    expect(work(count)).toBeLessThan(work(1) * 3);
  });
});

describe('ruby cell geometry (WORD_RUBY_ALIGN_PLACEMENT, review round 9)', () => {
  const lineTexts = (lines: readonly LineLayout[]) =>
    lines.map((line) => textPlacements(line).map((entry) => entry.text).join(''));
  const column = (pt: number) => ({ pageHeight: 60 + 80 + pt });
  const guide = { text: 'か', fontSizePt: 5, hpsRaisePt: 9, align: 'center' };

  it('aligns the guide of a right-to-left base on the cell however its runs are split', () => {
    const ruby = { text: 'かき', fontSizePt: 5, hpsRaisePt: 9, align: 'right' };
    const guideOf = (runs: readonly Run[]) => {
      const line = layoutParagraph([{ text: '前' }, ...runs, { text: '後' }]).lines[0]!;
      return textPlacements(line).find((entry) => entry.ruby)!.ruby!.paintOps.map((op) => op.origin.xPt);
    };
    const whole = guideOf([{ text: 'אבגד', ruby, extra: { rtl: true } }]);
    const split = guideOf([
      { text: 'אב', ruby, extra: { rtl: true } },
      { text: 'גד', extra: { rtl: true, rubyBaseContinuation: true, bold: true } },
    ]);
    expect(split).toHaveLength(whole.length);
    split.forEach((xPt, index) => expect(xPt).toBeCloseTo(whole[index]!, 6));
  });

  it('takes a character back with its variation selector', () => {
    const { lines } = layoutParagraph(
      [{ text: '前辻\u{E0100}' }, { text: 'ゃ京', ruby: guide }], 'tbRl', null, column(35),
    );
    expect(lineTexts(lines)).toEqual(['前', '辻\u{E0100}ゃ京']);
  });

  it('keeps a fitText region whole when its opener moves with the ruby cell', () => {
    const { lines } = layoutParagraph(
      [{ text: '前前前' }, { text: 'AB「', extra: { fitTextVal: 400, fitTextId: 7 } }, { text: '東京', ruby: guide }],
      'tbRl', null, column(55),
    );
    expect(lineTexts(lines)).toEqual(['前前前', 'AB「東京']);
  });

  it('keeps a tate-chu-yoko cell whole when a forbidden leader takes it back', () => {
    const { lines } = layoutParagraph(
      [{ text: '前前' }, { text: '12', eastAsianVert: true }, { text: 'ゃ京', ruby: guide }], 'tbRl', null, column(35),
    );
    expect(lineTexts(lines)).toEqual(['前前', '12ゃ京']);
  });

  it('allocates the retained head of a cut Latin run from its natural width', () => {
    const grid = { ...column(35), docGridType: 'snapToChars', docGridLinePitch: 18, docGridCharSpace: 0 } as Partial<SectionProps>;
    const narrow = { extra: { charScale: 0.6 } };
    const advances = (runs: readonly Run[]) => {
      const { lines } = layoutParagraph(runs, 'tbRl', null, grid);
      return [lineTexts(lines)[0], textPlacements(lines[0]!).map((entry) => entry.advancePt)];
    };
    const bold = { extra: { charScale: 0.6, bold: true } };
    const retracted = advances([{ text: 'X', ...narrow }, { text: 'AB', ...bold }, { text: 'ゃ京', ruby: guide }]);
    const control = advances([{ text: 'X', ...narrow }, { text: 'A', ...bold }]);
    expect(retracted[0]).toBe('XA');
    expect(retracted).toEqual(control);
  });

  it('finds the font span of each guide letter without rescanning the spans', () => {
    // Span reads made while a mixed-face horizontal guide is laid out, with
    // the text service's shapes observed.
    const visits = (count: number) => {
      const text = 'Aか'.repeat(count / 2);
      const fonts = { ascii: 'FaceA', highAnsi: 'FaceA', eastAsia: 'FaceB' };
      const ruby = { text, fontSizePt: 5, hpsRaisePt: 9, guideRuns: [{ text, fonts }] };
      const section = {
        pageWidth: 6000, pageHeight: 800, marginTop: 60, marginRight: 90, marginBottom: 80, marginLeft: 40,
        headerDistance: 36, footerDistance: 36, titlePage: false, evenAndOddHeaders: false,
      } as SectionProps;
      const document = {
        section, body: [paragraphOf([{ text: '漢字', ruby }])],
        headers: { default: null, first: null, even: null }, footers: { default: null, first: null, even: null },
        fontFamilyClasses: {},
      } as unknown as DocxDocumentModel;
      const services = createLayoutServices(document, { measureContext: makeCtx() });
      let reads = 0;
      const observed: typeof services.text = {
        ...services.text,
        shape: (request) => {
          const shape = services.text.shape(request);
          if (request.text !== text) return shape;
          return { ...shape, spans: shape.spans.map((span) => new Proxy(span, {
            get(target, key, receiver) {
              if (key === 'start' || key === 'end') reads += 1;
              return Reflect.get(target, key, receiver);
            },
          })) };
        },
      };
      const page = layoutDocument(document, { ...services, text: observed }, { currentDateMs: 0 }).pages[0]!;
      const line = page.layers.body.find((entry): entry is ParagraphLayout => entry.kind === 'paragraph')!.lines[0]!;
      expect(textPlacements(line).find((entry) => entry.ruby)!.ruby!.paintOps).toHaveLength(count);
      return reads;
    };
    const small = visits(100);
    expect(small).toBeGreaterThan(0);
    expect(visits(400)).toBeLessThan(small * 6);
  });
});

describe('ruby cell geometry (WORD_RUBY_ALIGN_PLACEMENT, review round 10)', () => {
  const lineTexts = (lines: readonly LineLayout[]) =>
    lines.map((line) => textPlacements(line).map((entry) => entry.text).join(''));
  const column = (pt: number) => ({ pageHeight: 60 + 80 + pt });
  const guide = { text: 'か', fontSizePt: 5, hpsRaisePt: 9, align: 'center' };
  const fit = { fitTextVal: 400, fitTextId: 7 };

  it('keeps a fitText region and the ruby cells in it as one unit', () => {
    // The region's opener is a ruby base: the region moves whole with it.
    const opener = layoutParagraph(
      [{ text: '前前前' }, { text: 'AB', extra: fit }, { text: '「', ruby: guide, extra: fit }, { text: '東京' }],
      'tbRl', null, column(55),
    );
    expect(lineTexts(opener.lines)).toEqual(['前前前', 'AB「東京']);
    // The region's first run is a ruby base: the region is admitted whole.
    const wide = { fitTextVal: 400, fitTextId: 8 };
    const first = layoutParagraph(
      [{ text: '前前前前' }, { text: '東', ruby: guide, extra: wide }, { text: '京', extra: wide }],
      'tbRl', null, column(55),
    );
    expect(lineTexts(first.lines)).toEqual(['前前前前', '東京']);
  });

  it('keeps a widened right-to-left base together however its runs are split', () => {
    const ruby = { text: 'かきくけこかきくけこ', fontSizePt: 5, hpsRaisePt: 9, align: 'right' };
    const baseOf = (runs: readonly Run[]) => {
      const line = layoutParagraph([{ text: '前' }, ...runs, { text: '後' }]).lines[0]!;
      const pieces = textPlacements(line).filter((entry) => entry.rubyBasePiece);
      // The ink of the base: each piece's glyph origins.
      return pieces.flatMap((entry) => entry.paintOps.map((op) => entry.origin.xPt + (op.offset?.xPt ?? 0)))
        .sort((left, right) => left - right);
    };
    const whole = baseOf([{ text: 'אבגד', ruby, extra: { rtl: true } }]);
    const split = baseOf([
      { text: 'אב', ruby, extra: { rtl: true } },
      { text: 'גד', extra: { rtl: true, rubyBaseContinuation: true, bold: true } },
    ]);
    // The split base starts where the whole one does and has no gap inside.
    expect(split[0]).toBeCloseTo(whole[0]!, 6);
    expect(split.at(-1)! - split[0]!).toBeCloseTo(20, 6);
  });

  it('selects a guide run\'s faces by its own script controls', () => {
    const faceOf = (ruby: Ruby, extra: Record<string, unknown> = {}) => {
      const line = layoutParagraph([{ text: '漢字', ruby, extra }]).lines[0]!;
      return JSON.stringify(textPlacements(line).find((entry) => entry.ruby)!.ruby!.paintOps[0]!.fontRoute);
    };
    const fonts = { ascii: 'AsciiFace', highAnsi: 'AsciiFace', eastAsia: 'EastFace', complexScript: 'ComplexFace' };
    // A right-to-left guide run on a left-to-right base.
    expect(faceOf({ text: 'سلام', fontSizePt: 5, hpsRaisePt: 9, guideRuns: [{ text: 'سلام', fonts, rtl: true }] }))
      .toContain('ComplexFace');
    // An East Asian guide run on a right-to-left base.
    expect(faceOf({ text: 'か', fontSizePt: 5, hpsRaisePt: 9, guideRuns: [{ text: 'か', fonts }] }, { rtl: true }))
      .toContain('EastFace');
  });

  it('takes a character back with a variation selector of another run', () => {
    const { lines } = layoutParagraph(
      [{ text: '前辻' }, { text: '\u{E0100}', extra: { bold: true } }, { text: 'ゃ京', ruby: guide }],
      'tbRl', null, column(35),
    );
    expect(lineTexts(lines)[0]).toBe('前');
  });

  it('paints a vertical guide\'s combining marks and variation selectors with their bases', () => {
    const opsOf = (text: string) => {
      const line = layoutParagraph([{ text: '漢字', ruby: { text, fontSizePt: 5, hpsRaisePt: 9 } }], 'tbRl').lines[0]!;
      return textPlacements(line).find((entry) => entry.ruby)!.ruby!.paintOps.map((op) => [op.text, op.glyphOrientation]);
    };
    expect(opsOf('か\u3099き')).toEqual([['か\u3099', 'upright'], ['き', 'upright']]);
    expect(opsOf('辻\u{E0100}')).toEqual([['辻\u{E0100}', 'upright']]);
  });
});

describe('ruby cell geometry (WORD_RUBY_ALIGN_PLACEMENT, review round 11)', () => {
  const lineTexts = (lines: readonly LineLayout[]) =>
    lines.map((line) => textPlacements(line).map((entry) => entry.text).join(''));
  const column = (pt: number) => ({ pageHeight: 60 + 80 + pt });
  const guideOps = (ruby: Ruby, direction?: string) => {
    const line = layoutParagraph([{ text: '漢字', ruby }], direction).lines[0]!;
    return textPlacements(line).find((entry) => entry.ruby)!.ruby!.paintOps;
  };

  it('keeps the selector that opens a run of other formatting with its base', () => {
    const { lines } = layoutParagraph(
      [{ text: '辻' }, { text: '\u{E0100}京', extra: { bold: true } }], 'tbRl', null, column(10),
    );
    expect(lineTexts(lines)).toEqual(['辻\u{E0100}', '京']);
  });

  it('paints a mark in a guide run of its own with its base, in the base\'s face', () => {
    const run = (text: string, eastAsia: string) => ({ text, fonts: { ascii: 'Latin', highAnsi: 'Latin', eastAsia } });
    const ruby = { text: 'か\u3099き', fontSizePt: 5, hpsRaisePt: 9, guideRuns: [run('か', 'FaceA'), run('\u3099き', 'FaceB')] };
    for (const direction of [undefined, 'tbRl']) {
      const ops = guideOps(ruby, direction);
      expect(ops.map((op) => op.text), direction).toEqual(['か\u3099', 'き']);
      expect(JSON.stringify(ops[0]!.fontRoute), direction).toContain('FaceA');
      expect(JSON.stringify(ops[1]!.fontRoute), direction).toContain('FaceB');
    }
  });

  it('joins guide runs that select one face whatever hint only one of them carries', () => {
    const fonts = { ascii: 'Latin', highAnsi: 'Latin', eastAsia: 'East', complexScript: 'Complex' };
    const ops = guideOps({
      text: 'سلام', fontSizePt: 5, hpsRaisePt: 9,
      guideRuns: [{ text: 'س', fonts, rtl: true, fontHint: 'eastAsia' }, { text: 'لام', fonts, rtl: true }],
    });
    expect(ops.map((op) => op.text)).toEqual(['سلام']);
  });

  it('gives a complex-script guide run its complex-script bold and italic', () => {
    const fonts = { ascii: 'Latin', highAnsi: 'Latin', eastAsia: 'East', complexScript: 'Complex' };
    const styleOf = (run: Partial<GuideRun>) => {
      const op = guideOps({ text: 'سلام', fontSizePt: 5, hpsRaisePt: 9, guideRuns: [{ text: 'سلام', fonts, ...run }] })[0]!;
      return [op.fontWeight, op.fontStyle];
    };
    expect(styleOf({ rtl: true, boldCs: true, italicCs: true })).toEqual([700, 'italic']);
    // The Latin axes do not style complex-script text.
    expect(styleOf({ rtl: true, bold: true, italic: true })).toEqual([400, 'normal']);
  });

  it('plans a vertical guide by extended grapheme clusters', () => {
    const emoji = '\u{1F469}\u200D\u{1F4BB}';
    const ops = guideOps({ text: `${emoji}か`, fontSizePt: 5, hpsRaisePt: 9 }, 'tbRl');
    expect(ops.map((op) => op.text)).toEqual([emoji, 'か']);
  });
});

describe('ruby cell geometry (WORD_RUBY_ALIGN_PLACEMENT, review round 12)', () => {
  const lineTexts = (lines: readonly LineLayout[]) =>
    lines.map((line) => textPlacements(line).map((entry) => entry.text).join(''));

  it('breaks before a cluster whose rest, in a run of other formatting, does not fit', () => {
    const flag = ['\u{1F1EF}', '\u{1F1F5}'];
    for (const direction of [undefined, 'tbRl']) {
      const section = (lengthPt: number) => (direction ? { pageHeight: 60 + 80 + lengthPt } : { pageWidth: 40 + 90 + lengthPt });
      const split = [{ text: `前${flag[0]}` }, { text: flag[1]!, extra: { bold: true } }];
      expect(lineTexts(layoutParagraph(split, direction, null, section(20)).lines), direction).toEqual(['前', flag.join('')]);
      // Nothing before the cluster can stay: it keeps its line whole.
      const alone = [{ text: flag[0]! }, { text: flag[1]!, extra: { bold: true } }];
      expect(lineTexts(layoutParagraph(alone, direction, null, section(10)).lines), direction).toEqual([flag.join('')]);
    }
  });

  it('paints the forms the vertical planner selects for a guide\'s punctuation', () => {
    const ops = (text: string) => {
      const line = layoutParagraph([{ text: '漢字', ruby: { text, fontSizePt: 5, hpsRaisePt: 9 } }], 'tbRl').lines[0]!;
      return textPlacements(line).find((entry) => entry.ruby)!.ruby!.paintOps.map((op) => op.text);
    };
    expect(ops('「A、か」')).toEqual(['\uFE41', 'A', '\uFE11', 'か', '\uFE42']);
    // A mark stays with the form selected for its base.
    expect(ops('、\u3099か')).toEqual(['\uFE11\u3099', 'か']);
  });
});

describe('ruby cell geometry (WORD_RUBY_ALIGN_PLACEMENT, review round 13)', () => {
  const lineTexts = (lines: readonly LineLayout[]) =>
    lines.map((line) => textPlacements(line).map((entry) => entry.text).join(''));

  it('judges the rest of a cluster by its character-grid cells', () => {
    const flag = ['\u{1F1EF}', '\u{1F1F5}'];
    for (const direction of [undefined, 'tbRl']) {
      const section = {
        ...(direction ? { pageHeight: 60 + 80 + 29 } : { pageWidth: 40 + 90 + 29 }),
        docGridType: 'snapToChars', docGridLinePitch: 18, docGridCharSpace: 0,
      } as Partial<SectionProps>;
      const { lines } = layoutParagraph(
        [{ text: '前' }, { text: flag[0]!, fontSize: 7 }, { text: flag[1]!, fontSize: 7, extra: { bold: true } }],
        direction, null, section,
      );
      expect(lineTexts(lines), direction).toEqual(['前', flag.join('')]);
    }
  });

  it('pads a widened base of mixed direction around the base as it stands on the line', () => {
    const ruby = { text: 'abcdefghijklmno', fontSizePt: 5, hpsRaisePt: 9, align: 'center' };
    const line = layoutParagraph([
      { text: '前' },
      { text: 'אב', ruby, extra: { rtl: true } },
      { text: 'גד', extra: { rtl: true, rubyBaseContinuation: true, bold: true } },
      { text: 'XY', extra: { rubyBaseContinuation: true } },
      { text: '後' },
    ]).lines[0]!;
    const inks = textPlacements(line).filter((entry) => entry.rubyBasePiece)
      .map((entry) => [entry.text, entry.origin.xPt + (entry.paintOps[0]?.offset?.xPt ?? 0)] as const)
      .sort((left, right) => left[1] - right[1]);
    // Bidirectional order on the line, with no gap between the pieces.
    expect(inks.map(([text]) => text)).toEqual(['גד', 'אב', 'XY']);
    expect(inks[1]![1] - inks[0]![1]).toBeCloseTo(20, 6);
    expect(inks[2]![1] - inks[1]![1]).toBeCloseTo(20, 6);
  });

  it('lays right-to-left guide runs of different formatting out in visual order', () => {
    const fonts = { ascii: 'Latin', highAnsi: 'Latin', eastAsia: 'East', complexScript: 'Complex' };
    const ruby = {
      text: 'سلام', fontSizePt: 5, hpsRaisePt: 9,
      guideRuns: [{ text: 'سل', fonts, rtl: true }, { text: 'ام', fonts, rtl: true, boldCs: true }],
    };
    const line = layoutParagraph([{ text: '漢字漢字', ruby }]).lines[0]!;
    const ops = textPlacements(line).find((entry) => entry.ruby)!.ruby!.paintOps;
    const xOf = (text: string) => ops.find((op) => op.text === text)!.origin.xPt;
    // The later run stands to the left of the earlier one.
    expect(xOf('ام')).toBeLessThan(xOf('سل'));
    expect(xOf('سل') - xOf('ام')).toBeCloseTo(10, 6);
  });
});

describe('ruby cells in bidirectional lines (WORD_RUBY_ALIGN_PLACEMENT)', () => {
  it('keeps a ruby base of mixed direction together beside right-to-left text', () => {
    const ruby = { text: 'かきくけ', fontSizePt: 5, hpsRaisePt: 9, align: 'center' };
    for (const [before, after] of [[true, false], [false, true], [true, true]]) {
      const line = layoutParagraph([
        { text: '前' },
        ...(before ? [{ text: 'שלום', extra: { rtl: true } }] : []),
        { text: 'אב', ruby, extra: { rtl: true } },
        { text: 'XY', extra: { rubyBaseContinuation: true } },
        ...(after ? [{ text: 'עברית', extra: { rtl: true } }] : []),
        { text: '後' },
      ]).lines[0]!;
      const placed = textPlacements(line).sort((left, right) => left.origin.xPt - right.origin.xPt);
      const cell = placed.map((entry, index) => (entry.rubyBasePiece ? index : -1)).filter((index) => index >= 0);
      // The two pieces stand side by side, nothing between them.
      expect(cell, `${before} ${after}`).toHaveLength(2);
      expect(cell[1]! - cell[0]!, `${before} ${after}`).toBe(1);
      // The guide stays over them.
      const pieces = placed.filter((entry) => entry.rubyBasePiece);
      const startPt = Math.min(...pieces.map((entry) => entry.bounds.xPt));
      const endPt = Math.max(...pieces.map((entry) => entry.bounds.xPt + entry.bounds.widthPt));
      for (const op of pieces.find((entry) => entry.ruby)!.ruby!.paintOps) {
        expect(op.origin.xPt, `${before} ${after}`).toBeGreaterThanOrEqual(startPt - 1e-6);
        expect(op.origin.xPt + 5, `${before} ${after}`).toBeLessThanOrEqual(endPt + 1e-6);
      }
    }
  });
});

describe('ruby cell geometry (WORD_RUBY_ALIGN_PLACEMENT, review round 14)', () => {
  it('keeps a right-to-left letter before a mixed ruby base out of the base', () => {
    const ruby = { text: 'abcdefghijklmno', fontSizePt: 5, hpsRaisePt: 9, align: 'center' };
    const line = layoutParagraph([
      { text: 'ק', extra: { rtl: true } },
      { text: 'אב', ruby, extra: { rtl: true } },
      { text: 'XY', extra: { rubyBaseContinuation: true } },
    ]).lines[0]!;
    const order = textPlacements(line).sort((left, right) => left.origin.xPt - right.origin.xPt).map((entry) => entry.text);
    expect(order).toEqual(['ק', 'אב', 'XY']);
  });

  it('orders a right-to-left guide by its characters, whatever formatting divides it', () => {
    const fonts = { ascii: 'Latin', highAnsi: 'Latin', eastAsia: 'East', complexScript: 'Complex' };
    const read = (guideRuns: readonly GuideRun[]) => {
      const ruby = { text: guideRuns.map((run) => run.text).join(''), fontSizePt: 5, hpsRaisePt: 9, guideRuns };
      const line = layoutParagraph([{ text: '漢字漢字', ruby }]).lines[0]!;
      return [...textPlacements(line).find((entry) => entry.ruby)!.ruby!.paintOps]
        .sort((left, right) => left.origin.xPt - right.origin.xPt).map((op) => op.text).join('');
    };
    // §17.3.2.30: European digits keep their direction in a right-to-left run.
    expect(read([{ text: '12', fonts, rtl: true }, { text: '34', fonts, rtl: true, boldCs: true }])).toBe('1234');
    // Letters of two right-to-left runs still stand in reverse.
    expect(read([{ text: 'سل', fonts, rtl: true }, { text: 'ام', fonts, rtl: true, boldCs: true }])).toBe('امسل');
  });
});

describe('grapheme clusters at kinsoku breaks (§17.3.1.16, UAX #29)', () => {
  const lineTexts = (lines: readonly LineLayout[]) =>
    lines.map((line) => textPlacements(line).map((entry) => entry.text).join(''));
  const column = (pt: number) => ({ pageHeight: 60 + 80 + pt });

  it('breaks text with zero-advance marks exactly as the same text without them', () => {
    // A variation selector or a combining mark adds no advance and no break
    // opportunity: the kinsoku retraction counts a character with its marks
    // as one, so it never lands inside the cluster and moves on from there.
    const texts = ['前前「辻ー後', '前「辻ー後', '前前前ゃー後', '前「「辻ー」後', '「辻ー後前', '前前辻ゃ「後'];
    for (const mark of ['\u{E0100}', '\u3099']) {
      for (const plain of texts) {
        for (const lengthPt of [25, 35, 45, 55]) {
          const marked = plain.replace(/[辻ゃ]/gu, (char) => char + mark);
          const linesOf = (text: string) => lineTexts(layoutParagraph([{ text }], 'tbRl', null, column(lengthPt)).lines)
            .map((line) => line.replaceAll(mark, ''));
          expect(linesOf(marked), `${plain} ${lengthPt}pt ${mark.codePointAt(0)!.toString(16)}`).toEqual(linesOf(plain));
        }
      }
    }
  });
});

describe('ruby cells in generated paragraphs (WORD_RUBY_ALIGN_PLACEMENT)', () => {
  // Deterministic generator. Every seed lays out one paragraph of ordinary
  // runs and ruby cells (split base runs, fitText pieces, kinsoku leaders and
  // openers) in an lrTb, tbRl or tbLrV column of varied length, and belongs
  // to one family that adds a feature to that base.
  const FAMILIES = [
    'base',
    // Guide runs with their own rFonts slots, some no character uses.
    'guideFaces',
    // A guide of many one-letter formatted runs.
    'guideRuns',
    // Cells led by a line-start-forbidden character, after little content.
    'leaderCell',
    // Consecutive openers of different runs before a cell.
    'openers',
    // Larger and raised runs and openers, which lengthen a line's extents.
    'raised',
    'justified',
    'tateChuYoko',
    'tabs',
    'list',
    'bidi',
    'float',
    // Characters with variation selectors, which no line edge may detach.
    'graphemes',
    // fitText regions of ordinary text, some ending with an opener.
    'fitRegion',
    // Right-to-left ruby bases in one or two runs.
    'rtlBase',
    // Condensed Latin runs of alternating weight on a character grid.
    'scaledLatin',
  ] as const;
  type Family = typeof FAMILIES[number];
  const SEEDS_PER_FAMILY = 100;

  const random = (seed: number) => () => {
    seed = (seed + 0x6d2b79f5) | 0;
    let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
  const KANJI = [...'一二三四五六七八九十百千万東京都庁前後中左右上下山川田林森村町市区円年月日火水木金土'];
  const LEADERS = [...'ゃょー」'];
  const OPENERS = [...'「『'];
  const LINE_START_FORBIDDEN = new Set(LEADERS);
  const LINE_END_FORBIDDEN = new Set(OPENERS);
  const ALIGNS = ['center', 'left', 'right', 'distributeLetter', undefined];
  const GUIDE_FACES = ['FaceA', 'FaceB', 'TallFace'];
  /** One source run of the paragraph, as the document model holds it. */
  type Source = Readonly<{ run: Record<string, unknown>; text: string; sizePt: number }>;
  type Cell = Readonly<{
    runs: readonly number[]; guide: string; guideRuns?: readonly GuideRun[];
    /** The text the guide paints, when it is not the guide text itself. */
    painted?: string;
  }>;
  const VARIATION_SELECTOR = '\u{E0100}';
  // A flag is one cluster of two regional indicators.
  const FLAG = ['\u{1F1EF}', '\u{1F1F5}'];
  // The forms the vertical planner paints for punctuation without a
  // vertical glyph of its own.
  const VERTICAL_FORMS: Readonly<Record<string, string>> = { '「': '\uFE41', '」': '\uFE42', '、': '\uFE11' };
  // An emoji joiner sequence: one cluster of three characters.
  const EMOJI = '\u{1F469}\u200D\u{1F4BB}';
  /** The UTF-16 offsets at which `value` may be divided between clusters. */
  const clusterBoundaries = (value: string) => new Set([
    0, ...[...new Intl.Segmenter(undefined, { granularity: 'grapheme' }).segment(value)].map((part) => part.index), value.length,
  ]);

  function generate(seed: number) {
    const family: Family = FAMILIES[seed % FAMILIES.length]!;
    const next = random(seed);
    const pick = <T,>(items: readonly T[]) => items[Math.floor(next() * items.length)]!;
    const pool = [...KANJI];
    const take = () => pool.splice(Math.floor(next() * pool.length), 1)[0]!;
    const direction = family === 'tateChuYoko' ? 'tbRl'
      : family === 'float' || family === 'bidi' || family === 'rtlBase' ? undefined : pick([undefined, 'tbRl', 'tbLrV']);
    const grid = family !== 'raised' && family !== 'float' && (family === 'scaledLatin' || next() < 0.3)
      ? { docGridType: 'snapToChars', docGridLinePitch: 18, docGridCharSpace: 0 } as Partial<SectionProps> : {};
    const template = paragraphOf([{ text: '' }]).runs[0] as unknown as Record<string, unknown>;
    const sources: Source[] = [];
    const cells: Cell[] = [];
    // Runs that are one indivisible unit each: a fitText region, with the
    // ruby cell that shares it.
    const regions: number[][] = [];

    // `advancePt`: the advance of one character when it is not the run's size
    // (a character-grid cell).
    const text = (value: string, extra: Record<string, unknown> = {}, advancePt?: number) => {
      sources.push({
        run: { ...template, text: value, ...extra }, text: value,
        sizePt: advancePt ?? (extra.fontSize as number | undefined) ?? 10,
      });
      return sources.length - 1;
    };
    const raised = () => family === 'raised' && next() < 0.5
      ? pick([{ fontSize: 20 }, { position: 18 }, { fontSize: 20, position: 18 }]) : {};
    let fitId = 100;
    // Whether an ordinary run may open with a leader: after a ruby cell, or
    // after two or more ordinary characters of its own format, so that the
    // leader, the character before it and an opener before that are never
    // one unbreakable unit longer than the shortest column.
    let leaderAllowed = false;
    if (family === 'float') {
      sources.push({
        run: {
          type: 'image', imagePath: 'word/media/image1.png', mimeType: 'image/png',
          widthPt: 15 + Math.floor(next() * 20), heightPt: 20 + Math.floor(next() * 30),
          anchor: true, anchorYFromPara: true, anchorYRelativeFrom: 'paragraph', anchorYPt: 0,
          anchorXPt: 0, anchorXFromMargin: true, anchorXRelativeFrom: 'margin',
          wrapMode: 'square', wrapSide: 'bothSides',
        },
        text: '￼', sizePt: 0,
      });
    }
    const items = family === 'leaderCell' ? 2 + Math.floor(next() * 3) : 3 + Math.floor(next() * 6);
    for (let index = 0; index < items; index++) {
      const cellItem = family === 'leaderCell' ? index > 0 : next() >= 0.5;
      // A fitText region just emitted that the ruby cell below shares.
      let sharedRegion: { runs: number[]; fit: Record<string, unknown> } | undefined;
      if (!cellItem) {
        if (family === 'tabs' && index > 0 && next() < 0.4) {
          sources.push({ run: { type: 'tab' }, text: '\t', sizePt: 0 });
          leaderAllowed = false;
        }
        // Right-to-left text beside the cells: in the rtlBase family it
        // would join a cell's right-to-left runs when the line is reordered.
        if ((family === 'bidi' || family === 'rtlBase') && next() < 0.4) {
          text(pick(['שלום', 'עברית', 'אב']), { rtl: true });
          leaderAllowed = false;
          continue;
        }
        if (family === 'fitRegion' && next() < 0.4) {
          const value = pick(['AB', 'ABC', 'A']) + (next() < 0.5 ? '「' : '');
          // Some regions go on into the ruby cell after them.
          const shared = next() < 0.4;
          const fit = { fitTextVal: 200 * ([...value].length + (shared ? 2 : 0)) + 40, fitTextId: fitId++ };
          const runs = [text(value, fit)];
          regions.push(runs);
          leaderAllowed = false;
          if (!shared) continue;
          sharedRegion = { runs, fit };
        } else if (family === 'graphemes' && next() < 0.3) {
          // A cluster divided between runs of different formatting: a
          // selector alone, a selector opening a run, an emoji sequence.
          const [head, tail] = pick([
            [take(), VARIATION_SELECTOR], [take(), VARIATION_SELECTOR + take()], ['\u{1F469}', '\u200D\u{1F4BB}'],
            // A flag: two regional indicators, each with an advance.
            [FLAG[0]!, FLAG[1]!],
          ]);
          // On a character grid a smaller cluster still takes whole cells:
          // its natural width fits where its cells do not.
          const small = grid.docGridType !== undefined && next() < 0.6 ? { fontSize: 7 } : {};
          text(head!, small, 10);
          text(tail!, { ...small, bold: true }, 10);
          leaderAllowed = false;
          continue;
        }
        if (!sharedRegion) {
          if (family === 'scaledLatin') {
            const letters = Array.from({ length: 1 + Math.floor(next() * 3) }, () => pick([...'ABCXYZ'])).join('');
            text(letters, { charScale: 0.6, bold: next() < 0.5 });
            leaderAllowed = false;
            continue;
          }
          const style = raised();
          const styled = Object.keys(style).length > 0;
          let value = Array.from({ length: family === 'leaderCell' ? 1 : 1 + Math.floor(next() * 3) }, take)
            .map((char) => char + (family === 'graphemes' && next() < 0.5 ? VARIATION_SELECTOR : '')).join('');
          if (!styled && leaderAllowed && next() < 0.3) value = pick(LEADERS) + value;
          const opener = !styled && next() < 0.25;
          if (opener) value += '「';
          text(value, style);
          leaderAllowed = !styled && !opener && [...value].length > 1;
          continue;
        }
      }
      if ((family === 'openers' || family === 'raised') && next() < 0.7) {
        // §17.3.1.16 openers before the cell, each its own run.
        const count = family === 'openers' ? 1 + Math.floor(next() * 3) : 1;
        for (let opener = 0; opener < count; opener++) {
          text(pick(OPENERS), { ...raised(), ...(opener % 2 === 1 ? { fontFamily: 'OtherFace' } : {}) });
        }
      }
      const rtlBase = family === 'rtlBase' && next() < 0.8;
      const leader = rtlBase ? ''
        : (family === 'leaderCell' || family === 'scaledLatin' || family === 'graphemes' ? next() < 0.7 : next() < 0.1)
          ? pick(LEADERS) : '';
      const tateChuYoko = family === 'tateChuYoko' && next() < 0.7;
      const pieces = rtlBase
        // Right-to-left runs, some followed or preceded by a left-to-right one.
        ? pick([['אבגד'], ['אב', 'גד'], ['א', 'בגד'], ['אבג', 'ד'], ['אב', 'גד', 'XY'], ['XY', 'אב', 'גד'], ['אב', 'XY', 'גד']])
        : tateChuYoko
          ? [...(next() < 0.5 ? [leader + take()] : []), pick(['12', '34', '56'])]
          : [leader + take() + (next() < 0.3 ? 'A' : ''), ...(next() < 0.5 ? [take()] : [])];
      const length = family === 'guideRuns' ? 12 + Math.floor(next() * 30)
        : 1 + Math.floor(next() * (rtlBase ? 16 : 6));
      const letters = family === 'guideFaces'
        // Latin, kana, and Arabic letters and digits (in right-to-left runs).
        ? Array.from({ length }, () => (next() < 0.4 ? pick([...'ABC']) : next() < 0.4 ? pick([...'سلم1234']) : pick([...'あいうえお']))).join('')
        : family === 'guideRuns' ? 'a'.repeat(length)
          : family === 'graphemes'
            // Kana with a combining voiced mark, kanji with a variation selector.
            ? Array.from({ length: Math.min(length, 3) }, () => pick(['か\u3099', 'き', '辻' + VARIATION_SELECTOR, 'く\u3099', EMOJI])).join('')
            : family === 'tateChuYoko'
              // Punctuation, which a vertical guide paints in its vertical forms.
              ? Array.from({ length }, () => pick([...'あいう「」、A'])).join('')
              : 'あいうえおかきくけこさしすせそたち'.slice(0, length);
      const guideRuns: GuideRun[] | undefined = family === 'guideFaces'
        ? [...letters].reduce<GuideRun[]>((runs, letter) => {
          // One to three letters per run, each run with its own slots.
          // An Arabic letter is in a right-to-left run of its own controls.
          const arabic = /\p{Script=Arabic}|\d/u.test(letter);
          if (runs.length === 0 || next() < 0.6 || arabic !== (runs.at(-1)!.rtl === true)) {
            const ascii = pick(GUIDE_FACES);
            runs.push({
              text: letter,
              fonts: { ascii, highAnsi: ascii, eastAsia: pick(GUIDE_FACES), complexScript: pick(GUIDE_FACES) },
              // Each axis styles only its own text (bCs / iCs for complex script).
              bold: next() < 0.3, italic: next() < 0.3, boldCs: next() < 0.3, italicCs: next() < 0.3,
              ...(arabic ? { rtl: true, ...(next() < 0.3 ? { fontHint: 'eastAsia' } : {}) } : {}),
            });
          } else {
            runs[runs.length - 1] = { ...runs.at(-1)!, text: runs.at(-1)!.text + letter };
          }
          return runs;
        }, [])
        : family === 'guideRuns'
          ? [...letters].map((letter, at) => ({ text: letter, fonts: { ascii: 'FaceA', highAnsi: 'FaceA', eastAsia: `Unused${at}` } }))
          : family === 'graphemes' && next() < 0.6
            // Run seams anywhere, between a base and its mark too.
            ? [...letters].reduce<GuideRun[]>((runs, letter) => {
              if (runs.length === 0 || next() < 0.5) {
                runs.push({ text: letter, fonts: { ascii: 'FaceA', highAnsi: 'FaceA', eastAsia: pick(GUIDE_FACES) } });
              } else {
                runs[runs.length - 1] = { ...runs.at(-1)!, text: runs.at(-1)!.text + letter };
              }
              return runs;
            }, [])
            : undefined;
      const ruby: Ruby = {
        text: letters, fontSizePt: 5, hpsRaisePt: next() < 0.5 ? 9 : 4, align: pick(ALIGNS),
        ...(guideRuns ? { guideRuns } : {}),
      };
      const fitPiece = !sharedRegion && !tateChuYoko && !rtlBase && next() < 0.3 ? Math.floor(next() * pieces.length) : -1;
      const runs = pieces.map((value, piece) => text(value, {
        ...(piece === 0 ? {
          ruby: { text: ruby.text, fontSizePt: ruby.fontSizePt, hpsRaisePt: ruby.hpsRaisePt },
          typographyInput: typographyInput(ruby),
        } : { rubyBaseContinuation: true }),
        ...(tateChuYoko && /^\d+$/.test(value) ? { eastAsianVert: true } : {}),
        ...(rtlBase ? { ...(/\p{Script=Hebrew}/u.test(value) ? { rtl: true } : {}), bold: piece % 2 === 1 } : {}),
        ...(piece === fitPiece ? { fitTextVal: 200 + Math.floor(next() * 400), fitTextId: fitId++ } : {}),
        ...(sharedRegion?.fit ?? {}),
      }));
      sharedRegion?.runs.push(...runs);
      cells.push({
        runs, guide: letters, ...(guideRuns ? { guideRuns } : {}),
        ...(family === 'tateChuYoko' ? { painted: [...letters].map((letter) => VERTICAL_FORMS[letter] ?? letter).join('') } : {}),
      });
      leaderAllowed = true;
    }
    const indentPt = family === 'list' ? 20 : 0;
    const paragraph = {
      ...paragraphOf([]),
      ...(family === 'justified' ? { alignment: pick(['both', 'distribute']) } : {}),
      ...(family === 'list' ? {
        numbering: { numId: 1, level: 0, format: 'decimal', text: '1.', indentLeft: indentPt, tab: indentPt, suff: 'tab' },
      } : {}),
      ...(family === 'tabs' ? { tabStops: [{ pos: 30, alignment: 'left', leader: 'none' }] } : {}),
    };
    const lengthPt = (family === 'leaderCell' ? 20 : family === 'fitRegion' || family === 'rtlBase' ? 45 : 25)
      + Math.floor(next() * 60) + indentPt + (family === 'float' ? 40 : 0);
    return { family, sources, cells, regions, direction, lengthPt, indentPt, grid, paragraph };
  }

  type Generated = ReturnType<typeof generate>;
  const sectionOf = ({ direction, lengthPt, grid }: Generated) =>
    ({ ...grid, ...(direction ? { pageHeight: 60 + 80 + lengthPt } : { pageWidth: 40 + 90 + lengthPt }) });
  const layoutRuns = (generated: Generated, runs: readonly unknown[]) =>
    layoutBody([{ ...generated.paragraph, runs } as unknown as DocParagraph & BodyElement],
      generated.direction, sectionOf(generated)).layers.body
      .find((entry): entry is ParagraphLayout => entry.kind === 'paragraph')!.lines;
  /** A line's content in source order (a bidi line places it in visual order). */
  const content = (line: LineLayout) => textPlacements(line)
    .filter((entry) => entry.role !== 'numbering-marker')
    .sort((left, right) => left.range.start - right.range.start);
  const lineText = (line: LineLayout) => content(line).map((entry) => entry.text).join('');

  /** The same runs with an explicit break wherever the layout broke. */
  function explicitlyBroken({ sources }: Generated, lines: readonly LineLayout[]): unknown[] {
    const ends = new Set(lines.slice(0, -1).map((line) => Math.max(...content(line).map((entry) => entry.range.end))));
    const out: unknown[] = [];
    let offset = 0;
    for (const { run, text } of sources) {
      let start = 0;
      for (let at = 1; at <= text.length; at++) {
        if (!ends.has(offset + at)) continue;
        out.push(run.type === 'text' ? { ...run, text: text.slice(start, at) } : run);
        out.push({ type: 'break', breakType: 'line' });
        start = at;
      }
      if (start < text.length) out.push(run.type === 'text' ? { ...run, text: text.slice(start) } : run);
      offset += text.length;
    }
    return out;
  }

  /** The characters measured to lay `generated` out with `runs`. */
  function work(generated: Generated, runs: readonly unknown[]): number {
    const before = measuredCharacters;
    layoutRuns(generated, runs);
    return measuredCharacters - before;
  }

  it.each(Array.from({ length: FAMILIES.length * SEEDS_PER_FAMILY }, (_, seed) => seed + 1))(
    'keeps every ruby cell whole and every line edge legal (seed %i)', (seed) => {
    const generated = generate(seed);
    const { family, sources, cells, regions, lengthPt, indentPt } = generated;
    const runs = sources.map((source) => source.run);
    const lines = layoutRuns(generated, runs);
    const texts = lines.map(lineText);
    const context = `${family} ${generated.direction ?? 'lrTb'} ${lengthPt}pt ${generated.grid.docGridType ?? ''} ${JSON.stringify(texts)}`;
    expect(texts.join(''), context).toBe(sources.filter((source) => source.run.type === 'text').map((source) => source.text).join(''));

    for (const cell of cells) {
      // One line holds every piece of the cell, and its guide is painted once.
      const holding = lines.filter((line) => content(line).some((entry) => cell.runs.includes(entry.sourceRunIndex!)));
      expect(holding, context).toHaveLength(1);
      const guides = content(holding[0]!).filter((entry) => entry.ruby && cell.runs.includes(entry.sourceRunIndex!));
      expect(guides.map((entry) => entry.ruby!.paintOps.map((op) => op.text).join('')), context)
        .toEqual([cell.painted ?? cell.guide]);
      // UAX #29: no paint operation of a guide parts a grapheme cluster (a
      // combining mark, a variation selector, an emoji sequence).
      const guideBoundaries = clusterBoundaries(cell.guide);
      let painted = 0;
      for (const op of guides[0]!.ruby!.paintOps) {
        expect(guideBoundaries.has(painted), `${cell.guide} ${painted} of ${context}`).toBe(true);
        painted += op.text.length;
      }
      if (!cell.guideRuns) continue;
      // §17.3.3.24: every guide letter is painted in the font its own run
      // names for the letter's slot, with the weight and slant of its axis
      // (bCs / iCs for complex script). A cluster follows its first letter.
      const letterFaces = cell.guideRuns.flatMap((run) => [...run.text].map((letter) => ({
        letter,
        font: run.rtl ? run.fonts.complexScript! : /[A-Za-z]/.test(letter) ? run.fonts.ascii! : run.fonts.eastAsia!,
        weight: (run.rtl ? run.boldCs : run.bold) ? 700 : 400,
        style: (run.rtl ? run.italicCs : run.italic) ? 'italic' : 'normal',
        rtl: run.rtl === true,
      })));
      let offset = 0;
      let base = letterFaces[0]!;
      const faceAt = letterFaces.map((face) => {
        if (guideBoundaries.has(offset)) base = face;
        offset += face.letter.length;
        return base;
      });
      // UAX #9: consecutive right-to-left guide runs painted apart stand in
      // reverse on a horizontal line.
      if (!generated.direction) {
        const ops = guides[0]!.ruby!.paintOps;
        let letter = 0;
        ops.forEach((op, index) => {
          const previous = ops[index - 1];
          if (previous && faceAt[letter]!.rtl && faceAt[letter - 1]!.rtl) {
            // §17.3.2.30: digits keep their direction, so two groups of
            // digits stand in reading order; letters stand in reverse.
            const digits = [previous, op].map((entry) => /^\d+$/.test(entry.text));
            if (digits[0] && digits[1]) {
              expect(op.origin.xPt, `${cell.guide} digits ${index} of ${context}`).toBeGreaterThan(previous.origin.xPt);
            } else if (!digits[0] && !digits[1]) {
              expect(op.origin.xPt, `${cell.guide} order ${index} of ${context}`).toBeLessThan(previous.origin.xPt);
            }
          }
          letter += [...op.text].length;
        });
      }
      let at = 0;
      let seam = 0;
      for (const op of guides[0]!.ruby!.paintOps) {
        // Letters of consecutive right-to-left runs that select one face
        // are shaped together: no operation begins between them.
        if (at > 0 && faceAt[at]!.rtl && faceAt[at - 1]!.rtl) {
          const same = (['font', 'weight', 'style'] as const).every((key) => faceAt[at]![key] === faceAt[at - 1]![key]);
          expect(same, `${cell.guide} seam ${seam} of ${context}`).toBe(false);
        }
        for (const letter of op.text) {
          // The stub face has no emoji, which fall back to another face.
          if (!/\p{Extended_Pictographic}|\u200d/u.test(letter)) {
            expect(JSON.stringify(op.fontRoute), `${cell.guide} ${at} of ${context}`).toContain(`\\"${faceAt[at]!.font}\\"`);
            expect([op.fontWeight, op.fontStyle], `${cell.guide} ${at} of ${context}`).toEqual([faceAt[at]!.weight, faceAt[at]!.style]);
          }
          at += 1;
        }
        seam += op.text.length;
      }
    }

    for (const region of regions) {
      // §17.3.2.14: a fitText region stays on one line, with the ruby cell
      // that shares it.
      expect(lines.filter((line) => content(line).some((entry) => region.includes(entry.sourceRunIndex!))), context).toHaveLength(1);
    }
    if (family === 'rtlBase') {
      // A guide no longer than its base stays over the cell, whichever way
      // the base runs and however its runs are split.
      for (const cell of cells) {
        const pieces = lines.flatMap(content).filter((entry) => cell.runs.includes(entry.sourceRunIndex!));
        const startPt = Math.min(...pieces.map((entry) => entry.bounds.xPt));
        const endPt = Math.max(...pieces.map((entry) => entry.bounds.xPt + entry.bounds.widthPt));
        // A guide that widens a cell with right-to-left runs leaves no gap
        // inside the base: each piece's glyphs follow those of the piece
        // before it on the line, in whatever order the directions put them.
        const inks = pieces.map((entry) => ({
          startPt: entry.origin.xPt + (entry.paintOps[0]?.offset?.xPt ?? 0), widthPt: [...entry.text].length * 10,
        })).sort((left, right) => left.startPt - right.startPt);
        if (pieces.some((entry) => entry.direction === 'rtl')) {
          inks.slice(1).forEach((ink, at) =>
            expect(ink.startPt, `${cell.guide} of ${context}`).toBeCloseTo(inks[at]!.startPt + inks[at]!.widthPt, 6));
        }
        if ([...cell.guide].length * 5 > endPt - startPt) continue;
        for (const op of pieces.find((entry) => entry.ruby)!.ruby!.paintOps) {
          expect(op.origin.xPt, `${cell.guide} of ${context}`).toBeGreaterThanOrEqual(startPt - 1e-6);
          expect(op.origin.xPt + 5, `${cell.guide} of ${context}`).toBeLessThanOrEqual(endPt + 1e-6);
        }
      }
    }
    const sizeAt = sources.flatMap((source) => source.run.type === 'text'
      ? [...source.text].map((char, at, all) =>
        (JOINED_RE.test(char) || all[at - 1] === '\u200d' ? 0 : source.sizePt)) : []);
    const characters = [...texts.join('')];
    let consumed = 0;
    const lineStarts = clusterBoundaries(texts.join(''));
    let startOffset = 0;
    texts.forEach((text, index) => {
      const chars = [...text];
      consumed += chars.length;
      // §17.3.1.16: no line starts with a line-start-forbidden character.
      if (index > 0) expect(LINE_START_FORBIDDEN.has(chars[0]!), `${index} of ${context}`).toBe(false);
      // A line that ends with a flag whose halves are separate runs holds it
      // within the line: the line broke before the flag rather than overrun.
      if (text.endsWith(FLAG.join('')) && chars.length > 2 && !LINE_END_FORBIDDEN.has(chars.at(-3)!)
        && !cells.some((cell) => content(lines[index]!).some((entry) => cell.runs.includes(entry.sourceRunIndex!)))) {
        const advance = content(lines[index]!).reduce((sum, entry) => sum + entry.advancePt, 0);
        expect(advance, `${index} of ${context}`).toBeLessThanOrEqual(lengthPt - indentPt + 1e-6);
      }
      // No line starts inside a grapheme cluster.
      if (index > 0) expect(lineStarts.has(startOffset), `${index} of ${context}`).toBe(true);
      startOffset += text.length;
      if (index === texts.length - 1 || !LINE_END_FORBIDDEN.has(chars.at(-1)!) || family === 'float') return;
      // A line ends with openers only when they and the unit after them
      // exceed a whole line. The openers are every trailing opener, with
      // the whole fitText region one closes; the unit is a whole ruby cell,
      // fitText region or Hebrew word, or one character with the variation
      // selector and the forbidden leaders glued to it.
      let openersPt = 0;
      const placed = content(lines[index]!);
      for (let at = placed.length - 1; at >= 0; at -= 1) {
        const entry = placed[at]!;
        const letters = [...entry.text];
        let trailing = 0;
        while (trailing < letters.length && LINE_END_FORBIDDEN.has(letters[letters.length - 1 - trailing]!)) trailing += 1;
        if (trailing === 0) break;
        const region = regions.find((candidate) => candidate.includes(entry.sourceRunIndex!));
        if (region) {
          // Every placed piece of the region.
          for (; at >= 0 && region.includes(placed[at]!.sourceRunIndex!); at -= 1) openersPt += placed[at]!.advancePt;
          at += 1;
          continue;
        }
        openersPt += entry.advancePt * trailing / letters.length;
        if (trailing < letters.length) break;
      }
      const following = content(lines[index + 1]!);
      const cell = cells.find((candidate) => candidate.runs[0] === following[0]?.sourceRunIndex);
      const unit = regions.find((region) => region.includes(following[0]?.sourceRunIndex ?? -1)) ?? cell?.runs ?? [];
      let glued = 1;
      for (let at = consumed + 1; sizeAt[at] === 0 || LINE_START_FORBIDDEN.has(characters[at]!); at += 1) {
        if (sizeAt[at] !== 0) glued += 1;
      }
      const unitPt = unit.length > 0 ? following.filter((entry) => unit.includes(entry.sourceRunIndex!))
        .reduce((sum, entry) => sum + entry.advancePt, 0)
        : sizeAt[consumed]! * Math.max(glued, /^\p{Script=Hebrew}*/u.exec(characters.slice(consumed).join(''))![0].length);
      expect(unitPt + openersPt > lengthPt - indentPt, `${index} of ${context}`).toBe(true);
    });

    if (family === 'guideRuns') {
      // Shaping a guide is linear in its runs: one run per letter costs a
      // bounded multiple of the same letters in one run per guide.
      const merged = sources.map((source) => {
        const input = source.run.typographyInput as ReturnType<typeof typographyInput> | undefined;
        return input ? {
          ...source.run,
          typographyInput: { ruby: { ...input.ruby, guideRuns: [{ ...input.ruby.guideRuns[0]!, text: (source.run.ruby as Ruby).text }] } },
        } : source.run;
      });
      expect(work(generated, runs), context).toBeLessThan(work(generated, merged) * 3);
    }

    // A float's exclusion depends on where each line falls, so only the
    // invariants above hold there.
    if (family === 'float') return;
    // Each line's extents are those of its own content: breaking the same
    // runs explicitly at the same places gives the same lines.
    const explicit = layoutRuns(generated, explicitlyBroken(generated, lines));
    expect(explicit.map(lineText), context).toEqual(texts);
    explicit.forEach((line, index) => {
      const laid = lines[index]!;
      expect(laid.advancePt, `${index} of ${context}`).toBeCloseTo(line.advancePt, 6);
      expect(laid.baselinePt - laid.bounds.yPt, `${index} of ${context}`).toBeCloseTo(line.baselinePt - line.bounds.yPt, 6);
      // On a character grid every piece takes the cells it takes when the
      // line is broken there explicitly.
      if (family === 'scaledLatin') {
        expect(content(laid).map((entry) => entry.advancePt), `${index} of ${context}`)
          .toEqual(content(line).map((entry) => entry.advancePt));
      }
    });
  });
});
