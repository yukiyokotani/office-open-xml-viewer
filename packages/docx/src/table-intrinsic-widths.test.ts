import { beforeAll, describe, expect, it } from 'vitest';
import { DEFAULT_KINSOKU_RULES } from '@silurus/ooxml-core';
import { layoutDocument } from './document-layout.js';
import { createLayoutServices } from './layout-runtime.js';
import { resolveColumnWidths } from './test-support/renderer-internals.test-support.js';
import {
  measureParagraphIntrinsicWidths,
  measureTableCellIntrinsicWidths,
} from './layout/intrinsic-width.js';
import { bodyAcquisitionInputProjections } from './parser-model.js';
import {
  resolveDocumentLayoutSettings,
  resolveSectionLayoutContext,
  type ParagraphLayoutContext,
} from './layout-context.js';
import type { TextLayoutService } from './layout/text.js';
import type {
  BodyElement,
  CellElement,
  DocParagraph,
  DocTable,
  DocTableCell,
  DocTableRow,
  DocxDocumentModel,
  SectionProps,
} from './types.js';

type ColumnState = Parameters<typeof resolveColumnWidths>[2];

function measuringContext(
  widthOf: (text: string) => number = (text) => [...text].length * 5,
): CanvasRenderingContext2D {
  let font = '10px serif';
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    fontKerning: 'auto',
    measureText(text: string) {
      const width = widthOf(text);
      return {
        width,
        actualBoundingBoxLeft: 0,
        actualBoundingBoxRight: /^[、。]$/u.test(text) ? width / 2 : width,
        fontBoundingBoxAscent: 8,
        fontBoundingBoxDescent: 2,
        actualBoundingBoxAscent: 8,
        actualBoundingBoxDescent: 2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {}, moveTo() {}, lineTo() {},
    stroke() {}, fill() {}, fillRect() {}, strokeRect() {}, clip() {}, rect() {},
    scale() {}, translate() {}, rotate() {}, setLineDash() {}, clearRect() {}, arc() {},
    quadraticCurveTo() {}, bezierCurveTo() {}, drawImage() {}, fillText() {}, strokeText() {},
    createLinearGradient() { return { addColorStop() {} }; },
    fillStyle: '#000000', strokeStyle: '#000000', lineWidth: 1,
    textAlign: 'left', direction: 'ltr',
  } as unknown as CanvasRenderingContext2D;
}

beforeAll(() => {
  (globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = class {
    private readonly context = measuringContext();
    getContext() { return this.context; }
  };
});

const borders = {
  top: null, right: null, bottom: null, left: null, insideH: null, insideV: null,
};

function textRun(text: string): DocParagraph['runs'][number] {
  return {
    type: 'text', text, fontSize: 10, fontFamily: 'serif',
    bold: false, italic: false, underline: false, strikethrough: false,
    color: null, isLink: false, background: null, vertAlign: null, hyperlink: null,
  } as DocParagraph['runs'][number];
}

function paragraph(
  runs: DocParagraph['runs'],
  overrides: Partial<DocParagraph> = {},
): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    widowControl: false, runs, defaultFontSize: 10, defaultFontFamily: 'serif',
    ...overrides,
  } as unknown as DocParagraph;
}

function cell(content: CellElement[]): DocTableCell {
  return {
    content, colSpan: 1, vMerge: null, borders, background: null, vAlign: 'top',
    widthPt: null, widthPct: null,
    marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
  } as unknown as DocTableCell;
}

function row(cells: DocTableCell[]): DocTableRow {
  return { cells, rowHeight: null, rowHeightRule: 'auto', isHeader: false } as DocTableRow;
}

function table(
  rows: DocTableRow[],
  colWidths: number[],
  layout?: 'fixed' | 'autofit',
): DocTable {
  return {
    type: 'table', rows, colWidths, borders,
    cellMarginTop: 0, cellMarginRight: 0, cellMarginBottom: 0, cellMarginLeft: 0,
    jc: 'left', ...(layout ? { layout } : {}),
  } as unknown as DocTable;
}

function model(body: BodyElement[]): DocxDocumentModel {
  const section = {
    pageWidth: 220, pageHeight: 300,
    marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
    headerDistance: 4, footerDistance: 4, titlePage: false, evenAndOddHeaders: false,
    sectionStart: 'nextPage', columns: null,
  } as SectionProps;
  return {
    section, body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: {}, footnotes: [],
  } as unknown as DocxDocumentModel;
}

function columnState(
  ctx: CanvasRenderingContext2D,
  services = createLayoutServices(model([]), { measureContext: ctx }),
): ColumnState {
  const document = model([]);
  const layoutSettings = resolveDocumentLayoutSettings(document);
  return {
    ctx, fontFamilyClasses: {}, layoutServices: services,
    pageWidth: 200, pageH: 300, pageIndex: 0, totalPages: 1,
    defaultTabPt: 36,
    acquisitionInputs: bodyAcquisitionInputProjections,
    layoutSettings,
    sectionLayout: resolveSectionLayoutContext(layoutSettings, document.section),
  } as unknown as ColumnState;
}

describe('table intrinsic content widths', () => {
  const intrinsicContext = (overrides: Partial<ParagraphLayoutContext> = {}): ParagraphLayoutContext => ({
    lineGrid: { active: false, pitchPt: null },
    characterGrid: { active: false, kind: null, pitchPt: null, deltaPt: 0 },
    rightIndentGrid: { pitchPt: null, paragraphAllowsAdjustment: true },
    physicalIndentLeftPt: 0,
    physicalIndentRightPt: 0,
    firstIndentPt: 0,
    lineSpacing: null,
    spaceBeforePt: 0,
    spaceAfterPt: 0,
    baseRtl: false,
    isJustified: false,
    stretchLastLine: false,
    tabStops: [],
    hasRuby: false,
    hasEastAsianText: true,
    kinsoku: DEFAULT_KINSOKU_RULES,
    defaultTabPt: 36,
    ...overrides,
  });

  it('ignores paragraph intrinsic width for an empty unnumbered cell paragraph', () => {
    const source = paragraph([]);
    const measured: DocParagraph[] = [];
    expect(measureTableCellIntrinsicWidths(
      cell([source as CellElement]),
      { left: 5.4, right: 5.4 },
      {
        paragraph: (value) => {
          measured.push(value as DocParagraph);
          return { minWidthPt: 10.8, maxWidthPt: 10.8 };
        },
        nestedTable: () => ({ minWidthPt: 0, maxWidthPt: 0 }),
      },
    )).toEqual({ minWidthPt: 10.8, maxWidthPt: 10.8 });
    expect(measured).toEqual([]);
  });

  it('retains paragraph intrinsic width when the cell has a visible run', () => {
    const source = paragraph([textRun('X')]);
    expect(measureTableCellIntrinsicWidths(
      cell([source as CellElement]),
      { left: 5.4, right: 5.4 },
      {
        paragraph: () => ({ minWidthPt: 15.8, maxWidthPt: 15.8 }),
        nestedTable: () => ({ minWidthPt: 0, maxWidthPt: 0 }),
      },
    )).toEqual({ minWidthPt: 26.6, maxWidthPt: 26.6 });
  });

  it('retains numbering-marker intrinsic width on an otherwise empty cell paragraph', () => {
    const source = paragraph([], {
      numbering: { numId: 1, level: 0 } as DocParagraph['numbering'],
    });
    const measured: DocParagraph[] = [];
    expect(measureTableCellIntrinsicWidths(
      cell([source as CellElement]),
      { left: 5.4, right: 5.4 },
      {
        paragraph: (value) => {
          measured.push(value as DocParagraph);
          return { minWidthPt: 15.8, maxWidthPt: 15.8 };
        },
        nestedTable: () => ({ minWidthPt: 0, maxWidthPt: 0 }),
      },
    )).toEqual({ minWidthPt: 26.6, maxWidthPt: 26.6 });
    expect(measured).toEqual([source]);
  });

  it('keeps empty paragraph indents out of the AutoFit column solver', () => {
    const withMargins = (source: DocParagraph): DocTableCell => ({
      ...cell([source as CellElement]),
      marginLeft: 5.4,
      marginRight: 5.4,
    });
    const source = table([row([
      withMargins(paragraph([], { indentRight: 10.8 })),
      withMargins(paragraph([], { indentLeft: 10.8 })),
      withMargins(paragraph([], { indentFirst: 10.8 })),
      withMargins(paragraph([], { indentFirst: -10.8, bidi: true })),
      withMargins(paragraph([textRun('X')], { indentRight: 10.8 })),
    ])], [12.8, 12.8, 12.8, 12.8, 12.8]);

    expect(resolveColumnWidths(source, 200, columnState(measuringContext())))
      .toEqual([10.8, 10.8, 10.8, 10.8, 26.6]);
  });

  it('uses visible sibling paragraphs without charging an empty paragraph indent', () => {
    const source = table([row([{
      ...cell([
        paragraph([], { indentLeft: 36, indentRight: 36 }) as CellElement,
        paragraph([textRun('XX')]) as CellElement,
      ]),
      marginLeft: 5.4,
      marginRight: 5.4,
    }])], [12.8]);

    expect(resolveColumnWidths(source, 200, columnState(measuringContext())))
      .toEqual([20.8]);
  });

  it('retains whitespace and non-breaking-space controls in AutoFit width', () => {
    const withMargins = (text: string): DocTableCell => ({
      ...cell([paragraph([textRun(text)]) as CellElement]),
      marginLeft: 5.4,
      marginRight: 5.4,
    });
    const source = table([row([
      withMargins(' '),
      withMargins('\u00a0'),
    ])], [0, 0]);

    expect(resolveColumnWidths(source, 200, columnState(measuringContext())))
      .toEqual([15.8, 15.8]);
  });

  it('includes the character-grid pitch in minimum content width', () => {
    const source = paragraph([textRun('漢字')]);

    expect(measureParagraphIntrinsicWidths(
      source,
      intrinsicContext({ characterGrid: { active: true, kind: 'linesAndChars', pitchPt: 12, deltaPt: 2 } }),
      200,
      { context: measuringContext(), fontFamilyClasses: {} },
      {
        pageIndex: 0, totalPages: 1, pageWritingMode: 'horizontal-tb',
        documentHasEastAsianText: true,
      },
    )).toEqual({ minWidthPt: 7, maxWidthPt: 14 });
  });

  it('does not merge runs with different character-grid participation', () => {
    const source = paragraph([
      { ...textRun('漢'), snapToGrid: true },
      { ...textRun('字'), snapToGrid: false },
    ] as DocParagraph['runs']);

    expect(measureParagraphIntrinsicWidths(
      source,
      intrinsicContext({ characterGrid: { active: true, kind: 'linesAndChars', pitchPt: 12, deltaPt: 2 } }),
      200,
      { context: measuringContext(), fontFamilyClasses: {} },
      {
        pageIndex: 0, totalPages: 1, pageWritingMode: 'horizontal-tb',
        documentHasEastAsianText: true,
      },
    )).toEqual({ minWidthPt: 7, maxWidthPt: 12 });
  });

  it('uses the balanced space cells and SBCS grid delta for intrinsic width', () => {
    const source = paragraph([textRun('AB  ')]);
    const context = measuringContext((text) => [...text].reduce(
      (sum, character) => sum + (character === ' ' ? 3 : character === '一' ? 10 : 5),
      0,
    ));

    expect(measureParagraphIntrinsicWidths(
      source,
      intrinsicContext({
        characterGrid: {
          active: true,
          kind: 'linesAndChars',
          pitchPt: 3,
          deltaPt: -2,
        },
      }),
      200,
      { context, fontFamilyClasses: {} },
      {
        pageIndex: 0,
        totalPages: 1,
        pageWritingMode: 'horizontal-tb',
        documentHasEastAsianText: true,
        balanceSingleByteDoubleByteWidth: true,
        layoutServices: createLayoutServices(model([]), { measureContext: context }),
      },
    )).toEqual({
      // The breakable trailing-space sequence is excluded from the minimum;
      // the unbroken maximum retains both balanced half-width cells.
      minWidthPt: 8,
      maxWidthPt: 16,
    });
  });

  it('uses per-grapheme whole-cell allocation for snapToChars intrinsic widths', () => {
    const source = paragraph([textRun('漢字')]);

    expect(measureParagraphIntrinsicWidths(
      source,
      intrinsicContext({
        characterGrid: {
          active: true,
          kind: 'snapToChars',
          pitchPt: 4,
          deltaPt: -6,
        },
      }),
      200,
      { context: measuringContext(), fontFamilyClasses: {} },
      {
        pageIndex: 0, totalPages: 1, pageWritingMode: 'horizontal-tb',
        documentHasEastAsianText: true,
      },
    )).toEqual({ minWidthPt: 8, maxWidthPt: 16 });
  });

  it('keeps Latin and East-Asian snapToChars allocation blocks distinct', () => {
    const source = paragraph([textRun('A漢')]);

    expect(measureParagraphIntrinsicWidths(
      source,
      intrinsicContext({
        characterGrid: {
          active: true,
          kind: 'snapToChars',
          pitchPt: 4,
          deltaPt: -6,
        },
      }),
      200,
      { context: measuringContext(), fontFamilyClasses: {} },
      {
        pageIndex: 0, totalPages: 1, pageWritingMode: 'horizontal-tb',
        documentHasEastAsianText: true,
      },
    )).toEqual({ minWidthPt: 8, maxWidthPt: 16 });
  });

  it('keeps adjacent tate-chu-yoko runs as separate one-em cells', () => {
    const source = paragraph([
      { ...textRun('12'), eastAsianVert: true },
      { ...textRun('34'), eastAsianVert: true },
    ] as DocParagraph['runs']);

    expect(measureParagraphIntrinsicWidths(
      source,
      intrinsicContext(),
      200,
      { context: measuringContext(), fontFamilyClasses: {} },
      {
        pageIndex: 0, totalPages: 1, pageWritingMode: 'vertical-rl',
        documentHasEastAsianText: true, verticalCJK: true,
      },
    )).toEqual({ minWidthPt: 20, maxWidthPt: 20 });
  });

  it('keeps an inline image as a minimum-content atom', () => {
    const source = table([row([cell([paragraph([{
      type: 'image', imagePath: 'word/media/image.png', mimeType: 'image/png',
      widthPt: 80, heightPt: 10, anchor: false,
    }]) as CellElement])])], [0]);

    expect(resolveColumnWidths(source, 200, columnState(measuringContext()))).toEqual([80]);
  });

  it('uses structural math metadata for minimum and maximum content width', () => {
    const resourceKey = 'math:body:0.0.0:inline';
    const run = {
      type: 'math', display: false, fontSize: 10, resourceKey,
      nodes: [{ kind: 'run', text: 'x', style: 'italic' }],
    } as unknown as DocParagraph['runs'][number];
    const source = table([row([cell([paragraph([run]) as CellElement])])], [0]);
    const ctx = measuringContext();
    const base = createLayoutServices(model([]), { measureContext: ctx });
    const services = Object.freeze({
      ...base,
      math: Object.freeze({
        fingerprint: 'table-intrinsic-math',
        resolve: (key: string) => ({
          resourceKey: key, widthEm: 5, ascentEm: 0.8, descentEm: 0.2, diagnostics: [],
        }),
      }),
    });

    expect(resolveColumnWidths(source, 200, columnState(ctx, services))).toEqual([50]);
  });

  it('retains a left-tab leader in minimum content width', () => {
    const source = table([row([cell([paragraph([textRun('\tX')], {
      tabStops: [{ pos: 60, alignment: 'left', leader: 'dot' }],
    }) as CellElement])])], [0]);

    expect(resolveColumnWidths(source, 200, columnState(measuringContext()))).toEqual([60]);
  });

  it('uses following content when resolving a right-tab maximum', () => {
    const tabbed = cell([paragraph([textRun('\tX')], {
      tabStops: [{ pos: 60, alignment: 'right', leader: 'none' }],
    }) as CellElement]);
    const source = table([row([tabbed, cell([])])], [0, 100]);

    expect(resolveColumnWidths(source, 100, columnState(measuringContext()))).toEqual([60, 40]);
  });

  it('includes retained numbering-marker geometry and paragraph indents', () => {
    const numbered = paragraph([textRun('X')], {
      indentLeft: 18,
      indentFirst: -9,
      numbering: {
        numId: 1, level: 0, format: 'decimal', text: '12345.',
        indentLeft: 18, tab: 18, suff: 'tab', jc: 'left',
      } as NonNullable<DocParagraph['numbering']>,
    });
    const source = table([row([cell([numbered as CellElement])])], [0]);

    // The six-glyph marker overruns the 9pt hanging area, so the suffix tab
    // advances the body to the next 36pt default stop: 18 + 54 + 5 = 77pt.
    expect(resolveColumnWidths(source, 200, columnState(measuringContext()))).toEqual([77]);
  });

  it('uses a fixed nested table as an outer minimum-content atom', () => {
    const nested = table([row([cell([])])], [80], 'fixed');
    const source = table([row([cell([nested as CellElement])])], [0]);

    expect(resolveColumnWidths(source, 200, columnState(measuringContext()))).toEqual([80]);
  });

  it('does not shrink a fixed nested table to its containing cell width', () => {
    const nested = table([
      row([cell([]), cell([])]),
    ], [60, 60], 'fixed');
    const outer = table([
      row([cell([nested as CellElement])]),
    ], [100], 'fixed');
    const document = model([outer as BodyElement]);
    const layout = layoutDocument(
      document,
      createLayoutServices(document, { measureContext: measuringContext() }),
      { currentDateMs: 0 },
    );
    const retainedOuter = layout.pages[0]?.layers.body.find((node) => node.kind === 'table');
    if (!retainedOuter || retainedOuter.kind !== 'table') {
      throw new Error('Expected outer retained table');
    }
    const retainedNested = retainedOuter.rows[0]?.cells[0]?.blocks[0]?.layout;
    if (!retainedNested || retainedNested.kind !== 'table') {
      throw new Error('Expected nested retained table');
    }

    // ECMA-376 §17.18.87 fixed layout resolves tblGrid/tcW against an authored
    // tblW; the containing cell is not an implicit preferred table width.
    // Word therefore permits a fixed nested table to overflow its cell.
    expect(retainedNested.columnWidthsPt).toEqual([60, 60]);
    expect(retainedNested.flowBounds.widthPt).toBe(120);
  });

  it('shapes identical formatting across a run seam as one proportional atom', () => {
    const ctx = measuringContext((text) => text === 'AV' ? 15 : [...text].length * 10);
    const source = table([row([cell([
      paragraph([textRun('A'), textRun('V')]) as CellElement,
    ])])], [0]);

    expect(resolveColumnWidths(source, 200, columnState(ctx))).toEqual([15]);
  });

  it('retains every rebased punctuation compression across compatible run seams', () => {
    const ctx = measuringContext();
    const document = model([]);
    const services = createLayoutServices(document, { measureContext: ctx });
    const source = paragraph([textRun('甲、'), textRun('乙。')]);

    const intrinsic = measureParagraphIntrinsicWidths(
      source,
      intrinsicContext(),
      200,
      { context: ctx, fontFamilyClasses: {} },
      {
        pageIndex: 0,
        totalPages: 1,
        pageWritingMode: 'horizontal-tb',
        documentHasEastAsianText: true,
        characterSpacingControl: 'compressPunctuation',
        layoutServices: services,
      },
    );

    expect(intrinsic).toEqual({
      minWidthPt: 7.5,
      maxWidthPt: 15,
    });
  });

  it('does not acquire content widths for fixed-layout columns', () => {
    const measured: string[] = [];
    const ctx = measuringContext((text) => {
      measured.push(text);
      return [...text].length * 5;
    });
    const source = table([row([cell([
      paragraph([textRun('fixed-only')]) as CellElement,
    ])])], [80], 'fixed');

    expect(resolveColumnWidths(source, 200, columnState(ctx))).toEqual([80]);
    expect(measured).toEqual([]);
  });
});

describe('table retained marker acquisition', () => {
  it('acquires numbering marker glyph geometry once', () => {
    const marker = '12345.';
    const numbered = paragraph([textRun('X')], {
      indentLeft: 18,
      indentFirst: -9,
      numbering: {
        numId: 1, level: 0, format: 'decimal', text: marker,
        indentLeft: 18, tab: 18, suff: 'tab', jc: 'left',
      } as NonNullable<DocParagraph['numbering']>,
    });
    const source = table([row([cell([numbered as CellElement])])], [100], 'fixed');
    const document = model([source as BodyElement]);
    const base = createLayoutServices(document, { measureContext: measuringContext() });
    let markerShapes = 0;
    const text: TextLayoutService = Object.freeze({
      fingerprint: base.text.fingerprint,
      localMetrics: base.text.localMetrics,
      resolve: (request: Parameters<TextLayoutService['resolve']>[0]) => base.text.resolve(request),
      shape: (request: Parameters<TextLayoutService['shape']>[0]) => {
        if (request.text === marker && request.clusterGeometry !== false) markerShapes += 1;
        return base.text.shape(request);
      },
    });

    layoutDocument(document, Object.freeze({ ...base, text }));

    expect(markerShapes).toBe(1);
  });
});

describe('page-owned story table width', () => {
  it.each([
    ['header', 'headers'],
    ['footer', 'footers'],
  ] as const)('keeps a negatively indented fixed %s table at its authored page-relative width', (
    story,
    collection,
  ) => {
    const storyTable = {
      ...table([
        row([
          cell([paragraph([textRun('L')]) as CellElement]),
          cell([paragraph([textRun('R')], { alignment: 'center' }) as CellElement]),
        ]),
      ], [150, 60], 'fixed'),
      tblInd: -10,
      widthPt: 210,
    } as DocTable;
    const base = model([paragraph([textRun('body')]) as BodyElement]);
    const document = {
      ...base,
      [collection]: {
        default: { body: [storyTable as BodyElement] },
        first: null,
        even: null,
      },
    } as DocxDocumentModel;

    const layout = layoutDocument(
      document,
      createLayoutServices(document, { measureContext: measuringContext() }),
      { currentDateMs: 0 },
    );
    const retained = layout.pages[0]?.layers[story].find((node) => node.kind === 'table');

    expect(retained).toBeDefined();
    if (retained?.kind !== 'table') return;
    expect(retained.flowBounds).toMatchObject({ xPt: 0, widthPt: 210 });
    expect(retained.columnWidthsPt).toEqual([150, 60]);
    expect(retained.rows[0]?.cells[1]?.flowBounds).toMatchObject({ xPt: 150, widthPt: 60 });
  });

  it('keeps a bidi fixed header table right-aligned while its negative leading indent reaches the page edge', () => {
    const storyTable = {
      ...table([
        row([
          cell([paragraph([]) as CellElement]),
          cell([paragraph([textRun('R')], { alignment: 'center' }) as CellElement]),
        ]),
      ], [150, 60], 'fixed'),
      borders: {
        ...borders,
        bottom: { width: 1, color: '#000000', style: 'single' },
      },
      tblInd: -10,
      widthPt: 210,
      bidiVisual: true,
    } as DocTable;
    const base = model([paragraph([textRun('body')]) as BodyElement]);
    const document = {
      ...base,
      headers: {
        default: { body: [storyTable as BodyElement] },
        first: null,
        even: null,
      },
    } as DocxDocumentModel;

    const layout = layoutDocument(
      document,
      createLayoutServices(document, { measureContext: measuringContext() }),
      { currentDateMs: 0 },
    );
    const retained = layout.pages[0]?.layers.header.find((node) => node.kind === 'table');

    expect(retained).toBeDefined();
    if (retained?.kind !== 'table') return;
    expect(retained.flowBounds).toMatchObject({ xPt: 10, widthPt: 210 });
    expect(retained.columnWidthsPt).toEqual([150, 60]);
    expect(retained.borders.filter((border) => border.edge === 'bottom')).toEqual([
      expect.objectContaining({
        from: { xPt: 10, yPt: expect.any(Number) },
        to: { xPt: 220, yPt: expect.any(Number) },
      }),
    ]);
  });
});
