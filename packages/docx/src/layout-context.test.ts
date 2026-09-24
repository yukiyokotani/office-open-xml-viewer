import { describe, expect, it } from 'vitest';
import type {
  DocParagraph,
  DocxDocumentModel,
  DocxTextRun,
  SectionProps,
} from './types.js';
import {
  enterTableCellStoryContext,
  resolveDocumentLayoutSettings,
  paragraphGridRightAdjustmentPt,
  resolveParagraphLayoutContext,
  resolveRunLayoutContext,
  resolveSectionLayoutContext,
  type StoryContext,
} from './layout-context.js';

const section = (overrides: Partial<SectionProps> = {}): SectionProps => ({
  pageWidth: 200,
  pageHeight: 300,
  marginTop: 20,
  marginRight: 20,
  marginBottom: 20,
  marginLeft: 20,
  headerDistance: 10,
  footerDistance: 10,
  titlePage: false,
  evenAndOddHeaders: false,
  ...overrides,
});

const paragraph = (overrides: Partial<DocParagraph> = {}): DocParagraph => ({
  alignment: 'left',
  indentLeft: 12,
  indentRight: 6,
  indentFirst: 2,
  spaceBefore: 3,
  spaceAfter: 4,
  lineSpacing: null,
  numbering: null,
  tabStops: [],
  runs: [],
  ...overrides,
});

const textRun = (overrides: Partial<DocxTextRun> = {}): DocxTextRun => ({
  text: 'x',
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  fontSize: 10,
  color: null,
  fontFamily: null,
  isLink: false,
  background: null,
  vertAlign: null,
  hyperlink: null,
  ...overrides,
});

const documentModel = (
  overrides: Partial<DocxDocumentModel> = {},
): DocxDocumentModel => ({
  section: section(),
  body: [],
  headers: { default: null, first: null, even: null },
  footers: { default: null, first: null, even: null },
  ...overrides,
});

const bodyStory: StoryContext = {
  story: 'body',
  containers: [],
  lineNumberingEligible: true,
};

const cellStory: StoryContext = {
  story: 'body',
  containers: [{ kind: 'tableCell' }],
  lineNumberingEligible: false,
};

describe('layout context resolvers', () => {
  it('retains the section column separator fact', () => {
    const settings = resolveDocumentLayoutSettings(documentModel());

    expect(resolveSectionLayoutContext(settings, section({
      columns: { count: 2, spacePt: 18, equalWidth: true, sep: true, cols: [] },
    })).columnSeparator).toBe(true);
    expect(resolveSectionLayoutContext(settings, section({
      columns: { count: 2, spacePt: 18, equalWidth: true, sep: false, cols: [] },
    })).columnSeparator).toBe(false);
  });

  it('preserves nested table-cell container depth', () => {
    const outerCell = enterTableCellStoryContext(bodyStory);
    const innerCell = enterTableCellStoryContext(outerCell);

    expect(outerCell.containers).toEqual([{ kind: 'tableCell' }]);
    expect(innerCell.containers).toEqual([
      { kind: 'tableCell' },
      { kind: 'tableCell' },
    ]);
    expect(innerCell.lineNumberingEligible).toBe(false);
  });

  it('normalizes document settings and detects East Asian body text once', () => {
    const run = textRun({ text: 'あ' });
    const settings = resolveDocumentLayoutSettings(documentModel({
      body: [{ type: 'paragraph', ...paragraph({ runs: [{ type: 'text', ...run }] }) }],
      settings: {
        kinsoku: false,
        defaultTabStop: 18,
        characterSpacingControl: 'compressPunctuation',
        mathDefJc: 'left',
        adjustLineHeightInTable: true,
        useFeLayout: true,
        balanceSingleByteDoubleByteWidth: true,
        lineWrapLikeWord6: true,
      },
    }));

    expect(settings.defaultTabPt).toBe(18);
    expect(settings.kinsoku.enabled).toBe(false);
    expect(settings.documentHasEastAsianText).toBe(true);
    expect(settings.characterSpacingControl).toBe('compressPunctuation');
    expect(settings.mathDefJc).toBe('left');
    expect(settings.compat).toEqual({
      adjustLineHeightInTable: true,
      useFeLayout: true,
      balanceSingleByteDoubleByteWidth: true,
      lineWrapLikeWord6: true,
    });
  });

  it('resolves absent settings and default docGrid without activating a grid', () => {
    const settings = resolveDocumentLayoutSettings(documentModel());
    const context = resolveSectionLayoutContext(
      settings,
      section({ docGridType: 'default', docGridLinePitch: 20 }),
    );

    expect(settings.defaultTabPt).toBe(36);
    expect(settings.compat).toEqual({
      adjustLineHeightInTable: false,
      useFeLayout: false,
      balanceSingleByteDoubleByteWidth: false,
      lineWrapLikeWord6: false,
    });
    expect(context.grid).toEqual({
      kind: 'none',
      linePitchPt: 20,
      charSpacePt: null,
    });
  });

  it('normalizes section geometry, columns, and snapToChars grid units', () => {
    const settings = resolveDocumentLayoutSettings(documentModel());
    const context = resolveSectionLayoutContext(settings, section({
      docGridType: 'snapToChars',
      docGridLinePitch: 20,
      docGridCharSpace: 4096,
      columns: {
        count: 2,
        spacePt: 10,
        equalWidth: true,
        sep: false,
        cols: [],
      },
    }));

    expect(context.geometry).toEqual({
      pageWidth: 200,
      pageHeight: 300,
      marginTop: 20,
      marginRight: 20,
      marginBottom: 20,
      marginLeft: 20,
      headerDistance: 10,
      footerDistance: 10,
    });
    expect(context.columns).toEqual([
      { xPt: 20, wPt: 75 },
      { xPt: 105, wPt: 75 },
    ]);
    expect(context.grid).toEqual({
      kind: 'snapToChars',
      linePitchPt: 20,
      charSpacePt: 1,
    });
  });

  it('keeps invalid line pitches inactive', () => {
    const settings = resolveDocumentLayoutSettings(documentModel());
    const sectionContext = resolveSectionLayoutContext(
      settings,
      section({ docGridType: 'lines', docGridLinePitch: Number.NaN }),
    );
    const paragraphContext = resolveParagraphLayoutContext(
      settings,
      sectionContext,
      bodyStory,
      paragraph(),
    );

    expect(paragraphContext.lineGrid).toEqual({ active: false, pitchPt: null });
  });

  it('resolves physical bidi indents without changing authored spacing', () => {
    const settings = resolveDocumentLayoutSettings(documentModel());
    const sectionContext = resolveSectionLayoutContext(
      settings,
      section({ docGridType: 'lines', docGridLinePitch: 20 }),
    );
    const context = resolveParagraphLayoutContext(
      settings,
      sectionContext,
      bodyStory,
      paragraph({ bidi: true }),
    );

    expect(context.physicalIndentLeftPt).toBe(6);
    expect(context.physicalIndentRightPt).toBe(12);
    expect(context.firstIndentPt).toBe(2);
    expect(context.spaceBeforePt).toBe(3);
    expect(context.spaceAfterPt).toBe(4);
    expect(context.baseRtl).toBe(true);
    expect(context.lineGrid).toEqual({ active: true, pitchPt: 20 });
  });

  it('measures an RTL suff=tab numbered first line at the indentLeft tab stop (§17.9.28)', () => {
    const settings = resolveDocumentLayoutSettings(documentModel());
    const sectionContext = resolveSectionLayoutContext(settings, section());
    const num = (overrides: Record<string, unknown> = {}) => ({
      numId: 1, level: 0, format: 'bullet', text: '•',
      indentLeft: 12, tab: 9, suff: 'tab', ...overrides,
    }) as unknown as DocParagraph['numbering'];

    // §17.9.28 + §17.3.1.6: a suff=tab marker's RTL first-line body advances to the
    // indentLeft tab stop, so the measured first-line indent is 0 — mirroring the
    // paint pass — NOT the raw −hanging (which would widen the first line and let a
    // split paragraph disagree with paint on line count).
    const rtlTab = resolveParagraphLayoutContext(settings, sectionContext, bodyStory,
      paragraph({ bidi: true, indentFirst: -9, numbering: num() }));
    expect(rtlTab.firstIndentPt).toBe(0);

    // LTR is byte-identical: keeps raw indentFirst (its pre-existing measure/paint
    // marker approximation is untouched by this RTL fix).
    const ltrTab = resolveParagraphLayoutContext(settings, sectionContext, bodyStory,
      paragraph({ bidi: false, indentFirst: -9, numbering: num() }));
    expect(ltrTab.firstIndentPt).toBe(-9);

    // suff=space/nothing under RTL is out of scope → keeps legacy raw indentFirst.
    const rtlNothing = resolveParagraphLayoutContext(settings, sectionContext, bodyStory,
      paragraph({ bidi: true, indentFirst: -9, numbering: num({ suff: 'nothing' }) }));
    expect(rtlNothing.firstIndentPt).toBe(-9);

    // An RTL numbering level with no marker glyph (empty text, no picture bullet) is
    // not a marker — the hanging indent applies to the body as usual.
    const rtlNoMarker = resolveParagraphLayoutContext(settings, sectionContext, bodyStory,
      paragraph({ bidi: true, indentFirst: -9, numbering: num({ text: '' }) }));
    expect(rtlNoMarker.firstIndentPt).toBe(-9);

    // A non-hanging (positive) first-line indent is not the §17.3.1.12 hanging-list
    // shape the tab-stop body placement targets — keep raw indentFirst (matches the
    // paint gate, which also requires indFirst < 0), so measure and paint agree.
    const rtlPositive = resolveParagraphLayoutContext(settings, sectionContext, bodyStory,
      paragraph({ bidi: true, indentFirst: 9, numbering: num() }));
    expect(rtlPositive.firstIndentPt).toBe(9);
  });

  it('gates only line-grid policy for paragraphs and table cells', () => {
    const noCompat = resolveDocumentLayoutSettings(documentModel());
    const withCompat = resolveDocumentLayoutSettings(documentModel({
      settings: { adjustLineHeightInTable: true },
    }));
    const sectionContext = resolveSectionLayoutContext(
      noCompat,
      section({
        docGridType: 'linesAndChars',
        docGridLinePitch: 20,
        docGridCharSpace: 2048,
      }),
    );

    const optedOut = resolveParagraphLayoutContext(
      noCompat,
      sectionContext,
      bodyStory,
      paragraph({ snapToGrid: false }),
    );
    expect(optedOut.lineGrid.active).toBe(false);
    expect(optedOut.characterGrid.active).toBe(true);

    const exact = resolveParagraphLayoutContext(
      noCompat,
      sectionContext,
      bodyStory,
      paragraph({ lineSpacing: { value: 18, rule: 'exact', explicit: true } }),
    );
    expect(exact.lineGrid.active).toBe(false);
    expect(exact.characterGrid.active).toBe(true);

    const cellWithoutCompat = resolveParagraphLayoutContext(
      noCompat,
      sectionContext,
      cellStory,
      paragraph(),
    );
    expect(cellWithoutCompat.lineGrid.active).toBe(false);
    expect(cellWithoutCompat.characterGrid.active).toBe(true);

    const compatibleSection = resolveSectionLayoutContext(
      withCompat,
      section({
        docGridType: 'linesAndChars',
        docGridLinePitch: 20,
        docGridCharSpace: 2048,
      }),
    );
    const cellWithCompat = resolveParagraphLayoutContext(
      withCompat,
      compatibleSection,
      cellStory,
      paragraph(),
    );
    expect(cellWithCompat.lineGrid.active).toBe(true);
  });

  it('aligns the paragraph right edge to the active character-grid pitch', () => {
    const model = documentModel();
    const settings = resolveDocumentLayoutSettings(model, { normalStyleFontSizePt: 9 });
    const charGrid = resolveSectionLayoutContext(settings, section({
      docGridType: 'linesAndChars',
      docGridLinePitch: 20,
      docGridCharSpace: 0,
    }));
    const context = resolveParagraphLayoutContext(
      settings,
      charGrid,
      bodyStory,
      paragraph(),
    );

    expect(settings.normalStyleFontSizePt).toBe(9);
    expect(context.rightIndentGrid).toEqual({
      pitchPt: 9,
      paragraphAllowsAdjustment: true,
    });
    expect(paragraphGridRightAdjustmentPt(context, 481.9)).toBeCloseTo(4.9, 9);

    const optedOut = resolveParagraphLayoutContext(
      settings,
      charGrid,
      bodyStory,
      paragraph({ adjustRightInd: false }),
    );
    expect(optedOut.rightIndentGrid).toEqual({
      pitchPt: 9,
      paragraphAllowsAdjustment: false,
    });
    expect(paragraphGridRightAdjustmentPt(optedOut, 481.9)).toBe(0);

    const lineOnly = resolveParagraphLayoutContext(
      settings,
      resolveSectionLayoutContext(settings, section({
        docGridType: 'lines',
        docGridLinePitch: 20,
      })),
      bodyStory,
      paragraph(),
    );
    expect(lineOnly.rightIndentGrid).toEqual({
      pitchPt: null,
      paragraphAllowsAdjustment: true,
    });

    const snapToChars = resolveParagraphLayoutContext(
      settings,
      resolveSectionLayoutContext(settings, section({
        docGridType: 'snapToChars',
        docGridCharSpace: 0,
      })),
      bodyStory,
      paragraph(),
    );
    expect(snapToChars.characterGrid).toEqual({
      active: true,
      kind: 'snapToChars',
      deltaPt: 0,
      pitchPt: 9,
    });
    expect(snapToChars.rightIndentGrid).toEqual({
      pitchPt: null,
      paragraphAllowsAdjustment: true,
    });

    const tableCell = resolveParagraphLayoutContext(
      settings,
      charGrid,
      cellStory,
      paragraph(),
    );
    expect(tableCell.rightIndentGrid).toEqual({
      pitchPt: 9,
      paragraphAllowsAdjustment: false,
    });
    expect(paragraphGridRightAdjustmentPt(tableCell, 481.9)).toBe(0);
  });

  it('adds the hanging indent as an implicit leading tab without hiding a bar rule', () => {
    const settings = resolveDocumentLayoutSettings(documentModel());
    const context = resolveParagraphLayoutContext(
      settings,
      resolveSectionLayoutContext(settings, section()),
      bodyStory,
      paragraph({
        indentLeft: 72,
        indentFirst: -18,
        tabStops: [
          { pos: 72, alignment: 'bar', leader: 'none' },
          { pos: 144, alignment: 'left', leader: 'none' },
        ],
      }),
    );

    expect(context.tabStops).toEqual([
      { pos: 72, alignment: 'left', leader: 'none' },
      { pos: 72, alignment: 'bar', leader: 'none' },
      { pos: 144, alignment: 'left', leader: 'none' },
    ]);
  });

  it('lets an authored alignment replace the implicit hanging stop at the same position', () => {
    const settings = resolveDocumentLayoutSettings(documentModel());
    const context = resolveParagraphLayoutContext(
      settings,
      resolveSectionLayoutContext(settings, section()),
      bodyStory,
      paragraph({
        indentLeft: 72,
        indentFirst: -18,
        tabStops: [
          { pos: 72, alignment: 'center', leader: 'none' },
          { pos: 144, alignment: 'left', leader: 'none' },
        ],
      }),
    );

    expect(context.tabStops).toEqual([
      { pos: 72, alignment: 'center', leader: 'none' },
      { pos: 144, alignment: 'left', leader: 'none' },
    ]);
  });

  it('lets a run opt out of character-grid spacing without changing line policy', () => {
    const settings = resolveDocumentLayoutSettings(documentModel());
    const sectionContext = resolveSectionLayoutContext(
      settings,
      section({
        docGridType: 'snapToChars',
        docGridLinePitch: 20,
        docGridCharSpace: 4096,
      }),
    );
    const paragraphContext = resolveParagraphLayoutContext(
      settings,
      sectionContext,
      bodyStory,
      paragraph(),
    );

    expect(paragraphContext.lineGrid.active).toBe(true);
    expect(resolveRunLayoutContext(
      paragraphContext,
      textRun({ snapToGrid: false }),
    ).characterGrid).toEqual({ active: false, kind: null, pitchPt: null, deltaPt: 0 });
  });
});
