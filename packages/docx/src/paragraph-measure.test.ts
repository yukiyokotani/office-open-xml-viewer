import { describe, expect, it } from 'vitest';
import { DEFAULT_KINSOKU_RULES } from '@silurus/ooxml-core';
import type { FloatRect } from './float-layout.js';
import {
  createFloatWrapOracle,
  measureParagraph,
  type ParagraphMeasurementEnvironment,
  type ParagraphPlacement,
  type TextMeasurer,
  type WrapOracle,
} from './paragraph-measure.js';
import { measureParagraphIntrinsicWidth } from './layout/frame.js';
import type { ParagraphLayoutContext } from './layout-context.js';
import type { LayoutTextSeg } from './line-layout.js';
import type { DocParagraph, DocxTextRun, FieldRun, ImageRun } from './types.js';

function makeContext(ascentRatio = 0.8, descentRatio = 0.2): CanvasRenderingContext2D {
  let font = '10px serif';
  const fontSize = (): number => Number.parseFloat(/([\d.]+)px/.exec(font)?.[1] ?? '10');
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    measureText: (text: string) => {
      const size = fontSize();
      return {
        width: [...text].length * size * 0.5,
        fontBoundingBoxAscent: size * ascentRatio,
        fontBoundingBoxDescent: size * descentRatio,
        actualBoundingBoxAscent: size * ascentRatio,
        actualBoundingBoxDescent: size * descentRatio,
      } as TextMetrics;
    },
  } as unknown as CanvasRenderingContext2D;
}

const measurer: TextMeasurer = {
  context: makeContext(),
  fontFamilyClasses: {},
};

const environment = (
  overrides: Partial<ParagraphMeasurementEnvironment> = {},
): ParagraphMeasurementEnvironment => ({
  pageIndex: 0,
  totalPages: 1,
  pageWritingMode: 'horizontal-tb',
  documentHasEastAsianText: false,
  ...overrides,
});

const RESOLVED_EA_FAMILY = 'Arbitrary Resolved EA';
const RESOLVED_EA_RATIO = 3269 / 2048;
const resolvedEaMetrics = (ratio = RESOLVED_EA_RATIO) => ({
  [RESOLVED_EA_FAMILY.toLowerCase()]: {
    family: RESOLVED_EA_FAMILY,
    eastAsianLineHeightRatio: ratio,
  },
});

const paragraph = (overrides: Partial<DocParagraph> = {}): DocParagraph => ({
  alignment: 'left',
  indentLeft: 0,
  indentRight: 0,
  indentFirst: 0,
  spaceBefore: 3,
  spaceAfter: 4,
  lineSpacing: null,
  numbering: null,
  tabStops: [],
  runs: [],
  ...overrides,
});

const textRun = (text: string, overrides: Partial<DocxTextRun> = {}): DocxTextRun => ({
  text,
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

const layoutContext = (
  overrides: Partial<ParagraphLayoutContext> = {},
): ParagraphLayoutContext => ({
  lineGrid: { active: false, pitchPt: null },
  characterGrid: { active: false, kind: null, pitchPt: null, deltaPt: 0 },
  rightIndentGrid: { pitchPt: null, paragraphAllowsAdjustment: true },
  physicalIndentLeftPt: 0,
  physicalIndentRightPt: 0,
  firstIndentPt: 0,
  lineSpacing: null,
  spaceBeforePt: 3,
  spaceAfterPt: 4,
  baseRtl: false,
  isJustified: false,
  stretchLastLine: false,
  tabStops: [],
  hasRuby: false,
  hasEastAsianText: false,
  kinsoku: DEFAULT_KINSOKU_RULES,
  defaultTabPt: 36,
  ...overrides,
});

const placement = (overrides: Partial<ParagraphPlacement> = {}): ParagraphPlacement => ({
  startYPt: 10,
  paragraphXPt: 0,
  availableWidthPt: 200,
  maximumYPt: 300,
  suppressSpaceBefore: false,
  ...overrides,
});

const measuredTextSequence = (
  measured: ReturnType<typeof measureParagraph>,
): string[] => measured.lines.map((line) => line.layout.segments
  .filter((segment): segment is LayoutTextSeg => 'text' in segment)
  .map((segment) => segment.text)
  .join(''));

describe('measureParagraph', () => {
  it('uses the same character-grid right-edge adjustment for line partitioning', () => {
    const source = paragraph({
      runs: [{ type: 'text', ...textRun('あ'.repeat(20)) }],
      spaceBefore: 0,
      spaceAfter: 0,
    });
    const grid = layoutContext({
      characterGrid: { active: true, kind: 'linesAndChars', pitchPt: 10, deltaPt: 0 },
      rightIndentGrid: { pitchPt: 9, paragraphAllowsAdjustment: true },
      spaceBeforePt: 0,
      spaceAfterPt: 0,
      hasEastAsianText: true,
    });
    const optedOut = {
      ...grid,
      rightIndentGrid: { pitchPt: 9, paragraphAllowsAdjustment: false },
    };

    const adjusted = measureParagraph(
      source,
      grid,
      placement({ availableWidthPt: 100 }),
      measurer,
      environment({ documentHasEastAsianText: true }),
    );
    const exact = measureParagraph(
      source,
      optedOut,
      placement({ availableWidthPt: 100 }),
      measurer,
      environment({ documentHasEastAsianText: true }),
    );

    expect(adjusted.lines).toHaveLength(2);
    expect(exact.lines).toHaveLength(1);
    expect(adjusted.placement.availableWidthPt).toBe(100);
  });

  it('chooses the globally widest gap from one oracle containing every float', () => {
    const float = (id: string, xLeft: number, xRight: number): FloatRect => ({
      kind: 'shape', mode: 'square', imageKey: id,
      imageX: xLeft, imageY: 0, imageW: xRight - xLeft, imageH: 20,
      xLeft, xRight, yTop: 0, yBottom: 20,
      side: 'bothSides', distLeft: 0, distRight: 0, distTop: 0, distBottom: 0,
      paraId: 0,
    });
    const oracle = createFloatWrapOracle([
      float('A', 40, 60),
      float('B', 0, 35),
    ]);

    expect(oracle.lineWindow({
      topYPt: 0, minimumStartWidthPt: 1, probeHeightPt: 10,
      paragraphXPt: 0, maximumWidthPt: 100,
      columnXPt: 0, columnWidthPt: 100,
    })).toEqual({ topYPt: 0, xOffsetPt: 60, maximumWidthPt: 40 });
  });

  it('snapshots compiled polygon geometry once at oracle acquisition', () => {
    const points = [
      { xPt: 10, yPt: 0 }, { xPt: 90, yPt: 0 },
      { xPt: 90, yPt: 80 }, { xPt: 80, yPt: 80 },
      { xPt: 80, yPt: 20 }, { xPt: 20, yPt: 20 },
      { xPt: 20, yPt: 80 }, { xPt: 10, yPt: 80 },
    ];
    const oracle = createFloatWrapOracle([{
      kind: 'shape', mode: 'square', authoredWrap: 'through', wrapPolygon: points,
      imageKey: 'compiled-through', imageX: 10, imageY: 0, imageW: 80, imageH: 80,
      xLeft: 10, xRight: 90, yTop: 0, yBottom: 80,
      side: 'bothSides', distLeft: 0, distRight: 0, distTop: 0, distBottom: 0,
      paraId: 0,
    }]);
    const query = () => oracle.lineWindow({
      topYPt: 30, minimumStartWidthPt: 1, probeHeightPt: 10,
      paragraphXPt: 0, maximumWidthPt: 100,
      columnXPt: 0, columnWidthPt: 100,
    });
    const acquired = query();

    points[4]!.xPt = 50;
    points[5]!.xPt = 50;

    expect(query()).toEqual(acquired);
    expect(acquired).toEqual({ topYPt: 30, xOffsetPt: 20, maximumWidthPt: 60 });
  });

  it('measures intrinsic width against the authored anchor band without a sentinel width', () => {
    const doc = paragraph({
      spaceBefore: 0,
      spaceAfter: 0,
      runs: [{ type: 'text', ...textRun('abc def') }],
    });
    const context = layoutContext({ spaceBeforePt: 0, spaceAfterPt: 0 });

    expect(measureParagraphIntrinsicWidth(doc, context, 200, measurer, environment())).toBe(35);
    // A normal 25pt layout wraps at the space into 15pt lines. Intrinsic mode
    // keeps the 35pt natural line intact, then caps the preferred width to the
    // real 25pt anchor band.
    expect(measureParagraphIntrinsicWidth(doc, context, 25, measurer, environment())).toBe(25);
  });

  it('includes paragraph indents, hanging numbering space, tabs, bidi, and inline resources', () => {
    const indented = layoutContext({
      spaceBeforePt: 0, spaceAfterPt: 0,
      physicalIndentLeftPt: 12, physicalIndentRightPt: 2, firstIndentPt: -6,
    });
    const numbered = paragraph({
      spaceBefore: 0, spaceAfter: 0,
      numbering: { numId: 1, level: 0, format: 'decimal', text: '1.', indentLeft: 12, tab: 6, suff: 'tab' } as never,
      runs: [{ type: 'text', ...textRun('A') }],
    });
    // The 6pt hanging zone remains inside the 12pt physical left indent.
    expect(measureParagraphIntrinsicWidth(numbered, indented, 100, measurer, environment())).toBe(13);

    const tabbed = paragraph({
      spaceBefore: 0, spaceAfter: 0,
      tabStops: [{ pos: 30, alignment: 'left', leader: 'none' }],
      runs: [{ type: 'text', ...textRun('A\tB') }],
    });
    expect(measureParagraphIntrinsicWidth(
      tabbed,
      layoutContext({
        spaceBeforePt: 0, spaceAfterPt: 0,
        tabStops: [{ pos: 30, alignment: 'left', leader: 'none' }],
      }),
      100,
      measurer,
      environment(),
    )).toBe(35);
    expect(measureParagraphIntrinsicWidth(
      { ...tabbed, bidi: true },
      layoutContext({
        spaceBeforePt: 0, spaceAfterPt: 0, baseRtl: true,
        tabStops: [{ pos: 50, alignment: 'left', leader: 'none' }],
      }),
      100,
      measurer,
      environment(),
    )).toBe(55);

    const image: ImageRun = {
      imagePath: 'word/media/inline.png', mimeType: 'image/png',
      widthPt: 20, heightPt: 10, anchor: false,
    };
    expect(measureParagraphIntrinsicWidth(
      paragraph({ spaceBefore: 0, spaceAfter: 0, runs: [{ type: 'image', ...image }] }),
      layoutContext({
        spaceBeforePt: 0, spaceAfterPt: 0,
        physicalIndentLeftPt: 3, physicalIndentRightPt: 4,
      }),
      100,
      measurer,
      environment(),
    )).toBe(27);
  });

  it('measures a no-float paragraph and excludes trailing spacing from contentEndYPt', () => {
    const result = measureParagraph(
      paragraph({ runs: [{ type: 'text', ...textRun('hello') }] }),
      layoutContext(),
      placement(),
      measurer,
      environment(),
    );

    expect(result.markOnly).toBe(false);
    expect(result.lines).toHaveLength(1);
    expect(result.lines[0].topYPt).toBe(13);
    expect(result.lines[0].advancePt).toBe(10);
    expect(result.contentStartYPt).toBe(13);
    expect(result.contentEndYPt).toBe(23);
    expect(result.requestedSpaceBeforePt).toBe(3);
    expect(result.requestedSpaceAfterPt).toBe(4);
    expect(result.uniformRubyAdvancePt).toBe(0);
    expect(result.placement).toEqual(placement());
  });

  it('uses a float-backed wrap window for line placement and width', () => {
    const float: FloatRect = {
      kind: 'shape', mode: 'square', imageKey: 'float',
      imageX: 0, imageY: 10, imageW: 80, imageH: 30,
      xLeft: 0, xRight: 80, yTop: 10, yBottom: 40,
      side: 'bothSides', distLeft: 0, distRight: 0, distTop: 0, distBottom: 0,
      paraId: 1,
    };

    const result = measureParagraph(
      paragraph({ spaceBefore: 0, runs: [{ type: 'text', ...textRun('wrapped') }] }),
      layoutContext({ spaceBeforePt: 0 }),
      placement({ wrap: createFloatWrapOracle([float]) }),
      measurer,
      environment(),
    );

    expect(result.lines[0].topYPt).toBe(10);
    expect(result.lines[0].layout.xOffset).toBe(80);
    expect(result.lines[0].layout.availWidth).toBe(120);
  });

  it('reserves one paragraph-mark line for an empty paragraph', () => {
    const result = measureParagraph(
      paragraph(), layoutContext(), placement(), measurer, environment(),
    );

    expect(result.markOnly).toBe(true);
    expect(result.lines).toEqual([]);
    expect(result.contentStartYPt).toBe(13);
    expect(result.contentEndYPt).toBe(23);
  });

  it('uses resolved context line spacing for an empty paragraph mark', () => {
    const authoredSpacing = { value: 6, rule: 'exact' as const, explicit: true };
    const resolvedSpacing = { value: 18, rule: 'exact' as const, explicit: true };
    const result = measureParagraph(
      paragraph({ spaceBefore: 0, lineSpacing: authoredSpacing }),
      layoutContext({ spaceBeforePt: 0, lineSpacing: resolvedSpacing }),
      placement({ startYPt: 0 }),
      measurer,
      environment(),
    );

    expect(result.markOnly).toBe(true);
    expect(result.contentEndYPt).toBe(18);
  });

  it('uses document-level East Asian metrics for an empty paragraph mark', () => {
    const tallMeasurer: TextMeasurer = {
      context: makeContext(0.9, 0.2),
      fontFamilyClasses: {},
    };
    const eastAsianMetrics = measureParagraph(
      paragraph({ defaultFontSize: 20 }),
      layoutContext({
        lineGrid: { active: true, pitchPt: 20 },
        spaceBeforePt: 0,
      }),
      placement({ startYPt: 0 }),
      tallMeasurer,
      environment({ documentHasEastAsianText: true }),
    );
    const paragraphOnlyMetrics = measureParagraph(
      paragraph({ defaultFontSize: 20 }),
      layoutContext({
        lineGrid: { active: true, pitchPt: 20 },
        spaceBeforePt: 0,
      }),
      placement({ startYPt: 0 }),
      tallMeasurer,
      environment({ documentHasEastAsianText: false }),
    );

    expect(eastAsianMetrics.contentEndYPt).toBe(40);
    expect(paragraphOnlyMetrics.contentEndYPt).toBe(22);
  });

  it('applies useFELayout grid-cell allocation to an empty paragraph mark', () => {
    const tallMeasurer: TextMeasurer = {
      context: makeContext(0.9, 0.2),
      fontFamilyClasses: {},
    };
    const result = measureParagraph(
      paragraph({ defaultFontSize: 20 }),
      layoutContext({
        lineGrid: { active: true, pitchPt: 20 },
        spaceBeforePt: 0,
      }),
      placement({ startYPt: 0 }),
      tallMeasurer,
      environment({ documentHasEastAsianText: false, useFeLayout: true }),
    );

    // Office compatibility evidence: useFELayout makes the content-less mark
    // participate in the same Far East whole-cell allocation as a CJK mark.
    expect(result.markOnly).toBe(true);
    expect(result.contentEndYPt).toBe(40);
  });

  it('uses resolved-resource Far East metrics for useFELayout empty marks', () => {
    const markAdvance = (
      fontSize: number,
      pitchPt: number,
      ratio?: number,
    ): number => measureParagraph(
      paragraph({
        defaultFontSize: fontSize,
        defaultFontFamily: RESOLVED_EA_FAMILY,
        defaultFontFamilyEastAsia: RESOLVED_EA_FAMILY,
        spaceBefore: 0,
      }),
      layoutContext({
        lineGrid: { active: true, pitchPt },
        spaceBeforePt: 0,
      }),
      placement({ startYPt: 0 }),
      measurer,
      environment({
        useFeLayout: true,
        resolvedLocalFonts: ratio === undefined ? {} : resolvedEaMetrics(ratio),
      }),
    ).contentEndYPt;

    // A parsed 1.651025-em resource crosses these grid boundaries; the same
    // arbitrary family without a resource metric stays on the generic path.
    const ratio = ((2171 + 430) * 1.3) / 2048;
    expect(markAdvance(10, 18, ratio)).toBe(18);
    expect(markAdvance(11, 18, ratio)).toBe(36);
    expect(markAdvance(12, 20, ratio)).toBe(20);
    expect(markAdvance(13, 20, ratio)).toBe(40);
    expect(markAdvance(11, 18)).toBe(18);
  });

  it('keeps positive atLeast useFELayout marks on the grid unless exact spacing overrides it', () => {
    const atLeast = { value: 18, rule: 'atLeast' as const, explicit: true };
    const exact = { value: 18, rule: 'exact' as const, explicit: true };
    const measure = (lineSpacing: typeof atLeast | typeof exact): number => measureParagraph(
      paragraph({
        defaultFontSize: 11,
        defaultFontFamily: RESOLVED_EA_FAMILY,
        defaultFontFamilyEastAsia: RESOLVED_EA_FAMILY,
        lineSpacing,
        spaceBefore: 0,
      }),
      layoutContext({
        lineGrid: { active: true, pitchPt: 18 },
        lineSpacing,
        spaceBeforePt: 0,
      }),
      placement({ startYPt: 0 }),
      measurer,
      environment({
        useFeLayout: true,
        resolvedLocalFonts: resolvedEaMetrics(((2171 + 430) * 1.3) / 2048),
      }),
    ).contentEndYPt;

    // §17.6.5 names exact spacing (not atLeast) as the grid-line override.
    expect(measure(atLeast)).toBe(36);
    expect(measure(exact)).toBe(18);
  });

  it.each([true, false])(
    'keeps an atLeast-zero empty mark at its design advance (explicit=%s)',
    (explicit) => {
      const atLeastZero = { value: 0, rule: 'atLeast' as const, explicit };
      const result = measureParagraph(
        paragraph({
          defaultFontSize: 10,
          defaultFontFamily: RESOLVED_EA_FAMILY,
          defaultFontFamilyEastAsia: RESOLVED_EA_FAMILY,
          lineSpacing: atLeastZero,
          spaceBefore: 0,
        }),
        layoutContext({
          lineGrid: { active: true, pitchPt: 14.55 },
          lineSpacing: atLeastZero,
          spaceBeforePt: 0,
        }),
        placement({ startYPt: 0 }),
        measurer,
        environment({ useFeLayout: true, resolvedLocalFonts: resolvedEaMetrics() }),
      );

      // Word's atLeast-zero compatibility path keeps a line whose design box
      // exceeds one grid pitch at its raw design advance instead of rounding
      // it to a second grid cell. Empty paragraph marks follow the same rule.
      expect(result.contentEndYPt).toBeCloseTo(10 * RESOLVED_EA_RATIO, 12);
    },
  );

  it.each([
    { value: -18, expected: 18 },
    { value: -0.05, expected: 0.05 },
    { value: 0.05, expected: 29.1 },
  ])('keeps the observed signed atLeast empty-mark boundary at $value pt', ({ value, expected }) => {
    const lineSpacing = { value, rule: 'atLeast' as const, explicit: true };
    const result = measureParagraph(
      paragraph({
        defaultFontSize: 10,
        defaultFontFamily: RESOLVED_EA_FAMILY,
        defaultFontFamilyEastAsia: RESOLVED_EA_FAMILY,
        lineSpacing,
        spaceBefore: 0,
      }),
      layoutContext({
        lineGrid: { active: true, pitchPt: 14.55 },
        lineSpacing,
        spaceBeforePt: 0,
      }),
      placement({ startYPt: 0 }),
      measurer,
      environment({ useFeLayout: true, resolvedLocalFonts: resolvedEaMetrics() }),
    );

    expect(result.contentEndYPt).toBeCloseTo(expected, 12);
  });

  it.each([
    { name: 'without a document grid', lineGrid: { active: false, pitchPt: null } },
    // snapToGrid=false is resolved by layout context into an inactive line axis
    // while retaining the section pitch for diagnostics.
    { name: 'when snapToGrid disables the line axis', lineGrid: { active: false, pitchPt: 14.55 } },
  ])('does not project negative atLeast grid compatibility $name', ({ lineGrid }) => {
    const lineSpacing = { value: -0.05, rule: 'atLeast' as const, explicit: true };
    const result = measureParagraph(
      paragraph({
        defaultFontSize: 10,
        defaultFontFamily: RESOLVED_EA_FAMILY,
        defaultFontFamilyEastAsia: RESOLVED_EA_FAMILY,
        lineSpacing,
        spaceBefore: 0,
      }),
      layoutContext({ lineGrid, lineSpacing, spaceBeforePt: 0 }),
      placement({ startYPt: 0 }),
      measurer,
      environment({ useFeLayout: true, resolvedLocalFonts: resolvedEaMetrics() }),
    );

    expect(result.contentEndYPt).toBeCloseTo(10 * RESOLVED_EA_RATIO, 12);
  });

  it('matches observed Word spacing for an explicit atLeast line on a body grid', () => {
    const designRatio = 3269 / 2048;
    const designMeasurer: TextMeasurer = {
      context: makeContext(designRatio * 0.8, designRatio * 0.2),
      fontFamilyClasses: {},
    };
    const explicitAtLeast = { value: 0, rule: 'atLeast' as const, explicit: true };
    const result = measureParagraph(
      paragraph({
        spaceBefore: 0,
        spaceAfter: 0,
        lineSpacing: explicitAtLeast,
        runs: [
          { type: 'text', ...textRun('あ', { fontSize: 14, fontFamilyEastAsia: 'Test CJK' }) },
          { type: 'break', breakType: 'line' },
          { type: 'text', ...textRun('い', { fontSize: 10, fontFamilyEastAsia: 'Test CJK' }) },
        ],
      }),
      layoutContext({
        lineGrid: { active: true, pitchPt: 20 },
        lineSpacing: explicitAtLeast,
        spaceBeforePt: 0,
        spaceAfterPt: 0,
        hasEastAsianText: true,
      }),
      placement({ startYPt: 0 }),
      designMeasurer,
      environment({ documentHasEastAsianText: true }),
    );

    // Windows Word leaves the first explicit-atLeast line at its raw 14pt
    // design height (slightly over one pitch), then keeps the ordinary line at
    // one 20pt pitch. This is a compatibility fixture, not a normative claim
    // that §17.3.1.33 or §17.6.5 defines this exception.
    expect(result.lines.map((line) => line.advancePt))
      .toEqual([14 * designRatio, 20]);
  });

  it('treats an anchor-only paragraph as a paragraph mark', () => {
    const anchor: ImageRun = {
      imagePath: 'word/media/anchor.png', mimeType: 'image/png',
      widthPt: 40, heightPt: 30, anchor: true,
    };
    const result = measureParagraph(
      paragraph({ runs: [{ type: 'image', ...anchor }] }),
      layoutContext(), placement(), measurer, environment(),
    );

    expect(result.markOnly).toBe(true);
    expect(result.lines).toEqual([]);
    expect(result.contentEndYPt).toBe(23);
  });

  it('uses one uniform snapped advance for every line in a ruby paragraph', () => {
    const result = measureParagraph(
      paragraph({
        spaceBefore: 0,
        runs: [{ type: 'text', ...textRun('aa aa', { ruby: { text: 'ruby', fontSizePt: 8 } }) }],
      }),
      layoutContext({
        lineGrid: { active: true, pitchPt: 10 },
        spaceBeforePt: 0,
        hasRuby: true,
      }),
      placement({ availableWidthPt: 12 }),
      measurer,
      environment(),
    );

    expect(result.lines.length).toBeGreaterThan(1);
    // Base ink is 8pt above/2pt below the baseline. The selected 8pt ruby face
    // contributes its exact 1.6pt descent above that base ink, so the natural
    // 10pt line plus the 9.6pt ruby reserve snaps once to the 10pt grid: 20pt.
    const baseNaturalPt = 10;
    const rubyReservePt = 8 + 1.6;
    const gridPitchPt = 10;
    const expectedAdvancePt = Math.ceil((baseNaturalPt + rubyReservePt) / gridPitchPt)
      * gridPitchPt;
    expect(new Set(result.lines.map((line) => line.advancePt)))
      .toEqual(new Set([expectedAdvancePt]));
  });

  it('carries the paragraph-wide ruby advance through continuations', () => {
    const doc = paragraph({
      spaceBefore: 0,
      runs: [
        { type: 'text', ...textRun('aa', { ruby: { text: 'ruby', fontSizePt: 12 } }) },
        { type: 'text', ...textRun(' bb cc dd ee ff gg') },
      ],
    });
    const context = layoutContext({
      lineGrid: { active: true, pitchPt: 10 },
      spaceBeforePt: 0,
      hasRuby: true,
    });
    const position = placement({ startYPt: 0, availableWidthPt: 18 });
    const full = measureParagraph(doc, context, position, measurer, environment());
    const uniformAdvancePt = full.lines[0].advancePt;

    expect(full.lines.length).toBeGreaterThanOrEqual(3);
    expect(new Set(full.lines.map((line) => line.advancePt)))
      .toEqual(new Set([uniformAdvancePt]));

    const continuation = measureParagraph(
      doc,
      context,
      position,
      measurer,
      environment(),
      {
        boundary: full.lines[0].layout.consumedEnd!,
        uniformRubyAdvancePt: uniformAdvancePt,
      },
    );

    expect(continuation.lines.every((line) => line.advancePt === uniformAdvancePt)).toBe(true);
    expect(full.uniformRubyAdvancePt).toBe(uniformAdvancePt);
    expect(continuation.uniformRubyAdvancePt).toBe(uniformAdvancePt);

    const secondContinuation = measureParagraph(
      doc,
      context,
      position,
      measurer,
      environment(),
      {
        boundary: continuation.lines[0].layout.consumedEnd!,
        uniformRubyAdvancePt: continuation.uniformRubyAdvancePt,
      },
    );

    expect(continuation.lines.length).toBeGreaterThan(1);
    expect(secondContinuation.lines.every((line) => line.advancePt === uniformAdvancePt)).toBe(true);
    expect(secondContinuation.uniformRubyAdvancePt).toBe(uniformAdvancePt);
  });

  it('passes bidi policy through to RTL tab layout', () => {
    const result = measureParagraph(
      paragraph({
        spaceBefore: 0,
        tabStops: [{ pos: 50, alignment: 'left', leader: 'none' }],
        runs: [{ type: 'text', ...textRun('A\tB') }],
        bidi: true,
      }),
      layoutContext({
        spaceBeforePt: 0,
        baseRtl: true,
        tabStops: [{ pos: 50, alignment: 'left', leader: 'none' }],
      }),
      placement({ availableWidthPt: 100 }),
      measurer,
      environment(),
    );

    const tab = result.lines[0].layout.segments.find((segment) => 'isTab' in segment);
    expect(tab?.measuredWidth).toBe(45);
  });

  it('includes an inline image in line height', () => {
    const image: ImageRun = {
      imagePath: 'word/media/inline.png', mimeType: 'image/png',
      widthPt: 20, heightPt: 24, anchor: false,
    };
    const result = measureParagraph(
      paragraph({ spaceBefore: 0, runs: [{ type: 'image', ...image }] }),
      layoutContext({ spaceBeforePt: 0 }), placement(), measurer, environment(),
    );

    expect(result.markOnly).toBe(false);
    expect(result.lines[0].advancePt).toBe(24);
  });

  it('preserves exact line spacing verbatim', () => {
    const exact = { value: 18, rule: 'exact' as const, explicit: true };
    const result = measureParagraph(
      paragraph({ spaceBefore: 0, lineSpacing: exact, runs: [{ type: 'text', ...textRun('exact') }] }),
      layoutContext({ spaceBeforePt: 0, lineSpacing: exact }),
      placement(), measurer, environment(),
    );

    expect(result.lines[0].advancePt).toBe(18);
  });

  it('resolves fields from the explicit line-layout environment', () => {
    const field: FieldRun = {
      fieldType: 'page', instruction: 'PAGE', fallbackText: '1',
      bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: null, background: null,
      vertAlign: null,
    };
    const result = measureParagraph(
      paragraph({ spaceBefore: 0, runs: [{ type: 'field', ...field }] }),
      layoutContext({ spaceBeforePt: 0 }), placement(),
      measurer,
      environment({ pageIndex: 8, totalPages: 12, displayPageNumber: 42 }),
    );

    expect(result.lines[0].layout.segments[0]).toMatchObject({ text: '42' });
  });

  it('hands the RAW column band (not the indented text band) to the wrap oracle', () => {
    // §20.4.2.20 / §17.6.4: the topAndBottom gate must see the COLUMN band, so
    // measure and paint scope a page-shared float to the same column. With a
    // physical left indent the indented text band (paragraphXPt) differs from the
    // column band (placement.paragraphXPt / availableWidthPt); both the one-time
    // pre-paragraph skip AND the per-line window must be handed the column band.
    const lineWindowColumns: Array<{
      columnXPt: number;
      columnWidthPt: number;
      paragraphXPt: number;
      maximumWidthPt: number;
    }> = [];
    const skipColumns: Array<{ columnXPt: number; columnWidthPt: number }> = [];
    const wrap: WrapOracle = {
      lineWindow: (input) => {
        lineWindowColumns.push({
          columnXPt: input.columnXPt,
          columnWidthPt: input.columnWidthPt,
          paragraphXPt: input.paragraphXPt,
          maximumWidthPt: input.maximumWidthPt,
        });
        return { topYPt: input.topYPt, xOffsetPt: 0, maximumWidthPt: input.maximumWidthPt };
      },
      skipTopAndBottomBands: (input) => {
        skipColumns.push({ columnXPt: input.columnXPt, columnWidthPt: input.columnWidthPt });
        return input.yPt;
      },
    };

    measureParagraph(
      paragraph({ spaceBefore: 0, runs: [{ type: 'text', ...textRun('hello world') }] }),
      layoutContext({ spaceBeforePt: 0, physicalIndentLeftPt: 100, physicalIndentRightPt: 0 }),
      placement({ paragraphXPt: 60, availableWidthPt: 228, wrap }),
      measurer,
      environment(),
    );

    // Column band = placement (60, 228). Indented text band = 60 + 100 = 160,
    // width 228 − 100 = 128. The oracle must be scoped to the COLUMN band.
    expect(skipColumns.length).toBeGreaterThan(0);
    for (const c of skipColumns) {
      expect(c.columnXPt).toBe(60);
      expect(c.columnWidthPt).toBe(228);
    }
    expect(lineWindowColumns.length).toBeGreaterThan(0);
    for (const c of lineWindowColumns) {
      expect(c.columnXPt).toBe(60);
      expect(c.columnWidthPt).toBe(228);
      // The indented text band handed alongside (for the square side-gap math) is
      // distinct — this is exactly the seam the finding is about.
      expect(c.paragraphXPt).toBe(160);
      expect(c.maximumWidthPt).toBe(128);
    }
  });

  it('remeasures at a changed start Y and records the exact placement', () => {
    const wrap: WrapOracle = {
      lineWindow: ({ topYPt, maximumWidthPt }) => topYPt < 50
        ? { topYPt, xOffsetPt: 10, maximumWidthPt: 20 }
        : { topYPt, xOffsetPt: 0, maximumWidthPt },
      skipTopAndBottomBands: ({ yPt }) => yPt,
    };
    const doc = paragraph({
      spaceBefore: 0,
      runs: [{ type: 'text', ...textRun('abcdefghijklmnopqrst') }],
    });
    const context = layoutContext({ spaceBeforePt: 0 });
    const firstPlacement = placement({ startYPt: 10, availableWidthPt: 100, wrap });
    const secondPlacement = placement({ startYPt: 60, availableWidthPt: 100, wrap });

    const first = measureParagraph(doc, context, firstPlacement, measurer, environment());
    const second = measureParagraph(doc, context, secondPlacement, measurer, environment());

    expect(first).not.toBe(second);
    expect(first.placement).toEqual(firstPlacement);
    expect(second.placement).toEqual(secondPlacement);
    expect(first.lines[0].layout.availWidth).toBe(20);
    expect(second.lines[0].layout.availWidth).toBe(100);
    expect(first.lines.length).toBeGreaterThan(second.lines.length);
  });

  it('reproduces the same-width suffix from a consumed line boundary', () => {
    const doc = paragraph({
      spaceBefore: 0,
      runs: [{
        type: 'text',
        ...textRun('alpha bravo charlie delta echo foxtrot golf hotel india juliet'),
      }],
    });
    const context = layoutContext({ spaceBeforePt: 0 });
    const position = placement({ startYPt: 0, availableWidthPt: 45 });
    const full = measureParagraph(doc, context, position, measurer, environment());
    const boundary = full.lines[0].layout.consumedEnd!;

    expect(full.lines.length).toBeGreaterThan(2);
    const continuation = measureParagraph(
      doc,
      context,
      position,
      measurer,
      environment(),
      { boundary },
    );

    expect(measuredTextSequence(continuation)).toEqual(measuredTextSequence(full).slice(1));
  });

  it('suppresses first-line indent when measuring a continuation', () => {
    const doc = paragraph({
      indentFirst: 20,
      spaceBefore: 0,
      runs: [{ type: 'text', ...textRun('a a a a a a a a a a a a a a a a a a') }],
    });
    const context = layoutContext({ firstIndentPt: 20, spaceBeforePt: 0 });
    const position = placement({ startYPt: 0, availableWidthPt: 50 });
    const full = measureParagraph(doc, context, position, measurer, environment());
    const boundary = full.lines[0].layout.consumedEnd!;
    const continuation = measureParagraph(
      doc,
      context,
      position,
      measurer,
      environment(),
      { boundary },
    );
    const fullText = measuredTextSequence(full);
    const continuationText = measuredTextSequence(continuation);

    expect(continuationText[0]).toBe(fullText[1]);
    expect(continuationText[0].length).toBeGreaterThan(fullText[0].length);
  });

  it('re-wraps a continuation at a narrower width without losing text', () => {
    const doc = paragraph({
      spaceBefore: 0,
      runs: [{ type: 'text', ...textRun('あ'.repeat(32)) }],
    });
    const context = layoutContext({ spaceBeforePt: 0 });
    const full = measureParagraph(
      doc,
      context,
      placement({ startYPt: 0, availableWidthPt: 40 }),
      measurer,
      environment(),
    );
    const boundary = full.lines[0].layout.consumedEnd!;
    const continuation = measureParagraph(
      doc,
      context,
      placement({ startYPt: 0, availableWidthPt: 20 }),
      measurer,
      environment(),
      { boundary },
    );

    for (const line of continuation.lines) {
      expect(line.layout.segments.reduce((sum, segment) => sum + segment.measuredWidth, 0))
        .toBeLessThanOrEqual(20);
    }
    expect(measuredTextSequence(continuation).join(''))
      .toBe(measuredTextSequence(full).slice(1).join(''));
    expect(continuation.lines.length).toBeGreaterThan(full.lines.length - 1);
  });

  it('composes continuation boundaries in original segment coordinates', () => {
    const doc = paragraph({
      spaceBefore: 0,
      runs: [{
        type: 'text',
        ...textRun('alpha bravo charlie delta echo foxtrot golf hotel india juliet'),
      }],
    });
    const context = layoutContext({ spaceBeforePt: 0 });
    const position = placement({ startYPt: 0, availableWidthPt: 45 });
    const full = measureParagraph(doc, context, position, measurer, environment());
    const first = measureParagraph(
      doc,
      context,
      position,
      measurer,
      environment(),
      { boundary: full.lines[0].layout.consumedEnd! },
    );
    const second = measureParagraph(
      doc,
      context,
      position,
      measurer,
      environment(),
      { boundary: first.lines[0].layout.consumedEnd! },
    );

    expect(first.lines.length).toBeGreaterThan(1);
    expect(measuredTextSequence(second)).toEqual(measuredTextSequence(first).slice(1));
  });
});
