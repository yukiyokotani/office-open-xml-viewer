import { describe, expect, it } from 'vitest';
import { DEFAULT_KINSOKU_RULES } from '@silurus/ooxml-core';
import type { FloatRect } from './float-layout.js';
import { polygonMeetsRect } from './layout/float-wrap.js';
import {
  createFloatWrapOracle,
  measureParagraph,
  type ParagraphMeasurementEnvironment,
  type ParagraphPlacement,
  type TextMeasurer,
  type WrapOracle,
} from './paragraph-measure.js';
import { measureParagraphIntrinsicWidth } from './layout/frame.js';
import { createFontResolver } from './layout/font-service.js';
import { createTextLayoutService } from './layout/text.js';
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
    sourceIdentity: 'test-resource:resolved-east-asia',
  },
});

function resolvedEaEnvironment(ratio?: number): ParagraphMeasurementEnvironment {
  const metrics = ratio === undefined ? {} : resolvedEaMetrics(ratio);
  const text = createTextLayoutService({
    fonts: createFontResolver(ratio === undefined ? [] : [{
      requestedFamily: RESOLVED_EA_FAMILY,
      resolvedFamily: RESOLVED_EA_FAMILY,
      source: 'local',
      resourceIdentity: 'test-resource:resolved-east-asia',
    }]),
    measurer: {
      fingerprint: 'resolved-east-asian-mark-test',
      measure: (request) => ({
        advancePt: [...request.text].length * request.fontSizePt * 0.5,
        ascentPt: request.fontSizePt * 0.8,
        descentPt: request.fontSizePt * 0.2,
      }),
    },
    fontMetrics: metrics,
    localMetrics: metrics,
  });
  return environment({
    useFeLayout: true,
    layoutServices: { text } as NonNullable<ParagraphMeasurementEnvironment['layoutServices']>,
  });
}

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
  it.each([
    ['Latin', 'word '.repeat(40), false],
    ['East Asian grid', '日本語の本文です。'.repeat(12), true],
  ] as const)('gives every fragment of a picture-bullet first line one advance (%s)', (_, text, grid) => {
    const float: FloatRect = { kind: 'shape', mode: 'square', authoredWrap: 'square',
      imageKey: 'center', imageX: 70, imageY: 0, imageW: 60, imageH: 200,
      xLeft: 70, xRight: 130, yTop: 0, yBottom: 200, side: 'bothSides',
      distLeft: 0, distRight: 0, distTop: 0, distBottom: 0, paraId: 0 };
    const result = measureParagraph(
      paragraph({ spaceBefore: 0, runs: [{ type: 'text', ...textRun(text) }] }),
      layoutContext({
        spaceBeforePt: 0,
        hasEastAsianText: grid,
        ...(grid ? { lineGrid: { active: true, pitchPt: 18 } } : {}),
      }),
      placement({ startYPt: 0, wrap: createFloatWrapOracle([float]) }),
      measurer,
      environment({ firstLinePictureMarkerHeightPt: 17, documentHasEastAsianText: grid }),
    );
    const firstTop = result.lines[0]!.topYPt;
    const first = result.lines.filter((line) => line.topYPt === firstTop);
    expect(first.length).toBeGreaterThan(1);
    expect(new Set(first.map((line) => line.advancePt)).size).toBe(1);
    const next = result.lines.find((line) => line.topYPt > firstTop)!;
    expect(next.topYPt).toBeGreaterThanOrEqual(firstTop + first[0]!.advancePt - 1e-8);
    expect(first[0]!.advancePt).toBeGreaterThanOrEqual(grid ? 36 : 17);
  });

  it('allocates the selected text marker and body line union before pagination', () => {
    const doc = paragraph({ runs: [{ type: 'text', ...textRun('List item') }] });
    const auto = layoutContext({
      spaceBeforePt: 0,
      lineSpacing: { rule: 'auto', value: 1.15, explicit: true },
    });
    const first = measureParagraph(
      doc, auto, placement({ startYPt: 0 }), measurer,
      environment({ firstLineNumberingMarkerBox: { ascentPt: 12, descentPt: 1 } }),
    );
    const plain = measureParagraph(
      doc, auto, placement({ startYPt: 0 }), measurer, environment(),
    );
    expect(plain.lines[0]?.advancePt).toBeCloseTo(11.5);
    expect(first.lines[0]?.advancePt).toBeCloseTo(15.5);
    expect(first.contentEndYPt).toBeCloseTo(15.5);
    expect(first.lines[0]?.layout.ascent).toBeCloseTo(12);

    const exact = measureParagraph(
      doc, layoutContext({ spaceBeforePt: 0, lineSpacing: { rule: 'exact', value: 12, explicit: true } }),
      placement({ startYPt: 0 }), measurer,
      environment({ firstLineNumberingMarkerBox: { ascentPt: 12, descentPt: 1 } }),
    );
    expect(exact.lines[0]?.advancePt).toBeCloseTo(12);
  });
  it.each([undefined, 12])('keeps marker descent out of a legacy list line (mode %s)', (compatibilityMode) => {
    const doc = paragraph({ runs: [{ type: 'text', ...textRun('List item') }] });
    const context = layoutContext({ spaceBeforePt: 0,
      lineSpacing: { rule: 'auto', value: 1.15, explicit: true } });
    const at = placement({ startYPt: 0 });
    const plain = measureParagraph(doc, context, at, measurer, environment());
    const marked = measureParagraph(doc, context, at, measurer,
      environment({ compatibilityMode, firstLineNumberingMarkerBox: {
        ascentPt: 8, descentPt: 6, intendedSinglePt: 14,
      } }));
    // Windows Word: a same-size marker can have a deeper selected font box
    // without raising the text ascent or moving the following paragraph.
    expect(marked.contentEndYPt).toBeCloseTo(plain.contentEndYPt);
  });
  it('treats implicit single spacing and explicit auto1 equally for a taller text marker', () => {
    const doc = paragraph({ runs: [{ type: 'text', ...textRun('List item') }] });
    const marker = environment({ firstLineNumberingMarkerBox: { ascentPt: 12, descentPt: 1 } });
    const implicit = measureParagraph(doc, layoutContext({ spaceBeforePt: 0, lineSpacing: null }),
      placement({ startYPt: 0 }), measurer, marker);
    const explicit = measureParagraph(doc, layoutContext({ spaceBeforePt: 0,
      lineSpacing: { rule: 'auto', value: 1, explicit: true } }),
      placement({ startYPt: 0 }), measurer, marker);
    // §17.3.1.33: absent line spacing is single spacing. The glyph union is
    // max(12, 8) + max(1, 2), without extra leading.
    expect(implicit.lines[0]?.advancePt).toBe(14);
    expect(implicit.contentEndYPt).toBe(explicit.contentEndYPt);
    expect(implicit.lines[0]?.layout.ascent).toBe(explicit.lines[0]?.layout.ascent);
  });
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

  it('chooses the first usable gap from one oracle containing every float', () => {
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
    })).toEqual({ topYPt: 0, xOffsetPt: 35, maximumWidthPt: 5 });
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
      topYPt: 30, minimumStartWidthPt: 20, probeHeightPt: 10,
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

  it('fingerprints anchor-host segments whose shaping script slot is unset', () => {
    // anchorHost runs acquire as metric-only empty segments (line-layout.ts),
    // which never pass through the shaping service, so their optional
    // LayoutTextSeg.script stays undefined. The intrinsic-merge fingerprint
    // must tolerate that instead of throwing "Cannot fingerprint undefined".
    const doc = paragraph({
      spaceBefore: 0,
      spaceAfter: 0,
      runs: [
        { type: 'anchorHost', fontSize: 10 },
        { type: 'text', ...textRun('abc') },
      ],
    });
    const context = layoutContext({ spaceBeforePt: 0, spaceAfterPt: 0 });

    expect(measureParagraphIntrinsicWidth(doc, context, 200, measurer, environment()))
      .toBe(15);
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

  it('retains the selected empty-mark gap for paragraph shading', () => {
    const float: FloatRect = { kind: 'shape', mode: 'square', authoredWrap: 'square',
      imageKey: 'empty-mark-gap', imageX: 40, imageY: 0, imageW: 120, imageH: 60,
      xLeft: 40, xRight: 160, yTop: 0, yBottom: 60, side: 'bothSides',
      distLeft: 0, distRight: 0, distTop: 0, distBottom: 0, paraId: 0 };
    const result = measureParagraph(paragraph({ spaceBefore: 0, shading: 'FFFF00' }), layoutContext({ spaceBeforePt: 0 }),
      placement({ paragraphXPt: 0, startYPt: 0, availableWidthPt: 200, wrap: createFloatWrapOracle([float]) }),
      measurer, environment());
    expect(result.markWrapBounds).toEqual({ xPt: 0, yPt: 0, widthPt: 40, heightPt: 10 });
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

  it('keeps useFELayout from selecting an East Asian mark face without a line grid', () => {
    const text = createTextLayoutService({
      fonts: createFontResolver(['Latin Mark', 'EA Mark'].map((family) => ({
        requestedFamily: family, resolvedFamily: family, source: 'local' as const,
      }))),
      measurer: {
        fingerprint: 'different-latin-and-east-asian-mark-boxes',
        measure: (request) => ({
          advancePt: request.fontSizePt * 0.5,
          ascentPt: request.fontSizePt * (request.fontRoute.familyList.includes('EA Mark') ? 1.2 : 0.8),
          descentPt: request.fontSizePt * 0.2,
        }),
      },
    });
    const advance = (useFeLayout: boolean, eastAsianHint = false): number => measureParagraph(
      paragraph({ defaultFontSize: 10, defaultFontFamily: 'Latin Mark', defaultFontFamilyEastAsia: 'EA Mark' }),
      layoutContext({ spaceBeforePt: 0 }),
      placement({ startYPt: 0 }),
      measurer,
      environment({
        useFeLayout,
        layoutServices: { text } as NonNullable<ParagraphMeasurementEnvironment['layoutServices']>,
        paragraphMarkShapeInput: {
          fonts: { ascii: 'Latin Mark', highAnsi: 'Latin Mark', eastAsia: 'EA Mark' },
          fontSizePt: 10, weight: 400, style: 'normal', complexScript: false,
          ...(eastAsianHint ? { fontHint: 'eastAsia' as const } : {}),
        },
      }),
    ).contentEndYPt;

    expect(advance(false)).toBe(10);
    expect(advance(true)).toBe(10);
    // An authored mark hint still selects the East Asian face without a grid.
    expect(advance(true, true)).toBe(14);
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
      resolvedEaEnvironment(ratio),
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
      resolvedEaEnvironment(((2171 + 430) * 1.3) / 2048),
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
        resolvedEaEnvironment(RESOLVED_EA_RATIO),
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
      resolvedEaEnvironment(RESOLVED_EA_RATIO),
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
      // Keep the East Asian content classification independent of useFELayout:
      // disabling the line grid removes its signed-spacing projection only.
      { ...resolvedEaEnvironment(RESOLVED_EA_RATIO), documentHasEastAsianText: true },
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
    const position = placement({ startYPt: 0, availableWidthPt: 18, wrap: createFloatWrapOracle([]) });
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
    expect(continuation.lines.map(line => line.topYPt))
      .toEqual(continuation.lines.map((_, index) => index * uniformAdvancePt));

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

  it('keeps a text-dominated line with a small inline picture on the text grid policy', () => {
    const result = measureParagraph(
      paragraph({ spaceBefore: 0, runs: [
        { type: 'text', ...textRun('Tall', { fontSize: 30, fontFamily: 'Synthetic Untabled Serif' }) },
        { type: 'image', imagePath: 'word/media/inline.png', mimeType: 'image/png',
          widthPt: 5, heightPt: 5, anchor: false },
      ] }),
      layoutContext({ spaceBeforePt: 0, lineGrid: { active: true, pitchPt: 20 } }),
      placement(), measurer, environment(),
    );
    expect(result.lines[0].advancePt).toBe(30);
  });

  it('takes an image-only line’s auto leading from its selected paragraph-mark face', () => {
    const multiple = 259 / 240;
    const imageHeightPt = 360000 / 12700;
    const text = createTextLayoutService({
      // The Word PDF comparison assumes a loaded Calibri paragraph-mark face.
      // An authored CSS name alone cannot establish that face in Canvas.
      fonts: createFontResolver([{
        requestedFamily: 'Calibri', resolvedFamily: 'Calibri', source: 'local',
        resourceIdentity: 'office-local:local("Calibri")',
      }]),
      measurer: {
        fingerprint: 'inline-picture-mark-face',
        measure: (request) => ({
          advancePt: [...request.text].length * request.fontSizePt * 0.5,
          ascentPt: request.fontSizePt * 0.8,
          descentPt: request.fontSizePt * 0.2,
        }),
      },
      fontMetrics: {}, localMetrics: {},
    });
    const source = paragraph({
      defaultFontFamily: 'Calibri', defaultFontSize: 11,
      spaceBefore: 0, spaceAfter: 0,
      lineSpacing: { rule: 'auto', value: multiple, explicit: true },
      runs: [{
        type: 'image', imagePath: 'word/media/inline.png', mimeType: 'image/png',
        widthPt: imageHeightPt, heightPt: imageHeightPt, anchor: false,
      }],
    });
    const result = measureParagraph(
      source,
      layoutContext({
        spaceBeforePt: 0, spaceAfterPt: 0,
        lineSpacing: { rule: 'auto', value: multiple, explicit: true },
      }),
      placement(), measurer,
      environment({
        layoutServices: { text } as NonNullable<ParagraphMeasurementEnvironment['layoutServices']>,
      }),
    );

    // Controlled Word PDF advances 29.363pt between successive image tops;
    // multiplying the 28.346pt picture by 259/240 would advance 30.591pt.
    expect(result.lines[0].advancePt).toBeCloseTo(29.363, 1);
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

  it('places an unwrapped (DrawingML wrap=none) line through the same exclusion fixed point', () => {
    // A full-band story float occupies y < 50 (e.g. a notBeside frame host).
    // wrap=none removes the box-edge break only, so the 30-glyph text (150pt)
    // stays one line wider than the 100pt band, and that line still starts
    // below the float with its own allocation.
    const wrap: WrapOracle = {
      lineWindow: ({ topYPt, maximumWidthPt }) => ({
        topYPt: Math.max(topYPt, 50), xOffsetPt: 0, maximumWidthPt,
      }),
      skipTopAndBottomBands: ({ yPt }) => yPt,
    };
    const doc = paragraph({
      spaceBefore: 0,
      runs: [{ type: 'text', ...textRun('abcdefghijklmnopqrstuvwxyzabcd') }],
    });
    const measured = measureParagraph(
      doc,
      layoutContext({ spaceBeforePt: 0 }),
      placement({ startYPt: 10, availableWidthPt: 100, noWrap: true, wrap }),
      measurer,
      environment(),
    );
    expect(measuredTextSequence(measured)).toEqual(['abcdefghijklmnopqrstuvwxyzabcd']);
    expect(measured.lines[0].topYPt).toBe(50);
    expect(measured.lines[0].layout.wrapAllocation).toEqual({
      physicalLineIndex: measured.lines[0].layout.physicalLineIndex,
      topYPt: 50,
      advancePt: measured.lines[0].advancePt,
    });
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

describe('ruby physical-line allocation', () => {
  const rubyParagraph = () => paragraph({
    spaceBefore: 0,
    runs: [
      { type: 'text', ...textRun('aa', { ruby: { text: 'ruby', fontSizePt: 12 } }) },
      { type: 'text', ...textRun(' bb cc dd ee ff gg') },
    ],
  });
  const rubyContext = () => layoutContext({
    lineGrid: { active: true, pitchPt: 10 }, spaceBeforePt: 0, hasRuby: true,
  });

  it('starts every physical line at its paragraph-wide allocated advance', () => {
    const measured = measureParagraph(rubyParagraph(), rubyContext(),
      placement({ startYPt: 0, availableWidthPt: 18, wrap: createFloatWrapOracle([]) }),
      measurer, environment());
    expect(measured.lines.map(line => line.topYPt)).toEqual([0, 30, 60, 90, 120, 150, 180]);
    expect(measured.contentEndYPt).toBe(210);
  });

  it('probes and starts wrapped gaps with one allocated advance per physical baseline', () => {
    const doc = rubyParagraph();
    doc.runs.push({ type: 'text', ...textRun(' hh ii jj kk ll mm nn oo pp qq rr ss tt uu vv ww xx yy zz') });
    const obstacle: FloatRect = {
      kind: 'shape', mode: 'square', authoredWrap: 'square', imageKey: 'test',
      imageX: 40, imageY: 0, imageW: 120, imageH: 100,
      xLeft: 40, xRight: 160, yTop: 0, yBottom: 100, side: 'bothSides',
      distLeft: 0, distRight: 0, distTop: 0, distBottom: 0, paraId: 0,
    };
    const oracle = createFloatWrapOracle([obstacle]);
    const probes: { topYPt: number; probeHeightPt: number }[] = [];
    const measured = measureParagraph(doc, rubyContext(),
      placement({ startYPt: 0, availableWidthPt: 200, wrap: {
        ...oracle,
        lineWindow: input => { probes.push(input); return oracle.lineWindow(input); },
      } }), measurer, environment());
    const physical = measured.lines.filter((line, index) => index === 0
      || measured.lines[index - 1].layout.physicalLineIndex !== line.layout.physicalLineIndex);
    expect(measured.lines.length).toBeGreaterThan(physical.length);
    for (let index = 1; index < physical.length; index += 1) {
      expect(physical[index].topYPt)
        .toBeGreaterThanOrEqual(physical[index - 1].topYPt + physical[index - 1].advancePt);
    }
    for (const line of measured.lines) {
      expect(probes).toContainEqual(expect.objectContaining({
        topYPt: line.topYPt, probeHeightPt: 30,
      }));
      expect(line.topYPt).toBe(physical[line.layout.physicalLineIndex!].topYPt);
    }
  });
});

function squareObstacle(left: number, right: number): FloatRect {
  return { kind: 'shape', mode: 'square', authoredWrap: 'square', imageKey: 'test',
    imageX: left, imageY: 0, imageW: right - left, imageH: 60,
    xLeft: left, xRight: right, yTop: 0, yBottom: 60, side: 'bothSides',
    distLeft: 0, distRight: 0, distTop: 0, distBottom: 0, paraId: 0 };
}

// Word controls longword-both-40 (#1670, modes 14/15): a word that fits no
// side gap moves below the object instead of an emergency split in the gap.
// A URL syntax opportunity is a legal break; this value has none that fits.
it.each(['text', 'hyperlink', 'field'] as const)('moves an unsplittable %s word below a sole float gap', (kind) => {
  const value = 'https://example.test/a-b/c';
  const run: DocParagraph['runs'][number] = kind === 'field'
    ? { type: 'field', ...textRun(''), fieldType: 'unknown', instruction: '', fallbackText: value }
    : { type: 'text', ...textRun(value, kind === 'hyperlink' ? { hyperlink: value } : {}) };
  const measured = measureParagraph(paragraph({ runs: [
    { type: 'text', ...textRun('ABC ') }, run,
  ] }), layoutContext({ spaceBeforePt: 0 }),
  placement({ startYPt: 0, wrap: createFloatWrapOracle([squareObstacle(0, 97)]) }), measurer, environment());
  expect(measured.lines.map(({ layout, topYPt }) => ({
    top: topYPt, left: layout.xOffset,
    text: layout.segments.map(segment => 'text' in segment ? segment.text : '').join(''),
  }))).toEqual([
    { top: 0, left: 97, text: 'ABC ' },
    { top: 60, left: 0, text: value },
  ]);
});

it('admits the complete fitText region with internal spaces beside a square float', () => {
  const doc = paragraph({ spaceBefore: 0, runs: [
    { type: 'text', ...textRun('AA BB', { fitTextVal: 1600, fitTextId: 1 }) },
  ] });
  const measured = measureParagraph(doc, layoutContext({ spaceBeforePt: 0 }),
    placement({ startYPt: 0, wrap: createFloatWrapOracle([squareObstacle(40, 100)]) }),
    measurer, environment());
  expect(measured.lines[0].layout.xOffset).toBe(100);
  expect(measured.lines[0].layout.availWidth).toBe(100);
  expect(measured.lines[0].layout.segments.reduce((width, segment) => width + segment.measuredWidth, 0)).toBe(80);
});

// ── WORD_FLOAT_GAP_FLOW model (#1670) ────────────────────────────────────
// Gaps of a physical line are candidates in reading order. A gap is admitted
// iff ordinary placement puts its head unit there without a forced break or
// overflow; otherwise the fragment rolls back and the next gap (or line) is
// tried. These reviewer and continuation repros pin the model directly.

function tightObstacle(left: number, right: number): FloatRect {
  const points = [
    { xPt: left, yPt: 0 }, { xPt: right, yPt: 0 }, { xPt: (left + right) / 2, yPt: 60 },
  ];
  return { ...squareObstacle(left, right), authoredWrap: 'tight', wrapPolygon: points };
}

/** Production supplies the paragraph's reading order to the oracle. */
const gapReference = (rtl: boolean) =>
  ({ xLeftPt: 0, xRightPt: 200, readingDirection: rtl ? 'rtl' : 'ltr' } as const);

function gapLines(runs: DocParagraph['runs'], obstacle: FloatRect, rtl = false) {
  return measureParagraph(paragraph({ runs }), layoutContext({ spaceBeforePt: 0, baseRtl: rtl }),
    placement({ startYPt: 0, wrap: createFloatWrapOracle([obstacle], gapReference(rtl)) }), measurer, environment())
    .lines.map(({ topYPt, layout }) => ({
      top: topYPt, left: layout.xOffset, width: layout.availWidth,
      text: layout.segments.map((segment) => 'text' in segment ? segment.text : '').join(''),
      advance: layout.segments.reduce((sum, segment) => sum + segment.measuredWidth, 0),
    }));
}

describe('float gap admission by placement', () => {
  it('starts in the first gap in reading order that admits the head unit', () => {
    // Main chose the widest gap (x=101); Word fills gaps left to right.
    expect(gapLines([{ type: 'text', ...textRun('A') }], squareObstacle(87, 101)))
      .toEqual([{ top: 0, left: 0, width: 87, text: 'A', advance: 5 }]);
  });

  it('measures a word joined across runs as one unit at the gap head', () => {
    // "NEXT" + "AA BB": the 30pt word NEXTAA misses the 24pt gap, fits the 35pt one.
    const lines = gapLines([
      { type: 'text', ...textRun('NEXT') }, { type: 'text', ...textRun('AA BB') },
    ], squareObstacle(24, 165));
    expect(lines[0]).toMatchObject({ top: 0, left: 165, text: 'NEXTAA ' });
    for (const line of lines) {
      expect(line.advance - (line.text.length - line.text.trimEnd().length) * 5)
        .toBeLessThanOrEqual(line.width);
    }
  });

  it('admits a fitText region only where its complete cell fits', () => {
    const lines = gapLines([
      { type: 'text', ...textRun('AA BB', { fitTextVal: 1600, fitTextId: 1 }) },
    ], squareObstacle(40, 100));
    expect(lines[0]).toMatchObject({ top: 0, left: 100, width: 100, advance: 80 });
  });

  it('fills successive gaps of one baseline before opening the next line', () => {
    const lines = gapLines([{ type: 'text', ...textRun('P008 [red] one two three') }],
      squareObstacle(40, 110));
    expect(lines.slice(0, 2)).toEqual([
      expect.objectContaining({ top: 0, left: 0, text: 'P008 ' }),
      expect.objectContaining({ top: 0, left: 110 }),
    ]);
  });

  // Review round 4: emergency splitting keeps one grapheme, so a whole
  // overwide grapheme must be rejected by its placed advance, not split index.
  it.each(['A', 'e\u0301', '😀', '👩\u200d💻'])('rejects a gap narrower than the single grapheme %s', (value) => {
    expect(gapLines([{ type: 'text', ...textRun(value) }], squareObstacle(2, 100))[0])
      .toMatchObject({ top: 0, left: 100 });
  });

  it('rejects a gap narrower than a joined one-grapheme follower', () => {
    const lines = gapLines([
      { type: 'text', ...textRun('AB') }, { type: 'text', ...textRun('C') },
    ], squareObstacle(12, 100));
    expect(lines[0]).toMatchObject({ top: 0, left: 100, text: 'ABC' });
  });

  it('rejects every narrower gap of a many-gap baseline with bounded work', () => {
    // Nine 10pt exclusions leave ten 10pt gaps; each 30pt word fits none.
    const obstacles = Array.from({ length: 9 }, (_, index) => squareObstacle(10 + index * 20, 20 + index * 20));
    const measured = measureParagraph(paragraph({ runs: [{ type: 'text', ...textRun('AAAAAA BBBBBB') }] }),
      layoutContext({ spaceBeforePt: 0 }),
      placement({ startYPt: 0, wrap: createFloatWrapOracle(obstacles) }), measurer, environment());
    expect(measured.lines.map(({ topYPt, layout }) => [topYPt, layout.xOffset, layout.availWidth]))
      .toEqual([[60, 0, 200]]);
  });

  // Review round 5: lines discovered by a pass are probed in that pass, so
  // convergence does not grow with the number of physical lines.
  it('converges a fixed-metric paragraph beside a tall exclusion in one pass sequence', () => {
    const tall = { ...squareObstacle(40, 460), yBottom: 400, imageH: 400 };
    const measured = measureParagraph(paragraph({ spaceBefore: 0, runs: [{ type: 'text', ...textRun('word '.repeat(48)) }] }),
      layoutContext({ spaceBeforePt: 0 }),
      placement({ startYPt: 0, availableWidthPt: 500, maximumYPt: 700,
        wrap: createFloatWrapOracle([tall], { xLeftPt: 0, xRightPt: 500, readingDirection: 'ltr' }) }),
      measurer, environment());
    expect(new Set(measured.lines.map(({ layout }) => layout.physicalLineIndex)).size).toBe(24);
    expect(measured.contentEndYPt).toBe(240);
  });

  it('keeps many lines under tall exclusions convergent and contained', () => {
    let seed = 0x1683;
    const random = (limit: number) => {
      seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0;
      return seed % limit;
    };
    for (let index = 0; index < 40; index += 1) {
      const left = 1 + random(60);
      const right = left + 20 + random(200 - left - 20);
      const bottom = 100 + random(400);
      const obstacle = random(2) === 0
        ? { ...squareObstacle(left, right), yBottom: bottom, imageH: bottom }
        : { ...tightObstacle(left, right), yBottom: bottom, imageH: bottom,
          wrapPolygon: [{ xPt: left, yPt: 0 }, { xPt: right, yPt: 0 }, { xPt: (left + right) / 2, yPt: bottom }] };
      const words = Array.from({ length: 20 + random(60) }, () => ['a', 'word', 'longer', 'W'][random(4)]).join(' ');
      const rtl = random(2) === 1;
      const runs: DocParagraph['runs'] = [{ type: 'text', ...textRun(words, rtl ? { rtl: true } : {}) }];
      const measured = measureParagraph(paragraph({ runs }), layoutContext({ spaceBeforePt: 0, baseRtl: rtl }),
        placement({ startYPt: 0, maximumYPt: 2000, wrap: createFloatWrapOracle([obstacle], gapReference(rtl)) }),
        measurer, environment());
      const lines = placedGeometry(measured);
      const label = `tall case ${index}: ${JSON.stringify({ obstacle, words, rtl })}`;
      expect(bandViolation(lines, obstacle, rtl), label).toBeNull();
      expect(physicalViolation(lines, rtl), label).toBeNull();
      expect(measuredTextSequence(measured).join('').replace(/\s/gu, ''), label).toBe(words.replace(/\s/gu, ''));
    }
  });

  it('fills gaps right to left in an RTL paragraph', () => {
    const lines = gapLines([{ type: 'text', ...textRun('אב גד הו זח טי כל', { rtl: true }) }],
      squareObstacle(60, 120), true);
    expect(lines[0]!.left).toBe(120);
    expect(lines[1]).toMatchObject({ top: 0, left: 0 });
  });
});

// Fixed seed keeps boundary failures reproducible without storing generated fixtures.
// Each construct also occurs after another cell, exercising flush/resume admission.
function atomicFloatPropertyCases() {
  let seed = 0x16831670;
  const random = (limit: number) => {
    seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0;
    return seed % limit;
  };
  const text = (value: string, options: Partial<DocxTextRun> = {}): DocParagraph['runs'][number] =>
    ({ type: 'text', ...textRun(value, options) });
  const constructs: Array<() => DocParagraph['runs']> = [
    () => [text('AA BB ', { fitTextVal: 400 + random(1600), fitTextId: 1 })],
    () => [text('AA', { fitTextVal: 800, fitTextId: 7 }), text('漢 字', { fitTextVal: 1600, fitTextId: 8 })],
    () => [text('AA ', { fitTextVal: 1600, fitTextId: 2 }), text('BB', { fitTextVal: 1600, fitTextId: 2, bold: true })],
    () => [text('漢字', { ruby: { text: 'かんじ', fontSizePt: 5 } })],
    () => [text('12', { eastAsianVert: true })],
    () => [text('合成文字', { eastAsianCombine: true })],
    () => [{ type: 'field', ...textRun(''), fieldType: 'unknown', instruction: 'EQ \\o(A,B)', fallbackText: 'AA BB' } as { type: 'field' } & FieldRun],
    () => [text('https://example.test/a-b/c', { hyperlink: 'https://example.test/a-b/c' })],
    () => [text('AA BB')],
    () => [text('AA‑BB')],
    () => [text('漢（字）文。')],
    () => [text('漢（'), text('字）文。')],
    () => [text('AA\tBB', { fitTextVal: 1000, fitTextId: 3 })],
    () => [text('AA\tBB CC')],
    () => [text('NEXT'), text('AA BB')],
    () => [text('ABCDEFGHIJKLMNOPQRSTUVWXYZABCDEFGH ')],
    () => [text('אב גד', { rtl: true, fitTextVal: 1200, fitTextId: 4 })],
    () => [text('אבג דהו זחט', { rtl: true })],
    () => [{ type: 'image', imagePath: 'test', mimeType: 'image/png', widthPt: 10 + random(80), heightPt: 10 } as { type: 'image' } & ImageRun],
    // Single graphemes, combining and emoji ZWJ sequences, joined one-grapheme
    // followers and a glyph wider than many gaps.
    () => [text('A')],
    () => [text('e\u0301')],
    () => [text('😀')],
    () => [text('👩\u200d💻')],
    () => [text('AB'), text('C')],
    () => [text('W', { fontSize: 40 })],
  ];
  return Array.from({ length: 304 }, (_, index) => {
    const rtl = random(2) === 1;
    const runs = [text('ABC '), ...constructs[index % constructs.length]!(), text(' NEXT'),
      ...constructs[random(constructs.length)]!()];
    // Every fourth float leaves a leading gap narrower than one glyph.
    const left = index % 4 === 0 ? 1 + random(4) : 5 + random(120);
    const right = left + 5 + random(195 - left);
    const float = random(3) === 0 ? tightObstacle(left, right) : squareObstacle(left, right);
    return { runs, rtl, float };
  });
}

interface PlacedLine {
  readonly top: number;
  readonly left: number;
  readonly width: number;
  readonly physical: number | undefined;
  /** Allocated advance of the physical line (its probed band height). */
  readonly height: number;
  readonly segments: readonly Readonly<{ text: string | null; width: number; visible: number }>[];
}

function placedGeometry(measured: ReturnType<typeof measureParagraph>): PlacedLine[] {
  return measured.lines.map(({ layout: line, topYPt, advancePt }) => ({
    top: topYPt, left: line.xOffset, width: line.availWidth, physical: line.physicalLineIndex,
    height: advancePt,
    segments: line.segments.map((segment) => {
      const isText = 'text' in segment;
      const fixed = isText && segment.fitTextRegionIndex !== undefined;
      // Collapsible line-end whitespace carries no ink; tabs are pen moves.
      const trailing = isText && !fixed
        ? [...segment.text].length - [...segment.text.replace(/[ 　]+$/u, '')].length : 0;
      const visible = 'isTab' in segment ? 0
        : segment.measuredWidth - (isText ? trailing * segment.fontSize * 0.5 : 0);
      return { text: isText ? segment.text : null, width: segment.measuredWidth, visible };
    }),
  }));
}

/** Every visible advance stays inside the paragraph band and outside the
 * exclusion, tested on the line band the solver probed. */
function bandViolation(lines: readonly PlacedLine[], obstacle: FloatRect, rtl: boolean): string | null {
  for (const line of lines) {
    const height = line.height;
    const total = line.segments.reduce((sum, segment) => sum + segment.width, 0);
    let pen = line.left + (rtl ? line.width - total : 0);
    for (const segment of line.segments) {
      const start = pen;
      const end = pen + segment.visible;
      pen += segment.width;
      if (!(segment.visible > 1e-9)) continue;
      // The full paragraph band keeps its established edge rules (§17.3.1.21
      // hanging punctuation, units wider than the column); a narrowed gap
      // owns its complete ink.
      const narrowed = line.width < 200 - 1e-9;
      // §17.3.1.21 permits one punctuation character past the paragraph edge
      // (never past an exclusion, which the check below covers).
      // The overflow is at the line end: right in LTR, left in RTL.
      const lineEnd = [...line.segments].reverse().find((item) => item.visible > 1e-9);
      const hanging = /\p{P}$/u.test(lineEnd?.text?.trimEnd() ?? '')
        && start >= -10 - 1e-9 && end <= 210 + 1e-9;
      const outside = rtl ? start < -1e-9 : end > 200 + 1e-9;
      if ((!rtl && start < -1e-9) || (rtl && end > 200 + 1e-9)
        || (narrowed && outside && !hanging)) return `band ${JSON.stringify(line)}`;
      const meets = obstacle.wrapPolygon
        ? polygonMeetsRect(obstacle.wrapPolygon, line.top, height, start + 1e-9, end - 1e-9)
        : line.top < obstacle.yBottom && line.top + height > obstacle.yTop
          && end > obstacle.xLeft + 1e-9 && start < obstacle.xRight - 1e-9;
      if (meets) return `exclusion ${JSON.stringify(line)}`;
    }
  }
  return null;
}

/** One physical line presents as one line: shared top, disjoint fragments in
 * reading order, and strictly increasing physical tops. */
function physicalViolation(lines: readonly PlacedLine[], rtl: boolean): string | null {
  for (let index = 1; index < lines.length; index += 1) {
    const previous = lines[index - 1]!;
    const line = lines[index]!;
    if (line.physical === previous.physical) {
      if (line.top !== previous.top) return `split physical line ${index}`;
      const ordered = rtl ? line.left + line.width <= previous.left + 1e-9
        : line.left >= previous.left + previous.width - 1e-9;
      if (!ordered) return `fragment order ${index}`;
    } else if (!(line.top > previous.top)) {
      return `physical tops ${index}`;
    }
  }
  return null;
}

it('keeps randomized content outside square and tight floats as physical lines', () => {
  for (const [index, item] of atomicFloatPropertyCases().entries()) {
    const measured = measureParagraph(paragraph({ runs: item.runs }),
      layoutContext({ spaceBeforePt: 0, baseRtl: item.rtl }),
      placement({ startYPt: 0, wrap: createFloatWrapOracle([item.float], gapReference(item.rtl)) }),
      measurer, environment());
    const lines = placedGeometry(measured);
    const label = `case ${index}: ${JSON.stringify(item)}`;
    expect(bandViolation(lines, item.float, item.rtl), label).toBeNull();
    expect(physicalViolation(lines, item.rtl), label).toBeNull();
  }
});

/** Classify the first difference from main using the model, never input IDs.
 * (a) main's own geometry leaves the band or meets the exclusion;
 * (b) the differing line exhibits a WORD_FLOAT_GAP_FLOW observation that
 *     main's widest-gap/one-inch/emergency rule contradicts. */
function classifyDifference(
  candidate: readonly PlacedLine[], previous: readonly PlacedLine[], obstacle: FloatRect, rtl: boolean,
): 'equal' | 'mainDefect' | 'gapOrder' | 'subInch' | 'continuation' | 'noGapEmergency' | null {
  const key = (line: PlacedLine) => JSON.stringify({ ...line, physical: undefined });
  const first = candidate.findIndex((line, index) =>
    previous[index] === undefined || key(line) !== key(previous[index]!));
  if (first < 0 && candidate.length === previous.length) return 'equal';
  if (bandViolation(previous, obstacle, rtl) !== null) return 'mainDefect';
  const index = first < 0 ? candidate.length : first;
  const changed = candidate[index];
  const old = previous[index];
  const narrowed = (line: PlacedLine | undefined) =>
    line !== undefined && line.top < obstacle.yBottom && line.width < 200 - 1e-9;
  const textOf = (line: PlacedLine | undefined) =>
    line?.segments.map((segment) => segment.text ?? '').join('') ?? '';
  // Main broke a word inside a narrowed gap at a non-opportunity; Word moves
  // that word below the object (controls longword-both-40).
  const isWordCharacter = (character: string | undefined) =>
    character !== undefined && /[\p{L}\p{N}:/.\-]/u.test(character)
      && !/[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}]/u.test(character);
  if (narrowed(old) && isWordCharacter([...textOf(old)].at(-1))
    && isWordCharacter([...textOf(previous[index + 1])][0])) return 'noGapEmergency';
  if (!narrowed(changed)) return null;
  if (candidate.filter((line) => line.top === changed!.top).length > 1) return 'continuation';
  if (old && old.top === changed!.top
    && (rtl ? changed!.left + changed!.width > old.left + old.width : changed!.left < old.left)) {
    return 'gapOrder';
  }
  if (changed!.width < 72 && (!old || old.top > changed!.top || old.left !== changed!.left)) return 'subInch';
  return null;
}

// Run with VRT_BASELINE_CHECKOUT naming the detached latest-main checkout.
// Ordinary unit runs retain the containment property; revision parity needs
// both production implementations and therefore explicitly skips without it.
const parityBaseline = process.env.VRT_BASELINE_CHECKOUT;
async function loadParityBaseline(): Promise<typeof import('./paragraph-measure.js')> {
  const BASELINE_PARAGRAPH_MEASURE = `${parityBaseline}/packages/docx/src/paragraph-measure.ts`;
  return await import(BASELINE_PARAGRAPH_MEASURE) as typeof import('./paragraph-measure.js');
}
it.skipIf(!parityBaseline)('classifies every float gap geometry difference from main', async () => {
  const baseline = await loadParityBaseline();
  const counts: Record<string, number> = {};
  for (const [index, item] of atomicFloatPropertyCases().entries()) {
    // Interior exclusions exercise multiple gaps; edge exclusions exercise sole
    // gaps on either side, including the historical one-inch admission boundary.
    const obstacles = item.float.wrapPolygon ? [item.float] : [item.float,
      squareObstacle(0, item.float.xRight), squareObstacle(item.float.xLeft, 200)];
    for (const obstacle of obstacles) {
      const doc = paragraph({ runs: item.runs });
      const context = layoutContext({ spaceBeforePt: 0, baseRtl: item.rtl });
      const render = (api: typeof baseline) => placedGeometry(api.measureParagraph(doc, context,
        placement({ startYPt: 0, wrap: api.createFloatWrapOracle([obstacle], gapReference(item.rtl)) }),
        measurer, environment()));
      const candidate = render({ measureParagraph, createFloatWrapOracle } as typeof baseline);
      const previous = render(baseline);
      const label = `case ${index}, obstacle ${obstacle.xLeft}..${obstacle.xRight}`;
      expect(bandViolation(candidate, obstacle, item.rtl), `${label} ${JSON.stringify({ runs: item.runs.map((run) => 'text' in run ? run.text : run.type), rtl: item.rtl, candidate })}`).toBeNull();
      const kind = classifyDifference(candidate, previous, obstacle, item.rtl);
      expect(kind, `${label}: unclassified diff: ${JSON.stringify({ runs: item.runs, rtl: item.rtl, candidate, previous })}`)
        .not.toBeNull();
      counts[kind!] = (counts[kind!] ?? 0) + 1;
    }
  }
  // Non-vacuity: the corpus exercises parity and the authorized differences.
  expect(counts.equal).toBeGreaterThan(0);
  expect(counts.mainDefect).toBeGreaterThan(0);
  expect((counts.gapOrder ?? 0) + (counts.continuation ?? 0) + (counts.subInch ?? 0)).toBeGreaterThan(0);
  console.info('Float gap parity classifications:', counts);
}, 60_000);

// Review round 6: mixed physical-line heights beside stacked exclusions. A
// short probe admitted a tall unit whose band then met the lower object; a
// tall probe excluded it again. Monotone per-line probe floors end that cycle.
function stackedObstacle(left: number, right: number, top: number, bottom: number): FloatRect {
  return { ...squareObstacle(left, right), imageY: top, imageH: bottom - top, yTop: top, yBottom: bottom };
}

function stackedLines(runs: DocParagraph['runs'], obstacles: readonly FloatRect[], rtl = false) {
  return measureParagraph(paragraph({ spaceBefore: 0, spaceAfter: 0, runs }),
    layoutContext({ spaceBeforePt: 0, spaceAfterPt: 0, baseRtl: rtl }),
    placement({ startYPt: 0, maximumYPt: 2000, wrap: createFloatWrapOracle(obstacles, gapReference(rtl)) }),
    measurer, environment());
}

it('converges mixed-height gap lines beside stacked square exclusions', () => {
  const runs: DocParagraph['runs'] = [10, 10, 24, 12, 24]
    .map((fontSize) => ({ type: 'text', ...textRun('word ', { fontSize }) }));
  const obstacles = [stackedObstacle(80, 120, 0, 115), stackedObstacle(40, 190, 45, 200)];
  const measured = stackedLines(runs, obstacles);
  const lines = placedGeometry(measured);
  for (const obstacle of obstacles) expect(bandViolation(lines, obstacle, false)).toBeNull();
  expect(physicalViolation(lines, false)).toBeNull();
  expect(measuredTextSequence(measured).join('')).toBe('word '.repeat(5));
});

it('keeps mixed sizes, pictures and ruby beside stacked exclusions convergent', () => {
  let seed = 0x16836;
  const random = (limit: number) => {
    seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0;
    return seed % limit;
  };
  for (let index = 0; index < 120; index += 1) {
    const runs: DocParagraph['runs'] = Array.from({ length: 4 + random(16) }, () => {
      const kind = random(6);
      if (kind === 0) {
        return { type: 'image', imagePath: 'test', mimeType: 'image/png',
          widthPt: 5 + random(40), heightPt: 5 + random(40) } as { type: 'image' } & ImageRun;
      }
      if (kind === 1) return { type: 'text', ...textRun('漢字', { ruby: { text: 'かんじ', fontSizePt: 5 + random(6) } }) };
      return { type: 'text', ...textRun(['a ', 'word ', 'longer ', 'W '][random(4)]!, { fontSize: [8, 10, 12, 18, 24, 36][random(6)]! }) };
    });
    const obstacles = Array.from({ length: 1 + random(3) }, () => {
      const left = random(150);
      const right = left + 10 + random(200 - left - 10);
      const top = random(150);
      return stackedObstacle(left, right, top, top + 10 + random(150));
    });
    const rtl = random(4) === 0;
    const label = `stacked case ${index}: ${JSON.stringify({ runs, obstacles, rtl })}`;
    const measured = stackedLines(runs, obstacles, rtl);
    const lines = placedGeometry(measured);
    for (const obstacle of obstacles) expect(bandViolation(lines, obstacle, rtl), label).toBeNull();
    expect(physicalViolation(lines, rtl), label).toBeNull();
  }
});

// With a single admitting gap (one exclusion at a paragraph edge leaving at
// least one inch) and words that fit it, the model and main coincide exactly.
it.skipIf(!parityBaseline)('matches main exactly when only one gap exists', async () => {
  const baseline = await loadParityBaseline();
  let seed = 0x1683a;
  const random = (limit: number) => {
    seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0;
    return seed % limit;
  };
  for (let index = 0; index < 120; index += 1) {
    const words = Array.from({ length: 5 + random(40) }, () => 'abcdefgh'.slice(0, 1 + random(8)));
    const runs: DocParagraph['runs'] = words.map((word) =>
      ({ type: 'text', ...textRun(`${word} `, { fontSize: [10, 12][random(2)]! }) }));
    const edge = 80 + random(40);
    const obstacle = random(2) === 0
      ? stackedObstacle(0, 200 - edge, 0, 20 + random(80))
      : stackedObstacle(edge, 200, 0, 20 + random(80));
    const render = (api: typeof baseline) => placedGeometry(api.measureParagraph(
      paragraph({ spaceBefore: 0, spaceAfter: 0, runs }), layoutContext({ spaceBeforePt: 0, spaceAfterPt: 0 }),
      placement({ startYPt: 0, maximumYPt: 2000, wrap: api.createFloatWrapOracle([obstacle], gapReference(false)) }),
      measurer, environment())).map((line) => ({ ...line, physical: undefined }));
    expect(render({ measureParagraph, createFloatWrapOracle } as typeof baseline),
      `one-gap case ${index}: ${JSON.stringify({ words, obstacle })}`).toEqual(render(baseline));
  }
}, 60_000);

// Review round 6 stress: a partition-dependent paragraph-wide allocation
// (ruby, grid, exact/atLeast spacing, inline pictures) beside a tall tight
// polygon and a stacked square. Cases 2485 and 4830 of this seeded generator
// cycled before the monotone advance rule; the generator is replayed to reach them.
it('converges partition-dependent allocations beside tall tight polygons', () => {
  let seed = 1683610;
  const random = (limit: number) => {
    seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0;
    return seed % limit;
  };
  for (let index = 0; index <= 4830; index += 1) {
    const width = [200, 300, 500][random(3)]!;
    const gap = [20, 40, 60, 80][random(4)]!;
    const count = 20 + random(150);
    const runs: DocParagraph['runs'] = Array.from({ length: count }, () =>
      ({ type: 'text', ...textRun('word ', { fontSize: [8, 10, 12, 18, 24][random(5)] }) }));
    if (index % 5 === 0) {
      runs.splice(random(count), 0,
        { type: 'image', imagePath: 'test', widthPt: 15, heightPt: 30 + random(80) } as { type: 'image' } & ImageRun);
    }
    const bottom = 60 + random(500);
    const polygon = [{ xPt: gap, yPt: 0 }, { xPt: width - gap, yPt: 0 },
      { xPt: width - gap - random(gap), yPt: bottom }, { xPt: gap + random(gap), yPt: bottom }];
    const obstacles: FloatRect[] = [{ ...squareObstacle(gap, width - gap), authoredWrap: 'tight',
      yBottom: bottom, imageH: bottom, wrapPolygon: polygon }];
    if (index % 3 === 0) obstacles.push(stackedObstacle(random(width - 30), width - 10, 40, 200));
    const context: { -readonly [K in keyof ParagraphLayoutContext]: ParagraphLayoutContext[K] } =
      layoutContext({ spaceBeforePt: 0 });
    let ruby = -1;
    if (index % 7 === 0) {
      context.hasRuby = true;
      ruby = random(count);
    }
    if (index % 4 === 0) context.lineSpacing = { rule: 'exact', value: 12 };
    if (index % 4 === 1) context.lineSpacing = { rule: 'atLeast', value: 12 };
    if (index % 4 === 2) context.lineGrid = { active: true, pitchPt: 20 };
    if (index !== 2485 && index !== 4830) continue;
    if (ruby >= 0) {
      const run = runs[ruby] as DocParagraph['runs'][number] & { ruby?: unknown };
      run.ruby = { text: 'ルビ', fontSizePt: 6, fontFamily: null, bold: false, italic: false, hpsRaisePt: 4, align: 'center' };
    }
    const measured = measureParagraph(paragraph({ spaceBefore: 0, runs }), context,
      placement({ startYPt: 0, availableWidthPt: width, maximumYPt: 700,
        wrap: createFloatWrapOracle(obstacles, { xLeftPt: 0, xRightPt: width, readingDirection: 'ltr' }) }),
      measurer, environment());
    const physical = measured.lines.filter((line, position) => position === 0
      || line.layout.physicalLineIndex !== measured.lines[position - 1]!.layout.physicalLineIndex);
    for (let position = 1; position < physical.length; position += 1) {
      expect(physical[position]!.topYPt + 1e-8)
        .toBeGreaterThanOrEqual(physical[position - 1]!.topYPt + physical[position - 1]!.advancePt);
    }
  }
});
