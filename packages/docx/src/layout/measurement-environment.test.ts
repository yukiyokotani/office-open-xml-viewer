import { describe, expect, it } from 'vitest';
import type { DocxDocumentModel } from '../types.js';
import type { ParagraphLayoutContext, SectionLayoutContext } from '../layout-context.js';
import type { LayoutSeg } from '../line-layout.js';
import type { BodyMeasurementContext } from './acquisition-context.js';
import {
  canonicalParagraphTextScaleEligible,
  docDefaultFontSizePt,
  gridForParagraphContext,
  paragraphMeasurementEnvironment,
  segmentEnvironmentOf,
  snapParaLineToGrid,
} from './measurement-environment.js';

describe('layout measurement environment', () => {
  it('admits only canonical body text without scale-aware compatibility paths', () => {
    const story = { story: 'body', containers: [], lineNumberingEligible: true } as const;
    const paragraphContext = { hasRuby: false, baseRtl: false };
    const paragraph = { alignment: 'left', numbering: null };
    const text = [{ text: 'plain', measuredWidth: 10 }] as LayoutSeg[];
    const eligible = (overrides: Readonly<{
      verticalCJK?: boolean;
      inFrame?: boolean;
      hasWrapContext?: boolean;
      paragraphContext?: typeof paragraphContext;
      paragraph?: typeof paragraph;
      segments?: LayoutSeg[];
    }> = {}) => canonicalParagraphTextScaleEligible(
      story,
      overrides.verticalCJK ?? false,
      overrides.inFrame ?? false,
      overrides.hasWrapContext ?? false,
      overrides.paragraphContext ?? paragraphContext,
      overrides.paragraph ?? paragraph,
      overrides.segments ?? text,
    );

    expect(eligible()).toBe(true);
    expect(canonicalParagraphTextScaleEligible(
      { ...story, containers: [{ kind: 'tableCell' }] },
      false, false, false, paragraphContext, paragraph, text,
    )).toBe(true);
    expect(canonicalParagraphTextScaleEligible(
      { story: 'header', containers: [], lineNumberingEligible: false },
      false, false, false, paragraphContext, paragraph, text,
    )).toBe(false);
    expect(eligible({ verticalCJK: true })).toBe(false);
    expect(eligible({ inFrame: true })).toBe(false);
    expect(eligible({ hasWrapContext: true })).toBe(false);
    expect(eligible({ paragraphContext: { hasRuby: true, baseRtl: false } })).toBe(false);
    expect(eligible({ paragraphContext: { hasRuby: false, baseRtl: true } })).toBe(false);
    expect(eligible({ paragraph: {
      alignment: 'left',
      numbering: { text: '1.' },
    } as unknown as typeof paragraph })).toBe(false);
    expect(eligible({ paragraph: { alignment: 'lowKashida', numbering: null } })).toBe(false);
    expect(eligible({ segments: [{ ...text[0], rtl: true }] as LayoutSeg[] })).toBe(false);
    expect(eligible({ segments: [{ isTab: true }] as LayoutSeg[] })).toBe(false);
    expect(eligible({
      segments: [{ ...text[0], math: true }] as unknown as LayoutSeg[],
    })).toBe(false);
    expect(eligible({ segments: [{ ...text[0], emphasisMark: 'dot' }] as LayoutSeg[] })).toBe(false);
  });

  it('resolves the parser-folded document font size before the first text run', () => {
    const document = (paragraph: object): DocxDocumentModel => ({
      section: {},
      body: [{ type: 'paragraph', runs: [], ...paragraph }],
      headers: {},
      footers: {},
    } as unknown as DocxDocumentModel);

    expect(docDefaultFontSizePt(document({
      defaultFontSize: 13,
      runs: [{ type: 'text', fontSize: 22 }],
    }))).toBe(13);
    expect(docDefaultFontSizePt(document({
      runs: [{ type: 'text', fontSize: 22 }],
    }))).toBe(22);
    expect(docDefaultFontSizePt(document({ runs: [] }))).toBe(10);
  });

  it('clears upright-vertical grouping only for the all-rotated btLr path', () => {
    const state = {
      pageIndex: 2,
      totalPages: 4,
      displayPageNumber: 3,
      pageNumberFormat: 'decimal',
      currentDateMs: 100,
      noteNumbers: new Map(),
      verticalCJK: true,
      verticalAllRotated: true,
      sectionLayout: { textDirection: 'btLr' },
      docEastAsian: true,
      layoutSettings: {
        characterSpacingControl: 'compressPunctuation',
        compat: {
          useFeLayout: true,
          balanceSingleByteDoubleByteWidth: true,
          lineWrapLikeWord6: true,
        },
      },
      resolvedLocalFonts: {},
      layoutServices: { text: {}, images: {}, math: {} },
      verticalGlyphMeasurement: {
        fingerprint: 'vertical:test',
        measureRunInkExtra: () => 0,
      },
    } as unknown as BodyMeasurementContext;

    expect(paragraphMeasurementEnvironment(state)).toMatchObject({
      pageIndex: 2,
      totalPages: 4,
      verticalCJK: false,
      pageWritingMode: 'vertical-rl',
      documentHasEastAsianText: true,
      useFeLayout: true,
      balanceSingleByteDoubleByteWidth: true,
      lineWrapLikeWord6: true,
    });
    const segments = segmentEnvironmentOf(state);
    expect(segments).not.toBe(state);
    expect(segments.verticalCJK).toBe(false);

    const upright = { ...state, verticalAllRotated: false } as BodyMeasurementContext;
    expect(segmentEnvironmentOf(upright)).toMatchObject({
      verticalCJK: true,
      characterSpacingControl: 'compressPunctuation',
      balanceSingleByteDoubleByteWidth: true,
      lineWrapLikeWord6: true,
    });
    expect(paragraphMeasurementEnvironment(upright).verticalCJK).toBe(true);
  });

  it('projects active grid axes and snaps only a positive line pitch', () => {
    const state = {
      sectionLayout: { grid: { kind: 'linesAndChars' } },
    } as unknown as Pick<BodyMeasurementContext, 'sectionLayout'>;
    const context = {
      lineGrid: { active: true, pitchPt: 12 },
      characterGrid: { active: true, kind: 'linesAndChars', deltaPt: 2, pitchPt: 14 },
    } as ParagraphLayoutContext;
    const grid = gridForParagraphContext(state, context);

    expect(grid).toEqual({
      type: 'linesAndChars',
      linePitchPt: 12,
      charSpacePt: 2,
      characterPitchPt: 14,
    });
    expect(snapParaLineToGrid(13, grid, 1)).toBe(24);
    expect(snapParaLineToGrid(8, grid, 1)).toBe(12);
    expect(snapParaLineToGrid(13, { ...grid, linePitchPt: 0 }, 1)).toBe(13);
    expect(snapParaLineToGrid(13, undefined, 1)).toBe(13);

    const noGridState = {
      sectionLayout: { grid: { kind: 'none' } } as SectionLayoutContext,
    } as Pick<BodyMeasurementContext, 'sectionLayout'>;
    expect(gridForParagraphContext(noGridState, {
      ...context,
      lineGrid: { active: false, pitchPt: null },
      characterGrid: { active: false, kind: null, pitchPt: null, deltaPt: 0 },
    })).toEqual({
      type: null,
      linePitchPt: null,
      charSpacePt: null,
      characterPitchPt: null,
    });
  });
});
