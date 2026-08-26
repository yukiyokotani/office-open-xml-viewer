import { describe, expect, it } from 'vitest';
import {
  defineCompatibilityRule,
  LINE_START_GAP_EPS_PT,
  WORD_FLOAT_DIFFERENT_PARAGRAPH_DISPLACEMENT,
  WORD_EMPTY_MARK_FLOAT_SIDE_GAP,
  WORD_PAGE_ANCHORED_TABLE_COLLISION_DEFERRAL,
  WORD_MIN_LINE_START_PT,
  WORD_SQUARE_LINE_START_ONE_INCH,
  WORD_SECTION_BTLR_TBRL_PAGE_FRAME,
  wordEmptyMarkMinimumStartWidthPx,
  wordMinLineStartPx,
} from './compatibility.js';
import {
  WORD_PRE_BREAK_ANCHOR_PARAGRAPH,
  WORD_PRE_BREAK_INLINE_DRAWING_GROUP,
  WORD_CONTINUOUS_SECTION_MARK_SPACING,
  WORD_CONTEXTUAL_SPACING_PER_SIDE,
  WORD_TRAILING_SPACE_AFTER_FIT_ADMISSION,
  WORD_VERTICAL_RL_FINAL_LINE_BASELINE_ADMISSION,
  WORD_TERMINAL_COLUMN_BREAK,
  wordFramePositionExtendsLineBox,
  wordFinalParagraphAdmissionExtentPt,
} from './body-pagination-compatibility.js';
import {
  WORD_FRAME_AUTO_WRAP_AROUND,
  WORD_LOWER_LAYER_SAME_PARAGRAPH_ANCHOR_COMPOSITION,
  WORD_PAGE_LEVEL_FLOAT_PRESCAN,
  WORD_PARAGRAPH_ANCHOR_PRE_SPACING_ORIGIN,
  WORD_VERTICAL_SECTION_PHYSICAL_HEADER_FOOTER,
  WORD_VERTICAL_SECTION_PHYSICAL_DRAWING_LAYER,
  WORD_ZERO_RELATIVE_SIZE_EXTENT_FALLBACK,
  wordPageLevelAnchorY,
  wordPreservesLowerLayerSameParagraphComposition,
  wordZeroRelativeSizeUsesExtent,
} from './anchor-compatibility.js';
import {
  WORD_AUTO_MULTIPLE_BASELINE_PIN,
  WORD_AUTHORED_CHARACTER_SPACING_PITCH_PRIORITY,
  WORD_BALANCED_CONSECUTIVE_SPACE_CELL,
  WORD_BALANCED_LINES_AND_CHARS_GRID_DELTA,
  WORD_CJK_BOTH_INTER_CHARACTER_EXPANSION,
  WORD_DEGENERATE_LINE_SPACING_SINGLE,
  WORD_DICTIONARY_SEA_ATOMIC_CHUNK,
  WORD_DICTIONARY_SEA_NATURAL_FIT,
  WORD_EAST_ASIAN_GRID_LINE_ALLOCATION,
  WORD_FAR_EAST_SINGLE_LINE_FACTOR,
  WORD_FIT_TEXT_INTER_CHARACTER_EXPANSION,
  WORD_FULL_WIDTH_CHARACTER_SPACING_SCOPE,
  WORD_GRID_AT_LEAST_TALL_LINE_UNSNAPPED,
  WORD_GRID_RIGHT_INDENT_PITCH_ALIGNMENT,
  WORD_HANGING_TAB_SAME_POSITION_PRECEDENCE,
  WORD_JAPANESE_PUNCTUATION_COMPRESSION_CELL,
  WORD_JUSTIFICATION_LEADING_INDENT_EXCLUSION,
  WORD_JUSTIFIED_CANDIDATE_SEPARATOR_FIT,
  WORD_MIXED_ANCHOR_VISIBLE_LINE_METRICS,
  WORD_MS_MINCHO_EMPTY_EAST_ASIAN_MARK_HEIGHT,
  WORD_NUMBERING_MARKER_OVERFLOW_TAB_ADVANCE,
  WORD_NUMBERING_MARKER_PARAGRAPH_MARK_FALLBACK,
  WORD_NUMBERING_SUFFIX_COINCIDENT_LIST_TAB,
  WORD_NUMERIC_DECIMAL_TAB_INFERENCE,
  WORD_OVERFLOW_PUNCTUATION_LANGUAGE_SETS,
  WORD_OVERLONG_TOKEN_EMERGENCY_BREAK,
  WORD_RUBY_PARAGRAPH_UNIFORM_LINE_ADVANCE,
  WORD_RTL_DECIMAL_TAB_PHYSICAL_ALIGNMENT,
  WORD_SNAP_TO_CHARS_EAST_ASIAN_CELL_FIT,
  WORD_SOURCE_RUN_SPACE_SEQUENCE,
  WORD_TAB_STOP_PAGE_EDGE_CLAMP,
  WORD_TABLE_CELL_IGNORES_GRID_RIGHT_INDENT_ADJUSTMENT,
  WORD_THAI_DISTRIBUTE_CLUSTER_POLICY,
  WORD_UNIFORM_RUN_POSITION_LEADING,
  wordAutoMultipleCenterBoxPx,
  wordBalancedConsecutiveSpaceCellApplies,
  wordBalancedLinesAndCharsGridDeltaFactor,
  wordBalancedSpaceCellAdjustmentApplies,
  wordDegenerateLineSpacingIsSingle,
  wordEastAsianGridLineCells,
  wordFarEastSingleLinePx,
  wordFirstJustifiedContentSegment,
  wordGridAtLeastLineHeightPx,
  wordGridRightIndentAdjustmentPt,
  wordUseFeLayoutParagraphMarkGridAdvancePx,
  wordUseFeLayoutInheritedGridHeightPx,
  wordCandidateFitWidthPx,
  wordDocumentCharacterCompressionApplies,
  wordJustifiedCandidateFitAllowancePx,
  wordJapanesePunctuationRetainedExtentPt,
  wordMsMinchoEmptyEastAsianMarkSingleLinePx,
  wordNumberingSuffixAcceptsCoincidentListTab,
  wordRubyUniformLineHeightPx,
  wordSnapToCharsEastAsianCellCount,
  wordSourceRunSpaceContinuesSequence,
  wordContainerAllowsGridRightIndentAdjustment,
  wordUniformRunPositionPaintPt,
  wordVisibleLineMetricPx,
} from './line-compatibility.js';
import {
  WORD_AUTO_TEXT_CONTRAST_EFFECTIVE_BACKGROUND,
  WORD_PARAGRAPH_BORDER_FLOW_RESERVATION,
  WORD_PARAGRAPH_SHADING_BORDER_BOX,
  WORD_RUN_DECORATION_JUSTIFIED_ADVANCE,
  WORD_SNAP_TO_CHARS_TERMINAL_UNDERLINE,
} from './paint-compatibility.js';
import {
  WORD_KASHIDA_FINAL_FORM_PRIORITY,
  WORD_NEUTRAL_SCRIPT_ATTACHMENT,
  WORD_RTL_COMPLEX_SCRIPT_EUROPEAN_DIGITS_AN,
  WORD_RTL_RUN_AMBIGUOUS_CLASS_OVERRIDE,
  WORD_VERTICAL_TU_CORNER_PLACEMENT,
  wordKashidaFinalFormApplies,
  wordNeutralAttachesToActiveScript,
  wordPreservesVerticalTuCorner,
  wordRtlAmbiguousCharacter,
} from './script-compatibility.js';
import {
  WORD_BOOK_FOLD_GUTTER_RIGHT_EDGE,
  WORD_CONTINUOUS_SECTION_PAGE_NUMBER_RESTART,
  WORD_DEFAULT_LINE_NUMBER_DISTANCE,
  WORD_SECTION_MARK_BLANK_PAGE_SUPPRESSION,
  WORD_DEFAULT_LINE_NUMBER_DISTANCE_PT,
  WORD_TRAILING_EMPTY_MARK_BASELINE_ADMISSION,
  wordBookFoldGutterEdge,
  wordContinuousSectionRestartDisplayNumber,
  wordLineNumberDistancePt,
  wordTrailingEmptyMarkAdmissionAllowancePt,
} from './section-compatibility.js';
import {
  WORD_AUTOFIT_EMPTY_PARAGRAPH_CONTENT_WIDTH,
  WORD_CELL_VERTICAL_ALIGNMENT_INK_BLOCK,
  WORD_EXACT_ROW_HEIGHT_BOTTOM_PADDING,
  WORD_EXACT_ROW_VERTICAL_CLIP_ONLY,
  WORD_AUTHORED_AUTO_ROW_HEIGHT_FLOOR,
  WORD_EFFECTIVE_FLOATING_TABLE_POSITIONING,
  WORD_FIRST_ROW_TABLE_EXCEPTION_SCOPE,
  WORD_OMITTED_ROW_HEIGHT_RULE_AT_LEAST,
  WORD_OVER_PAGE_CANT_SPLIT_CLIP,
  WORD_POSITIONED_TABLE_ADJACENCY_EXCLUSION,
  WORD_SPACED_CELL_INSIDE_BORDER_CONFLICT,
  WORD_TABLE_BORDER_LAYER_CASCADE,
  WORD_TABLE_BORDER_WEIGHT_PRECEDENCE,
  WORD_TABLE_BORDER_STYLE_PRECEDENCE,
  WORD_TABLE_CELL_SPACING_SCOPE_SHADOW,
  WORD_TABLE_INDENT_ALL_ALIGNMENTS,
  WORD_TABLE_MARGIN_SCOPE_SHADOW,
  WORD_TRAILING_STRUCTURAL_CELL_MARKER,
  WORD_VERTICAL_MERGE_TERMINAL_BORDER,
  WORD_VERTICAL_SECTION_UPRIGHT_BLOCK_TABLE,
  wordAlignedTableOriginPt,
  wordAutofitEmptyParagraphHasNoIntrinsicContent,
  wordAuthoredAutoRowHeightUsesFloor,
  wordAuthoredBorderParticipates,
  wordClipsOverPageCantSplitRow,
  wordDropsTrailingStructuralCellMarker,
  wordExactRowFloorPt,
  wordExactRowVerticalClipBounds,
  wordSpacedCellInsideBorderOverridesTable,
  wordTableBorderWeight,
  wordTableCellSpacingValuePt,
  wordTableMarginValuePt,
  wordTableRowHeightRule,
} from './table-compatibility.js';

describe('defineCompatibilityRule', () => {
  it('retains explicit Microsoft evidence as immutable data', () => {
    const rule = defineCompatibilityRule({
      id: 'word-example',
      evidence: { kind: 'microsoft-note', reference: '[MS-OE376] §2.1' },
      description: 'Synthetic compatibility boundary',
    });

    expect(rule.evidence).toEqual({ kind: 'microsoft-note', reference: '[MS-OE376] §2.1' });
    expect(Object.isFrozen(rule)).toBe(true);
    expect(Object.isFrozen(rule.evidence)).toBe(true);
  });

  it('retains a concrete regression-test reference without claiming an Office observation', () => {
    const rule = defineCompatibilityRule({
      id: 'regression-example',
      evidence: {
        kind: 'regression-test',
        reference: 'packages/docx/src/example.test.ts#named behavior',
      },
      description: 'Synthetic compatibility boundary',
    });

    expect(rule.evidence).toEqual({
      kind: 'regression-test',
      reference: 'packages/docx/src/example.test.ts#named behavior',
    });
  });

  it('rejects observation evidence without a reproducible fixture identity', () => {
    expect(() => defineCompatibilityRule({
      id: 'word-observation',
      evidence: {
        kind: 'office-observation',
        syntheticFixtureId: '',
        application: 'Word',
        version: 'current',
        platform: 'Windows',
      },
      description: 'Unreproducible observation',
    })).toThrow(/syntheticFixtureId/);
  });

  it('rejects unstructured evidence and unstable rule identities', () => {
    expect(() => defineCompatibilityRule({
      id: 'Not Stable',
      evidence: { kind: 'microsoft-note', reference: '[MS-OE376] §2.1' },
      description: 'Invalid identity',
    })).toThrow(/kebab-case/);
    expect(() => defineCompatibilityRule({
      id: 'bad-microsoft-reference',
      evidence: { kind: 'microsoft-note', reference: 'MS-OE376 somewhere' },
      description: 'Invalid Microsoft evidence',
    })).toThrow(/Microsoft specification section/);
    expect(() => defineCompatibilityRule({
      id: 'bad-regression-reference',
      evidence: { kind: 'regression-test', reference: 'a test somewhere' },
      description: 'Invalid regression evidence',
    })).toThrow(/path#test-title/);
  });

  it('records the approved Word section btLr page-frame difference', () => {
    expect(WORD_SECTION_BTLR_TBRL_PAGE_FRAME).toEqual({
      id: 'word-section-btlr-tbrl-page-frame',
      evidence: {
        kind: 'regression-test',
        reference: 'packages/docx/src/layout/coordinate-space.test.ts#maps Transitional text direction %s to %s',
      },
      description: expect.stringMatching(
        /Issue #988 comment 4950296007.*normative.*lr.*page frame.*glyph orientation.*paint-owned/i,
      ),
    });
    expect(Object.isFrozen(WORD_SECTION_BTLR_TBRL_PAGE_FRAME)).toBe(true);
    expect(Object.isFrozen(WORD_SECTION_BTLR_TBRL_PAGE_FRAME.evidence)).toBe(true);
  });

  it('pins every regression-test reference to a concrete repository test', () => {
    const rules = [
      WORD_SECTION_BTLR_TBRL_PAGE_FRAME,
      WORD_TERMINAL_COLUMN_BREAK,
      WORD_PRE_BREAK_ANCHOR_PARAGRAPH,
      WORD_PRE_BREAK_INLINE_DRAWING_GROUP,
      WORD_CONTINUOUS_SECTION_MARK_SPACING,
      WORD_CONTEXTUAL_SPACING_PER_SIDE,
    ];

    expect(rules.map((rule) => (
      rule.evidence.kind === 'regression-test' ? rule.evidence.reference : null
    ))).toEqual([
      'packages/docx/src/layout/coordinate-space.test.ts#maps Transitional text direction %s to %s',
      'packages/docx/src/pagination.test.ts#ignores a terminal last-column break before a hard page boundary',
      'packages/docx/src/pagination.test.ts#does not push an anchor-only pre-break paragraph to a new page just for its empty mark',
      'packages/docx/src/pagination.test.ts#moves a preceding image with its pre-break callout when the pair only fits fresh',
      'packages/docx/src/body-layout-input.test.ts#projects mutually exclusive collapsed-mark and drop-previous-after roles',
      'packages/docx/src/contextual-spacing-body-paint.test.ts#paints the adjudicated six-case gap table',
    ]);
  });

  it('limits positioned-run line-box suppression to observed drop caps', () => {
    expect(wordFramePositionExtendsLineBox('none')).toBe(true);
    expect(wordFramePositionExtendsLineBox('drop')).toBe(false);
    expect(wordFramePositionExtendsLineBox('margin')).toBe(false);
  });
});

describe('Word table AutoFit observations', () => {
  it('records the empty-paragraph content-width boundary', () => {
    expect(WORD_AUTOFIT_EMPTY_PARAGRAPH_CONTENT_WIDTH.evidence).toEqual({
      kind: 'office-observation',
      syntheticFixtureId: 'autofit-empty-paragraph-boundary-matrix',
      application: 'Microsoft Word',
      version: '16.111.1',
      platform: 'macOS 26.5.2',
    });
    expect(WORD_AUTOFIT_EMPTY_PARAGRAPH_CONTENT_WIDTH.description)
      .not.toMatch(/sample|private|\.docx|\.pdf/i);
    expect(wordAutofitEmptyParagraphHasNoIntrinsicContent({
      runs: [], numbering: null,
    })).toBe(true);
    expect(wordAutofitEmptyParagraphHasNoIntrinsicContent({
      runs: [{ type: 'text', text: ' ' }] as never, numbering: null,
    })).toBe(false);
    expect(wordAutofitEmptyParagraphHasNoIntrinsicContent({
      runs: [], numbering: { numId: 1, level: 0 } as never,
    })).toBe(false);
  });
});

describe('float compatibility evidence', () => {
  it('keeps the measured square line-start threshold behind one named rule', () => {
    expect(WORD_SQUARE_LINE_START_ONE_INCH.evidence).toMatchObject({
      kind: 'regression-test',
    });
    expect(WORD_MIN_LINE_START_PT).toBe(72);
    expect(LINE_START_GAP_EPS_PT).toBe(0.05);
    expect(wordMinLineStartPx(1)).toBeCloseTo(71.95, 10);
  });

  it('names the established different-paragraph displacement policy', () => {
    expect(WORD_FLOAT_DIFFERENT_PARAGRAPH_DISPLACEMENT).toMatchObject({
      id: 'word-float-different-paragraph-displacement',
      evidence: { kind: 'regression-test' },
    });
  });

  it('names absolute floating-table collision deferral as compatibility behavior', () => {
    expect(WORD_PAGE_ANCHORED_TABLE_COLLISION_DEFERRAL).toMatchObject({
      id: 'word-page-anchored-table-collision-deferral',
      evidence: { kind: 'regression-test' },
    });
  });
});

describe('layout compatibility inventory', () => {
  it("uses Word's observed 18pt line-number distance only when omitted", () => {
    expect(WORD_DEFAULT_LINE_NUMBER_DISTANCE_PT).toBe(18);
    expect(wordLineNumberDistancePt(undefined)).toBe(18);
    expect(wordLineNumberDistancePt(null)).toBe(18);
    expect(wordLineNumberDistancePt(0)).toBe(0);
    expect(wordLineNumberDistancePt(12)).toBe(12);
  });

  it('keeps each layout observation behind explicit immutable evidence', () => {
    const rules = [
      WORD_ZERO_RELATIVE_SIZE_EXTENT_FALLBACK,
      WORD_DEFAULT_LINE_NUMBER_DISTANCE,
      WORD_CONTINUOUS_SECTION_PAGE_NUMBER_RESTART,
      WORD_TRAILING_EMPTY_MARK_BASELINE_ADMISSION,
      WORD_SECTION_MARK_BLANK_PAGE_SUPPRESSION,
      WORD_BOOK_FOLD_GUTTER_RIGHT_EDGE,
      WORD_EXACT_ROW_HEIGHT_BOTTOM_PADDING,
      WORD_TABLE_BORDER_LAYER_CASCADE,
      WORD_SPACED_CELL_INSIDE_BORDER_CONFLICT,
      WORD_TABLE_INDENT_ALL_ALIGNMENTS,
      WORD_EXACT_ROW_VERTICAL_CLIP_ONLY,
      WORD_OVER_PAGE_CANT_SPLIT_CLIP,
      WORD_POSITIONED_TABLE_ADJACENCY_EXCLUSION,
      WORD_EAST_ASIAN_GRID_LINE_ALLOCATION,
      WORD_SNAP_TO_CHARS_EAST_ASIAN_CELL_FIT,
      WORD_TABLE_CELL_IGNORES_GRID_RIGHT_INDENT_ADJUSTMENT,
      WORD_GRID_RIGHT_INDENT_PITCH_ALIGNMENT,
      WORD_HANGING_TAB_SAME_POSITION_PRECEDENCE,
      WORD_RTL_DECIMAL_TAB_PHYSICAL_ALIGNMENT,
      WORD_GRID_AT_LEAST_TALL_LINE_UNSNAPPED,
      WORD_DEGENERATE_LINE_SPACING_SINGLE,
      WORD_AUTO_MULTIPLE_BASELINE_PIN,
      WORD_MIXED_ANCHOR_VISIBLE_LINE_METRICS,
      WORD_JUSTIFICATION_LEADING_INDENT_EXCLUSION,
      WORD_JUSTIFIED_CANDIDATE_SEPARATOR_FIT,
      WORD_OVERFLOW_PUNCTUATION_LANGUAGE_SETS,
      WORD_FULL_WIDTH_CHARACTER_SPACING_SCOPE,
      WORD_AUTHORED_CHARACTER_SPACING_PITCH_PRIORITY,
      WORD_RUBY_PARAGRAPH_UNIFORM_LINE_ADVANCE,
      WORD_FIT_TEXT_INTER_CHARACTER_EXPANSION,
      WORD_CJK_BOTH_INTER_CHARACTER_EXPANSION,
      WORD_THAI_DISTRIBUTE_CLUSTER_POLICY,
      WORD_NUMERIC_DECIMAL_TAB_INFERENCE,
      WORD_NUMBERING_MARKER_OVERFLOW_TAB_ADVANCE,
      WORD_NUMBERING_MARKER_PARAGRAPH_MARK_FALLBACK,
      WORD_NUMBERING_SUFFIX_COINCIDENT_LIST_TAB,
      WORD_TAB_STOP_PAGE_EDGE_CLAMP,
      WORD_DICTIONARY_SEA_NATURAL_FIT,
      WORD_DICTIONARY_SEA_ATOMIC_CHUNK,
      WORD_OVERLONG_TOKEN_EMERGENCY_BREAK,
      WORD_UNIFORM_RUN_POSITION_LEADING,
      WORD_NEUTRAL_SCRIPT_ATTACHMENT,
      WORD_RTL_RUN_AMBIGUOUS_CLASS_OVERRIDE,
      WORD_RTL_COMPLEX_SCRIPT_EUROPEAN_DIGITS_AN,
      WORD_KASHIDA_FINAL_FORM_PRIORITY,
      WORD_VERTICAL_TU_CORNER_PLACEMENT,
      WORD_PARAGRAPH_SHADING_BORDER_BOX,
      WORD_AUTO_TEXT_CONTRAST_EFFECTIVE_BACKGROUND,
      WORD_RUN_DECORATION_JUSTIFIED_ADVANCE,
      WORD_SNAP_TO_CHARS_TERMINAL_UNDERLINE,
      WORD_PARAGRAPH_BORDER_FLOW_RESERVATION,
      WORD_EMPTY_MARK_FLOAT_SIDE_GAP,
      WORD_VERTICAL_SECTION_PHYSICAL_DRAWING_LAYER,
      WORD_PAGE_LEVEL_FLOAT_PRESCAN,
      WORD_PARAGRAPH_ANCHOR_PRE_SPACING_ORIGIN,
      WORD_VERTICAL_SECTION_PHYSICAL_HEADER_FOOTER,
      WORD_FRAME_AUTO_WRAP_AROUND,
      WORD_LOWER_LAYER_SAME_PARAGRAPH_ANCHOR_COMPOSITION,
      WORD_TABLE_BORDER_WEIGHT_PRECEDENCE,
      WORD_OMITTED_ROW_HEIGHT_RULE_AT_LEAST,
      WORD_AUTHORED_AUTO_ROW_HEIGHT_FLOOR,
      WORD_EFFECTIVE_FLOATING_TABLE_POSITIONING,
      WORD_TABLE_CELL_SPACING_SCOPE_SHADOW,
      WORD_TABLE_MARGIN_SCOPE_SHADOW,
      WORD_FIRST_ROW_TABLE_EXCEPTION_SCOPE,
      WORD_TRAILING_STRUCTURAL_CELL_MARKER,
      WORD_CELL_VERTICAL_ALIGNMENT_INK_BLOCK,
      WORD_VERTICAL_MERGE_TERMINAL_BORDER,
      WORD_VERTICAL_SECTION_UPRIGHT_BLOCK_TABLE,
      WORD_TRAILING_SPACE_AFTER_FIT_ADMISSION,
      WORD_VERTICAL_RL_FINAL_LINE_BASELINE_ADMISSION,
    ];

    expect(new Set(rules.map((rule) => rule.id)).size).toBe(rules.length);
    expect(rules.every(Object.isFrozen)).toBe(true);
    expect(rules.every((rule) => Object.isFrozen(rule.evidence))).toBe(true);
  });

  it('isolates uniform w:position leading without changing mixed-run displacement', () => {
    expect(wordUniformRunPositionPaintPt(6, 6)).toBe(3);
    expect(wordUniformRunPositionPaintPt(-6, -6)).toBe(-3);
    expect(wordUniformRunPositionPaintPt(6, 0)).toBe(6);
    expect(WORD_UNIFORM_RUN_POSITION_LEADING.description)
      .not.toMatch(/sample|private|\.docx|\.pdf/i);
  });

  it('records the anonymized Word observation for lower-layer anchor composition', () => {
    expect(WORD_LOWER_LAYER_SAME_PARAGRAPH_ANCHOR_COMPOSITION.evidence).toEqual({
      kind: 'office-observation',
      syntheticFixtureId: 'lower-layer-same-paragraph-anchor-composition',
      application: 'Microsoft Word',
      version: '16.111.1',
      platform: 'macOS 26.5.2',
    });
    expect(WORD_LOWER_LAYER_SAME_PARAGRAPH_ANCHOR_COMPOSITION.description)
      .not.toMatch(/sample|private|\.docx|\.pdf/i);
  });

  it('separates normative punctuation eligibility from the observed compression amount', () => {
    expect(WORD_OVERFLOW_PUNCTUATION_LANGUAGE_SETS.evidence).toEqual({
      kind: 'microsoft-note',
      reference: '[MS-OE376] §2.1.56',
    });
    expect(WORD_FULL_WIDTH_CHARACTER_SPACING_SCOPE.evidence).toEqual({
      kind: 'microsoft-note',
      reference: '[MS-OE376] §2.1.562',
    });
    expect(WORD_FULL_WIDTH_CHARACTER_SPACING_SCOPE.description)
      .toMatch(/only which characters are eligible.*not define.*compression amount/i);
    expect(WORD_JAPANESE_PUNCTUATION_COMPRESSION_CELL.evidence).toEqual({
      kind: 'office-observation',
      syntheticFixtureId: 'japanese-fullwidth-punctuation-compression-cell',
      application: 'Microsoft Word',
      version: '16.111.1',
      platform: 'macOS 26.5.2',
    });
    expect(WORD_JAPANESE_PUNCTUATION_COMPRESSION_CELL.description)
      .not.toMatch(/sample|private|\.docx|\.pdf/i);
    expect(wordJapanesePunctuationRetainedExtentPt({
      punctuationAdvancePt: 8,
      punctuationInkEndPt: 1,
      ideographicCellAdvancePt: 10,
    })).toBe(8);
    expect(wordJapanesePunctuationRetainedExtentPt({
      punctuationAdvancePt: 10,
      punctuationInkEndPt: 1,
      ideographicCellAdvancePt: 10,
    })).toBe(5);
    expect(wordJapanesePunctuationRetainedExtentPt({
      punctuationAdvancePt: 10,
      punctuationInkEndPt: 6,
      ideographicCellAdvancePt: 10,
    })).toBe(6);
    expect(WORD_AUTHORED_CHARACTER_SPACING_PITCH_PRIORITY.evidence).toEqual({
      kind: 'office-observation',
      syntheticFixtureId: 'authored-character-spacing-punctuation-pitch',
      application: 'Microsoft Word',
      version: '16.111.1',
      platform: 'macOS 26.5.2',
    });
    expect(WORD_AUTHORED_CHARACTER_SPACING_PITCH_PRIORITY.description)
      .not.toMatch(/sample|private|\.docx|\.pdf/i);
    expect(WORD_SOURCE_RUN_SPACE_SEQUENCE.evidence).toEqual({
      kind: 'office-observation',
      syntheticFixtureId: 'source-run-space-sequence-wrap-matrix',
      application: 'Microsoft Word',
      version: '16.111.1',
      platform: 'macOS 26.5.2',
    });
    expect(wordSourceRunSpaceContinuesSequence('deadline ', ' ')).toBe(true);
    expect(wordSourceRunSpaceContinuesSequence('deadline', ' ')).toBe(false);
    expect(WORD_BALANCED_CONSECUTIVE_SPACE_CELL.evidence).toEqual({
      kind: 'office-observation',
      syntheticFixtureId: 'single-double-byte-width-space-grid-matrix',
      application: 'Microsoft Word',
      version: '16.111.1',
      platform: 'macOS 26.5.2',
    });
    expect(WORD_BALANCED_CONSECUTIVE_SPACE_CELL.description)
      .not.toMatch(/sample|private|\.docx|\.pdf/i);
    expect(wordBalancedConsecutiveSpaceCellApplies(1)).toBe(false);
    expect(wordBalancedConsecutiveSpaceCellApplies(2)).toBe(true);
    expect(wordBalancedConsecutiveSpaceCellApplies(8)).toBe(true);
    expect(wordBalancedSpaceCellAdjustmentApplies(undefined)).toBe(true);
    expect(wordBalancedSpaceCellAdjustmentApplies('linesAndChars')).toBe(true);
    expect(wordBalancedSpaceCellAdjustmentApplies('snapToChars')).toBe(false);
    expect(WORD_BALANCED_LINES_AND_CHARS_GRID_DELTA.evidence).toEqual({
      kind: 'office-observation',
      syntheticFixtureId: 'single-double-byte-width-grid-observation-matrix',
      application: 'Microsoft Word',
      version: '16.111.1',
      platform: 'macOS 26.5.2',
    });
    expect(wordBalancedLinesAndCharsGridDeltaFactor('A', 'ascii')).toBe(0.5);
    expect(wordBalancedLinesAndCharsGridDeltaFactor(' ', 'ascii')).toBe(0.5);
    expect(wordBalancedLinesAndCharsGridDeltaFactor('\u3000', 'eastAsia')).toBe(0.5);
    expect(wordBalancedLinesAndCharsGridDeltaFactor('日', 'eastAsia')).toBe(1);
    expect(wordBalancedLinesAndCharsGridDeltaFactor('Ａ', 'eastAsia')).toBe(1);
    expect(WORD_TABLE_CELL_IGNORES_GRID_RIGHT_INDENT_ADJUSTMENT.evidence).toEqual({
      kind: 'office-observation',
      syntheticFixtureId: 'table-cell-adjust-right-indent-width-position-matrix',
      application: 'Microsoft Word',
      version: '16.111.1',
      platform: 'macOS 26.5.2',
    });
    expect(wordContainerAllowsGridRightIndentAdjustment(false)).toBe(true);
    expect(wordContainerAllowsGridRightIndentAdjustment(true)).toBe(false);
    expect(WORD_GRID_RIGHT_INDENT_PITCH_ALIGNMENT.evidence).toEqual({
      kind: 'office-observation',
      syntheticFixtureId: 'grid-right-indent-character-pitch-boundary-matrix',
      application: 'Microsoft Word',
      version: '16.111.1',
      platform: 'macOS 26.5.2',
    });
    expect(wordGridRightIndentAdjustmentPt(81, 9)).toBe(0);
    expect(wordGridRightIndentAdjustmentPt(85.9, 9)).toBeCloseTo(4.9, 9);
    expect(wordGridRightIndentAdjustmentPt(80.999, 9)).toBeCloseTo(8.999, 9);
    expect(wordGridRightIndentAdjustmentPt(81.001, 9)).toBeCloseTo(0.001, 9);
    expect(wordGridRightIndentAdjustmentPt(0, 9)).toBe(0);
    expect(wordGridRightIndentAdjustmentPt(85.9, 0)).toBe(0);
    expect(WORD_SNAP_TO_CHARS_EAST_ASIAN_CELL_FIT.evidence).toEqual({
      kind: 'office-observation',
      syntheticFixtureId: 'snap-to-chars-east-asian-cell-fit-matrix',
      application: 'Microsoft Word',
      version: '16.111.1',
      platform: 'macOS 26.5.2',
    });
    expect(wordSnapToCharsEastAsianCellCount(8, 10)).toBe(1);
    expect(wordSnapToCharsEastAsianCellCount(10, 10)).toBe(1);
    expect(wordSnapToCharsEastAsianCellCount(10.01, 10)).toBe(2);
    expect(WORD_SNAP_TO_CHARS_EAST_ASIAN_CELL_FIT.description)
      .not.toMatch(/sample|private|\.docx|\.pdf/i);
    expect(wordDocumentCharacterCompressionApplies(undefined)).toBe(true);
    expect(wordDocumentCharacterCompressionApplies(0)).toBe(true);
    expect(wordDocumentCharacterCompressionApplies(0.4)).toBe(false);
    expect(wordDocumentCharacterCompressionApplies(-0.5)).toBe(true);
    expect(WORD_NUMBERING_MARKER_PARAGRAPH_MARK_FALLBACK.evidence).toEqual({
      kind: 'office-observation',
      syntheticFixtureId: 'numbering-marker-paragraph-mark-formatting',
      application: 'Microsoft Word',
      version: '16.111.1',
      platform: 'macOS 26.5.2',
    });
    expect(WORD_NUMBERING_MARKER_PARAGRAPH_MARK_FALLBACK.description)
      .not.toMatch(/sample|private|\.docx|\.pdf/i);
  });

  it('scopes the observed MS Mincho height to empty East-Asian marks', () => {
    expect(WORD_MS_MINCHO_EMPTY_EAST_ASIAN_MARK_HEIGHT.evidence).toEqual({
      kind: 'office-observation',
      syntheticFixtureId: 'ms-mincho-empty-east-asian-paragraph-mark',
      application: 'Microsoft Word',
      version: '16.111.1',
      platform: 'macOS 26.5.2',
    });
    expect(WORD_MS_MINCHO_EMPTY_EAST_ASIAN_MARK_HEIGHT.description)
      .not.toMatch(/sample|private|\.docx|\.pdf/i);
    expect(wordMsMinchoEmptyEastAsianMarkSingleLinePx('MS Mincho', 12, true)).toBeCloseTo(15.6, 5);
    expect(wordMsMinchoEmptyEastAsianMarkSingleLinePx('ＭＳ 明朝', 12, true)).toBeCloseTo(15.6, 5);
    expect(wordMsMinchoEmptyEastAsianMarkSingleLinePx('MS Mincho', 12, false)).toBe(0);
    expect(wordMsMinchoEmptyEastAsianMarkSingleLinePx('MS PMincho', 12, true)).toBe(0);
  });

  it('records the anonymized Word observation for vertical final-line admission', () => {
    expect(WORD_VERTICAL_RL_FINAL_LINE_BASELINE_ADMISSION.evidence).toEqual({
      kind: 'office-observation',
      syntheticFixtureId: 'vertical-rl-final-line-baseline-admission',
      application: 'Microsoft Word',
      version: '16.111.1',
      platform: 'macOS 26.5.2',
    });
    expect(WORD_VERTICAL_RL_FINAL_LINE_BASELINE_ADMISSION.description)
      .not.toMatch(/sample|private|\.docx|\.pdf/i);
  });

  it('pins behavior-preserving compatibility decision helpers', () => {
    expect(wordZeroRelativeSizeUsesExtent(0)).toBe(true);
    expect(wordZeroRelativeSizeUsesExtent(-0.1)).toBe(false);
    expect(wordPreservesLowerLayerSameParagraphComposition('page', 1, 2)).toBe(true);
    expect(wordPreservesLowerLayerSameParagraphComposition('page', 2, 1)).toBe(false);
    expect(wordPreservesLowerLayerSameParagraphComposition('page', 1, undefined)).toBe(false);
    expect(wordPreservesLowerLayerSameParagraphComposition('host', 1, 2)).toBe(false);
    expect(wordExactRowFloorPt(12, [2, 4, 3])).toBe(16);
    expect(wordAuthoredBorderParticipates('none')).toBe(false);
    expect(wordAuthoredBorderParticipates('nil')).toBe(true);
    expect(wordAlignedTableOriginPt(100, 8, false)).toBe(108);
    expect(wordAlignedTableOriginPt(100, 8, true)).toBe(92);
    expect(wordSpacedCellInsideBorderOverridesTable({
      spacingPt: 0,
      directStyle: null,
      conditionalInsideStyle: 'single',
    })).toBe(false);
    expect(wordSpacedCellInsideBorderOverridesTable({
      spacingPt: 1,
      directStyle: 'none',
      conditionalInsideStyle: 'single',
    })).toBe(true);
    expect(wordSpacedCellInsideBorderOverridesTable({
      spacingPt: 1,
      directStyle: 'nil',
      conditionalInsideStyle: 'single',
    })).toBe(false);
    expect(wordExactRowVerticalClipBounds(
      { xPt: 20, yPt: 30, widthPt: 40, heightPt: 50 },
      { xPt: 5, yPt: 10, widthPt: 100, heightPt: 200 },
    )).toEqual({ xPt: 5, yPt: 30, widthPt: 100, heightPt: 50 });
    expect(wordExactRowVerticalClipBounds(
      { xPt: -5, yPt: 30, widthPt: 40, heightPt: 50 },
      { xPt: 5, yPt: 10, widthPt: 100, heightPt: 200 },
    )).toEqual({ xPt: -5, yPt: 30, widthPt: 110, heightPt: 50 });
    expect(wordClipsOverPageCantSplitRow({
      compatibility: 'word',
      availableHeightPt: 99.99995,
      freshPageHeightPt: 100,
      epsilonPt: 0.0001,
    })).toBe(true);
    expect(wordClipsOverPageCantSplitRow({
      compatibility: 'standard',
      availableHeightPt: 100,
      freshPageHeightPt: 100,
      epsilonPt: 0.0001,
    })).toBe(false);
    expect(wordContinuousSectionRestartDisplayNumber(4, 8, 6)).toBe(6);
    expect(wordBookFoldGutterEdge()).toBe('right');
    const trailingMark = {
      hasContinuationBoundary: false,
      inkless: true,
      undecorated: true,
      keepNext: false,
      markReservePt: 0,
      pageBottomIsUnreserved: true,
      physicalRegionBottomIsActive: true,
      hasFollowingInk: true,
      followsNextPageSectionBoundary: false,
      markExtentPt: 12,
      markBelowBaselinePt: 3,
    };
    expect(wordTrailingEmptyMarkAdmissionAllowancePt(trailingMark)).toBe(3);
    expect(wordTrailingEmptyMarkAdmissionAllowancePt({
      ...trailingMark,
      hasFollowingInk: false,
    })).toBe(0);
    expect(wordTrailingEmptyMarkAdmissionAllowancePt({
      ...trailingMark,
      hasFollowingInk: false,
      followsNextPageSectionBoundary: true,
    })).toBe(12);
    expect(WORD_SECTION_MARK_BLANK_PAGE_SUPPRESSION.evidence).toEqual({
      kind: 'office-observation',
      syntheticFixtureId: 'next-page-section-mark-bottom-edge-admission',
      application: 'Microsoft Word',
      version: '16.111.1',
      platform: 'macOS 26.5.2',
    });
    expect(wordFinalParagraphAdmissionExtentPt({
      advancePt: 30,
      retainedSpaceAfterPt: 10,
      authoredSpaceAfterPt: 10,
    })).toBe(20);
    expect(wordNumberingSuffixAcceptsCoincidentListTab(
      20,
      { pos: 20, alignment: 'num' },
    )).toBe(true);
    expect(wordNumberingSuffixAcceptsCoincidentListTab(
      20,
      { pos: 20, alignment: 'left' },
    )).toBe(false);
    expect(wordCandidateFitWidthPx({
      widthPx: 60,
      trailingSpacePx: 12,
      lineWillJustify: true,
    })).toBe(60);
    expect(wordCandidateFitWidthPx({
      widthPx: 60,
      trailingSpacePx: 12,
      lineWillJustify: true,
      wrapNarrowed: true,
    })).toBe(48);
    expect(wordJustifiedCandidateFitAllowancePx({
      biasBudgetPx: 5.25,
      resolvedMeasurementRouteCount: 1,
    })).toBe(5.25);
    expect(wordJustifiedCandidateFitAllowancePx({
      biasBudgetPx: 5.25,
      resolvedMeasurementRouteCount: 2,
    })).toBe(0);
    expect(wordCandidateFitWidthPx({
      widthPx: 60,
      trailingSpacePx: 12,
      lineWillJustify: false,
    })).toBe(48);
  });

  it('pins East Asian grid allocation and the untabled Far East metric factor', () => {
    expect(WORD_FAR_EAST_SINGLE_LINE_FACTOR).toBe(1.3);
    expect(wordEastAsianGridLineCells(17.19, 18)).toBe(1);
    expect(wordEastAsianGridLineCells(20.06, 18)).toBe(2);
    expect(wordEastAsianGridLineCells(0, 18)).toBe(1);
    expect(wordFarEastSingleLinePx(22, 10)).toBe(22);
    expect(wordFarEastSingleLinePx(0, 10)).toBe(13);
    expect(wordUseFeLayoutInheritedGridHeightPx(36, 18, 1.15)).toBe(36);
    expect(wordUseFeLayoutInheritedGridHeightPx(18, 18, 1.15)).toBeCloseTo(20.7, 12);
  });

  it('keeps neutral characters attached to the active script slice', () => {
    expect(wordNeutralAttachesToActiveScript(' ')).toBe(true);
    expect(wordNeutralAttachesToActiveScript('1')).toBe(true);
    expect(wordNeutralAttachesToActiveScript('.')).toBe(true);
    expect(wordNeutralAttachesToActiveScript('A')).toBe(false);
    expect(wordNeutralAttachesToActiveScript('ش')).toBe(false);
  });

  it('pins paint and legacy-text compatibility helper branches', () => {
    expect(wordEmptyMarkMinimumStartWidthPx(12, 2)).toBe(24);
    expect(wordPageLevelAnchorY('margin', false)).toBe(true);
    expect(wordPageLevelAnchorY('line', false)).toBe(false);
    expect(wordPageLevelAnchorY(null, false)).toBe(true);
    expect(wordPageLevelAnchorY(undefined, true)).toBe(false);
    expect(wordGridAtLeastLineHeightPx(28, 18, 18)).toBe(28);
    const markGridAdvance = (
      lineSpacing: Parameters<typeof wordUseFeLayoutParagraphMarkGridAdvancePx>[0]['lineSpacing'],
      gridAllocationActive = true,
      scale = 1,
    ): number => wordUseFeLayoutParagraphMarkGridAdvancePx({
      ordinaryAdvancePx: 16,
      allocatedGridAdvancePx: 30,
      atLeastZeroAdvancePx: 18,
      lineSpacing,
      gridAllocationActive,
      scale,
    });
    expect(markGridAdvance(null)).toBe(30);
    expect(markGridAdvance({ rule: 'atLeast', value: 18, explicit: true })).toBe(30);
    expect(markGridAdvance({ rule: 'atLeast', value: 0, explicit: false })).toBe(18);
    expect(markGridAdvance({ rule: 'atLeast', value: -1, explicit: true })).toBe(1);
    expect(markGridAdvance({ rule: 'atLeast', value: -18, explicit: true })).toBe(18);
    expect(markGridAdvance({ rule: 'atLeast', value: -1, explicit: true }, true, 2)).toBe(2);
    expect(markGridAdvance({ rule: 'atLeast', value: 0, explicit: true })).toBe(18);
    expect(markGridAdvance({ rule: 'exact', value: 18, explicit: true })).toBe(16);
    expect(markGridAdvance({ rule: 'atLeast', value: -1, explicit: true }, false)).toBe(16);
    expect(wordDegenerateLineSpacingIsSingle('exact', 0)).toBe(true);
    expect(wordDegenerateLineSpacingIsSingle('atLeast', 0)).toBe(false);
    expect(wordAutoMultipleCenterBoxPx(true, false, 10, 12, 24)).toBe(12);
    expect(wordAutoMultipleCenterBoxPx(true, true, 10, 12, 6)).toBe(6);
    expect(wordVisibleLineMetricPx(20, 8)).toBe(8);
    expect(wordVisibleLineMetricPx(20, undefined)).toBe(20);
    expect(wordFirstJustifiedContentSegment(
      [{ text: '  ' }, { text: 'body' }],
      false,
    )).toBe(1);
    expect(wordFirstJustifiedContentSegment(
      [{ text: '  ' }, { text: 'body' }],
      true,
    )).toBe(0);
    expect(wordRubyUniformLineHeightPx(true, [12, 18, 14])).toBe(18);
    expect(wordRubyUniformLineHeightPx(false, [12, 18, 14])).toBe(0);
    expect(wordRtlAmbiguousCharacter('.')).toBe(true);
    expect(wordRtlAmbiguousCharacter('A')).toBe(false);
    expect(wordKashidaFinalFormApplies(3, 3)).toBe(true);
    expect(wordKashidaFinalFormApplies(2, 3)).toBe(false);
    expect(wordPreservesVerticalTuCorner(0xfe11)).toBe(true);
    expect(wordPreservesVerticalTuCorner(null)).toBe(false);
  });

  it('pins supported table-model compatibility helper branches', () => {
    expect(WORD_TABLE_BORDER_STYLE_PRECEDENCE[0]).toBe('single');
    expect(WORD_TABLE_BORDER_STYLE_PRECEDENCE.at(-1)).toBe('inset');
    expect(Object.isFrozen(WORD_TABLE_BORDER_STYLE_PRECEDENCE)).toBe(true);
    expect(wordTableBorderWeight('single', 1.5)).toBe(12);
    expect(wordTableBorderWeight('dotted', 99)).toBe(1);
    expect(wordTableBorderWeight('unknown', 2)).toBe(0);
    expect(wordTableRowHeightRule('auto', false)).toBe('atLeast');
    expect(wordTableRowHeightRule('auto', true)).toBe('auto');
    expect(wordAuthoredAutoRowHeightUsesFloor('auto', 0)).toBe(true);
    expect(wordAuthoredAutoRowHeightUsesFloor('auto', Number.POSITIVE_INFINITY)).toBe(true);
    expect(wordAuthoredAutoRowHeightUsesFloor('atLeast', 10)).toBe(false);
    expect(wordTableCellSpacingValuePt('pct', null)).toBe(0);
    expect(wordTableCellSpacingValuePt('dxa', 4)).toBe(4);
    expect(wordTableMarginValuePt({
      kind: 'pct', dxaValuePt: null, scope: 'table', edge: 'start',
    })).toBe(0);
    expect(wordTableMarginValuePt({
      kind: 'nil', dxaValuePt: null, scope: 'table', edge: 'top',
    })).toBeNull();
    expect(wordTableMarginValuePt({
      kind: 'dxa', dxaValuePt: 3, scope: 'cell', edge: 'bottom',
    })).toBe(3);
  });

  it('drops only an empty trailing paragraph after a non-paragraph cell block', () => {
    expect(wordDropsTrailingStructuralCellMarker({
      contentLength: 2,
      previousKind: 'table',
      lastKind: 'paragraph',
      lastParagraphRunCount: 0,
    })).toBe(true);
    expect(wordDropsTrailingStructuralCellMarker({
      contentLength: 2,
      previousKind: 'paragraph',
      lastKind: 'paragraph',
      lastParagraphRunCount: 0,
    })).toBe(false);
    expect(wordDropsTrailingStructuralCellMarker({
      contentLength: 2,
      previousKind: 'table',
      lastKind: 'paragraph',
      lastParagraphRunCount: 1,
    })).toBe(false);
    expect(wordDropsTrailingStructuralCellMarker({
      contentLength: 1,
      previousKind: undefined,
      lastKind: 'paragraph',
      lastParagraphRunCount: 0,
    })).toBe(false);
  });
});
