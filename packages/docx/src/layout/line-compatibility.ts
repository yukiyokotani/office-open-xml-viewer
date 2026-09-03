import { defineCompatibilityRule } from './compatibility.js';
import type { LineSpacing, TabStop } from '../types.js';

export const WORD_EAST_ASIAN_GRID_LINE_ALLOCATION = defineCompatibilityRule({
  id: 'word-east-asian-grid-line-allocation',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/layout/compatibility.test.ts#pins East Asian grid allocation and the untabled Far East metric factor',
  },
  description: 'For an East Asian single-spaced line on a document grid, preserve the measured whole-cell allocation from the intended face design height and use the established 1.3-times-em fallback only when that design height is unavailable.',
});

export const WORD_TABLE_CELL_IGNORES_GRID_RIGHT_INDENT_ADJUSTMENT = defineCompatibilityRule({
  id: 'word-table-cell-ignores-grid-right-indent-adjustment',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'table-cell-adjust-right-indent-width-position-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'In the observed linesAndChars matrix, paragraphs inside fixed-width table cells retain the same line breaks for omitted (default true) and explicit-false w:adjustRightInd across four boundary widths and both left/right cell positions. Scope this Word-only exception to table-cell containers; ordinary body paragraphs retain the ECMA-376 §17.3.1.1 adjustment.',
});

export const WORD_SNAP_TO_CHARS_EAST_ASIAN_CELL_FIT = defineCompatibilityRule({
  id: 'word-snap-to-chars-east-asian-cell-fit',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'snap-to-chars-east-asian-cell-fit-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'For snapToChars, Word centers each East-Asian grapheme independently in the smallest whole number of character-pitch units that contains its natural advance. A grapheme that fits uses the one-unit placement described by [MS-OI29500] §2.1.534; an undersized authored pitch expands only that grapheme to additional units.',
});

export const WORD_SNAP_TO_CHARS_SCRIPT_BLOCK_ALLOCATION = defineCompatibilityRule({
  id: 'word-snap-to-chars-script-block-allocation',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §2.1.534',
  },
  description: 'Allocate snapToChars Latin text in contiguous blocks centered across the required grid units, complex-script blocks from their leading edge, and East-Asian graphemes independently by character cell.',
});

/** Word compatibility projection governed by
 * {@link WORD_SNAP_TO_CHARS_EAST_ASIAN_CELL_FIT}. */
export function wordSnapToCharsEastAsianCellCount(
  naturalAdvancePt: number,
  pitchPt: number,
): number {
  if (!(pitchPt > 0) || !Number.isFinite(naturalAdvancePt)) return 1;
  return Math.max(1, Math.ceil(Math.max(0, naturalAdvancePt) / pitchPt - 1e-9));
}

/** Compatibility projection governed by
 * {@link WORD_TABLE_CELL_IGNORES_GRID_RIGHT_INDENT_ADJUSTMENT}. */
export function wordContainerAllowsGridRightIndentAdjustment(
  insideTableCell: boolean,
): boolean {
  return !insideTableCell;
}

export const WORD_GRID_RIGHT_INDENT_PITCH_ALIGNMENT = defineCompatibilityRule({
  id: 'word-grid-right-indent-pitch-alignment',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'grid-right-indent-character-pitch-boundary-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'For body paragraphs whose ECMA-376 §17.3.1.1 adjustment is enabled on a linesAndChars character grid, Word reduces the physical line width to the greatest whole character-pitch multiple not exceeding the available width. The observed matrix covers exact and non-exact widths, zero and negative charSpace, explicit opt-out, line-only control, both physical indent sides, and the separately registered table-cell exception.',
});

/** Word compatibility projection governed by
 * {@link WORD_GRID_RIGHT_INDENT_PITCH_ALIGNMENT}. */
export function wordGridRightIndentAdjustmentPt(
  availableWidthPt: number,
  pitchPt: number,
): number {
  if (!(pitchPt > 0) || !Number.isFinite(availableWidthPt) || availableWidthPt <= 0) {
    return 0;
  }
  const remainder = ((availableWidthPt % pitchPt) + pitchPt) % pitchPt;
  const epsilon = 1e-9;
  return remainder <= epsilon || pitchPt - remainder <= epsilon ? 0 : remainder;
}

export const WORD_HANGING_TAB_SAME_POSITION_PRECEDENCE = defineCompatibilityRule({
  id: 'word-hanging-tab-same-position-precedence',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'hanging-indent-authored-tab-collision-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'When the implicit tab created by a hanging indent shares its coordinate with an authored center, end, or start stop, Word resolves one advancing stop at that coordinate using the authored alignment. An authored bar remains an independent drawing rule, so the implicit advancing stop survives beside it. If center/end alignment would place following text before the current pen, the tab contributes zero advance.',
});

/** Compatibility projection governed by
 * {@link WORD_HANGING_TAB_SAME_POSITION_PRECEDENCE}. */
export function wordAuthoredTabReplacesImplicitHangingStop(
  alignment: TabStop['alignment'],
): boolean {
  return alignment !== 'bar' && alignment !== 'clear';
}

export const WORD_RTL_DECIMAL_TAB_PHYSICAL_ALIGNMENT = defineCompatibilityRule({
  id: 'word-rtl-decimal-tab-physical-alignment',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'rtl-decimal-tab-run-boundary-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'For LTR numeric cells embedded in a bidi paragraph, Word aligns the physical left edge of the first halfwidth period to the decimal stop across source-run boundaries. When no period exists, it aligns the numeric cell\'s physical right edge to the stop.',
});

export const WORD_DECIMAL_TAB_SEPARATOR_RESOLUTION = defineCompatibilityRule({
  id: 'word-decimal-tab-separator-resolution',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §2.1.556',
  },
  description: 'Use the first explicit halfwidth period as the decimal-tab alignment point; when absent, use the implicit separator after the final digit of the first Unicode decimal-number sequence.',
});

export const WORD_USE_FE_LAYOUT_INHERITED_GRID_MINIMUM = defineCompatibilityRule({
  id: 'word-use-fe-layout-inherited-grid-minimum',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'use-fe-layout-visible-script-grid-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'With useFELayout enabled, a visible Latin line with a resolved eastAsia font axis participates in Far East grid metrics even when w:rFonts@hint is absent; inherited automatic spacing keeps the larger of its whole-cell design allocation and one grid pitch multiplied by the inherited spacing value.',
});

export const WORD_USE_FE_LAYOUT_EMPTY_MARK_GRID_ALLOCATION = defineCompatibilityRule({
  id: 'word-use-fe-layout-empty-mark-grid-allocation',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'use-fe-layout-empty-mark-grid-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'With useFELayout enabled, a content-less paragraph mark participates in Far East whole-cell document-grid allocation even when the document contains no literal East Asian text. Its face-specific Far East design height governs the cell count; exact spacing and snapToGrid=false remain the document-grid overrides named by ECMA-376 §17.6.5. Observed Word output gives signed atLeast spacing a discontinuous boundary on an active grid: negative values use their absolute magnitude as the mark advance, zero keeps the ordinary atLeast-zero advance regardless of inheritance source, and positive values retain whole-cell allocation.',
});

/** Compatibility projection governed by the useFELayout empty-mark allocation
 * and {@link WORD_GRID_AT_LEAST_TALL_LINE_UNSNAPPED}. The caller supplies the
 * ordinary line-spacing result and the mark's whole-cell grid allocation.
 * Exact spacing is the normative §17.6.5 override. Observed Word output gives
 * signed atLeast values an empty-mark-specific negative/zero/positive boundary. */
export function wordUseFeLayoutParagraphMarkGridAdvancePx(
  input: Readonly<{
    ordinaryAdvancePx: number;
    allocatedGridAdvancePx: number;
    atLeastZeroAdvancePx: number;
    lineSpacing: LineSpacing | null;
    gridAllocationActive: boolean;
    scale: number;
  }>,
): number {
  const {
    ordinaryAdvancePx,
    allocatedGridAdvancePx,
    atLeastZeroAdvancePx,
    lineSpacing,
    gridAllocationActive,
    scale,
  } = input;
  if (!gridAllocationActive) return ordinaryAdvancePx;
  if (lineSpacing?.rule === 'atLeast' && lineSpacing.value < 0) {
    return Math.abs(lineSpacing.value) * scale;
  }
  if (lineSpacing?.rule === 'atLeast' && lineSpacing.value === 0) {
    return atLeastZeroAdvancePx;
  }
  return lineSpacing?.rule === 'exact'
    ? ordinaryAdvancePx
    : Math.max(ordinaryAdvancePx, allocatedGridAdvancePx);
}

export const WORD_CONTIGUOUS_UNDERLINE_GEOMETRY = defineCompatibilityRule({
  id: 'word-contiguous-underline-geometry',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/layout/paragraph.test.ts#keeps a solid underline continuous across floating-precision retained run seams',
  },
  description: 'Adjacent compatible underlined source runs share one safe baseline and continuous authored cadence while style, color, and thickness boundaries remain distinct.',
});

export const WORD_GRID_AT_LEAST_TALL_LINE_UNSNAPPED = defineCompatibilityRule({
  id: 'word-grid-at-least-tall-line-unsnapped',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/line-box-height.test.ts#does not round tall East Asian content up to an additional grid cell',
  },
  description: 'An explicitly authored atLeast line on an active document grid keeps the maximum of its natural height, authored minimum, and one pitch instead of rounding tall content to another whole cell.',
});

export const WORD_DEGENERATE_LINE_SPACING_SINGLE = defineCompatibilityRule({
  id: 'word-degenerate-line-spacing-single',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-DOC] §2.9.146',
  },
  description: 'Preserve a non-collapsing single-line fallback for exact or automatic line spacing at or below zero, consistent with the native LSPD representation.',
});

export const WORD_AUTO_MULTIPLE_BASELINE_PIN = defineCompatibilityRule({
  id: 'word-auto-multiple-baseline-pin',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'auto-multiple-baseline-pin',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'Paint a positive automatic line-spacing multiplier with its glyph baseline pinned inside the single design line, placing extra leading or compressed overflow toward block-end; this is draw-only and does not replace the centered trailing-mark pagination metric.',
});

export const WORD_MIXED_ANCHOR_VISIBLE_LINE_METRICS = defineCompatibilityRule({
  id: 'word-mixed-anchor-visible-line-metrics',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/anchor-host-metrics.test.ts#reserves host line height without using its zero-ink box for a visible run baseline',
  },
  description: 'A zero-ink drawing anchor host reserves its line and grid height while visible neighboring glyphs retain their own ascent, descent, and design-line baseline.',
});

export const WORD_JUSTIFICATION_LEADING_INDENT_EXCLUSION = defineCompatibilityRule({
  id: 'word-justification-leading-indent-exclusion',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/text-distribute.test.ts#forwards (segs, slack, firstContentSi, lastDrawnSi) positionally',
  },
  description: 'Keep leading whitespace used as a first-line text indent fixed while distributing justified-line slack across content in a left-to-right line.',
});

export const WORD_JUSTIFIED_CANDIDATE_SEPARATOR_FIT = defineCompatibilityRule({
  id: 'word-justified-candidate-separator-fit',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/justify-shrink-overshoot.test.ts#counts a candidate trailing space when the prospective line will justify',
  },
  description: 'On a full paragraph-width line that will be fully justified, include the candidate word separator in its wrap-fit width; lines narrowed by DrawingML wrap exclusions retain collapsible line-end separator fit behavior.',
});

export const WORD_OVERFLOW_PUNCTUATION_LANGUAGE_SETS = defineCompatibilityRule({
  id: 'word-overflow-punctuation-language-sets',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OE376] §2.1.56',
  },
  description: 'Apply the language-specific punctuation sets documented for Word in [MS-OE376] §2.1.56, and let overflowPunct override kinsoku when both rules affect the same character.',
});

export const WORD_FULL_WIDTH_CHARACTER_SPACING_SCOPE = defineCompatibilityRule({
  id: 'word-full-width-character-spacing-scope',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OE376] §2.1.562',
  },
  description: 'Interpret ST_CharacterSpacing as applying whitespace compression to full-width punctuation characters. This rule establishes only which characters are eligible; it does not define a universal compression amount.',
});

export const WORD_JAPANESE_PUNCTUATION_COMPRESSION_CELL = defineCompatibilityRule({
  id: 'word-japanese-punctuation-compression-cell',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'japanese-fullwidth-punctuation-compression-cell',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'In the observed Japanese compatibility matrix, 、。 ，． and the closing forms 」』】）］｝ on a full ideographic-cell advance retain at least half of that cell. U+3017 and full-width !/? remain full-cell. A fontTable w:pitch value classifies the authored face for font selection; it is not a switch for document-level characterSpacingControl. Punctuation that the selected face already exposes on a smaller proportional advance is retained as measured rather than compressed a second time. Tight adjacent glyph ink can require a larger retained extent to prevent collision. This is an Office-observed compression amount, not a normative interpretation of ST_CharacterSpacing.',
});

export const WORD_AUTHORED_CHARACTER_SPACING_PITCH_PRIORITY = defineCompatibilityRule({
  id: 'word-authored-character-spacing-pitch-priority',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'authored-character-spacing-punctuation-pitch',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'When a run authors a positive w:spacing character pitch, Word preserves that expanded pitch instead of additionally applying the document-level punctuation whitespace compression. Omitted, zero, or overlapping run spacing leaves characterSpacingControl active.',
});

export const WORD_SOURCE_RUN_SPACE_SEQUENCE = defineCompatibilityRule({
  id: 'word-source-run-space-sequence',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'source-run-space-sequence-wrap-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'At a source-run boundary, Word keeps a space-only continuation attached when the preceding run already ends in a space. A single leading space in a distinct run without a preceding space remains a break opportunity. This isolates source-boundary compatibility from the ordinary UAX #14 LB7 handling within one authored run.',
});

export const WORD_CONSECUTIVE_SPACE_NATURAL_ADVANCE = defineCompatibilityRule({
  id: 'word-consecutive-space-natural-advance',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'consecutive-space-wrap-grid-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'When visible text follows two or more authored consecutive spaces, Word preserves the sequence at natural advance instead of using it as Knuth-Plass inter-word shrink capacity. The result is invariant across linesAndChars with negative/zero charSpace and a line-only grid; source-run boundaries remain governed separately by the source-space-sequence rule.',
});

export const WORD_BALANCED_CONSECUTIVE_SPACE_CELL = defineCompatibilityRule({
  id: 'word-balanced-consecutive-space-cell',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'single-double-byte-width-space-grid-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'With ECMA-376 §17.15.3.3 balanceSingleByteDoubleByteWidth enabled, Word retains one ordinary inter-word U+0020 at its proportional natural advance, while a sequence of two or more authored U+0020 spaces advances each space by half of the selected East-Asian ideographic cell. The observed matrix covers one, two, four, and eight spaces; same-run and source-run boundaries; proportional and fixed-pitch faces; linesAndChars with negative/zero charSpace; and a line-only grid.',
});

/** Compatibility projection governed by
 * {@link WORD_BALANCED_CONSECUTIVE_SPACE_CELL}. */
export function wordBalancedConsecutiveSpaceCellApplies(spaceCount: number): boolean {
  return Number.isInteger(spaceCount) && spaceCount >= 2;
}

/** Evidence-bounded grid scope governed by
 * {@link WORD_BALANCED_CONSECUTIVE_SPACE_CELL}. `snapToChars` has a separate
 * Microsoft-documented block/cell allocator and is outside this observation. */
export function wordBalancedSpaceCellAdjustmentApplies(
  gridType: string | null | undefined,
): boolean {
  return gridType !== 'snapToChars';
}

export const WORD_BALANCED_LINES_AND_CHARS_GRID_DELTA = defineCompatibilityRule({
  id: 'word-balanced-lines-and-chars-grid-delta',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'single-double-byte-width-grid-observation-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'With balanceSingleByteDoubleByteWidth enabled on linesAndChars, Word applies half of the authored charSpace delta to ASCII SBCS text and to U+0020/U+3000 space characters, while applying the full delta to CJK ideographs and full-width ASCII forms. The Word-output evidence covers ASCII digits, letters, punctuation, spaces, CJK, full-width ASCII, mixed text, proportional/fixed-pitch faces, negative/zero/positive charSpace, and line-only controls. Non-ASCII high-ANSI and complex-script text are outside the observed matrix and retain the preexisting grid behavior.',
});

export const WORD_IDEOGRAPHIC_SPACE_LINE_END_ALLOWANCE = defineCompatibilityRule({
  id: 'word-ideographic-space-line-end-allowance',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'ideographic-space-line-end-count-and-run-boundary-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'Word keeps a single U+3000 immediately following visible East-Asian text on that line when the visible glyph is force-fitted into a narrow table cell. A paragraph-final sequence of two or more U+3000 characters remains authored width-bearing content and may form blank continuation lines. The observed matrix covers single and trailing multiple spaces, linesAndChars with negative/positive charSpace, line-only grids, and snapToGrid opt-out.',
});

/** Compatibility projection governed by
 * {@link WORD_IDEOGRAPHIC_SPACE_LINE_END_ALLOWANCE}. */
export function wordIdeographicSpaceLineEndAllowanceCount(
  hasEastAsianVisiblePredecessor: boolean,
  consecutiveSpaceCount: number,
): 0 | 1 {
  return hasEastAsianVisiblePredecessor && consecutiveSpaceCount === 1 ? 1 : 0;
}

/** Compatibility projection governed by
 * {@link WORD_BALANCED_LINES_AND_CHARS_GRID_DELTA}. Script-slot acquisition
 * has already separated ordinary East-Asian and ASCII SBCS text; the explicit
 * space branch retains Word's observed U+3000 exception without reclassifying
 * other East-Asian glyphs. Non-ASCII high-ANSI/complex text stays outside the
 * observed projection. */
export function wordBalancedLinesAndCharsGridDeltaFactor(
  text: string,
  script: 'ascii' | 'highAnsi' | 'eastAsia' | 'complexScript',
): 0.5 | 1 | undefined {
  if (script === 'complexScript') return undefined;
  const spaceOnly = text.length > 0 && [...text].every(
    (character) => character === ' ' || character === '\u3000',
  );
  if (spaceOnly) return 0.5;
  if (script === 'eastAsia') return 1;
  return [...text].every((character) => (character.codePointAt(0) ?? 0x80) <= 0x7f)
    ? 0.5
    : undefined;
}

export const WORD_MS_MINCHO_EMPTY_EAST_ASIAN_MARK_HEIGHT = defineCompatibilityRule({
  id: 'word-ms-mincho-empty-east-asian-mark-height',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'ms-mincho-empty-east-asian-paragraph-mark',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'In the observed compatibility fixture, an empty 12-point East-Asian paragraph mark routed to MS Mincho occupies a 15.6-point single-line box. Scope this 1.3-em floor to empty East-Asian paragraph marks; ordinary MS Mincho text lines and Latin marks retain their independently measured metrics.',
});

/** Compatibility projection governed by
 * {@link WORD_JAPANESE_PUNCTUATION_COMPRESSION_CELL}. */
export function wordJapanesePunctuationRetainedExtentPt(input: Readonly<{
  punctuationAdvancePt: number;
  punctuationInkEndPt: number;
  ideographicCellAdvancePt: number;
}>): number {
  const advancePt = Math.max(0, input.punctuationAdvancePt);
  const cellAdvancePt = Math.max(0, input.ideographicCellAdvancePt);
  if (advancePt < cellAdvancePt) return advancePt;
  return Math.min(
    advancePt,
    Math.max(
      0,
      input.punctuationInkEndPt,
      cellAdvancePt / 2,
    ),
  );
}

/** Compatibility projection governed by
 * {@link WORD_AUTHORED_CHARACTER_SPACING_PITCH_PRIORITY}. */
export function wordDocumentCharacterCompressionApplies(
  authoredCharacterSpacingPt: number | undefined,
): boolean {
  return authoredCharacterSpacingPt === undefined || authoredCharacterSpacingPt <= 0;
}

/** Compatibility projection governed by {@link WORD_SOURCE_RUN_SPACE_SEQUENCE}. */
export function wordSourceRunSpaceContinuesSequence(
  previousText: string,
  currentText: string,
): boolean {
  return previousText.endsWith(' ') && currentText.startsWith(' ');
}

const WORD_OVERFLOW_PUNCTUATION = {
  ja: new Set([...',.’”、。」』】），．］｝｡､']),
  zhHans: new Set([...`!%),.:;>?]}¢°·ˇ’”‰′″℃∶、。〃〉》」』】〗〕〞﹚﹜﹞！＂％＇），．：；？］｝￠`]),
  zhHant: new Set([...`!),.:;?]}’”′、。〉》」』】〕〞﹚﹜﹞！），．：；？］｝`]),
  ko: new Set([...`!%),.:;?]}¢°’”′″℃〉》」』】〕！％），．：；？］｝￠`]),
} as const;
const ALL_WORD_OVERFLOW_PUNCTUATION = new Set([
  ...WORD_OVERFLOW_PUNCTUATION.ja,
  ...WORD_OVERFLOW_PUNCTUATION.zhHans,
  ...WORD_OVERFLOW_PUNCTUATION.zhHant,
  ...WORD_OVERFLOW_PUNCTUATION.ko,
]);

/** Compatibility projection governed by
 * {@link WORD_OVERFLOW_PUNCTUATION_LANGUAGE_SETS}. */
export function wordIsOverflowPunctuation(
  character: string,
  language: string | undefined,
): boolean {
  const normalized = language?.toLowerCase();
  if (normalized?.startsWith('ja')) return WORD_OVERFLOW_PUNCTUATION.ja.has(character);
  if (normalized?.startsWith('ko')) return WORD_OVERFLOW_PUNCTUATION.ko.has(character);
  if (normalized?.startsWith('zh')) {
    return (/(?:^|-)(?:tw|hk|mo)(?:-|$)|hant/u.test(normalized)
      ? WORD_OVERFLOW_PUNCTUATION.zhHant
      : WORD_OVERFLOW_PUNCTUATION.zhHans).has(character);
  }
  return ALL_WORD_OVERFLOW_PUNCTUATION.has(character);
}

/** Compatibility projection governed by {@link WORD_JUSTIFIED_CANDIDATE_SEPARATOR_FIT}. */
export function wordCandidateFitWidthPx(input: Readonly<{
  widthPx: number;
  trailingSpacePx: number;
  lineWillJustify: boolean;
  wrapNarrowed?: boolean;
}>): number {
  return input.lineWillJustify && input.wrapNarrowed !== true
    ? input.widthPx
    : input.widthPx - input.trailingSpacePx;
}

/** A calibrated same-route allowance cannot be projected across a line whose
 * characters resolve to different measurement routes. */
export function wordJustifiedCandidateFitAllowancePx(input: Readonly<{
  biasBudgetPx: number;
  resolvedMeasurementRouteCount: number;
}>): number {
  return input.resolvedMeasurementRouteCount === 1 ? input.biasBudgetPx : 0;
}

export const WORD_RUBY_PARAGRAPH_UNIFORM_LINE_ADVANCE = defineCompatibilityRule({
  id: 'word-ruby-paragraph-uniform-line-advance',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/paragraph-measure.test.ts#uses one uniform snapped advance for every line in a ruby paragraph',
  },
  description: 'Every line in a ruby-bearing paragraph uses the paragraph-wide maximum snapped line advance so its baseline rhythm remains uniform.',
});

export const WORD_FIT_TEXT_INTER_CHARACTER_EXPANSION = defineCompatibilityRule({
  id: 'word-fit-text-inter-character-expansion',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/fit-text.test.ts#distributes (val − Σnatural)/(n−1) as the inter-character gap, no trailing gap',
  },
  description: 'Expand a multi-character fitText region to its authored width by distributing the residual evenly across interior character gaps.',
});

export const WORD_CJK_BOTH_INTER_CHARACTER_EXPANSION = defineCompatibilityRule({
  id: 'word-cjk-both-inter-character-expansion',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/text-distribute.test.ts#§17.18.44: fills a wrapped pure-CJK line via inter-CJK pitch (expansion default)',
  },
  description: 'Treat inter-CJK boundaries as eligible inter-word gaps when expanding a non-final both-justified line that contains no spaces.',
});

export const WORD_THAI_DISTRIBUTE_CLUSTER_POLICY = defineCompatibilityRule({
  id: 'word-thai-distribute-cluster-policy',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/thai-distribute.test.ts#fills non-final lines to the right margin under thaiDistribute',
  },
  description: 'Expand non-final thaiDistribute lines at Thai grapheme-cluster boundaries while retaining a natural-width final line.',
});

export const WORD_NUMERIC_DECIMAL_TAB_INFERENCE = defineCompatibilityRule({
  id: 'word-numeric-decimal-tab-inference',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/decimal-tab-autoalign.test.ts#right-aligns numbers of different digit counts at the decimal tab',
  },
  description: 'Right-align an otherwise tab-less numeric paragraph at its leading decimal tab while leaving non-numeric and no-decimal-tab paragraphs unchanged.',
});

export const WORD_NUMBERING_MARKER_OVERFLOW_TAB_ADVANCE = defineCompatibilityRule({
  id: 'word-numbering-marker-overflow-tab-advance',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/numbered-marker-tab-advance.test.ts#advances the body past the marker to the next tab stop, not onto indentLeft',
  },
  description: 'When a numbering marker overruns its hanging-indent budget, advance the body to the next reachable tab stop beyond the marker edge.',
});

export const WORD_NUMBERING_SUFFIX_COINCIDENT_LIST_TAB = defineCompatibilityRule({
  id: 'word-numbering-suffix-coincident-list-tab',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/layout/numbering-marker.test.ts#keeps a suffix tab on the list stop coincident with the marker end',
  },
  description: 'For the tab synthesized by a numbering suffix, accept an authored numeric list tab coincident with the shaped marker end instead of advancing to the next automatic tab stop.',
});

export const WORD_NUMBERING_MARKER_PARAGRAPH_MARK_FALLBACK = defineCompatibilityRule({
  id: 'word-numbering-marker-paragraph-mark-fallback',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'numbering-marker-paragraph-mark-formatting',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'When numbering-level rPr omits a marker formatting axis, Word takes that axis from the effective paragraph-mark rPr rather than a content run. A numbering-level concrete value or explicit auto remains authoritative, and body and text-box stories use the same cascade.',
});

/** Compatibility projection governed by {@link WORD_NUMBERING_SUFFIX_COINCIDENT_LIST_TAB}. */
export function wordNumberingSuffixAcceptsCoincidentListTab(
  markerEndPt: number,
  stop: Readonly<{ pos: number; alignment: string }>,
): boolean {
  return stop.alignment === 'num' && Math.abs(stop.pos - markerEndPt) <= 1e-6;
}

export const WORD_TAB_STOP_PAGE_EDGE_CLAMP = defineCompatibilityRule({
  id: 'word-tab-stop-page-edge-clamp',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/rtl-tab-stops.test.ts#pins a page number to the left text margin when the stop is past it',
  },
  description: 'Clamp content assigned to a tab stop beyond the trailing text edge back onto that edge instead of placing ink outside the page content band.',
});

export const WORD_DICTIONARY_SEA_NATURAL_FIT = defineCompatibilityRule({
  id: 'word-dictionary-sea-natural-fit',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/sea-justified-fit.test.ts#Rule 1: wraps the paragraph-final Thai word on a thaiDistribute closing line (zero space-shrink)',
  },
  description: 'Do not admit a dictionary Southeast-Asian word by compressing preceding inter-word spaces when its natural advance exceeds the remaining line width.',
});

export const WORD_DICTIONARY_SEA_ATOMIC_CHUNK = defineCompatibilityRule({
  id: 'word-dictionary-sea-atomic-chunk',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/sea-justified-fit.test.ts#Rule 2: a no-space chunk that fits a full line moves whole instead of splitting',
  },
  description: 'Move a glued dictionary Southeast-Asian chunk to a fresh line whole when it fits that full line, using dictionary breaks only when the chunk itself is overlong.',
});

export const WORD_OVERLONG_TOKEN_EMERGENCY_BREAK = defineCompatibilityRule({
  id: 'word-overlong-token-emergency-break',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/run-inline-formatting.test.ts#breaks a no-space token wider than the line at the character level',
  },
  description: 'Emergency-break an overlong token at grapheme-safe character boundaries on an empty line so the complete token remains inside the content band.',
});

export const WORD_EXTERNAL_LINK_SYNTAX_BREAKS = defineCompatibilityRule({
  id: 'word-external-link-syntax-breaks',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'external-link-syntax-formatting-seam-matrix',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'Treat readable separators in the path and query of displayed external URLs as line-break opportunities, while keeping the scheme and authority intact and preserving authored no-break hyphens and grapheme clusters.',
});

/** Compatibility projection governed by {@link WORD_EXTERNAL_LINK_SYNTAX_BREAKS}.
 * `graphemeBoundaries` and `authoredNoBreakOffsets` are UTF-16 offsets in the
 * complete displayed link token, not in an individual formatting run. */
export function wordExternalLinkSyntaxBreakOffsets(
  text: string,
  graphemeBoundaries: ReadonlySet<number>,
  authoredNoBreakOffsets: ReadonlySet<number>,
): readonly number[] {
  const scheme = /^[A-Za-z][A-Za-z0-9+.-]*:\/\//u.exec(text);
  if (!scheme) return [];
  const authorityStart = scheme[0].length;
  const authorityEndCandidate = text.slice(authorityStart).search(/[/?#]/u);
  const authorityEnd = authorityEndCandidate < 0
    ? text.length
    : authorityStart + authorityEndCandidate;
  const offsets: number[] = [];
  for (let index = authorityEnd; index < text.length; index += 1) {
    const character = text[index]!;
    const offset = index + 1;
    const readable =
      (character === '/' && index > authorityEnd)
      || character === '-'
      || character === '?'
      || character === '&';
    if (
      readable
      && graphemeBoundaries.has(offset)
      && !authoredNoBreakOffsets.has(offset)
    ) offsets.push(offset);
  }
  return offsets;
}

export const WORD_RUN_VERTICAL_ALIGN_BASELINE_SHIFT = defineCompatibilityRule({
  id: 'word-run-vertical-align-baseline-shift',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/run-char-metrics-render.test.ts#w:vertAlign raises superscript, lowers subscript, and leaves ordinary baselines unchanged',
  },
  description: 'Retain the established run-level baseline displacement for vertically aligned text: superscript rises by 0.35 of its authored font size and subscript falls by 0.15, while the separately authored w:position remains additive.',
});

export const WORD_UNIFORM_RUN_POSITION_LEADING = defineCompatibilityRule({
  id: 'word-uniform-run-position-leading',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'uniform-run-position-leading',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'When every metric-bearing item on a line has the same non-zero w:position, Word preserves the enlarged line extent but shares the resulting surplus above and below the glyphs. A line containing a differently-positioned item retains the full relative displacement.',
});

/** Paint-relative baseline position governed by
 * {@link WORD_UNIFORM_RUN_POSITION_LEADING}. */
export function wordUniformRunPositionPaintPt(
  authoredPositionPt: number,
  commonLinePositionPt: number,
): number {
  return commonLinePositionPt === 0
    ? authoredPositionPt
    : authoredPositionPt - commonLinePositionPt / 2;
}

/** Compatibility projection governed by
 * {@link WORD_RUN_VERTICAL_ALIGN_BASELINE_SHIFT}. */
export function wordRunVerticalAlignRaisePt(
  verticalAlign: string | null | undefined,
  authoredFontSizePt: number,
): number {
  if (verticalAlign === 'super') return authoredFontSizePt * 0.35;
  if (verticalAlign === 'sub') return -authoredFontSizePt * 0.15;
  return 0;
}

export const WORD_FAR_EAST_SINGLE_LINE_FACTOR = 1.3;

/** Compatibility projection governed by
 * {@link WORD_MS_MINCHO_EMPTY_EAST_ASIAN_MARK_HEIGHT}. */
export function wordMsMinchoEmptyEastAsianMarkSingleLinePx(
  family: string | null | undefined,
  emPx: number,
  eastAsianMark: boolean,
): number {
  if (!eastAsianMark || !family) return 0;
  const normalized = family.trim().toLowerCase();
  return normalized === 'ms mincho' || normalized === 'ｍｓ 明朝'
    ? emPx * WORD_FAR_EAST_SINGLE_LINE_FACTOR
    : 0;
}

export function wordEastAsianGridLineCells(
  naturalHeightPx: number,
  pitchPx: number,
): number {
  return pitchPx > 0 ? Math.max(1, Math.ceil(naturalHeightPx / pitchPx)) : 1;
}

export function wordFarEastSingleLinePx(
  intendedSinglePx: number,
  emPx: number,
): number {
  return intendedSinglePx > 0
    ? intendedSinglePx
    : emPx * WORD_FAR_EAST_SINGLE_LINE_FACTOR;
}

/** Compatibility projection governed by
 * {@link WORD_USE_FE_LAYOUT_INHERITED_GRID_MINIMUM}. */
export function wordUseFeLayoutInheritedGridHeightPx(
  allocatedCellHeightPx: number,
  pitchPx: number,
  inheritedMultiple: number,
): number {
  return Math.max(allocatedCellHeightPx, pitchPx * inheritedMultiple);
}

export function wordGridAtLeastLineHeightPx(
  naturalPx: number,
  authoredMinimumPx: number,
  gridMinimumPx: number,
): number {
  return Math.max(naturalPx, authoredMinimumPx, gridMinimumPx);
}

export function wordDegenerateLineSpacingIsSingle(
  rule: string,
  value: number,
): boolean {
  return (rule === 'exact' || rule === 'auto') && value <= 0;
}

export function wordAutoMultipleCenterBoxPx(
  autoMultiple: boolean,
  compressedAuto: boolean,
  glyphNaturalPx: number,
  intendedSinglePx: number,
  lineHeightPx: number,
): number {
  return autoMultiple && !compressedAuto
    ? Math.max(glyphNaturalPx, intendedSinglePx)
    : lineHeightPx;
}

export function wordVisibleLineMetricPx(
  reservedMetricPx: number,
  visibleMetricPx: number | undefined,
): number {
  return visibleMetricPx ?? reservedMetricPx;
}

export function wordFirstJustifiedContentSegment(
  segments: readonly object[],
  bidi: boolean,
): number {
  if (bidi) return 0;
  for (let index = 0; index < segments.length; index += 1) {
    const segment = segments[index];
    const text = 'text' in segment && typeof segment.text === 'string'
      ? segment.text
      : undefined;
    if (text === undefined || /\S/.test(text)) return index;
  }
  return 0;
}

export function wordRubyUniformLineHeightPx(
  hasRuby: boolean,
  lineHeightsPx: readonly number[],
): number {
  return hasRuby ? Math.max(0, ...lineHeightsPx) : 0;
}
