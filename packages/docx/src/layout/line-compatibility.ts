import { defineCompatibilityRule } from './compatibility.js';

export const WORD_EAST_ASIAN_GRID_LINE_ALLOCATION = defineCompatibilityRule({
  id: 'word-east-asian-grid-line-allocation',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/layout/compatibility.test.ts#pins East Asian grid allocation and the untabled Far East metric factor',
  },
  description: 'For an East Asian single-spaced line on a document grid, preserve the measured whole-cell allocation from the intended face design height and use the established 1.3-times-em fallback only when that design height is unavailable.',
});

export const WORD_USE_FE_LAYOUT_INHERITED_GRID_MINIMUM = defineCompatibilityRule({
  id: 'word-use-fe-layout-inherited-grid-minimum',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'far-east-hinted-latin-grid-multiple',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'With useFELayout enabled, a Latin line carrying an eastAsia-hinted run participates in Far East grid metrics; inherited automatic spacing keeps the larger of its whole-cell design allocation and one grid pitch multiplied by the inherited spacing value.',
});

export const WORD_USE_FE_LAYOUT_EMPTY_MARK_GRID_ALLOCATION = defineCompatibilityRule({
  id: 'word-use-fe-layout-empty-mark-grid-allocation',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/paragraph-measure.test.ts#applies useFELayout grid-cell allocation to an empty paragraph mark',
  },
  description: 'With useFELayout enabled, a content-less paragraph mark participates in Far East whole-cell document-grid allocation even when the document contains no literal East Asian text.',
});

export const WORD_CONTIGUOUS_UNDERLINE_GEOMETRY = defineCompatibilityRule({
  id: 'word-contiguous-underline-geometry',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/layout/paragraph.test.ts#uses one safe baseline for a solid underline spanning adjacent source runs',
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
    kind: 'regression-test',
    reference: 'packages/docx/src/line-spacing-baseline.test.ts#2.0× keeps the baseline at top + ascent (extra 1.0× leading below, NOT centred)',
  },
  description: 'Paint an automatic line-spacing multiplier at or above one with its glyph baseline pinned inside the single design line and place multiplier leading below it; this is draw-only and does not replace the centered trailing-mark pagination metric.',
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
  description: 'In the observed Japanese compatibility fixture, compressed full-width punctuation retains at least half of the ideographic cell measured through the selected font route. Tight adjacent glyph ink can require a larger retained extent to prevent collision. This is an Office-observed compression amount, not a normative interpretation of ST_CharacterSpacing.',
});

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
  return Math.min(
    advancePt,
    Math.max(
      0,
      input.punctuationInkEndPt,
      input.ideographicCellAdvancePt / 2,
    ),
  );
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

export const WORD_NUMBERING_MARKER_ALIGNED_FIRST_LINE = defineCompatibilityRule({
  id: 'word-numbering-marker-aligned-first-line',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'numbering-marker-aligned-first-line',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'A numbering marker follows the retained first-line alignment. For centered paragraphs Word centers the union of marker, suffix gap, and body; physical end alignment keeps the same composite beside the aligned body in both LTR and RTL paragraphs.',
});

/** Center-only translation governed by
 * {@link WORD_NUMBERING_MARKER_ALIGNED_FIRST_LINE}. The body-only slack remains
 * unchanged because it also owns line breaking, compression, and justification. */
export function wordNumberingCompositeCenterDeltaPt(input: Readonly<{
  baseRtl: boolean;
  bodyOffsetPt: number;
  bodyWidthPt: number;
  markerInterval?: Readonly<{
    startPt: number;
    endPt: number;
  }>;
}>): number {
  if (!input.markerInterval) return 0;
  const bodyEndPt = input.bodyOffsetPt + input.bodyWidthPt;
  const compositeStartPt = Math.min(input.bodyOffsetPt, input.markerInterval.startPt);
  const compositeEndPt = Math.max(bodyEndPt, input.markerInterval.endPt);
  const logicalDeltaPt = (bodyEndPt - compositeStartPt - compositeEndPt) / 2;
  return input.baseRtl ? -logicalDeltaPt : logicalDeltaPt;
}

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
  description: 'Emergency-break a non-CJK token that is wider than an empty line at grapheme-safe character boundaries so it remains inside the content band.',
});

export const WORD_RUN_VERTICAL_ALIGN_BASELINE_SHIFT = defineCompatibilityRule({
  id: 'word-run-vertical-align-baseline-shift',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/run-char-metrics-render.test.ts#w:vertAlign raises superscript, lowers subscript, and leaves ordinary baselines unchanged',
  },
  description: 'Retain the established run-level baseline displacement for vertically aligned text: superscript rises by 0.35 of its authored font size and subscript falls by 0.15, while the separately authored w:position remains additive.',
});

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
