import type {
  BodyLayoutInput,
  BodyLayoutSequenceEntryFor,
  BodyParagraphSourceInput,
} from './body-layout-input.js';
import { defineCompatibilityRule } from './compatibility.js';
import type {
  LineLayout,
  ParagraphLayout,
  ParagraphPlacement,
  RetainedGlyphPaintOperation,
  WritingMode,
} from './types.js';

export const WORD_PRE_BREAK_ANCHOR_PARAGRAPH = defineCompatibilityRule({
  id: 'word-pre-break-anchor-paragraph',
  evidence: {
    kind: 'library-policy',
    reference: 'packages/docx/src/pagination.test.ts#does not push an anchor-only pre-break paragraph to a new page just for its empty mark',
  },
  description: 'The established pagination regression contract keeps an anchor-only paragraph immediately before an authored page break in the pre-break flow region without charging its otherwise visible paragraph mark. Library policy, no Office measurement: keeps an anchor-only pre-break paragraph in the pre-break region (PR #898 corpus fix, nothing measured).',
});

export const WORD_PRE_BREAK_INLINE_DRAWING_GROUP = defineCompatibilityRule({
  id: 'word-pre-break-inline-drawing-group',
  evidence: {
    kind: 'library-policy',
    reference: 'packages/docx/src/pagination.test.ts#moves a preceding image with its pre-break callout when the pair only fits fresh',
  },
  description: 'The established pagination regression contract relocates a preceding inline DrawingML resource with an immediately following host-owned anchor paragraph before an authored page break when the pair fits only in a fresh flow region. Library policy, no Office measurement: keeps a preceding inline drawing with its pre-break callout (PR #898 corpus fix, nothing measured).',
});

export const WORD_CONTINUOUS_SECTION_MARK_SPACING = defineCompatibilityRule({
  id: 'word-continuous-section-mark-spacing',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/body-layout-input.test.ts#projects mutually exclusive collapsed-mark and drop-previous-after roles',
  },
  description: 'The retained body input projects the established continuous-section empty-mark spacing behavior into one mutually exclusive role before pagination. Evidence (observed Office behaviour): Word PDF of a corpus document in PR #595, where Word and LibreOffice agree; no ECMA-376 or Microsoft note covers it.',
});

export const WORD_CONTEXTUAL_SPACING_PER_SIDE = defineCompatibilityRule({
  id: 'word-contextual-spacing-per-side',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/contextual-spacing-body-paint.test.ts#paints the adjudicated six-case gap table',
  },
  description: 'For same-style adjacent paragraphs, contextualSpacing removes only the contribution owned by each toggling side; a current-only toggle preserves the previous paragraph spaceAfter contribution. Classification: normative, ECMA-376 §17.3.1.9 contextualSpacing; a six-case synthetic Word (macOS) control in issue #1015 / PR #1016 confirms it (application version not recorded).',
});

export const WORD_EMPTY_KEEP_NEXT_BRIDGE = defineCompatibilityRule({
  id: 'word-empty-keep-next-bridge',
  evidence: {
    kind: 'library-policy',
    reference: 'packages/docx/src/layout/body-paginator-production.test.ts#bridges an undecorated empty keepNext mark through the following paragraph',
  },
  description: 'Word print pagination treats an undecorated empty keep-with-next paragraph as a bridge: the following paragraph is admitted completely with the first indivisible content of its successor. Library policy, no Office measurement: keeps an undecorated empty keepNext mark from stranding (PR #1080, corpus acceptance gate only).',
});

export const WORD_AUTOMATIC_KEEP_NEXT_START_SPACING = defineCompatibilityRule({
  id: 'word-automatic-keep-next-start-spacing',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'automatic-paragraph-top-spacing',
    application: 'Microsoft Word',
    version: '16.0.20430',
    platform: 'Windows',
  },
  description: 'When automatic overflow relocates a keep-with-next unit to a fresh physical page, suppress the leading paragraph space-before for that grouped relocation. Standalone authored hard page breaks use a separate rule. Evidence: the trr3 family C controls (keepNext off/on) recorded for word-automatic-paragraph-top-spacing; origin PR #1080.',
});

export const WORD_AUTOMATIC_PARAGRAPH_TOP_SPACING = defineCompatibilityRule({
  id: 'word-automatic-paragraph-top-spacing',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'automatic-paragraph-top-spacing',
    application: 'Microsoft Word',
    version: '16.0.20430',
    platform: 'Windows',
  },
  description: 'trr3 family C controls: before=0/15pt pairs put After at the same position after automatic page overflow for empty, visible-text and line-break-only paragraphs, in modes 14/15 with keepNext off/on. This supersedes the earlier Mac empty-mark inference. ECMA-376 §17.3.1.33 defines the authored spacing, not this page-top suppression. Image-only and mixed-object paragraphs remain outside the controls and keep their existing ordinary-overflow spacing. Keep-next groups, table/float overflow, same-page columns and continuations have separate ownership. The auto=1.15/grid pitch discrepancies in these controls are not settled by the spacing pairs.',
});

export const WORD_STANDALONE_HARD_PAGE_BREAK_TOP_SPACING = defineCompatibilityRule({
  id: 'word-standalone-hard-page-break-top-spacing',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'standalone-hard-page-break-top-spacing',
    application: 'Microsoft Word',
    version: '16.0.20430',
    platform: 'Windows',
  },
  description: 'Windows trr3 family C before=0/15pt pairs (empty, visible text, line-break-only; modes 14/15; keepNext off/on) confirm suppression after a standalone hard page break. pageBreakBefore retains before-spacing in mode 14 and suppresses it in mode 15; the latter is selected only for that explicit mode. Without pageBreakBefore, the first paragraph of a document or a new section retains before-spacing. Standalone suppression applies only to parser-proven authored breaks, excluding synthetic Cover Pages breaks, parity section breaks and inline breaks after visible content. Word caret Y is not paragraph top: the paired After positions establish the spacing effect; these controls do not establish exact twip rounding or automatic grid pitch.',
});

export const WORD_TRAILING_SPACE_AFTER_FIT_ADMISSION = defineCompatibilityRule({
  id: 'word-trailing-space-after-fit-admission',
  evidence: {
    kind: 'library-policy',
    reference: 'packages/docx/src/layout/paragraph-pagination.test.ts#admits final visible content when only authored spaceAfter crosses the region edge',
  },
  description: 'Admit the final visible paragraph content at a flow-region edge when only its authored trailing space crosses the edge, while retaining that space for placement and paint. Library policy, no Office measurement: trailing space-after admission assumption from PR #1080; nothing measured.',
});

export const WORD_GRID_LINE_CONTENT_FIT_ADMISSION = defineCompatibilityRule({
  id: 'word-grid-line-content-fit-admission',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'grid-line-page-edge-admission',
    application: 'Microsoft Word',
    version: '16.0.20430',
    platform: 'Windows',
  },
  description: 'Windows Word 16.0 build 20430 admits a centered text box inside a 20pt line-grid cell when the natural box ends inside the body, even when the cell overhangs the bottom margin. A later-positioned caption whose natural box crosses the body bottom moves to the next page. These inside/outside observations support content-box admission, not baseline-only admission or a fixed overhang tolerance. ECMA-376 §§17.6.5 and 17.3.1.32 define grid participation, not this page-edge rule. trr3 family D further observes the same page boundary for 11pt MS Mincho empty marks and single-character paragraphs on a 16.4pt grid (modes 14/15, lines/linesAndChars, zero/one preceding empty mark): 15pt moves and 15.5pt stays. Empty marks carry the equivalent metric-only allowance through word-trailing-empty-mark-baseline-admission. Only leading added beyond its own natural/authored height is eligible; authored atLeast minima and auto multiples remain charged. The allowance is restricted to a positive whitelist of plain glyph paint in undecorated paragraphs: paragraph/run borders (including between/bar rules), shading/highlight, underline/strike/emphasis, ruby, fields, tabs, objects and unknown retained paint retain the full cell, even when a continuation slice omits paragraph edges. Off-grid and exact lines retain full-box admission. This observation applies only at the physical body bottom with no intervening reservation; note bands, wrap exclusions and shortened column regions retain the whole grid cell. Evidence: issue #1674 / PR #1796 controls.',
});

// Positive paint whitelist for the bounded plain-caption observation. Unknown
// retained properties fail closed: adding a decoration must not silently widen
// this compatibility rule. Geometry, source ownership and ordinary glyph paint
// are safe. §17.3.2.4 requires run borders; their retained geometry surrounds
// the full allocated cell. Shading, underline/strike/emphasis, ruby, fields,
// tabs and objects are outside the evidence.
const GRID_PLAIN_PARAGRAPH_KEYS = new Set([
  'kind', 'id', 'source', 'flowDomainId', 'ordinaryFlow', 'sectionFlowOwnership',
  'flowBounds', 'inkBounds', 'clipBounds', 'advancePt', 'diagnostics',
  'paragraphId', 'styleId', 'bookmarkStarts', 'spacing', 'contextualSpacing',
  'lines', 'resources', 'events', 'exclusions', 'cellContainmentBounds', 'anchorCollisions',
  'anchorFrames',
  'continuation',
]);
const GRID_PLAIN_LINE_KEYS = new Set([
  'range', 'bounds', 'baselinePt', 'advancePt', 'gridAfterLeadingPt', 'placements',
  'wrapBounds', 'wrapFragments',
]);
const GRID_PLAIN_TEXT_KEYS = new Set([
  'kind', 'text', 'optionalHyphenGlyph', 'semanticSlotSpans', 'sourceRunIndex',
  'range', 'origin', 'bounds', 'highlightBounds', 'advancePt', 'clusters',
  'trailingSpaceCompressionPt', 'color', 'fontRoute', 'fontSizePt',
  'fontWeight', 'fontStyle', 'direction', 'writingMode', 'characterSpacingPt',
  'characterScale', 'fitText', 'kerning', 'ownedTrailingSlackPt', 'hyperlink',
]);
// Acquisition may still carry basePaintOps beside the final glyph operations.
// Both contain glyph paint only; a future effect must opt in explicitly.
const GRID_PLAIN_GLYPH_KEYS = new Set([
  'text', 'range', 'offset', 'letterSpacingPt', 'scaleX', 'scaleY', 'direction',
  'kerning', 'writingMode', 'glyphOrientation', 'verticalFeature', 'glyphOffsetPt',
  'sourceMapping', 'inkBounds', 'blockAxisInkBounds', 'naturalAdvancePt',
]);
const GRID_PLAIN_TYPOGRAPHY_KEYS = new Set([
  'caps', 'smallCaps', 'strike', 'doubleStrike', 'verticalAlign', 'positionPt', 'emphasis',
]);

function wordGridLineHasPlainPaint(paragraph: ParagraphLayout, line: LineLayout): boolean {
  return Object.entries(paragraph).every(([key, value]) => value === undefined
    || GRID_PLAIN_PARAGRAPH_KEYS.has(key)
    || (['borders', 'drawings', 'textBoxes'].includes(key)
      && Array.isArray(value) && value.length === 0))
    && paragraph.lines.every(candidate => (candidate.barTabRules?.length ?? 0) === 0)
    && Object.entries(line).every(([key, value]) => value === undefined
      || GRID_PLAIN_LINE_KEYS.has(key)
      || (key === 'barTabRules' && Array.isArray(value) && value.length === 0))
    && line.placements.every(placement => placement.kind === 'text'
      && Object.entries(placement).every(([key, value]) => value === undefined
        || GRID_PLAIN_TEXT_KEYS.has(key)
        || ((key === 'paintOps' || key === 'basePaintOps') && Array.isArray(value)
          && value.every(operation => Object.keys(operation).every(name => GRID_PLAIN_GLYPH_KEYS.has(name))))
        || (key === 'role' && value === 'content')
        || (key === 'sourceRuns' && placement.sourceRuns?.every(owner =>
          owner.role === undefined && owner.dependency === undefined))
        || (key === 'decorations' && placement.decorations.length === 0)
        || (key === 'unsupportedGeometry' && placement.unsupportedGeometry?.length === 0)
        || (key === 'typography' && placement.typography
          && Object.keys(placement.typography).every(name => GRID_PLAIN_TYPOGRAPHY_KEYS.has(name))
          && !placement.typography.strike && !placement.typography.doubleStrike
          && placement.typography.verticalAlign.status === 'missing'
          && placement.typography.positionPt.status === 'missing'
          && placement.typography.emphasis.status === 'missing')));
}

/** Omit only eligible blank terminal grid leading, preserving paragraph trailing extents
 * and all retained placement/paint geometry. Applies to complete and continuing
 * fragments; earlier lines still consume their full grid cells. The caller
 * supplies a physical body bottom only when the active region reaches it
 * without a footer, committed note band or column-balancing reservation. */
export function wordGridLineAdmissionExtentPt(
  paragraph: ParagraphLayout,
  logicalLineBoxExtentPt: number,
  unreservedBodyBottomPt: number | undefined,
  additionalReservePt = 0,
  blockOffsetPt = 0,
): number {
  void WORD_GRID_LINE_CONTENT_FIT_ADMISSION;
  const lastLine = paragraph.lines.at(-1);
  const leadingPt = lastLine?.gridAfterLeadingPt ?? 0;
  if (leadingPt === 0) return logicalLineBoxExtentPt;
  // The page-margin observation provides no allowance against a retained
  // reservation (§17.11.21 note bands; §20.4 anchored wrap exclusions).
  // This is a scope restriction, not a new Office compatibility inference.
  if (unreservedBodyBottomPt === undefined || additionalReservePt > 0) {
    return logicalLineBoxExtentPt;
  }
  // Keep-next lookahead acquires successors at the current cursor; only
  // its terminal slice moves by the preceding keep members' full advances;
  // host-owned exclusions move with it, while page-owned exclusions stay put.
  if (lastLine && paragraph.exclusions.some(({ bounds, verticalOwnership }) => {
    const topPt = bounds.yPt + (verticalOwnership === 'host' ? blockOffsetPt : 0);
    return topPt < unreservedBodyBottomPt
      && topPt + bounds.heightPt > lastLine.bounds.yPt + blockOffsetPt
      && bounds.xPt < paragraph.flowBounds.xPt + paragraph.flowBounds.widthPt
      && bounds.xPt + bounds.widthPt > paragraph.flowBounds.xPt;
  })) return logicalLineBoxExtentPt;
  // Authored paragraph paint also disables gridAfterLeadingPt at acquisition,
  // before continuation/keep-next slicing can remove a top/bottom/between edge.
  if (!lastLine || !wordGridLineHasPlainPaint(paragraph, lastLine)) {
    return logicalLineBoxExtentPt;
  }
  return Math.max(0, logicalLineBoxExtentPt - leadingPt);
}

export const WORD_VERTICAL_RL_FINAL_LINE_BASELINE_ADMISSION = defineCompatibilityRule({
  id: 'word-vertical-rl-final-line-baseline-admission',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'vertical-rl-final-line-baseline-admission',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'In a vertical-rl section, Word admits the final visible text column when its transformed baseline and retained visible ink remain inside the block-end edge even if the complete logical line box crosses that edge. The complete retained advance remains authoritative after admission.',
});

export const WORD_LOWERED_DROP_CAP_ANCHOR_LEADING = defineCompatibilityRule({
  id: 'word-lowered-drop-cap-anchor-leading',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'lowered-drop-cap-anchor-leading',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'Word keeps the following anchor text below a baseline-lowered drop-cap glyph while preserving the drop cap exclusion height authored by framePr lines. ECMA-376 specifies those two inputs independently but does not prescribe this interaction.',
});

/** Compatibility projection governed by
 * {@link WORD_LOWERED_DROP_CAP_ANCHOR_LEADING}. */
export function wordLoweredDropCapAnchorLeadingPt(
  maximumBaselineLoweringPt: number,
): number {
  return Math.max(0, maximumBaselineLoweringPt);
}

/** Whether a framed paragraph's baseline shift contributes to its line box.
 * Only the observed lowered-drop-cap interaction suppresses that contribution;
 * an ordinary `w:framePr` remains governed by §17.3.2.24 metrics. */
export function wordFramePositionExtendsLineBox(dropCap: string): boolean {
  return dropCap === 'none';
}

/** Compatibility projection governed by {@link WORD_TRAILING_SPACE_AFTER_FIT_ADMISSION}. */
export function wordFinalParagraphAdmissionExtentPt(input: Readonly<{
  advancePt: number;
  retainedSpaceAfterPt: number;
  authoredSpaceAfterPt: number;
}>): number {
  return Math.max(
    0,
    input.advancePt - Math.min(input.authoredSpaceAfterPt, input.retainedSpaceAfterPt),
  );
}

function glyphBlockEndPt(operation: RetainedGlyphPaintOperation): number {
  return operation.origin.yPt + (operation.inkBounds?.descentPt ?? 0);
}

function placementVisibleBlockEndPt(placement: ParagraphPlacement): number | null {
  if (placement.kind === 'resource' || placement.kind === 'drawing') {
    return placement.bounds.yPt + placement.bounds.heightPt;
  }
  if (placement.kind === 'anchor-host') return null;
  if (placement.kind === 'tab') {
    const glyphs = placement.leaderGlyphs ?? [];
    return glyphs.length > 0 ? Math.max(...glyphs.map(glyphBlockEndPt)) : null;
  }
  const paintOps = placement.paintOps ?? [];
  const visibleEnds = paintOps.length > 0
    ? paintOps.map((operation) => placement.origin.yPt + operation.offset.yPt
      + (operation.blockAxisInkBounds?.endPt
        ?? operation.inkBounds?.descentPt
        ?? 0))
    : [placement.origin.yPt];
  for (const decoration of placement.decorations) {
    const strokeExtentPt = decoration.widthPt / 2;
    visibleEnds.push(
      decoration.from.yPt + strokeExtentPt,
      decoration.to.yPt + strokeExtentPt,
    );
    for (const point of decoration.path ?? []) visibleEnds.push(point.yPt + strokeExtentPt);
  }
  for (const fragment of placement.highlightFragments ?? []) {
    visibleEnds.push(fragment.rect.yPt + fragment.rect.heightPt);
  }
  for (const border of placement.runBorderFragments ?? []) {
    const strokeExtentPt = border.widthPt / 2;
    visibleEnds.push(border.from.yPt + strokeExtentPt, border.to.yPt + strokeExtentPt);
  }
  for (const glyph of placement.emphasis?.glyphs ?? []) {
    visibleEnds.push(glyphBlockEndPt(glyph));
  }
  for (const glyph of placement.ruby?.paintOps ?? []) {
    visibleEnds.push(glyphBlockEndPt(glyph));
  }
  for (const path of placement.emphasis?.paths ?? []) {
    const strokeExtentPt = path.stroke === null ? 0 : path.strokeWidthPt / 2;
    for (const point of path.points) visibleEnds.push(point.yPt + strokeExtentPt);
  }
  return Math.max(...visibleEnds);
}

/** Compatibility projection governed by
 * {@link WORD_VERTICAL_RL_FINAL_LINE_BASELINE_ADMISSION}.
 *
 * The vertical-rl page transform is `physicalX = pageWidth - logicalY`.
 * Therefore the retained final line's largest visible logical y-coordinate is
 * its transformed physical trailing edge. The transform is isometric, so its
 * origin-relative logical extent can be compared directly with the block
 * extent without a scale-dependent tolerance. */
export function wordVerticalRlFinalLineAdmissionExtentPt(input: Readonly<{
  paragraph: ParagraphLayout;
  writingMode: WritingMode;
  logicalLineBoxExtentPt: number;
  availableBlockExtentPt: number;
}>): number {
  if (input.writingMode !== 'vertical-rl') return input.logicalLineBoxExtentPt;
  // This rule is a narrow reduced-extent admission when the retained line box
  // itself does not fit. Ordinary line boxes that fit keep their established
  // admission extent; glyph ink may overhang a text column into its margin.
  if (input.logicalLineBoxExtentPt <= input.availableBlockExtentPt) {
    return input.logicalLineBoxExtentPt;
  }
  const finalLine = input.paragraph.lines.at(-1);
  if (!finalLine) return input.logicalLineBoxExtentPt;
  const hasUnprovedTransformedInk = finalLine.placements.some((placement) =>
    placement.kind === 'text'
    && (placement.paintOps ?? []).some((operation) =>
      operation.glyphOrientation !== undefined
      && operation.blockAxisInkBounds === undefined));
  if (hasUnprovedTransformedInk) return input.logicalLineBoxExtentPt;
  const visibleEnds = finalLine.placements.flatMap((placement) => {
    const endPt = placementVisibleBlockEndPt(placement);
    return endPt === null ? [] : [endPt];
  });
  if (visibleEnds.length === 0) return input.logicalLineBoxExtentPt;
  if (input.paragraph.shading) return input.logicalLineBoxExtentPt;
  for (const border of input.paragraph.borders) {
    const strokeExtentPt = border.widthPt / 2;
    visibleEnds.push(border.from.yPt + strokeExtentPt, border.to.yPt + strokeExtentPt);
  }
  if (input.paragraph.paragraphMark && !input.paragraph.paragraphMark.hidden) {
    const mark = input.paragraph.paragraphMark.bounds;
    visibleEnds.push(mark.yPt + mark.heightPt);
  }
  const visibleExtentPt = Math.max(0,
    Math.max(...visibleEnds) - input.paragraph.flowBounds.yPt);
  return visibleExtentPt;
}

/** Compatibility projection governed by {@link WORD_EMPTY_KEEP_NEXT_BRIDGE}. */
export function wordEmptyKeepNextBridgesSuccessor(input: Readonly<{
  keepNext: boolean;
  inkless: boolean;
  undecoratedMark: boolean;
}>): boolean {
  return input.keepNext && input.inkless && input.undecoratedMark;
}

/** Compatibility projection governed by {@link WORD_CONTINUOUS_SECTION_MARK_SPACING}. */
export function wordContinuousSectionRole(
  sequence: readonly BodyLayoutSequenceEntryFor<Readonly<{ startType: string }>>[],
  index: number,
): BodyParagraphSourceInput['continuousSectionRole'] {
  const entry = sequence[index];
  if (entry?.kind !== 'body-block' || entry.block.kind !== 'paragraph') return undefined;
  const next = sequence[index + 1];
  const afterNext = sequence[index + 2];
  const suppressBefore = entry.block.inkless === true
    && next?.kind === 'begin-section'
    && next.section.startType === 'continuous';
  if (suppressBefore && entry.block.spaceBeforePt === 0) return 'collapse-mark';
  if (suppressBefore) return 'suppress-before';
  return entry.block.inkless === true
    && next?.kind === 'body-block'
    && next.block.kind === 'paragraph'
    && next.block.inkless === true
    && next.block.spaceBeforePt === 0
    && afterNext?.kind === 'begin-section'
    && afterNext.section.startType === 'continuous'
    ? 'drop-previous-after'
    : undefined;
}

function paragraphHasOnlyDrawingAnchors(layout: ParagraphLayout): boolean {
  return layout.drawings.length > 0 && layout.lines.every((line) =>
    line.placements.every((placement) =>
      placement.kind === 'drawing' || placement.kind === 'anchor-host'));
}

/** A positive-size inline DrawingML resource that participates in line flow. */
export function wordPreBreakInlineDrawingResource(layout: ParagraphLayout): boolean {
  return layout.lines.some((line) => line.placements.some((placement) =>
    ((placement.kind === 'resource'
      && (placement.resourceKind === 'image' || placement.resourceKind === 'chart'))
      || placement.kind === 'drawing')
    && placement.advancePt > 0
    && placement.bounds !== undefined
    && placement.bounds.widthPt > 0
    && placement.bounds.heightPt > 0));
}

/** Host-relative drawing extent below the following paragraph's flow cursor. */
export function wordPreBreakHostAnchorExtentPt(
  layout: ParagraphLayout,
  anchorTopPt: number,
): number | null {
  if (!paragraphHasOnlyDrawingAnchors(layout)) return null;
  const drawings = layout.drawings.filter((drawing) => (
    drawing.anchorLayer?.verticalOwnership === 'host'
    && Number.isFinite(drawing.flowBounds.xPt)
    && Number.isFinite(drawing.flowBounds.yPt)
    && Number.isFinite(drawing.flowBounds.widthPt)
    && Number.isFinite(drawing.flowBounds.heightPt)
    && drawing.flowBounds.widthPt > 0
    && drawing.flowBounds.heightPt > 0
  ));
  if (drawings.length !== layout.drawings.length) return null;
  const bottomPt = Math.max(...drawings.map((drawing) =>
    drawing.flowBounds.yPt + drawing.flowBounds.heightPt));
  return Math.max(0, bottomPt - anchorTopPt);
}

/** Returns the retained anchor paragraph with its compatibility-only flow charge removed. */
export function wordFlowNeutralPreBreakAnchorParagraph(
  layout: ParagraphLayout,
): ParagraphLayout | null {
  if (!paragraphHasOnlyDrawingAnchors(layout)) return null;
  const { paragraphMark: _paragraphMark, ...withoutMark } = layout;
  return Object.freeze({
    ...withoutMark,
    advancePt: 0,
    flowBounds: Object.freeze({ ...layout.flowBounds, heightPt: 0 }),
  });
}

export const WORD_PARAGRAPH_FOOTNOTE_CONTINUATION = defineCompatibilityRule({
  id: 'word-paragraph-footnote-continuation',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'paragraph-footnote-continuation',
    application: 'Microsoft Word',
    version: '16.113.3',
    platform: 'macOS',
  },
  description: 'Existing Word-produced controls cover a completed body paragraph with a long note, following body text and a following hard page/section break without an intervening note. Registered earlier controls kept a short-note/font-size case and a long body paragraph with references on opposite pages whole; their inputs are not retained, so later classes do not claim to reproduce them. Public Word 16.113.3 controls (exact lines, one note of two or four 10pt or two 8pt lines, 6pt authored after on the paragraph whose last line closes with the reference, widowControl off/on and keepLines) split the note when the whole note missed only because of that authored after and kept it whole at the exact charged fit or without after; a one-line note there has no head and stays on the unchanged fallback. This continuation-mode policy decides only whether a note may split; where its head is reserved is governed by word-reference-line-footnote-allocation. It does not claim general Word pagination equivalence or support width-changing continuation. Notes that cannot fit with their first reference line on a fresh page use the separate content-retention policy. Evidence: PR #1726 controls.',
});

/** Decide continuation admission within the documented input classes.
 * Evaluated once for each reference-line note plan. `wholeNoteFitsOnly-
 * WithoutTrailingAfter` is the after-only class: the whole note misses beside
 * a unit that completes its paragraph but fits when only that paragraph's
 * admissible authored after (WORD_TRAILING_SPACE_AFTER_FIT_ADMISSION) is
 * excluded. It permits a split even when the whole note fits a fresh page and
 * nothing follows the reference; it does not decide where the head goes.
 * `pageOpensWithIncomingTail` is the carried-tail class: the page's band
 * already opens with real retained lines of a note continued from an earlier
 * page. Public Word 16.113.3 controls (single column, exact 10pt lines, two
 * notes; a two- or four-line second note whose reference closes its
 * paragraph, with or without authored after; and the automatic-line K03
 * document) split such a whole-missing note beside the tail instead of moving
 * the reference line, and kept it whole where the charged total fits. Pages
 * without such a tail keep the exclusion above; this is not a universal Word
 * rule. */
export function wordParagraphFootnoteMayContinue(input: Readonly<{
  wholeNoteFitsWithReference: boolean;
  wholeNoteFitsFreshPage: boolean;
  textAfterReference: boolean;
  hardBreakBeforeNextNote: boolean;
  wholeNoteFitsOnlyWithoutTrailingAfter: boolean;
  pageOpensWithIncomingTail: boolean;
  /** The note joins a column band (WORD_PER_COLUMN_FOOTNOTE_BANDS). */
  columnBand: boolean;
}>): boolean {
  return !input.wholeNoteFitsWithReference
    && (input.columnBand || !input.wholeNoteFitsFreshPage || input.textAfterReference
      || input.hardBreakBeforeNextNote
      || input.wholeNoteFitsOnlyWithoutTrailingAfter || input.pageOpensWithIncomingTail);
}

/**
 * Footnote reservation order inside one paragraph fragment (continuation
 * mode only). ECMA-376 §17.11.21 / §17.18.34 fix only ownership: a
 * note starts on the page that paints its reference. They do not prescribe
 * how a page's height is divided between a note and later body lines.
 *
 * Library policy: at each line that first references notes on the page, the
 * paginator takes one immutable note plan from that line's remaining capacity
 * (earlier plans already charged; preceding new notes whole, only the last
 * new note partitionable to at least one real line plus any native notice).
 * Later lines of the paragraph are admitted only beside the committed plans.
 * A later reference line can add reserve but never repartition, shrink or
 * drop an earlier plan; if its own plan does not fit, that line and every
 * later line are rejected together and no part of its plan is committed.
 *
 * Observed basis (finite, not a universal Word algorithm): Word 16.113.3
 * source-open controls with one column, exact 10pt lines, widowControl off,
 * single-paragraph notes and the reference on the first of several lines.
 * Authored exact-6pt separator stories at 36pt and 35.95pt remaining kept
 * two and one note lines beside the reference line and moved the next body
 * line; bare and missing separator stories and a fourteen-line continuing
 * paragraph also moved the body line after the note head. Extending the same
 * commitment to later reference lines is owner-approved library policy.
 */
export const WORD_REFERENCE_LINE_FOOTNOTE_ALLOCATION = defineCompatibilityRule({
  id: 'word-reference-line-footnote-allocation',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'footnote-reference-line-head-boundary',
    application: 'Microsoft Word',
    version: '16.113.3',
    platform: 'macOS',
  },
  description: 'With footnote continuation enabled, a referenced note is planned once at its reference line and later body lines of the paragraph are admitted only against that committed plan. Observed in exact-line single-column controls with first-line references (36pt and 35.95pt boundaries kept two and one head lines and moved the following line; a long continuing paragraph moved after the head). Later reference lines follow the same library policy and are rejected atomically when their own plan cannot fit. Separator band charges and other classes (columns, keep/widow, non-exact lines) are not established by these controls.',
});

/** Saved print controls: only a hard page/section boundary before the next
 * reference qualifies; a column break or intervening note does not. */
export function wordFootnoteLookaheadBoundary(entry: BodyLayoutInput['sequence'][number]): boolean {
  return (entry.kind === 'authored-break' && entry.break !== 'column')
    || (entry.kind === 'begin-section' && entry.section.startType !== 'continuous')
    || (entry.kind === 'body-block' && entry.block.kind === 'paragraph' && entry.block.pageBreakBefore);
}

export const WORD_PER_COLUMN_FOOTNOTE_BANDS = defineCompatibilityRule({
  id: 'word-per-column-footnote-bands',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'multi-column-footnote-ieee-controls',
    application: 'Microsoft Word',
    version: 'Microsoft 365',
    platform: 'Windows (ExportAsFixedFormat)',
  },
  description: 'ECMA-376 §17.11.21 puts pageBottom footnotes at the bottom of the page without defining multi-column note geometry. In Windows Word controls of two- and three-column sections (equal and 144/288pt columns, exact 12pt body and 10pt note lines), compatibility modes 11, 12 and 14 give each column its own note band, independent of footnoteLayoutLikeWW8 and the other legacy compatibility flags: a column\'s band holds the notes referenced in that column, wraps them at that column\'s width, starts with its own one-line separator band (also for an empty separator story) and ends at the body bottom, even when a column break ends the column early; a column without references keeps its full height. A note that does not fit continues at the top of the next column\'s band, before that column\'s own notes, and then on the next page; the reference line stays with as much of the note as fits, even when the whole note would fit the next column or a fresh page with nothing after the reference (one-line paragraphs closed by the reference, 10- and 20-line notes), unlike the single-column fallback of WORD_PARAGRAPH_FOOTNOTE_CONTINUATION. Endnotes follow the text in the column flow. Compatibility mode 15 (with or without footnoteLayoutLikeWW8) instead flows one page-wide note area through the columns; see WORD_PAGE_WIDE_FOOTNOTE_AREA. An omitted compatibility mode keeps the page band.',
});

/** Gate of {@link WORD_PER_COLUMN_FOOTNOTE_BANDS} for the active region. */
export function wordPerColumnFootnoteBandsApply(
  compatibilityMode: number | undefined,
  regionColumnCount: number,
): boolean {
  return compatibilityMode !== undefined && compatibilityMode < 15 && regionColumnCount > 1;
}


export const WORD_PAGE_WIDE_FOOTNOTE_AREA = defineCompatibilityRule({
  id: 'word-page-wide-footnote-area',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'mode15-footnote-area-controls',
    application: 'Microsoft Word',
    version: 'Microsoft 365',
    platform: 'Windows (ExportAsFixedFormat)',
  },
  description: 'ECMA-376 §17.11.21 does not define multi-column pageBottom geometry. Windows mode15 controls with two and three columns, equal and unequal widths, exact12pt body and exact10pt notes flow a source-ordered note prefix through one common-height area. Every body column is shortened by that height; later references can reflow earlier body columns. The largest prefix compatible with its retained references is selected, then the smallest common height retaining that prefix balances the area. Each destination wraps its source suffix at its own width. Only the first separator paints ink; the other columns reserve its measured advance. One-to-two and two-to-one continuous-section controls, including references on both sides, retain the first reference section as area owner. Compatibility flags do not change this mode15 observation. Omitted modes retain the page band. Single-column owners retain the independently measured physical-page split policy of WORD_PARAGRAPH_FOOTNOTE_CONTINUATION.',
});


export const WORD_PAGE_WIDE_ENDNOTE_AREA = defineCompatibilityRule({
  id: 'word-page-wide-endnote-area',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'mode15-endnote-area-controls',
    application: 'Microsoft Word',
    version: 'Microsoft 365',
    platform: 'Windows (ExportAsFixedFormat)',
  },
  description: 'Windows Word controls with compatibility mode 15, two equal columns, exact 12pt body lines and exact 10pt note lines balance the final body columns before document-end notes. The endnotes start below their common body edge, flow through the section columns in source order and continue on subsequent pages. The smallest common height retaining the selected note prefix balances the area; only its first separator occurrence paints ink. Thirty- and hundred-line body controls with multiple endnotes, plus a combined footnote/endnote control, bound the observation. Earlier modes retain sequential endnote column flow; absent modes are not inferred.',
});


export const WORD_CONTINUOUS_SECTION_FOOTNOTE_PAGE = defineCompatibilityRule({
  id: 'word-continuous-section-footnote-page',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'continuous-section-footnote-controls',
    application: 'Microsoft Word',
    version: 'Microsoft 365',
    platform: 'Windows (ExportAsFixedFormat)',
  },
  description: 'ECMA-376 Part4 section14.8.3.20 allows reference-free paragraphs after a continuous section break on a page containing earlier footnotes when footnoteLayoutLikeWW8 is true. Windows controls of exact12pt single-column sections keep the following section on the same page in mode15, with or without the flag, and in mode11 with the flag; mode11 without the flag starts another page. Mode15 also keeps one-to-two and two-to-one column transitions on the page, including a later reference. A mode11 two-to-one control with a column-owned footnote band starts another page despite the flag, so that observed column-band class retains its physical-page boundary. With the normative legacy flag, a later new reference starts another page; mode15 is the separately observed exception.',
});

export function wordContinuousSectionMayShareFootnotePage(
  compatibilityMode: number | undefined,
  footnoteLayoutLikeWW8: boolean | undefined,
  hasColumnBands: boolean,
): boolean {
  return compatibilityMode === 15 || (footnoteLayoutLikeWW8 === true && !hasColumnBands);
}
