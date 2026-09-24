import type {
  BodyLayoutInput,
  BodyParagraphSourceInput,
} from './body-layout-input.js';
import { defineCompatibilityRule } from './compatibility.js';
import type {
  ParagraphLayout,
  ParagraphPlacement,
  RetainedGlyphPaintOperation,
  WritingMode,
} from './types.js';

export const WORD_TERMINAL_COLUMN_BREAK = defineCompatibilityRule({
  id: 'word-terminal-column-break',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/pagination.test.ts#ignores a terminal last-column break before a hard page boundary',
  },
  description: 'The established pagination regression contract does not materialize a column transition when no body flow content occurs before the next forced page or non-continuous section boundary.',
});

export const WORD_PRE_BREAK_ANCHOR_PARAGRAPH = defineCompatibilityRule({
  id: 'word-pre-break-anchor-paragraph',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/pagination.test.ts#does not push an anchor-only pre-break paragraph to a new page just for its empty mark',
  },
  description: 'The established pagination regression contract keeps an anchor-only paragraph immediately before an authored page break in the pre-break flow region without charging its otherwise visible paragraph mark.',
});

export const WORD_PRE_BREAK_INLINE_DRAWING_GROUP = defineCompatibilityRule({
  id: 'word-pre-break-inline-drawing-group',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/pagination.test.ts#moves a preceding image with its pre-break callout when the pair only fits fresh',
  },
  description: 'The established pagination regression contract relocates a preceding inline DrawingML resource with an immediately following host-owned anchor paragraph before an authored page break when the pair fits only in a fresh flow region.',
});

export const WORD_CONTINUOUS_SECTION_MARK_SPACING = defineCompatibilityRule({
  id: 'word-continuous-section-mark-spacing',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/body-layout-input.test.ts#projects mutually exclusive collapsed-mark and drop-previous-after roles',
  },
  description: 'The retained body input projects the established continuous-section empty-mark spacing behavior into one mutually exclusive role before pagination.',
});

export const WORD_CONTEXTUAL_SPACING_PER_SIDE = defineCompatibilityRule({
  id: 'word-contextual-spacing-per-side',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/contextual-spacing-body-paint.test.ts#paints the adjudicated six-case gap table',
  },
  description: 'For same-style adjacent paragraphs, contextualSpacing removes only the contribution owned by each toggling side; a current-only toggle preserves the previous paragraph spaceAfter contribution.',
});

export const WORD_EMPTY_KEEP_NEXT_BRIDGE = defineCompatibilityRule({
  id: 'word-empty-keep-next-bridge',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/layout/body-paginator-production.test.ts#bridges an undecorated empty keepNext mark through the following paragraph',
  },
  description: 'Word print pagination treats an undecorated empty keep-with-next paragraph as a bridge: the following paragraph is admitted completely with the first indivisible content of its successor.',
});

export const WORD_AUTOMATIC_KEEP_NEXT_START_SPACING = defineCompatibilityRule({
  id: 'word-automatic-keep-next-start-spacing',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/layout/body-paginator-production.test.ts#suppresses leading spacing when a $kind keepNext unit moves to an automatic page',
  },
  description: 'When automatic overflow relocates a keep-with-next unit to a fresh physical page, suppress the leading paragraph space-before for that grouped relocation. Standalone authored hard page breaks use a separate rule.',
});

export const WORD_AUTOMATIC_PARAGRAPH_TOP_SPACING = defineCompatibilityRule({
  id: 'word-automatic-paragraph-top-spacing',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'automatic-paragraph-top-spacing',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'A controlled ordinary-text exact-line paragraph with before=6pt does not retain that before spacing when ordinary overflow moves the complete paragraph to a fresh physical page. A before=0pt paragraph defines the page-top control. Empty mark-only paragraphs retain authored before spacing in observed Word output. Image-only and mixed-object paragraphs are outside these controls, so the library conservatively retains their authored spacing. Table and float overflow, same-page columns, and paragraph continuations have separate ownership.',
});

export const WORD_STANDALONE_HARD_PAGE_BREAK_TOP_SPACING = defineCompatibilityRule({
  id: 'word-standalone-hard-page-break-top-spacing',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'standalone-hard-page-break-top-spacing',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'Controlled before=0/6pt documents in compatibility modes 14 and 15 place the first paragraph after a standalone hard page break at the same page-top origin: its before spacing is suppressed. The first paragraph of a document or a new section retains before spacing. A pageBreakBefore paragraph differs by mode (retained in 14, suppressed in 15); compatibilityMode is not yet represented in the parser model. This rule applies only to parser-proven authored breaks, excluding synthetic Cover Pages breaks, parity section breaks, and inline breaks after visible content.',
});

export const WORD_TRAILING_SPACE_AFTER_FIT_ADMISSION = defineCompatibilityRule({
  id: 'word-trailing-space-after-fit-admission',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/layout/paragraph-pagination.test.ts#admits final visible content when only authored spaceAfter crosses the region edge',
  },
  description: 'Admit the final visible paragraph content at a flow-region edge when only its authored trailing space crosses the edge, while retaining that space for placement and paint.',
});

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
  sequence: BodyLayoutInput['sequence'],
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

/**
 * Plans which authored column breaks have body flow content before the next
 * forced boundary. ECMA-376 §§17.3.3.1 and 17.18.4 define the column advance;
 * suppressing a terminal transition is Word compatibility behavior recorded by
 * {@link WORD_TERMINAL_COLUMN_BREAK}.
 */
export function wordActiveColumnBreakIndexes(
  sequence: BodyLayoutInput['sequence'],
): ReadonlySet<number> {
  const active = new Set<number>();
  let hasFlowContentBeforeBoundary = false;
  for (let index = sequence.length - 1; index >= 0; index -= 1) {
    const entry = sequence[index]!;
    if (entry.kind === 'body-block') {
      hasFlowContentBeforeBoundary = entry.block.kind !== 'paragraph'
        || !entry.block.pageBreakBefore;
      continue;
    }
    if (entry.kind === 'adjacent-table-group') {
      hasFlowContentBeforeBoundary = true;
      continue;
    }
    if (entry.kind === 'authored-break') {
      if (entry.break === 'column') {
        if (hasFlowContentBeforeBoundary) active.add(index);
      } else {
        hasFlowContentBeforeBoundary = false;
      }
      continue;
    }
    if (entry.kind === 'begin-section' && entry.section.startType !== 'continuous') {
      hasFlowContentBeforeBoundary = false;
    }
  }
  return active;
}
