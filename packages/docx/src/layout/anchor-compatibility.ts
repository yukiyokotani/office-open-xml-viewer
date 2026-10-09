import { defineCompatibilityRule } from './compatibility.js';
import type { ParagraphLayout, StoryLayout } from './types.js';

/** Evidence: Word for Mac PDF exports of synthetic controls (issue #1615):
 * margin-, page- and paragraph-relative pictures; square, full-width square
 * and topAndBottom wrap; allowOverlap on and off; single anchors, anchors in
 * adjacent paragraphs and a later anchor sharing a page with an accepted one;
 * multi-line anchor paragraphs with the anchor run on the first or last line;
 * widow control; keepNext on the preceding paragraph; a picture taller than
 * any page remainder; two newspaper columns; a full-width picture mid-body;
 * and a 400-paragraph document with six clustered anchors. */
export const WORD_PAGE_ANCHOR_LINE_DEFERRAL = defineCompatibilityRule({
  id: 'word-page-anchor-line-deferral',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'page-anchor-line-deferral',
    application: 'Microsoft Word',
    version: '16.113.2',
    platform: 'macOS 27.0',
  },
  description: 'ECMA-376 does not say where a page-owned drawing goes when its own wrap exclusion pushes its anchor line off the page. Word tests anchor lines in flow order on each page: with the page-owned anchors already accepted on page N whose lines precede line L, plus the anchors on L, Word lays out N again; if L stays on N (any column of the region) those anchors are accepted and the earlier lines wrap around them. Otherwise page N keeps its layout without them, L starts a later page as if the page ended just above L (keepNext and widow control act as for an overflow; a mid-paragraph anchor line splits its paragraph), and the anchors go to the page where L lands. Counterexamples in the same controls: a line that still fits keeps its picture on N, and the first line of a page stays even when its picture pushes it past the body bottom. Unmeasured: a deferral whose line would otherwise fall in an earlier column of a multi-column page moves to the next page, not the next column.',
});

export const WORD_PAGE_ANCHOR_REGION_REGISTRATION = defineCompatibilityRule({
  id: 'word-page-anchor-region-registration',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'page-anchor-line-deferral',
    application: 'Microsoft Word',
    version: '16.113.2',
    platform: 'macOS 27.0',
  },
  description: 'A page-owned drawing belongs to its page, not to the newspaper column its anchor line reaches: in a two-column control Word keeps column 1 wrapped below a margin-relative topAndBottom picture over column 1 while the anchor paragraph moves into column 2. Register such drawings where their section region opens on the page.',
});

export const WORD_ZERO_RELATIVE_SIZE_EXTENT_FALLBACK = defineCompatibilityRule({
  id: 'word-zero-relative-size',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/layout/anchor-frame.test.ts#uses wp:extent when Word does not support an exact-zero relative size',
  },
  description: 'Word 2010 accepts only positive wp14:pctWidth and wp14:pctHeight values under [MS-ODRAWXML] notes 125/126. Preserve an authored zero as acquisition evidence while resolving the object from wp:extent.',
});

export const WORD_VERTICAL_SECTION_PHYSICAL_DRAWING_LAYER = defineCompatibilityRule({
  id: 'word-vertical-section-physical-drawing-layer',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/anchor-vertical-physical.test.ts#lands an upright-section anchor at the recorded physical centroid',
  },
  description: 'Resolve anchored drawings in an upright vertical section against the physical page frame independently of the rotated text-flow coordinate space.',
});

export const WORD_PAGE_LEVEL_FLOAT_PRESCAN = defineCompatibilityRule({
  id: 'word-page-level-float-prescan',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/page-anchor-prescan.test.ts#pre-scan REGISTERS a page-level (relativeFrom="margin") wrap float on an earlier-scanned paragraph',
  },
  description: 'A wrapping drawing whose vertical reference is page-level participates from page start so source-earlier paragraphs on that page see its exclusion.',
});

export const WORD_PARAGRAPH_ANCHOR_PRE_SPACING_ORIGIN = defineCompatibilityRule({
  id: 'word-paragraph-anchor-pre-spacing-origin',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/anchor-paragraph-spacebefore.test.ts#anchors a wrapSquare paragraph float at the pre-spaceBefore paragraph top',
  },
  description: 'Resolve a paragraph-relative anchored drawing from the paragraph top before applying the paragraph spaceBefore contribution.',
});

export const WORD_VERTICAL_SECTION_PHYSICAL_HEADER_FOOTER = defineCompatibilityRule({
  id: 'word-vertical-section-physical-header-footer',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/vertical-header-footer.test.ts#recovers the physical page box + margins from the logical (swapped) section',
  },
  description: 'Paint a vertical section header and footer in the unrotated physical page frame rather than rotating them with the body text flow.',
});

export const WORD_FRAME_AUTO_WRAP_AROUND = defineCompatibilityRule({
  id: 'word-frame-auto-wrap-around',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/frame-geometry.test.ts#wrap="around" and "auto" → square float (auto ≡ around in Word)',
  },
  description: 'Resolve an authored frame wrap value of auto through the same square side-wrap path as around.',
});

export const WORD_LOWER_LAYER_SAME_PARAGRAPH_ANCHOR_COMPOSITION = defineCompatibilityRule({
  id: 'word-lower-layer-same-paragraph-anchor-composition',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'lower-layer-same-paragraph-anchor-composition',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'Word preserves a source-later, lower-z, page-owned drawing at its authored position when it belongs to the same anchor paragraph as already composed higher layers. This is a Word-observed compatibility override to ECMA-376 §20.4.2.3, not a normative OOXML rule.',
});

export const WORD_TEXTBOX_VISIBLE_ANCHOR_EXTENT = defineCompatibilityRule({
  id: 'word-textbox-visible-anchor-extent',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'textbox-visible-anchor-extent',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'For DrawingML middle and bottom text anchoring, derive the positioned extent through the last visible retained block while preserving structural trailing empty paragraphs and terminal paragraph spacing in the complete story.',
});

export const WORD_OVERLAPPING_LAYOUT_IN_CELL_OVERLAY = defineCompatibilityRule({
  id: 'word-overlapping-layout-in-cell-overlay',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/table-cell-anchor-reflow.test.ts#does not grow an automatic row for an overlapping layoutInCell wrapNone object',
  },
  description: 'Word leaves an overlap-permitted, non-wrapping layoutInCell drawing as an overlay instead of growing its automatic table row. This is an Office compatibility exception to the general resize behavior in ECMA-376 §20.4.2.3 layoutInCell; non-overlap drawings retain normative cell containment.',
});

/** Compatibility projection governed by
 * {@link WORD_OVERLAPPING_LAYOUT_IN_CELL_OVERLAY}. */
export function wordLayoutInCellOwnsRowContainment(
  allowOverlap: boolean,
  wrapKind: 'none' | 'square' | 'tight' | 'through' | 'topAndBottom',
): boolean {
  return !allowOverlap || wrapKind !== 'none';
}

function paragraphContributesTextBoxAnchorExtent(paragraph: ParagraphLayout): boolean {
  if (
    paragraph.shading
    || paragraph.borders.length > 0
    || paragraph.resources.length > 0
    || paragraph.drawings.length > 0
    || paragraph.textBoxes.length > 0
    || paragraph.lineNumbers?.some((line) => line.paintOps.length > 0)
  ) {
    return true;
  }
  return paragraph.lines.some((line) => line.placements.some((placement) => {
    if (placement.kind === 'text' || placement.kind === 'resource' || placement.kind === 'drawing') {
      return true;
    }
    return placement.kind === 'tab' && (placement.leaderGlyphs?.length ?? 0) > 0;
  }));
}

/** Compatibility projection governed by {@link WORD_TEXTBOX_VISIBLE_ANCHOR_EXTENT}. */
export function wordTextBoxVisibleAnchorExtentPt(story: StoryLayout): number {
  const startPt = story.flowBounds.yPt;
  let visibleEndPt: number | undefined;
  for (const block of story.blocks) {
    if (block.kind === 'table') {
      visibleEndPt = Math.max(
        visibleEndPt ?? startPt,
        block.flowBounds.yPt + block.advancePt,
      );
      continue;
    }
    if (block.kind !== 'paragraph' || !paragraphContributesTextBoxAnchorExtent(block)) continue;
    visibleEndPt = Math.max(
      visibleEndPt ?? startPt,
      block.flowBounds.yPt + Math.max(0, block.advancePt - block.spacing.afterPt),
    );
  }
  return visibleEndPt === undefined ? 0 : Math.max(0, visibleEndPt - startPt);
}

export function wordZeroRelativeSizeUsesExtent(fraction: number): boolean {
  return fraction === 0;
}

export function wordPageLevelAnchorY(
  relativeFrom: string | null | undefined,
  paragraphRelativeFallback: boolean,
): boolean {
  if (relativeFrom == null) return !paragraphRelativeFallback;
  return relativeFrom !== 'paragraph'
    && relativeFrom !== 'line'
    && relativeFrom !== 'character';
}

export function wordPreservesLowerLayerSameParagraphComposition(
  movingOwnership: 'page' | 'host',
  movingRelativeHeight: number | null,
  blockerRelativeHeight: number | undefined,
): boolean {
  return movingOwnership === 'page'
    && movingRelativeHeight !== null
    && blockerRelativeHeight !== undefined
    && movingRelativeHeight < blockerRelativeHeight;
}


/** Evidence: issue #1623 Word controls (compatibility modes 14 and 15; column,
 * margin and page-sized references; tight and square wrap; empty and text
 * anchor paragraphs; free gaps left, right and on both sides; paragraph
 * indents, first-line and hanging indents and centred anchor paragraphs). */
export const WORD_MODE14_COLUMN_LINE_START_ORIGIN = defineCompatibilityRule({
  id: 'word-mode14-column-line-start-origin',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'cross-paragraph-float-overlap',
    application: 'Microsoft Word',
    version: '16.113.2',
    platform: 'macOS 27.0',
  },
  description: 'In compatibility mode 14 a positionH posOffset relative to the column is measured from the left edge of the first free gap of the anchor paragraph first-line band (at the paragraph content start, before any float pushes the line down), around the floats of other paragraphs only; paragraph indents and alignment do not move it. The resolved position is kept when the line later moves. Mode 15 measures from the column edge. Alignment and percentage offsets are unmeasured and keep the column edge.',
});

export const WORD_TIGHT_WRAP_BOTTOM_EDGE = defineCompatibilityRule({
  id: 'word-tight-wrap-bottom-edge',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'cross-paragraph-float-overlap',
    application: 'Microsoft Word',
    version: '16.113.2',
    platform: 'macOS 27.0',
  },
  description: 'A line whose top lies exactly on the bottom edge of a wrapTight polygon still wraps around that edge (modes 14 and 15); a square object and a line whose bottom touches a polygon top do not.',
});

export const WORD_TIGHT_WRAP_LINE_STEP_ADVANCE = defineCompatibilityRule({
  id: 'word-tight-wrap-line-step-advance',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'cross-paragraph-float-overlap',
    application: 'Microsoft Word',
    version: '16.113.2',
    platform: 'macOS 27.0',
  },
  description: 'When no gap beside a wrapTight polygon can hold a line, Word moves the line down by its own height and tries again (modes 14 and 15) instead of jumping to the polygon bottom as it does for square wrap.',
});

export const WORD_MODE14_TIGHT_ANCHOR_LINE_REWRAP = defineCompatibilityRule({
  id: 'word-mode14-tight-anchor-line-rewrap',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'cross-paragraph-float-overlap',
    application: 'Microsoft Word',
    version: '16.113.2',
    platform: 'macOS 27.0',
  },
  description: 'In compatibility mode 14 the line holding a wrapTight anchor is first laid out without that object; it is laid out again around the object (with distL/distR) only when the unpadded polygon intersects the line content, an empty paragraph mark counting as one paragraph-mark em. Mode 15 and square wrap always wrap the anchor line.',
});

export const WORD_LATER_ANCHOR_EARLIER_LINE_WRAP = defineCompatibilityRule({
  id: 'word-later-anchor-earlier-line-wrap',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'cross-paragraph-float-overlap',
    application: 'Microsoft Word',
    version: '16.113.2',
    platform: 'macOS 27.0',
  },
  description: 'A paragraph-relative wrapping drawing also wraps the lines of earlier paragraphs on its page that it overlaps (modes 14 and 15, tight and square). The drawing keeps the position resolved when its anchor paragraph was first laid out; the earlier lines move, not the drawing.',
});

export const WORD_MODE14_TIGHT_ANCHOR_TOP_TOUCH = defineCompatibilityRule({
  id: 'word-mode14-tight-anchor-top-touch',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'cross-paragraph-float-overlap',
    application: 'Microsoft Word',
    version: '16.113.2',
    platform: 'macOS 27.0',
  },
  description: 'In compatibility mode 14, when the first line of the anchor paragraph of a wrapTight drawing cannot be placed at its start position, an earlier line whose bottom lies exactly on the polygon top also wraps around it. Otherwise a line touching the polygon top from above is not wrapped.',
});

export const WORD_GRID_PICTURE_LINE_ORIGIN = defineCompatibilityRule({
  id: 'word-grid-picture-line-origin',
  evidence: { kind: 'office-observation', syntheticFixtureId: 'float-grid-picture-origin',
    application: 'Microsoft Word', version: '16.113.2', platform: 'macOS 27.0' },
  description: 'Issue #1674 controlled exports (74 picture-origin cases within the 142-case gap/grid suite) in modes 14/15 keep paragraph-relative pictures at the pre-before-spacing paragraph top. On an active line grid the first line-relative picture uses that same origin before authored before-spacing, and later line references use their actual line tops. Grid pitches 12–36pt, font sizes 10/20/30pt, offsets, phases, snap overrides, exact/atLeast/auto spacing, empty/multiline paragraphs, page/margin references and square/tight/none wrap disprove a universal half-pitch picture offset. Non-grid line references, aligned line references (no wp:align control) and non-picture payloads retain their previous reference frames.',
});

export function wordGridPictureLineOriginPt(lineTop: number, paragraphTop: number, contentStart: number): number {
  return lineTop - (contentStart - paragraphTop);
}
