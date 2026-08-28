import { defineCompatibilityRule } from './compatibility.js';

export const WORD_DEFAULT_LINE_NUMBER_DISTANCE = defineCompatibilityRule({
  id: 'word-default-line-number-distance',
  evidence: {
    kind: 'regression-test',
    reference: "packages/docx/src/layout/compatibility.test.ts#uses Word's observed 18pt line-number distance only when omitted",
  },
  description: 'ECMA-376 §17.6.8 leaves an omitted line-number distance implementation-defined. Preserve Word-compatible 18pt placement only when the authored distance is absent.',
});

export const WORD_CONTINUOUS_SECTION_PAGE_NUMBER_RESTART = defineCompatibilityRule({
  id: 'word-continuous-section-page-number-restart',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/page-number-field-render.test.ts#restarts a spilling continuous section after its shared first page',
  },
  description: 'Issue #804 records that Word anchors a continuous section page-number restart to the first physical page containing that section body content, even when another section owns the page top. An empty same-page region does not consume the restart.',
});

export const WORD_TRAILING_EMPTY_MARK_BASELINE_ADMISSION = defineCompatibilityRule({
  id: 'word-trailing-empty-mark-baseline-admission',
  evidence: {
    kind: 'regression-test',
    reference: 'packages/docx/src/paginate-trailing-empty-mark-fit.test.ts#KEEPS an inkless empty paragraph on the page when ink-bearing content follows and only its below-baseline whitespace overflows',
  },
  description: 'At the unreserved physical body edge, Word admits an undecorated non-terminal empty paragraph mark by its baseline when later ink follows in the same flow.',
});

export const WORD_SECTION_MARK_BLANK_PAGE_SUPPRESSION = defineCompatibilityRule({
  id: 'word-section-mark-blank-page-suppression',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'next-page-section-mark-bottom-edge-admission',
    application: 'Microsoft Word',
    version: '16.111.1',
    platform: 'macOS 26.5.2',
  },
  description: 'For the observed nextPage boundary, Word consumes an undecorated inkless paragraph that carries the section boundary on the outgoing page instead of creating an otherwise empty intermediate page solely for that non-painting mark. The following section still starts its requested page. Other section-start kinds remain outside this observation.',
});

export const WORD_BOOK_FOLD_GUTTER_RIGHT_EDGE = defineCompatibilityRule({
  id: 'word-book-fold-gutter-right-edge',
  evidence: {
    kind: 'microsoft-note',
    reference: '[MS-OI29500] §§2.1.389, 2.1.391',
  },
  description: 'For book-fold printing Word places the automatic gutter at the right-margin bisector edge, including reverse book-fold mode.',
});

export const WORD_DEFAULT_LINE_NUMBER_DISTANCE_PT = 18;

export function wordLineNumberDistancePt(authoredDistancePt: number | null | undefined): number {
  return authoredDistancePt ?? WORD_DEFAULT_LINE_NUMBER_DISTANCE_PT;
}

export function wordContinuousSectionRestartDisplayNumber(
  authoredStart: number,
  physicalPageIndex: number,
  firstAppearancePageIndex: number,
): number {
  return authoredStart + physicalPageIndex - firstAppearancePageIndex;
}

export function wordTrailingEmptyMarkAdmissionAllowancePt(input: Readonly<{
  hasContinuationBoundary: boolean;
  inkless: boolean;
  undecorated: boolean;
  keepNext: boolean;
  markReservePt: number;
  pageBottomIsUnreserved: boolean;
  physicalRegionBottomIsActive: boolean;
  hasFollowingInk: boolean;
  followsNextPageSectionBoundary: boolean;
  markExtentPt: number;
  markBelowBaselinePt: number;
}>): number {
  const eligible = !input.hasContinuationBoundary
    && input.inkless
    && input.undecorated
    && !input.keepNext
    && input.markReservePt === 0
    && input.pageBottomIsUnreserved
    && input.physicalRegionBottomIsActive;
  if (!eligible) return 0;
  if (input.followsNextPageSectionBoundary) return Math.max(0, input.markExtentPt);
  return input.hasFollowingInk ? input.markBelowBaselinePt : 0;
}

export function wordBookFoldGutterEdge(): 'right' {
  return 'right';
}
