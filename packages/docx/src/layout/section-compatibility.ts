import { defineCompatibilityRule } from './compatibility.js';

export const WORD_DEFAULT_LINE_NUMBER_DISTANCE = defineCompatibilityRule({
  id: 'word-default-line-number-distance',
  evidence: {
    kind: 'library-policy',
    reference: "packages/docx/src/layout/compatibility.test.ts#uses Word's observed 18pt line-number distance only when omitted",
  },
  description: 'ECMA-376 §17.6.8 leaves an omitted line-number distance implementation-defined. Preserve Word-compatible 18pt placement only when the authored distance is absent. Library policy, no Office measurement: 18pt default line-number distance; ECMA-376 §17.6.8 leaves it implementation-defined (8e3212767).',
});

export const WORD_CONTINUOUS_SECTION_PAGE_NUMBER_RESTART = defineCompatibilityRule({
  id: 'word-continuous-section-page-number-restart',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'continuous-section-page-number-restart',
    application: 'Microsoft Word',
    version: '16.0.20430',
    platform: 'Windows',
  },
  description: 'Issue #804 records that Word anchors a continuous section page-number restart to the first physical page containing that section body content, even when another section owns the page top. An empty same-page region does not consume the restart. Evidence: issue #804 / PR #818 fixture, and the #1896 controls in Windows Word 16.0.20430 (a shared first page, a section starting near the page bottom, a section first painting on the next page, and a one-page document).',
});

export const WORD_TRAILING_EMPTY_MARK_BASELINE_ADMISSION = defineCompatibilityRule({
  id: 'word-trailing-empty-mark-baseline-admission',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'trailing-empty-mark-grid-page-edge',
    application: 'Microsoft Word',
    version: '16.0.20430',
    platform: 'Windows',
  },
  description: 'At the unreserved physical body edge, Word admits an undecorated non-terminal empty paragraph mark by its baseline when later ink follows in the same flow and the mark does not occupy a document-grid line cell. Windows Word 16.0.20430 trr3 family D (272 controls: modes 14/15, lines/linesAndChars, 16.4pt pitch, 11pt MS Mincho, zero/one preceding empty mark) instead admits grid marks by their centered natural/authored content box: 15pt remaining moves, 15.5–16.35pt stays, and 16.4pt fully fits. The single-character counterexamples follow the existing plain-text content-box rule too. Grid-only trailing leading may overhang; the complete advance remains retained. Other fonts/pitches use the same resolved metrics, not a measured per-font allowance. Next-page section marks retain their separate blank-page suppression. Evidence: issue #981 / PR #989 and trr3 family D.',
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

export const WORD_TWO_PAGE_SHEET_IMPOSITION = defineCompatibilityRule({
  id: 'word-two-page-sheet-imposition',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'book-fold-two-on-one-imposition',
    application: 'Microsoft Word',
    version: '16.0.20527',
    platform: 'Windows',
  },
  description: 'ECMA-376 §§17.15.1.11/.13/.64 describe dividing the section page in half. Word instead keeps the authored pgSz (§17.6.13) as every editing page and uses pgSz/@orient for the sheet arrangement. Margins L/R/T/B, gutter G; page 1, 3, … is the mirrored member (page 2, 4, … under bookFoldRevPrinting). Landscape sheet, book fold or two-up: the normal page has left L and right R+G, the mirrored page left R+G and right L; with gutterAtTop the top becomes T+G, the normal page has left L and right R, and the mirrored page keeps that width but starts at R+G (right L-G). Bottom stays B. Book-fold headers and footers keep the single-page extent (gutter on the left, mirrorMargins and gutterAtTop as usual); two-up headers and footers follow the imposed page without gutterAtTop. Portrait sheet, two-up: left L and right R on every page, the mirrored page has top B+G and bottom T, the other top T and bottom B+G; gutterAtTop and mirrorMargins have no effect; headers and footers use L/R. Portrait sheet, book fold: the body uses L/R/T/B with no gutter or mirroring; headers and footers keep the single-page extent. mirrorMargins left the imposed body unchanged in both measured cases (landscape-sheet book fold, portrait-sheet two-up). Evidence: 28 Windows Word PDF exports (two runs; margins chosen so no two candidate rules coincide) covering book fold and reverse book fold on landscape sheets of portrait and landscape shape, gutterAtTop with both, book fold on a portrait sheet, two-up on portrait and landscape sheets of either shape with and without gutterAtTop, mirrorMargins with book fold and two-up, and single-page controls; plus a corpus document whose half-Legal booklet pages export at the authored 7in x 8.5in. Not measured: reverse book fold or gutterAtTop on a portrait sheet, two-up on a landscape sheet with mirrorMargins, rtlGutter with any imposition, and page parity after a page-number restart (physical page parity is used).',
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
  markOnLineGrid: boolean;
  markGridAfterLeadingPt?: number;
}>): number {
  const eligible = !input.hasContinuationBoundary
    && input.inkless
    && input.undecorated
    && !input.keepNext
    && input.markReservePt === 0
    && input.pageBottomIsUnreserved
    && input.physicalRegionBottomIsActive;
  if (!eligible) return 0;
  // `word-section-mark-blank-page-suppression`
  if (input.followsNextPageSectionBoundary) return Math.max(0, input.markExtentPt);
  // Windows trr3 family D contradicts the earlier full-cell empty-mark
  // inference. Keep the centered natural/authored content box inside the body,
  // just as WORD_GRID_LINE_CONTENT_FIT_ADMISSION does for plain text. The full
  // grid advance still positions successors; baseline-only admission is too lax.
  if (input.markOnLineGrid) {
    return input.hasFollowingInk ? input.markGridAfterLeadingPt ?? 0 : 0;
  }
  return input.hasFollowingInk ? input.markBelowBaselinePt : 0;
}
