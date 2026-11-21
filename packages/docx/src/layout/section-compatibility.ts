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
  description: 'ECMA-376 §§17.15.1.11/.13/.64 describe dividing the section page in half. Word instead keeps the authored pgSz (§17.6.13) as every editing page and uses pgSz/@orient for the sheet arrangement. Margins L/R/T/B, gutter G; page 1, 3, … is the mirrored member (page 2, 4, … under bookFoldRevPrinting). Landscape sheet, book fold or two-up: the normal page has left L and right R+G, the mirrored page left R+G and right L; with gutterAtTop the top becomes T+G, the normal page has left L and right R, and the mirrored page keeps that width but starts at R+G (right L-G). Bottom stays B. When G exceeds L that right margin is negative and Word runs the mirrored column past the page edge; library policy keeps body columns inside the page box (page-factory invariant), so the right margin stops at 0 and only those lines wrap earlier. Book-fold headers and footers keep the single-page extent (gutter on the left, mirrorMargins and gutterAtTop as usual); two-up headers and footers follow the imposed page without gutterAtTop. Portrait sheet, two-up: left L and right R on every page, the mirrored page has top B+G and bottom T, the other top T and bottom B+G; gutterAtTop and mirrorMargins have no effect; headers and footers use L/R. Portrait sheet, book fold: the body uses L/R/T/B with no gutter or mirroring; headers and footers keep the single-page extent. mirrorMargins left the imposed body unchanged in both measured cases (landscape-sheet book fold, portrait-sheet two-up). A tbRl section binds on the right: its single-page gutter joins the right margin (left L, right R+G, header and footer alike); book fold leaves it on those single-page margins on every page; two-up on a landscape sheet mirrors page 2, 4, … instead (page 1 left L and right R+G, page 2 left R+G and right L, header and footer following); two-up on a portrait sheet is unchanged. A tbLrV section binds on the right in the same way, although its lines advance from the left margin: single page left L and right R+G (gutterAtTop moves the gutter to the top instead), mirrorMargins page 2 left R+G, landscape-sheet two-up page 1 left L and page 2 left R+G, landscape-sheet book fold left L on every page. Vertical-section headers and footers take these physical margins. Evidence: 44 Windows Word PDF exports in three runs, with margins chosen so that no two candidate rules coincide: book fold and reverse book fold on landscape sheets of portrait and landscape shape, gutterAtTop with both, book fold on a portrait sheet, two-up on portrait and landscape sheets of either shape with and without gutterAtTop, mirrorMargins with book fold and two-up, gutters smaller than, equal to and larger than the left margin under book fold and landscape two-up with and without gutterAtTop, tbRl sections under book fold, landscape and portrait two-up and no imposition, tbLrV sections (gutter 36pt, margins L40 R90 T60 B80) with no imposition, mirrorMargins, gutterAtTop, landscape two-up and landscape book fold, and single-page controls; plus a corpus document whose half-Legal booklet pages export at the authored 7in x 8.5in. Not measured: reverse book fold or gutterAtTop on a portrait sheet, two-up on a landscape sheet with mirrorMargins, rtlGutter with any imposition, gutterAtTop or mirrorMargins in a tbRl section, gutterAtTop with mirrorMargins and portrait-sheet imposition of a tbLrV section, book fold of a tbRl section on a portrait sheet, other vertical text directions (btLr, tbRlV, lrTbV keep the left-bound rules), and page parity after a page-number restart (physical page parity is used).',
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

export const WORD_SECTION_TBRLV_TBLRV_PAGE_FRAME = defineCompatibilityRule({
  id: 'word-section-tbrlv-tblrv-page-frame',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'vertical-section-upright-block-table',
    application: 'Microsoft Word',
    version: '16.0.20527',
    platform: 'Windows',
  },
  description: 'Part 4 §14.11.7 names tbRlV as lines flowing from right to left with rotated glyphs. Word lays a section-level tbRlV out in the page frame of tbLrV (word-section-tblrv-left-page-frame): the first line strip starts at the physical left margin and later lines advance rightward. Evidence: a Windows Word 16.0.20527 PDF of a synthetic A4 one-column tbRlV section (margins left 40, right 90, top 60, bottom 80pt, 18pt line grid) holding a paragraph, a 180pt-wide block table and two paragraphs: the strips are at x 40–58, 58–76 and 76–94pt, the table at x 58–238pt, and the glyphs (upright East Asian, Latin turned clockwise) are those of the tbLrV control with the same content, coordinate for coordinate. Outside this evidence: gutter and mirrored margins, headers and footers, columns and two-page imposition of a tbRlV section, and tbRlV on table cells, which keeps its own handling.',
});

export const WORD_SECTION_TBLRV_LEFT_PAGE_FRAME = defineCompatibilityRule({
  id: 'word-section-tblrv-left-page-frame',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'vertical-section-body-margins',
    application: 'Microsoft Word',
    version: '16.0.20527',
    platform: 'Windows',
  },
  description: 'A section-level tbLrV (Part 4 §14.11.7, equivalent to lrV: lines flow from left to right) lays its body out in the vertical-lr page frame: the first line strip starts at the physical left margin and the top margin, and later lines advance rightward, so the logical top/right/bottom/left margins are the physical left/bottom/right/top. Evidence: Windows Word PDF exports of A4 one-column sections with an 18pt line grid and distinct margins (left 40, right 90, top 60, bottom 80pt): tbLrV put its first line strip at x 40–58pt with its top at y 60pt and advanced right by 18pt per line, while the tbRl and btLr controls started at the right margin (strip 487.3–505.3pt). Line-strip content is word-section-tblrv-line-strip; gutter and imposition are word-two-page-sheet-imposition; upright tables, the line grid of headers and footers, and column-break paragraph marks of a tbLrV section are outside this evidence.',
});

export const WORD_SECTION_TBLRV_LINE_STRIP = defineCompatibilityRule({
  id: 'word-section-tblrv-line-strip',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'vertical-section-line-strips',
    application: 'Microsoft Word',
    version: '16.0.20527',
    platform: 'Windows',
  },
  description: 'Inside each line strip a tbLrV section has the arrangement of a tbRl section: East Asian glyphs upright, Latin text turned clockwise, ruby and emphasis marks on the physical right of the base text, and the underline at the same place across the glyph column; only the order of the strips differs (left to right from the left margin instead of right to left). The vertical-lr page frame is a reflection, so each tbLrV line is reflected across its own block band before that frame applies, which composes to the tbRl quarter turn translated onto the strip. Evidence: Windows Word PDF exports of paired tbLrV and tbRl A4 sections on a 36pt line grid with Latin text, full-width brackets, an underlined run, dot emphasis marks and a ruby run: every glyph, ruby and emphasis mark and the underline sat at the same offset from its strip edge in both directions (base glyph columns centred in the strip, a ruby line\'s base column 2.6pt left of centre), and the glyph and Latin text directions matched. Where the underline of vertical text lies across the glyph column (the physical right edge of the column in these controls) is the same question in both directions and is not settled by this rule. Inline pictures of both directions are word-vertical-inline-picture-upright-extent.',
});

export const WORD_VERTICAL_INLINE_PICTURE_UPRIGHT_EXTENT = defineCompatibilityRule({
  id: 'word-vertical-inline-picture-upright-extent',
  evidence: {
    kind: 'office-observation',
    syntheticFixtureId: 'vertical-section-inline-picture',
    application: 'Microsoft Word',
    version: '16.0.20527',
    platform: 'Windows',
  },
  description: 'An inline picture in a tbRl or tbLrV section stays upright at its §20.4.2.7 wp:extent: its physical height is its advance along the line and its physical width is its extent across the line, so the line strip grows to hold the width. Evidence: Windows Word PDF exports of a 60 x 40pt inline picture in tbRl and tbLrV A4 sections on an 18pt line grid: the picture painted upright and unmirrored at 60 x 40pt, centred in a 72pt (four-pitch) strip, and the following text started 40pt after the picture start.',
});
