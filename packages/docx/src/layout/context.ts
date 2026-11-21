import type {
  ColumnGeom,
  ColumnsSpec,
  HeadersFooters,
  LineNumbering,
  PageNumType,
  SectionGeom,
  SectionProps,
} from '../types.js';
import {
  computeSectionColumns,
  type SectionLayoutContext,
  type SectionGridContext,
} from '../layout-context.js';
import type { DeepReadonly, SectionFlow, WritingMode } from './types.js';
import type { AnchorFrameRect } from './anchor-frame.js';
import {
  logicalPageExtent,
  logicalToPhysicalMatrix,
  physicalToLogicalMatrix,
  sectionWritingMode,
  transformRectEdges,
  uprightPhysicalExtent,
} from './coordinate-space.js';

/**
 * Section facts that must change atomically at a page-flow boundary. Keeping the
 * occurrence identity beside geometry prevents two equal-looking consecutive
 * sections from being mistaken for the same section by page-number/header logic.
 */
export interface PageFlowSectionContext {
  readonly sectionOccurrenceId: string;
  readonly geometry: Readonly<SectionGeom>;
  readonly columns: readonly Readonly<ColumnGeom>[];
  readonly textDirection: string;
  /** Canonical producer section flow; absent for every OOXML section. */
  readonly sectionFlow?: SectionFlow;
  /** §17.6.1 section direction controls newspaper-column population order. */
  readonly sectionBidi: boolean;
  readonly grid: Readonly<SectionGridContext>;
}

export function createPageFlowSectionContext(input: Readonly<{
  sectionOccurrenceId: string;
  geometry: SectionGeom;
  columns: readonly Readonly<ColumnGeom>[];
  textDirection: string;
  sectionFlow?: SectionFlow;
  sectionBidi?: boolean;
  grid?: Readonly<SectionGridContext>;
}>): PageFlowSectionContext {
  if (input.sectionOccurrenceId.length === 0) {
    throw new RangeError('Section occurrence id must not be empty');
  }
  if (input.columns.length === 0) {
    throw new RangeError('A page-flow section requires at least one column');
  }
  return Object.freeze({
    sectionOccurrenceId: input.sectionOccurrenceId,
    geometry: Object.freeze({ ...input.geometry }),
    columns: Object.freeze(input.columns.map((column) => Object.freeze({ ...column }))),
    textDirection: input.textDirection,
    ...(input.sectionFlow ? { sectionFlow: input.sectionFlow } : {}),
    sectionBidi: input.sectionBidi ?? false,
    grid: Object.freeze(input.grid ?? {
      kind: 'none',
      linePitchPt: null,
      charSpacePt: null,
    }),
  });
}

/** §17.6.11 permits signed top/bottom margins, but body flow uses their distance
 * from the page edge; the sign controls header/footer overlap separately. */
export function sectionContentStartBlockPt(section: PageFlowSectionContext): number {
  return sectionBodyInsetPt(section.geometry.marginTop);
}

export function sectionContentEndBlockPt(section: PageFlowSectionContext): number {
  return section.geometry.pageHeight - sectionBodyInsetPt(section.geometry.marginBottom);
}

/** Signed top/bottom margins retain overlap policy; body placement uses distance. */
export function sectionBodyInsetPt(marginPt: number): number {
  return Math.abs(marginPt);
}

/**
 * Physical-to-logical mapping of a section's page box and §17.6.11 `w:pgMar`
 * margins, through the section frame's own canonical matrix
 * ({@link physicalToLogicalMatrix}). The writing mode is the only authority:
 * the frame that lays the body out is the frame that paints it, so a physical
 * margin owns the logical edge it is painted on.
 *
 * - vertical-rl (tbRl, tbRlV, and btLr under
 *   `word-section-btlr-tbrl-page-frame`): lines advance right to left, so the
 *   logical top/right/bottom/left margins are the physical right/bottom/left/top.
 * - vertical-lr (tbLrV, Part 4 §14.11.7 ≡ Part 1 §17.18.93 lrV, lines flowing
 *   from left to right): the logical top/right/bottom/left margins are the
 *   physical left/bottom/right/top (`word-section-tblrv-left-page-frame`).
 * - sideways-lr (a bottomToTop section flow): logical left/right/top/bottom
 *   are the physical bottom/top/left/right margins.
 * - horizontal-tb: the identity.
 */
export function logicalSectionGeometry(
  physical: SectionGeom,
  writingMode: WritingMode,
): SectionGeom {
  const page = { widthPt: physical.pageWidth, heightPt: physical.pageHeight };
  const extent = logicalPageExtent(page, writingMode);
  const edges = transformRectEdges(physicalToLogicalMatrix(writingMode, page), {
    top: physical.marginTop,
    right: physical.marginRight,
    bottom: physical.marginBottom,
    left: physical.marginLeft,
  });
  return {
    pageWidth: extent.widthPt,
    pageHeight: extent.heightPt,
    marginLeft: edges.left,
    marginTop: edges.top,
    marginRight: edges.right,
    marginBottom: edges.bottom,
    headerDistance: physical.headerDistance,
    footerDistance: physical.footerDistance,
  };
}

/** Exact inverse of {@link logicalSectionGeometry} for the same writing mode. */
export function physicalSectionGeometry(
  logical: SectionGeom,
  writingMode: WritingMode,
): SectionGeom {
  const page = uprightPhysicalExtent(
    { widthPt: logical.pageWidth, heightPt: logical.pageHeight },
    writingMode,
  );
  const edges = transformRectEdges(logicalToPhysicalMatrix(writingMode, page), {
    top: logical.marginTop,
    right: logical.marginRight,
    bottom: logical.marginBottom,
    left: logical.marginLeft,
  });
  return {
    pageWidth: page.widthPt,
    pageHeight: page.heightPt,
    marginTop: edges.top,
    marginRight: edges.right,
    marginBottom: edges.bottom,
    marginLeft: edges.left,
    headerDistance: logical.headerDistance,
    footerDistance: logical.footerDistance,
  };
}

/**
 * The upright physical `page` and `margin` anchor reference frames of a
 * vertical section, rebuilt from its section-logical frames.
 *
 * ECMA-376 §20.4.3.4/§20.4.3.5 (ST_RelFromH/ST_RelFromV `page`, `margin`) name
 * the page and its §17.6.11 `w:pgMar` text margins. Under
 * `word-vertical-section-physical-drawing-layer` positionH/positionV of a
 * vertical section resolve on the upright physical page, while retained body
 * acquisition holds the page box and margins in the section's logical frame.
 * The physical margins are recovered with {@link physicalSectionGeometry}, the
 * exact inverse of the {@link logicalSectionGeometry} mapping that produced
 * the logical frame, for every vertical direction (the clockwise edge mapping
 * of tbRl/tbRlV/tbLrV/btLr, and the counter-clockwise one of a bottomToTop
 * section flow). The paint matrix of a direction is not that inverse in
 * general: tbLrV paints through the vertical-lr reflection, which would put
 * the physical right margin on the left. The derived left/right/top/bottom
 * and inside/outside margin edges follow from these two frames.
 *
 * Column, character, paragraph and line frames are outside this mapping.
 */
export function uprightPhysicalPageFrames<
  T extends { readonly page: AnchorFrameRect | null; readonly margin: AnchorFrameRect | null },
>(frames: T, sectionFlow?: SectionFlow | null): T {
  const { page, margin } = frames;
  if (!page) return frames;
  const logical: SectionGeom = {
    pageWidth: page.widthPt,
    pageHeight: page.heightPt,
    marginLeft: margin ? margin.xPt - page.xPt : 0,
    marginTop: margin ? margin.yPt - page.yPt : 0,
    marginRight: margin ? page.xPt + page.widthPt - margin.xPt - margin.widthPt : 0,
    marginBottom: margin ? page.yPt + page.heightPt - margin.yPt - margin.heightPt : 0,
    headerDistance: 0,
    footerDistance: 0,
  };
  const physical = physicalSectionGeometry(logical, sectionFlow);
  return {
    ...frames,
    page: { xPt: 0, yPt: 0, widthPt: physical.pageWidth, heightPt: physical.pageHeight },
    margin: margin && {
      xPt: physical.marginLeft,
      yPt: physical.marginTop,
      widthPt: Math.max(0, physical.pageWidth - physical.marginLeft - physical.marginRight),
      heightPt: Math.max(0, physical.pageHeight - physical.marginTop - physical.marginBottom),
    },
  };
}

export function isVerticalSectionDirection(textDirection: string): boolean {
  return textDirection === 'tbRl'
    || textDirection === 'tbRlV'
    || textDirection === 'tbLrV'
    || textDirection === 'btLr';
}

/** ECMA-376 §17.18.65 ST_PageOrientation of the authored pgSz. */
export type PageSheetOrientation = 'portrait' | 'landscape';

export interface SectionPageLayoutPolicy {
  readonly physicalGeometry: Readonly<SectionGeom>;
  readonly columns: DeepReadonly<ColumnsSpec> | null;
  readonly textDirection: string;
  readonly gutterPt: number;
  readonly rtlGutter: boolean;
  /** §17.6.13 pgSz/@orient: the sheet arrangement of a two-page imposition. */
  readonly sheetOrientation: PageSheetOrientation;
  readonly mirrorMargins: boolean;
  readonly gutterAtTop: boolean;
  readonly bookFoldPrinting: boolean;
  readonly bookFoldRevPrinting: boolean;
  readonly printTwoOnOne: boolean;
}

interface PageMargins {
  readonly marginTop: number;
  readonly marginRight: number;
  readonly marginBottom: number;
  readonly marginLeft: number;
}

interface ImposedPageMargins {
  readonly body: PageMargins;
  /** Inline extent of the header and footer stories. */
  readonly headerFooter: Pick<PageMargins, 'marginLeft' | 'marginRight'>;
}

function requirePhysicalPageIndex(pageIndex: number): void {
  if (!Number.isInteger(pageIndex) || pageIndex < 0) {
    throw new RangeError('Physical page index must be a non-negative integer');
  }
}

/** §17.6.11 margins with the §17.15.1.50/§17.6.16 gutter and §17.15.1.57
 * mirrored margins, for a section printed one page per sheet. */
function singlePageMargins(policy: SectionPageLayoutPolicy, pageIndex: number): PageMargins {
  let { marginTop, marginRight, marginBottom, marginLeft } = policy.physicalGeometry;
  if (policy.gutterAtTop && !policy.mirrorMargins) marginTop += policy.gutterPt;
  // A tbRl section binds on the right: its gutter joins the right margin.
  // `word-two-page-sheet-imposition`
  else if (policy.rtlGutter || rightBoundSection(policy)) marginRight += policy.gutterPt;
  else marginLeft += policy.gutterPt;
  if (policy.mirrorMargins && pageIndex % 2 === 1) {
    [marginLeft, marginRight] = [marginRight, marginLeft];
  }
  return { marginTop, marginRight, marginBottom, marginLeft };
}

/** A §17.6.20 tbRl section, whose lines advance right to left. */
function rightBoundSection(policy: SectionPageLayoutPolicy): boolean {
  return policy.textDirection === 'tbRl';
}

/**
 * ECMA-376 §§17.15.1.11/.13 (book fold) and §17.15.1.64 (two pages per sheet)
 * arrange two editing pages on one printed sheet. The authored pgSz (§17.6.13)
 * stays the editing page and pgSz/@orient selects how the two pages share the
 * sheet; every branch below is the compatibility rule
 * `word-two-page-sheet-imposition`, whose evidence and limits live with it.
 */
function imposedPageMargins(policy: SectionPageLayoutPolicy, pageIndex: number): ImposedPageMargins {
  const single = singlePageMargins(policy, pageIndex);
  const bookFold = policy.bookFoldPrinting || policy.bookFoldRevPrinting;
  if (!bookFold && !policy.printTwoOnOne) return { body: single, headerFooter: single };
  const {
    marginTop: top, marginRight: right, marginBottom: bottom, marginLeft: left,
  } = policy.physicalGeometry;
  const gutter = policy.gutterPt;
  const rightBound = rightBoundSection(policy);
  // A right-bound section is not folded: book fold leaves it on its
  // single-page margins.
  if (bookFold && rightBound) return { body: single, headerFooter: single };
  // Page 1, 3, … of a forward booklet or two-up sheet is the mirrored member;
  // reverse book fold mirrors page 2, 4, … instead.
  const mirrored = (pageIndex % 2 === 0) !== (bookFold && policy.bookFoldRevPrinting);
  if (policy.sheetOrientation === 'portrait') {
    if (bookFold) {
      // Book fold on a portrait sheet: the body takes no gutter and no
      // mirroring; header and footer keep their single-page extent.
      return {
        body: { marginTop: top, marginRight: right, marginBottom: bottom, marginLeft: left },
        headerFooter: single,
      };
    }
    // Two pages stacked on a portrait sheet: the horizontal bisector carries
    // the gutter, so top and bottom mirror while left and right stay as
    // authored. gutterAtTop and mirrorMargins do not apply.
    const horizontal = { marginLeft: left, marginRight: right };
    return {
      body: mirrored
        ? { marginTop: bottom + gutter, marginBottom: top, ...horizontal }
        : { marginTop: top, marginBottom: bottom + gutter, ...horizontal },
      headerFooter: horizontal,
    };
  }
  // Two pages side by side on a landscape sheet: the vertical bisector is the
  // authored right-margin side (`word-book-fold-gutter-right-edge`), which
  // lies on the left of a mirrored page. A right-bound section orders the
  // pair right to left, so its mirrored member is page 2, 4, …. With
  // gutterAtTop the gutter moves to the top, yet a mirrored page still starts
  // at right + gutter and keeps the unmirrored width, so a gutter wider than
  // the left margin runs the column past the page edge. Library policy: body
  // columns stay inside the page box (page-factory invariant), so that right
  // margin stops at the page edge and only those lines wrap earlier.
  const sideBySideMirrored = mirrored !== rightBound;
  const sideBySide = (gutterAtTop: boolean) => sideBySideMirrored
    ? { marginLeft: right + gutter, marginRight: Math.max(0, left - (gutterAtTop ? gutter : 0)) }
    : { marginLeft: left, marginRight: right + (gutterAtTop ? 0 : gutter) };
  return {
    body: {
      marginTop: top + (policy.gutterAtTop ? gutter : 0),
      marginBottom: bottom,
      ...sideBySide(policy.gutterAtTop),
    },
    // Book-fold headers and footers keep the single-page extent; two-up
    // headers and footers follow the imposed page, without gutterAtTop.
    headerFooter: bookFold ? single : sideBySide(false),
  };
}

export function effectivePhysicalSectionGeometry(
  policy: SectionPageLayoutPolicy,
  pageIndex: number,
): SectionGeom {
  requirePhysicalPageIndex(pageIndex);
  return { ...policy.physicalGeometry, ...imposedPageMargins(policy, pageIndex).body };
}

/** Left/right margins bounding the header and footer stories of a physical
 * page; they differ from the body margins only under two-page imposition. */
export function effectiveHeaderFooterInlineMargins(
  policy: SectionPageLayoutPolicy,
  pageIndex: number,
): Readonly<Pick<SectionGeom, 'marginLeft' | 'marginRight'>> {
  requirePhysicalPageIndex(pageIndex);
  return imposedPageMargins(policy, pageIndex).headerFooter;
}

export function resolveSectionContextForPage(
  base: SectionLayoutContext,
  policy: SectionPageLayoutPolicy,
  pageIndex: number,
): SectionLayoutContext {
  const physical = effectivePhysicalSectionGeometry(policy, pageIndex);
  const geometry = isVerticalSectionDirection(policy.textDirection)
    ? logicalSectionGeometry(physical, sectionWritingMode({
        textDirection: policy.textDirection,
        sectionFlow: base.sectionFlow,
      }))
    : physical;
  return Object.freeze({
    ...base,
    geometry: Object.freeze(geometry),
    columns: Object.freeze(computeSectionColumns({
      ...geometry,
      titlePage: false,
      evenAndOddHeaders: false,
      columns: policy.columns,
    } as SectionProps).map((column) => Object.freeze(column))),
  });
}

export interface SectionPlacementFacts {
  readonly sectionId: string;
  readonly sectionBidi: boolean;
  readonly vAlign: string | null;
  readonly lineNumbering: Readonly<LineNumbering> | null;
  readonly docGridType: string | null;
  readonly docGridLinePitch: number | null;
  readonly docGridCharSpace: number | null;
  readonly gutterPt: number | null;
  readonly rtlGutter: boolean | null;
  readonly pageBordersAuthored: boolean;
  readonly pageBorders: Readonly<import('../types.js').PageBorders> | null;
  readonly pageGeometry: Readonly<Partial<SectionGeom>> | null;
}

/**
 * One lexical section occurrence in body order. Equal section properties do not
 * make two occurrences interchangeable: page numbering, title-page selection,
 * and line-number restart rules are occurrence-sensitive.
 */
export interface BodySectionOccurrence {
  readonly sectionOccurrenceId: string;
  readonly ordinal: number;
  /** First body item owned by this occurrence. */
  readonly startBodyIndex: number;
  /** Last body item owned by this occurrence, inclusive. */
  readonly endBodyIndex: number;
  /** The paragraph-owned sectPr marker which terminates this occurrence. */
  readonly markerBodyIndex: number | null;
  readonly final: boolean;
  /** ECMA-376 §17.6.22: how this section starts relative to its predecessor. */
  readonly startType: string;
  readonly columns: ColumnsSpec | null;
  /** Physical §17.6.13/§17.6.11 page box; writing-mode transforms happen later. */
  readonly geometry: SectionGeom;
  readonly textDirection: string | null;
  /** Canonical section flow normalized at the parser-model boundary; absent for
   * OOXML sections. Layout never sees the raw parser wire. */
  readonly sectionFlow?: SectionFlow;
  readonly pageNumType: PageNumType | null;
  readonly headers: HeadersFooters;
  readonly footers: HeadersFooters;
  readonly titlePage: boolean;
  readonly sectionBidi: boolean;
  readonly vAlign: string | null;
  readonly lineNumbering: Readonly<LineNumbering> | null;
  readonly docGridType: string | null;
  readonly docGridLinePitch: number | null;
  readonly docGridCharSpace: number | null;
  readonly gutterPt: number;
  readonly rtlGutter: boolean;
  /** §17.6.13 pgSz/@orient after page-box inheritance. */
  readonly sheetOrientation: PageSheetOrientation;
  readonly pageBordersAuthored: boolean;
  readonly pageBorders: Readonly<import('../types.js').PageBorders> | null;
  readonly placement: SectionPlacementFacts;
}

/** Complete parser-boundary projection; layout does not scan document nodes. */
export interface BodySectionIndexInput {
  readonly bodyLength: number;
  readonly occurrences: readonly BodySectionOccurrence[];
}

export interface BodySectionIndex {
  readonly occurrences: readonly BodySectionOccurrence[];
  /** Accepts body.length as the insertion point owned by the final section. */
  sectionAtBodyIndex(bodyIndex: number): BodySectionOccurrence;
}

export function sectionPageBox(section: SectionProps): SectionGeom {
  return {
    pageWidth: section.pageWidth,
    pageHeight: section.pageHeight,
    marginTop: section.marginTop,
    marginRight: section.marginRight,
    marginBottom: section.marginBottom,
    marginLeft: section.marginLeft,
    headerDistance: section.headerDistance,
    footerDistance: section.footerDistance,
  };
}

export function defaultSectionGeometry(): SectionGeom {
  return {
    pageWidth: 612, pageHeight: 792,
    marginTop: 72, marginRight: 72, marginBottom: 72, marginLeft: 72,
    headerDistance: 36, footerDistance: 36,
  };
}

/** Resolve section-owned facts after parser acquisition, without retaining the
 * document wrapper across the canonical layout boundary. */
export function resolveAcquiredSectionLayoutContext(
  section: SectionProps,
  sectionBidi = false,
  sectionFlow?: SectionFlow,
): SectionLayoutContext {
  const gridKind = section.docGridType === 'lines'
    || section.docGridType === 'linesAndChars'
    || section.docGridType === 'snapToChars'
    ? section.docGridType
    : 'none';
  return Object.freeze({
    geometry: Object.freeze(sectionPageBox(section)),
    columns: Object.freeze(computeSectionColumns(section).map((column) => Object.freeze(column))),
    columnSeparator: section.columns?.sep === true,
    grid: Object.freeze({
      kind: gridKind,
      linePitchPt: section.docGridLinePitch ?? null,
      charSpacePt: section.docGridCharSpace == null ? null : section.docGridCharSpace / 4096,
    }),
    textDirection: section.textDirection ?? 'lrTb',
    ...(sectionFlow ? { sectionFlow } : {}),
    sectionBidi,
    verticalAlignment: section.vAlign ?? 'top',
    ...(section.lineNumbering === null || section.lineNumbering === undefined
      ? {}
      : { lineNumbering: Object.freeze({ ...section.lineNumbering }) }),
  });
}

/**
 * Index a complete §17.6.18/§17.6.17 occurrence projection. Construction is
 * O(occurrences + body), and subsequent source-index lookup is one array access.
 */
export function createBodySectionIndex(input: BodySectionIndexInput): BodySectionIndex {
  if (!Number.isInteger(input.bodyLength) || input.bodyLength < 0 || input.occurrences.length === 0) {
    throw new RangeError('A body section index requires a non-negative length and occurrences');
  }
  const occurrenceOrdinalByBodyIndex = new Array<number>(input.bodyLength + 1);
  let expectedStart = 0;
  input.occurrences.forEach((occurrence, ordinal) => {
    const last = ordinal === input.occurrences.length - 1;
    if (
      occurrence.ordinal !== ordinal
      || occurrence.startBodyIndex !== expectedStart
      || occurrence.endBodyIndex !== (last ? input.bodyLength - 1 : occurrence.markerBodyIndex)
      || occurrence.final !== last
      || (last ? occurrence.markerBodyIndex !== null : occurrence.markerBodyIndex === null)
    ) {
      throw new RangeError(`Invalid section occurrence ${ordinal}`);
    }
    for (
      let ownedIndex = occurrence.startBodyIndex;
      ownedIndex <= occurrence.endBodyIndex;
      ownedIndex += 1
    ) {
      occurrenceOrdinalByBodyIndex[ownedIndex] = ordinal;
    }
    expectedStart = occurrence.endBodyIndex + 1;
  });
  const finalOrdinal = input.occurrences.length - 1;
  occurrenceOrdinalByBodyIndex[input.bodyLength] = finalOrdinal;
  const retainedOccurrences = Object.freeze([...input.occurrences]);
  const retainedOrdinals = Object.freeze(occurrenceOrdinalByBodyIndex);
  return Object.freeze({
    occurrences: retainedOccurrences,
    sectionAtBodyIndex(bodyIndex: number): BodySectionOccurrence {
      if (!Number.isInteger(bodyIndex) || bodyIndex < 0 || bodyIndex >= retainedOrdinals.length) {
        throw new RangeError(`Body index ${bodyIndex} is outside the retained section index`);
      }
      return retainedOccurrences[retainedOrdinals[bodyIndex]!]!;
    },
  });
}
