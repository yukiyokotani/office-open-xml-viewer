import type { SectionLayoutContext } from '../layout-context.js';
import type { LayoutOptions } from './options.js';
import type {
  ParagraphFragmentation,
  ParagraphFragmentCursor,
} from './paragraph-pagination.js';
import type { TableFragmentCursor } from './table-pagination.js';
import type { BodyAdjacentTableGroupInput } from './body-layout-input.js';
import type {
  BodyFlowRegistryDeltaPt,
  BodyFlowRegistrySnapshotPt,
  DeepReadonly,
  LayoutRect,
  LayoutServices,
  ParagraphLayout,
  PointPt,
  SourceRef,
  StoryLayout,
  FlowContainer,
  NoteLayout,
  TableLayout,
} from './types.js';

export class NoteCapacityExceededError extends Error {
  readonly code = 'NOTE_CAPACITY_EXCEEDED' as const;

  constructor(
    readonly kind: 'footnote' | 'endnote',
    readonly pageIndex: number,
    readonly containerId: string,
  ) {
    super(`${kind} story exceeds ${containerId} on page ${pageIndex}`);
    this.name = 'NoteCapacityExceededError';
  }
}

/** Page-local acquisition coordinates supplied by the immutable paginator.
 * This value cannot open a page or choose a transition. */
export interface BodyAcquisitionLocation {
  readonly pageIndex: number;
  readonly columnIndex: number;
  readonly flowDomainId: string;
  readonly section: DeepReadonly<SectionLayoutContext>;
  readonly cursorPt: PointPt;
  readonly availableBounds: LayoutRect;
}

export interface BodyLayoutSessionInput {
  readonly source: SourceRef;
  readonly section: DeepReadonly<SectionLayoutContext>;
  readonly initialLocation: BodyAcquisitionLocation;
}

export interface BodyParagraphAcquisitionInput {
  readonly input: Readonly<{ kind: 'paragraph'; source: SourceRef }>;
  readonly location: BodyAcquisitionLocation;
  readonly availableInlineExtentPt: number;
  readonly suppressSpaceBefore: boolean;
  readonly continuation: ParagraphFragmentCursor;
}

export interface AdjacentTableGroupCursor {
  readonly tableIndex: number;
  readonly sourceRowIndex: number;
  readonly tableCursor?: TableFragmentCursor;
}

/**
 * Entry into the next owner segment of a table projected into cell-owner runs
 * (table-owner-runs.ts). `same-region`: the previous segment completed and
 * this one starts in the same flow region, so the paginator must not open a
 * new region. `fresh-region`: this segment's start moved to a new region, so
 * it continues the source table there and repeats its header rows.
 */
export type OwnerSegmentEntry = 'same-region' | 'fresh-region';

export type BodyTableContinuationCursor =
  | Readonly<{
      kind: 'table';
      cursor: TableFragmentCursor;
      floatingContinuationFrame?: 'fresh-text' | 'authored';
      ownerSegmentEntry?: OwnerSegmentEntry;
    }>
  | Readonly<{
      kind: 'adjacent-table-group';
      cursor: AdjacentTableGroupCursor;
      floatingContinuationFrame?: 'fresh-text';
      ownerSegmentEntry?: OwnerSegmentEntry;
    }>;

export interface BodyTableAcquisitionInput {
  readonly input: Readonly<{ kind: 'table'; source: SourceRef }> | BodyAdjacentTableGroupInput;
  readonly location: BodyAcquisitionLocation;
  /** Initial insertion reference of this fragment in this flow region.
   * Same-region wrap retries advance location without moving an empty
   * paragraph's nominal anchor. New regions/fragments acquire a new reference. */
  readonly unwrappedLocation?: BodyAcquisitionLocation;
  readonly availableInlineExtentPt: number;
  readonly availableBlockExtentPt: number;
  readonly freshPageBlockExtentPt: number;
  readonly cursor?: BodyTableContinuationCursor;
}

export interface AcquiredParagraphBlock {
  readonly layout: ParagraphLayout;
  readonly blockExtentPt: number;
  readonly fragmentation: ParagraphFragmentation;
  readonly uniformRubyAdvancePt?: number;
  readonly markBelowBaselinePt?: number;
  /** True when this mark's line box occupies a §17.6.5 document-grid cell. */
  readonly markOnLineGrid?: boolean;
  readonly flowRegistryDelta?: BodyFlowRegistryDeltaPt;
  readonly placement?: Readonly<{
    coordinateSpace: 'logical-body';
    xPt: number;
    yPt: number;
    sectionFlowOwnership: 'host-flow' | 'page';
  }>;
  /** §17.3.1.11 admits identical adjacent framePr members with their owner. */
  readonly retainedFootnoteReferenceIds?: readonly string[];
  readonly relocationBlockExtentPt?: number;
}

export interface AcquiredTableBlock {
  readonly layout: TableLayout;
  readonly blockExtentPt: number;
  /** Retained table height beyond an Office-clipped page band. */
  readonly unpaintedOverflowPt?: number;
  readonly nextCursor?: BodyTableContinuationCursor | null;
  readonly flowRegistryDelta?: BodyFlowRegistryDeltaPt;
  /** Host-flow extent below the cursor occupied by a zero-advance owner,
   * charged with its footnote reserve exactly like a placed paragraph frame. */
  readonly relocationBlockExtentPt?: number;
  readonly requiresFreshFlowRegion?: boolean;
  readonly retryAtBlockStartPt?: number;
  readonly placement?: Readonly<{
    coordinateSpace: 'logical-body' | 'upright-physical';
    xPt: number;
    yPt: number;
    sectionFlowOwnership?: 'host-flow' | 'page';
  }>;
}

export interface StoryLayoutAcquisitionInput {
  readonly source: SourceRef;
  readonly pageIndex: number;
  readonly section: DeepReadonly<SectionLayoutContext>;
  readonly container: FlowContainer;
  /**
   * The translation its composer applies to the laid-out story (a header,
   * footer or note band), when known before layout. Content only a page
   * position places then resolves on the page through it: a header/footer
   * root cell-owner host anchored to the page or margin states its wrap
   * exclusion in story coordinates through its inverse, and the §17.4.57
   * positioned tables of story tables are given final frames through it.
   */
  readonly bandTranslationPt?: Readonly<{ xPt: number; yPt: number }>;
  /**
   * The destination page's page and margin rectangles stated in this story's
   * own coordinates, when its composer knows the transform that places the
   * laid-out story (a text box: its anchor offset, autofit shift, orientation
   * and final drawing placement; story-page-frames.ts). Positioned tables of
   * its tables then resolve against them in story coordinates and move with
   * every later transform of the story, like its other content.
   */
  readonly pageFrames?: Readonly<{ page: LayoutRect; margin: LayoutRect }>;
}

export interface NoteLayoutAcquisitionInput {
  readonly kind: 'footnote' | 'endnote';
  readonly referenceIds: readonly string[];
  readonly pageIndex: number;
  readonly section: DeepReadonly<SectionLayoutContext>;
  readonly container: FlowContainer;
  readonly firstOnPage: boolean;
  /** The first note on this page resumes a note begun on an earlier page. */
  readonly continuing?: boolean;
  /**
   * Present when the container already states the notes' final page
   * position (document-end endnotes are laid out where they are painted):
   * each note story's band translation, the identity there, so the
   * positioned tables of its tables resolve on the page
   * (`StoryLayoutAcquisitionInput.bandTranslationPt`). Footnotes are stacked
   * after layout and omit it. Note-root tables elect no cell-owner carrier
   * (table-owner-runs.ts tableRowsElectCarriers).
   */
  readonly bandTranslationPt?: Readonly<{ xPt: number; yPt: number }>;
  /**
   * Footnotes: the page-final flow top a previous pass composed for a note
   * whose story is band dependent (header-footer-reserve.ts
   * `footnoteTopsPt`; for a continued tail, moved by its source cut). That
   * note's story is laid out with the band that moves its flow top there,
   * the translation composition then applies to its retained content.
   */
  readonly plannedTopsPt?: Readonly<Record<string, number>>;
}

export interface FollowingBodyBlockMeasurementInput {
  readonly input: Readonly<{ kind: 'paragraph'; source: SourceRef }>
    | BodyTableAcquisitionInput['input'];
  readonly location: BodyAcquisitionLocation;
  readonly availableInlineExtentPt: number;
}

export interface FollowingBodyBlockMeasurement {
  readonly fullExtentPt: number;
  readonly leadContentExtentPt: number;
  /** References painted when the complete block is retained with a keep chain. */
  readonly fullFootnoteReferenceIds?: readonly string[];
  /** References painted by the first indivisible content admitted with keepNext. */
  readonly leadFootnoteReferenceIds?: readonly string[];
  /** Paragraph only: page-owned drawing occurrence keys anchored on each
   * measured line, so keepNext can honour an anchor-line deferral. */
  readonly pageOwnedAnchorKeysByLine?: readonly (readonly string[])[];
}

/** First-placement geometry of a carried paragraph-relative drawing. */
export type CarriedHostAnchor = Readonly<{
  bounds: Readonly<{ xPt: number; yPt: number; widthPt: number; heightPt: number }>;
  exclusionBounds: Readonly<{ xPt: number; yPt: number; widthPt: number; heightPt: number }>;
  horizontalOwnership?: 'page' | 'host';
  verticalOwnership?: 'page' | 'host';
  wrap: 'square' | 'tight' | 'through' | 'topAndBottom';
  wrapSide?: string | null;
  wrapDistances?: Readonly<{ topPt: number; rightPt: number; bottomPt: number; leftPt: number }>;
  wrapPolygon?: readonly Readonly<{ xPt: number; yPt: number }>[];
  topEdgeInclusiveFromYPt?: number;
  anchorLineExemptTopPt?: number;
}>;

export interface PageAnchorPrescanInput {
  readonly anchors: readonly (
    | Readonly<{
        kind: 'drawing';
        occurrenceId: string;
        paragraphSource: SourceRef;
      }>
    | Readonly<{
        kind: 'floating-table';
        occurrenceId: string;
        tableSource: SourceRef;
        bounds: Readonly<{ xPt: number; yPt: number; widthPt: number; heightPt: number }>;
      }>
    | Readonly<{
        /** WORD_LATER_ANCHOR_EARLIER_LINE_WRAP: a paragraph-relative drawing
         * carried from a previous pass with its first-placement geometry. */
        kind: 'host-drawing';
        occurrenceId: string;
        paragraphSource: SourceRef;
        carry: CarriedHostAnchor;
      }>
  )[];
  readonly location: BodyAcquisitionLocation;
  readonly availableInlineExtentPt: number;
}

export interface LineNumberGlyphMetrics {
  readonly widthPt: number;
  readonly ascentPt: number;
  readonly descentPt: number;
  readonly font?: string;
}

export interface BodyLayoutSession {
  readonly hasPaginationFields: boolean;
  measureParagraph(request: BodyParagraphAcquisitionInput): AcquiredParagraphBlock;
  measureTable(request: BodyTableAcquisitionInput): AcquiredTableBlock;
  /** B1 story acquisition. Optional only for synthetic body-only kernels. */
  layoutStory?(request: StoryLayoutAcquisitionInput): StoryLayout;
  /** B1 note acquisition. Optional only for synthetic body-only kernels. */
  layoutNotes?(request: NoteLayoutAcquisitionInput): readonly NoteLayout[];
  measureFollowingBlock(request: FollowingBodyBlockMeasurementInput): FollowingBodyBlockMeasurement;
  prescanPageAnchors?(request: PageAnchorPrescanInput): BodyFlowRegistryDeltaPt | null;
  measureLineNumberGlyph(text: string): LineNumberGlyphMetrics;
  resetPageAcquisition(location: BodyAcquisitionLocation): void;
  moveAcquisitionCursor(location: BodyAcquisitionLocation): void;
  flowRegistrySnapshot(): BodyFlowRegistrySnapshotPt;
  commitFlowRegistryDelta(delta: BodyFlowRegistryDeltaPt): void;
}

/** Document-private acquisition adapter. Page construction and transition
 * policy are intentionally absent from this interface. */
export interface BodyLayoutKernel {
  openBodyLayoutSession(
    input: BodyLayoutSessionInput,
    services: LayoutServices,
    options: LayoutOptions,
  ): BodyLayoutSession;
}
