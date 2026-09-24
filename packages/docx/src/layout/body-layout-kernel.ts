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

export type BodyTableContinuationCursor =
  | Readonly<{
      kind: 'table';
      cursor: TableFragmentCursor;
      floatingContinuationFrame?: 'fresh-text' | 'authored';
    }>
  | Readonly<{ kind: 'adjacent-table-group'; cursor: AdjacentTableGroupCursor }>;

export interface BodyTableAcquisitionInput {
  readonly input: Readonly<{ kind: 'table'; source: SourceRef }> | BodyAdjacentTableGroupInput;
  readonly location: BodyAcquisitionLocation;
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
}

export interface NoteLayoutAcquisitionInput {
  readonly kind: 'footnote' | 'endnote';
  readonly referenceIds: readonly string[];
  readonly pageIndex: number;
  readonly section: DeepReadonly<SectionLayoutContext>;
  readonly container: FlowContainer;
  readonly firstOnPage: boolean;
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
}

export interface PageAnchorPrescanInput {
  readonly anchors: readonly Readonly<{
    occurrenceId: string;
    paragraphSource: SourceRef;
  }>[];
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
