import type { NoteNumbering } from '../line-layout.js';
import type {
  KinsokuRules,
  NumberFormat,
  ResolvedFontMetric,
} from '@silurus/ooxml-core';
import type { FloatRect } from '../float-layout.js';
import type {
  DocumentLayoutSettings,
  SectionLayoutContext,
  StoryContext,
} from '../layout-context.js';
import type { ParagraphLayoutSource } from './text.js';
import type {
  MeasurementTextContext,
  VerticalGlyphMeasurementService,
} from './measurement-capabilities.js';
import type { BodyAcquisitionInputProjections } from './acquisition-input-projections.js';
import type { CompleteTextBoxStoryAcquirer } from './paragraph.js';
import type {
  RetainedTableAcquisition,
  RetainedTableAcquisitionDependencies,
} from './table-acquisition.js';
import type { LayoutRect, LayoutServices, NativeSectionFlow } from './types.js';

/** One acquired body-table occurrence and the point-space placement facts that
 * bind it to the current retained-layout session. */
export interface RetainedTableRecord {
  readonly sourceIndex: number;
  readonly acquisition: RetainedTableAcquisition;
  readonly contentWidthPt: number;
  /** Precomputed retainedTableAcquisitionIsReusableAcrossPages(acquisition);
   *  stored so the per-flow-region reuse check stays O(1). */
  readonly reusableAcrossPages: boolean;
  readonly anchorYPt: number;
}

/** Physical page facts needed while projecting DrawingML anchors from a
 * vertical section's upright page into its logical acquisition frame. */
export interface PhysicalAnchorFrame {
  readonly pageWidth: number;
  readonly pageHeight: number;
  readonly marginLeft: number;
  readonly marginRight: number;
  readonly marginTop: number;
  readonly marginBottom: number;
  readonly physicalPageWidthPt: number;
  /** Native counter-clockwise frame; absent for the Transitional vertical
   * frame, whose established clockwise projection is unchanged. */
  readonly nativeSectionFlow?: NativeSectionFlow;
}

/** The authority a nested text-box story inherits from the state that owns it. */
export interface CompleteTextBoxStoryOwner {
  readonly sectionLayout: SectionLayoutContext;
  readonly pageIndex: number;
}

/** Read-only page/container geometry consumed by DrawingML anchor placement. */
export interface AnchorGeometryContext {
  /** All geometry is expressed in authored points. */
  readonly contentX: number;
  readonly contentW: number;
  readonly pageH: number;
  readonly marginLeft: number;
  readonly marginRight: number;
  readonly marginTop: number;
  readonly marginBottom: number;
  readonly pageWidth: number;
}

/** Mutable exclusion registry owned by one layout-acquisition flow domain. */
export interface FloatRegistrationState extends AnchorGeometryContext {
  floats: FloatRect[];
  floatParaSeq: number;
}

/** DrawingML anchor capability: float registration plus vertical-page
 * projection and page-start pre-scan ownership. */
export interface AnchorFloatRegistrationState extends FloatRegistrationState {
  pageAnchorPrescanned?: Set<ParagraphLayoutSource>;
  /** WORD_LATER_ANCHOR_EARLIER_LINE_WRAP: object frames of paragraph-relative
   * drawings carried to this page, keyed by anchor occurrence. The anchor
   * paragraph keeps the carried frame instead of re-resolving it. */
  frozenAnchorFrames?: Map<string, LayoutRect>;
  verticalCJK?: boolean;
  verticalAllRotated?: boolean;
  verticalPhys?: PhysicalAnchorFrame;
}

/**
 * Mutable cursor owned by retained-layout acquisition.
 *
 * This state may measure text and register exclusions, but it has no paint
 * resources or drawing-mode switch. Body paint consumes the retained result
 * produced from this cursor; it never reuses or mutates the cursor itself.
 */
export interface BodyAcquisitionState extends AnchorFloatRegistrationState {
  /** Synchronous text metrics with no backing-canvas or paint surface. */
  ctx: MeasurementTextContext;
  /** Vertical glyph metrics bound to the same concrete measurement context,
   * without exposing its backing canvas to acquisition consumers. */
  verticalGlyphMeasurement: VerticalGlyphMeasurementService;
  /** Required parser-to-layout fact projections. */
  acquisitionInputs: BodyAcquisitionInputProjections;
  /** Current logical text container in authored points. */
  contentX: number;
  contentW: number;
  y: number;
  pageH: number;
  pageIndex: number;
  totalPages: number;
  displayPageNumber?: number;
  pageNumberFormat?: NumberFormat;
  marginLeft: number;
  marginRight: number;
  marginTop: number;
  marginBottom: number;
  pageWidth: number;
  layoutSettings: DocumentLayoutSettings;
  sectionLayout: SectionLayoutContext;
  storyContext: StoryContext;
  docEastAsian: boolean;
  fontFamilyClasses: Record<string, string>;
  resolvedLocalFonts: Readonly<Record<string, ResolvedFontMetric>>;
  layoutServices?: LayoutServices;
  retainedTableAcquisition:
    RetainedTableAcquisitionDependencies<BodyAcquisitionState>;
  /** Session-level nested story acquisition. The section and page of the
   * state that contains the text box (body location, table cell or story
   * candidate) are passed explicitly, so they, and the frame derived from
   * them, govern the nested story; callers bind it per owner. */
  acquireCompleteTextBoxStory?: (
    owner: CompleteTextBoxStoryOwner,
    request: Parameters<CompleteTextBoxStoryAcquirer>[0],
  ) => ReturnType<CompleteTextBoxStoryAcquirer>;
  /** A table cell paragraph re-acquired by its table's page placement (its
   * text boxes hold page-placed content): the page translation its host flow
   * receives there (ParagraphAcquisitionOptions.hostFlowPageTranslationPt). */
  cellHostFlowPageTranslationPt?: Readonly<{ xPt: number; yPt: number }>;
  cellParagraphAnchorReferenceDeltaPt?: number;
  /** A text box story and its tables: the page frames its page-owned anchor
   * axes keep and the translation its flow receives to reach them
   * (story-page-frames.ts storyAnchorPageFrames); null when its box carries
   * no page band into it. Absent outside text box stories. */
  textBoxStoryHostFrames?: Readonly<{
    frames: Readonly<{ page: LayoutRect; margin: LayoutRect }>;
    flowPt: Readonly<{ xPt: number; yPt: number }>;
  }> | null;
  retainedTablesBySourceIndex: Map<number, RetainedTableRecord>;
  /** Set only while acquiring a body table that is placed upright in the
   * physical page (identity paint root); its cells take that frame. */
  uprightPhysicalTable?: boolean;
  kinsoku: KinsokuRules;
  defaultTabPt: number;
  currentDateMs?: number;
  /** ECMA-376 §17.13.5 tracked-change view (from the selected LayoutOptions):
   * true = markup view, absent/false = final view (deletions hidden). */
  showTrackedChanges?: boolean;
  /** Markup-view author → stable palette colour resolver; built once per
   * layout session (only when showTrackedChanges is set). */
  revisionAuthorColor?: (author?: string) => string;
  noteNumbers?: Map<string, number>;
  /** §17.11.17/.18/.20 display format and start for each note kind. */
  noteNumbering?: NoteNumbering;
  noteReferenceNumber?: number;
  containerShading?: string | null;
}

/** Immutable measurement authority passed to text/table measurement helpers.
 * Cursor movement and float mutation stay on {@link BodyAcquisitionState}. */
export type BodyMeasurementContext = Readonly<Pick<
  BodyAcquisitionState,
  | 'ctx'
  | 'verticalGlyphMeasurement'
  | 'acquisitionInputs'
  | 'contentX'
  | 'pageH'
  | 'pageWidth'
  | 'pageIndex'
  | 'totalPages'
  | 'displayPageNumber'
  | 'pageNumberFormat'
  | 'layoutSettings'
  | 'sectionLayout'
  | 'storyContext'
  | 'docEastAsian'
  | 'fontFamilyClasses'
  | 'resolvedLocalFonts'
  | 'layoutServices'
  | 'kinsoku'
  | 'defaultTabPt'
  | 'currentDateMs'
  | 'showTrackedChanges'
  | 'revisionAuthorColor'
  | 'noteNumbers'
  | 'noteNumbering'
  | 'noteReferenceNumber'
  | 'verticalCJK'
  | 'verticalAllRotated'
  | 'containerShading'
>>;
