import type {
  KinsokuRules,
  NumberFormat,
  ResolvedLocalFontMetric,
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
import type { LayoutServices } from './types.js';

/** One acquired body-table occurrence and the point-space placement facts that
 * bind it to the current retained-layout session. */
export interface RetainedTableRecord {
  readonly sourceIndex: number;
  readonly acquisition: RetainedTableAcquisition;
  readonly contentWidthPt: number;
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
  resolvedLocalFonts: Readonly<Record<string, ResolvedLocalFontMetric>>;
  layoutServices?: LayoutServices;
  retainedTableAcquisition:
    RetainedTableAcquisitionDependencies<BodyAcquisitionState>;
  acquireCompleteTextBoxStory?: CompleteTextBoxStoryAcquirer;
  retainedTablesBySourceIndex: Map<number, RetainedTableRecord>;
  kinsoku: KinsokuRules;
  defaultTabPt: number;
  currentDateMs?: number;
  noteNumbers?: Map<string, number>;
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
  | 'noteNumbers'
  | 'noteReferenceNumber'
  | 'verticalCJK'
  | 'verticalAllRotated'
  | 'containerShading'
>>;
