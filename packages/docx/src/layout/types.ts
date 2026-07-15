import type { SectionLayoutContext } from '../layout-context.js';
import type {
  TextFontSlotPresence,
  TextFontSlots,
  TextLayoutService,
} from './text.js';
import type { ImageMetadataService, MathMetadataService } from './resources.js';
import type { AnchorFrameResult } from './anchor-frame.js';
import type {
  CanvasFontRoute,
  ChartModel,
  DrawingMLShapePaintPlan,
  Duotone,
  Fill,
} from '@silurus/ooxml-core';

export type { TextLayoutService } from './text.js';
export type { ImageMetadataService, MathMetadataService } from './resources.js';

export type LayoutNodeId = string;

export type SourceRef = Readonly<{
  story: 'body' | 'header' | 'footer' | 'footnote' | 'endnote' | 'textbox';
  storyInstance: string;
  path: readonly number[];
}>;

export type DeepReadonly<T> =
  T extends (...args: never[]) => unknown ? T
  : T extends readonly (infer U)[] ? readonly DeepReadonly<U>[]
  : T extends object ? { readonly [K in keyof T]: DeepReadonly<T[K]> }
  : T;

export interface PointPt {
  readonly xPt: number;
  readonly yPt: number;
}

export interface LayoutRect extends PointPt {
  readonly widthPt: number;
  readonly heightPt: number;
}

export type FlowDomainKind =
  | 'body'
  | 'header'
  | 'footer'
  | 'footnote'
  | 'endnote'
  | 'textbox'
  | 'tableCell';

export interface FlowDomain {
  readonly id: string;
  readonly kind: FlowDomainKind;
  readonly bounds: LayoutRect;
}

export interface PageGeometry extends LayoutRect {
  readonly contentTopPt: number;
  readonly contentBottomPt: number;
}

export interface Matrix2DData {
  readonly a: number;
  readonly b: number;
  readonly c: number;
  readonly d: number;
  readonly e: number;
  readonly f: number;
}

export type ClipPathData =
  | Readonly<{ kind: 'rect'; rect: LayoutRect }>
  | Readonly<{ kind: 'polygon'; points: readonly PointPt[] }>;

export interface FlowOwnership {
  readonly flowDomainId: string;
  readonly flowBounds: LayoutRect;
  readonly inkBounds: LayoutRect;
  readonly clipBounds?: LayoutRect;
  readonly advancePt: number;
  readonly ordinaryFlow: boolean;
}

interface LayoutNodeBase extends FlowOwnership {
  readonly id: LayoutNodeId;
  readonly source: SourceRef;
}

export type DrawingPaintCommand =
  | Readonly<{
      kind: 'noop';
    }>
  | Readonly<{
      kind: 'drawingml-shape';
      plan: DeepReadonly<DrawingMLShapePaintPlan>;
    }>
  | Readonly<{
      kind: 'fill-rect';
      rect: LayoutRect;
      fill: string;
    }>
  | Readonly<{
      kind: 'stroke-rect';
      rect: LayoutRect;
      stroke: string;
      lineWidthPt: number;
      dashPt: readonly number[];
    }>
  | Readonly<{
      kind: 'text';
      rect: LayoutRect;
      text: string;
      fill: string;
      fontRoute: CanvasFontRoute;
      fontSizePt: number;
      fontWeight: number;
      fontStyle: 'normal' | 'italic';
      align: 'start' | 'center' | 'end';
      baseline: 'top' | 'middle' | 'alphabetic' | 'bottom';
    }>
  | Readonly<{
      kind: 'watermark-text';
      rect: LayoutRect;
      text: string;
      fill: DeepReadonly<Fill> | null;
      opacity: number;
      rotationDeg: number;
      /** True applies §19.1.2.23 fitshape; false preserves authored font size. */
      fitShape: boolean;
      fontSizePt: number;
      /** Glyph source box relative to span origin x=0 / alphabetic baseline y=0. */
      sourceBounds: LayoutRect;
      spans: readonly Readonly<{
        text: string;
        advancePt: number;
        fontRoute: CanvasFontRoute;
        fontWeight: number;
        fontStyle: 'normal' | 'italic';
      }>[];
    }>
  | Readonly<{
      kind: 'resource';
      resourceKey: string;
      resourceKind: PaintResourceKind;
      rect: LayoutRect;
    }>;

export interface DrawingLayout extends LayoutNodeBase {
  readonly kind: 'drawing';
  readonly transform?: Matrix2DData;
  readonly clip?: ClipPathData;
  readonly commands: readonly DrawingPaintCommand[];
  readonly anchorLayer?: Readonly<{
    occurrenceId: string;
    behindDoc: boolean;
    relativeHeight: number;
    sourceOrder: number;
    horizontalOwnership: 'page' | 'host';
    verticalOwnership: 'page' | 'host';
  }>;
  readonly textBoxIds?: readonly LayoutNodeId[];
}

/** Clone-safe transitional VML facts projected at the parser/model boundary.
 * Presence of the boolean controls distinguishes parser-created false defaults
 * from the stable public ShapeRun compatibility surface. */
export interface VmlTextPathAcquisitionInput {
  readonly string: string;
  readonly fontFamily?: string | null;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly textPathOk?: boolean;
  readonly on?: boolean;
  readonly fitShape?: boolean;
  readonly fitPath?: boolean;
  readonly trim?: boolean;
  readonly xScale?: boolean;
  readonly fontSizePt?: number;
}

export interface TextRange {
  readonly start: number;
  readonly end: number;
}

export type TextDirection = 'ltr' | 'rtl';
export type WritingMode = 'horizontal-tb' | 'vertical-rl' | 'vertical-lr';

export type TextDecorationLayout = Readonly<{
  kind: 'underline' | 'strikethrough' | 'overline';
  /** Original ECMA-376 ST_Underline token when this is a w:u operation. */
  authoredStyle?: string;
  from: PointPt;
  to: PointPt;
  color: string;
  widthPt: number;
  style: 'solid' | 'double' | 'dotted' | 'dashed' | 'wavy';
  /** Final acquired path. Multi-stroke/dash/wave expansion belongs to layout. */
  path?: readonly PointPt[];
  readonly dashPatternPt?: readonly number[];
}>;

export interface RetainedGlyphPaintOperation {
  readonly text: string;
  readonly origin: PointPt;
  readonly fontRoute: CanvasFontRoute;
  readonly fontSizePt: number;
  readonly fontWeight: number;
  readonly fontStyle: 'normal' | 'italic';
  readonly color: TextColorPolicy;
  /** Tight selected-face ink relative to this operation's baseline origin. */
  readonly inkBounds?: Readonly<{
    xMinPt: number;
    xMaxPt: number;
    ascentPt: number;
    descentPt: number;
  }>;
}

export type RetainedMarkPath = Readonly<{
  kind: 'polyline';
  points: readonly PointPt[];
  fill: string | null;
  stroke: string | null;
  strokeWidthPt: number;
}>;

export interface RetainedRunBorderFacts {
  readonly val: string;
  readonly color: string;
  readonly widthPt: number;
  readonly spacePt: number;
  readonly themeColor?: string;
  readonly themeTint?: string;
  readonly themeShade?: string;
  readonly shadow?: boolean;
  readonly frame?: boolean;
}

export interface TextClusterLayout {
  readonly range: TextRange;
  readonly offset: PointPt;
  readonly advancePt: number;
}

export interface TextPaintOp {
  readonly text: string;
  readonly range: TextRange;
  readonly offset: PointPt;
  readonly letterSpacingPt: number;
  readonly scaleX: number;
  readonly direction: TextDirection;
  readonly kerning: 'auto' | 'normal' | 'none';
  readonly writingMode: WritingMode;
  readonly glyphOrientation?: 'sideways' | 'upright';
  /** `kashida` permits acquisition-inserted U+0640 glyphs over one source range. */
  readonly sourceMapping?: 'identity' | 'kashida';
}

export type TextColorPolicy =
  | Readonly<{ kind: 'explicit'; color: string }>
  | Readonly<{ kind: 'auto'; background?: string }>
  | Readonly<{ kind: 'default' }>;

export type RetainedTypographyValue<T> = Readonly<{
  status: 'missing' | 'invalid' | 'valid';
  raw: string | null;
  value: T | null;
}>;

export interface RetainedRunTypographyFacts {
  readonly caps: boolean;
  readonly smallCaps: boolean;
  readonly strike: boolean;
  readonly doubleStrike: boolean;
  readonly verticalAlign: RetainedTypographyValue<'super' | 'sub'>;
  readonly positionPt: RetainedTypographyValue<number>;
  readonly emphasis: RetainedTypographyValue<string>;
  readonly underline?: Readonly<{
    val: RetainedTypographyValue<string>;
    color: RetainedTypographyValue<string>;
    themeColor: RetainedTypographyValue<string>;
    themeTint: RetainedTypographyValue<string>;
    themeShade: RetainedTypographyValue<string>;
  }>;
}

export interface TextPlacement {
  readonly kind: 'text';
  readonly text: string;
  /** Parsed run occurrence retained for destination-page field convergence. */
  readonly sourceRunIndex?: number;
  readonly role?: 'content' | 'numbering-marker' | 'field-result';
  readonly dependency?: 'page' | 'total-pages' | 'date' | 'time' | 'document';
  readonly range: TextRange;
  readonly origin: PointPt;
  readonly bounds: LayoutRect;
  readonly advancePt: number;
  /** Shaped cluster geometry for selection/hit testing. Always covers `range`. */
  readonly clusters: readonly TextClusterLayout[];
  /** Immutable contextual paint operations. Normally one whole-run operation. */
  readonly paintOps: readonly TextPaintOp[];
  readonly color: TextColorPolicy;
  readonly fontRoute: CanvasFontRoute;
  readonly fontSizePt: number;
  readonly fontWeight: number;
  readonly fontStyle: 'normal' | 'italic';
  readonly direction: TextDirection;
  readonly writingMode?: WritingMode;
  readonly characterSpacingPt?: number;
  readonly characterScale?: number;
  readonly fitText?: Readonly<{ regionIndex: number; perGapPt: number; trailingPadPt: number }>;
  readonly kerning?: boolean;
  readonly positionPt?: number;
  readonly verticalAlign?: 'super' | 'sub';
  readonly tateChuYoko?: boolean;
  readonly tateChuYokoCompress?: boolean;
  readonly ruby?: Readonly<{
    text: string;
    advancePt: number;
    authored: Readonly<{
      align?: string;
      baseFontSizePt?: number;
      raisePt?: number;
      language?: string;
    }>;
    readonly paintOps: readonly RetainedGlyphPaintOperation[];
  }>;
  readonly emphasisMark?: string;
  readonly emphasis?: Readonly<{
    authored: string;
    /** Selected authored mark glyphs, one per non-space source cluster. */
    glyphs?: readonly RetainedGlyphPaintOperation[];
    /** Authoritative outline paths remain representable when supplied by a font service. */
    paths?: readonly RetainedMarkPath[];
  }>;
  readonly highlight?: string;
  readonly highlightFragments?: readonly Readonly<{ rect: LayoutRect; color: string }>[];
  readonly background?: string;
  /** Justification width owned after this visual fragment. */
  readonly ownedTrailingSlackPt?: number;
  readonly runBorder?: RetainedRunBorderFacts;
  readonly runBorderFragments?: readonly BorderSegment[];
  readonly revision?: Readonly<{ kind: string; author?: string }>;
  readonly typography?: RetainedRunTypographyFacts;
  readonly unsupportedGeometry?: readonly (
    | 'underline'
    | 'strikethrough'
    | 'double-strikethrough'
    | 'emphasis'
  )[];
  readonly decorations: readonly TextDecorationLayout[];
  readonly hyperlink?: string;
  readonly bookmark?: string;
}

export interface TabPlacement {
  readonly kind: 'tab';
  readonly range: TextRange;
  readonly bounds?: LayoutRect;
  readonly advancePt: number;
  readonly leader: 'none' | 'dot' | 'hyphen' | 'underscore' | 'heavy' | 'middleDot';
  /** Fully repeated and positioned during acquisition; paint never measures. */
  readonly leaderGlyphs?: readonly RetainedGlyphPaintOperation[];
}

export interface AnchorHostPlacement {
  readonly kind: 'anchor-host';
  readonly range: TextRange;
  readonly bounds: LayoutRect;
  readonly baselinePt: number;
  readonly sourceMetrics?: Readonly<{ ascentPt: number; descentPt: number }>;
  readonly anchorOccurrenceId?: string;
}

export type InlineResourceKind = 'image' | 'chart' | 'math' | 'picture-bullet';
export type PaintResourceKind = InlineResourceKind;

export type PaintResourceDescriptorKind = InlineResourceKind;

export type ImagePaintResourceDescriptor = Readonly<{
  kind: 'image' | 'picture-bullet';
  resourceKey: string;
  partPath: string;
  mimeType: string;
  intrinsicSize: Readonly<{ widthPt: number; heightPt: number }>;
  svgImagePath?: string;
  srcRect?: Readonly<{ l: number; t: number; r: number; b: number }>;
  rotation?: number;
  flipH?: boolean;
  flipV?: boolean;
  alpha?: number;
  colorReplaceFrom?: string;
  duotone?: DeepReadonly<Duotone>;
}>;

export type ChartPaintResourceDescriptor = Readonly<{
  kind: 'chart';
  resourceKey: string;
  intrinsicSize: Readonly<{ widthPt: number; heightPt: number }>;
  model: DeepReadonly<ChartModel>;
}>;

export type MathPaintResourceDescriptor = Readonly<{
  kind: 'math';
  resourceKey: string;
}>;

export type PaintResourceDescriptor =
  | ImagePaintResourceDescriptor
  | ChartPaintResourceDescriptor
  | MathPaintResourceDescriptor;

export interface PaintResourceRegistry {
  readonly keys: readonly string[];
  readonly descriptors: readonly DeepReadonly<PaintResourceDescriptor>[];
  resolve<K extends PaintResourceDescriptorKind>(
    resourceKey: string,
    expectedKind: K,
  ): DeepReadonly<Extract<PaintResourceDescriptor, { kind: K }>>;
}

export interface ResourcePlacement {
  readonly kind: 'resource';
  readonly range: TextRange;
  readonly resourceKey: string;
  readonly resourceKind: InlineResourceKind;
  readonly bounds: LayoutRect;
  readonly advancePt: number;
}

export interface DrawingPlacement {
  readonly kind: 'drawing';
  readonly range: TextRange;
  readonly drawingId: LayoutNodeId;
  readonly bounds: LayoutRect;
  readonly advancePt: number;
}

export type ParagraphPlacement =
  | TextPlacement
  | TabPlacement
  | AnchorHostPlacement
  | ResourcePlacement
  | DrawingPlacement;

export interface LineLayout {
  readonly range: TextRange;
  readonly bounds: LayoutRect;
  readonly baselinePt: number;
  readonly advancePt: number;
  readonly placements: readonly ParagraphPlacement[];
}

export type InlineResourceLayout = Readonly<{
  kind: InlineResourceKind;
  resourceKey: string;
  intrinsicSize: Readonly<{ widthPt: number; heightPt: number }>;
}>;

export interface BorderSegment {
  readonly edge?: 'top' | 'right' | 'bottom' | 'left' | 'between';
  readonly from: PointPt;
  readonly to: PointPt;
  readonly color: string;
  readonly widthPt: number;
  /** Exact authored ST_Border token. Kept independently of paint normalization. */
  readonly authoredStyle: string;
  readonly style: 'solid' | 'double' | 'dotted' | 'dashed' | 'wavy';
  /** Final ST_Border cadence in point-space; empty for continuous/double rails. */
  readonly dashPatternPt?: readonly number[];
}

export type FillPaint = Readonly<{ color: string }>;

export interface WrapExclusion {
  readonly id: string;
  readonly wrap: 'square' | 'tight' | 'through' | 'topAndBottom';
  readonly bounds: LayoutRect;
  readonly polygon: readonly PointPt[];
  readonly anchorOccurrenceId?: string;
  readonly verticalOwnership?: 'page' | 'host';
}

export interface ParagraphFlowEvent {
  readonly kind: 'break';
  readonly breakKind: 'line' | 'page' | 'column';
  readonly offset: number;
}

/** Plain frame placement geometry consumed by layout and renderer adapters. */
export interface FrameGeometryState {
  readonly scale: number;
  readonly contentX: number;
  readonly contentW: number;
  readonly pageWidth: number;
  readonly pageH: number;
  readonly marginLeft: number;
  readonly marginRight: number;
  readonly marginTop: number;
  readonly marginBottom: number;
  readonly y: number;
}

export interface ParagraphMarkLayout {
  readonly hidden: boolean;
  readonly bounds: LayoutRect;
}

export interface LineNumberPaintOperation {
  readonly kind: 'text';
  readonly text: string;
  readonly origin: PointPt;
  readonly font: string;
  readonly color: string;
  readonly textAlign: 'right';
}

/** ECMA-376 §17.6.8 retained line counter and its optional paint operation. */
export interface LineNumberLayout {
  readonly lineIndex: number;
  readonly counterValue: number;
  readonly bounds: LayoutRect;
  readonly paintOps: readonly LineNumberPaintOperation[];
}

export interface ParagraphSpacingLayout {
  readonly beforePt: number;
  readonly afterPt: number;
}

export interface ParagraphLayout extends LayoutNodeBase {
  readonly kind: 'paragraph';
  readonly styleId?: string | null;
  readonly spacing: ParagraphSpacingLayout;
  readonly contextualSpacing: boolean;
  readonly lines: readonly LineLayout[];
  readonly borders: readonly BorderSegment[];
  readonly shading?: FillPaint;
  readonly resources: readonly InlineResourceLayout[];
  readonly drawings: readonly DrawingLayout[];
  readonly textBoxes: readonly TextBoxLayout[];
  readonly events: readonly ParagraphFlowEvent[];
  readonly exclusions: readonly WrapExclusion[];
  readonly anchorFrames?: readonly AnchorFrameResult[];
  readonly paragraphMark?: ParagraphMarkLayout;
  readonly lineNumbers?: readonly LineNumberLayout[];
  readonly continuation?: Readonly<{
    lineStart: number;
    lineEnd: number;
    continuesFromPrevious: boolean;
    continuesOnNext: boolean;
  }>;
}

export type ResolvedBorderSegment = BorderSegment;

export interface TableCellBlockLayout {
  readonly layout: ParagraphLayout | TableLayout;
  /** Final block origin from the cell border-box top. */
  readonly offsetPt: number;
  readonly advancePt: number;
}

export interface TableCellLayout extends LayoutNodeBase {
  readonly kind: 'table-cell';
  readonly contentBounds: LayoutRect;
  readonly verticalMerge: 'none' | 'restart' | 'continue';
  readonly vAlign: 'top' | 'center' | 'bottom';
  readonly background?: FillPaint;
  readonly blocks: readonly TableCellBlockLayout[];
}

export interface TableRowLayout extends LayoutNodeBase {
  readonly kind: 'table-row';
  readonly cells: readonly TableCellLayout[];
  /** Resolved row track height. Kept explicit while the A5 page-slice adapter
   * still consumes the historical row-height projection. */
  readonly heightPt: number;
  readonly contentHeightPt: number;
  readonly repeatedHeader?: boolean;
}

export interface TableLayout extends LayoutNodeBase {
  readonly kind: 'table';
  readonly columnWidthsPt: readonly number[];
  readonly rows: readonly TableRowLayout[];
  readonly borders: readonly ResolvedBorderSegment[];
}

/**
 * Page-local ownership of an out-of-flow nested table. Absolute page/margin
 * coordinates remain unresolved here; the wrapper retains the anchor fragment
 * bounds needed by the later placement stage and reuses the acquired child.
 */
export interface FloatingTablePlacementLayout {
  readonly kind: 'floating-table-placement';
  readonly occurrenceId: string;
  readonly ownership: 'source' | 'repeated-header';
  readonly physicalPageIndex: number;
  readonly displayPageNumber: number;
  readonly hostCellId: LayoutNodeId;
  readonly sourceBlockIndex: number;
  readonly anchorBlockIndex: number;
  readonly tableId: LayoutNodeId;
  readonly overlap: 'never' | 'overlap';
  readonly positioning: FloatingTablePositionInput;
  readonly acquiredTextOffsetPt?: Readonly<{ xPt: number; yPt: number }>;
  /** Final host-cell text column, distinct from the paragraph's vertical anchor. */
  readonly columnBounds?: LayoutRect;
  readonly anchorBounds: LayoutRect;
  readonly child: TableLayout;
}

/** Explicit point-space anchor frames supplied at the page/column adapter. */
export interface FloatingTableReferenceFramesPt {
  readonly page: LayoutRect;
  readonly margin: LayoutRect;
  readonly text: LayoutRect;
}

/** Point-space snapshot used while final table-fragment float placement is probed. */
export interface FloatRegistryEntryPt {
  readonly kind: 'table' | 'shape' | 'frame';
  readonly occurrenceId: string;
  readonly paragraphId: number;
  readonly bounds: LayoutRect;
  readonly exclusionBounds: LayoutRect;
}

export type FloatRegistryCoordinateSpace =
  | 'logical-page-points'
  | 'upright-physical-page-points';

export interface FloatRegistrySnapshotPt {
  readonly coordinateSpace: FloatRegistryCoordinateSpace;
  readonly flowDomainId: string;
  readonly entries: readonly FloatRegistryEntryPt[];
  readonly nextParagraphId: number;
}

export interface FloatRegistryDeltaPt {
  readonly coordinateSpace: FloatRegistryCoordinateSpace;
  readonly flowDomainId: string;
  readonly baseNextParagraphId: number;
  readonly nextParagraphId: number;
  readonly entries: readonly FloatRegistryEntryPt[];
}

/** Paint-ready result of resolving a page-local floating-table occurrence. */
export interface ResolvedFloatingTablePlacementLayout {
  readonly kind: 'resolved-floating-table-placement';
  readonly occurrenceId: string;
  readonly xPt: number;
  readonly yPt: number;
  readonly bounds: LayoutRect;
  readonly exclusionBounds: LayoutRect;
  readonly overlap: 'never' | 'overlap';
  readonly child: TableLayout;
  readonly source: FloatingTablePlacementLayout;
}

export interface TextBoxLayout extends LayoutNodeBase {
  readonly kind: 'textbox';
  readonly paragraphs: readonly ParagraphLayout[];
  readonly writingMode: WritingMode;
  readonly verticalMode?: 'vert' | 'vert270' | 'eaVert' | 'mongolianVert';
  readonly contentBounds?: LayoutRect;
  readonly insets: Readonly<{ topPt: number; rightPt: number; bottomPt: number; leftPt: number }>;
}

export interface NoteLayout extends LayoutNodeBase {
  readonly kind: 'note';
}

export type PaintNode = ParagraphLayout | TableLayout | DrawingLayout | TextBoxLayout | NoteLayout;

export type PageLayerId =
  | 'background'
  | 'behindText'
  | 'header'
  | 'body'
  | 'notes'
  | 'front'
  | 'footer';

export interface PagePaintEntry {
  readonly layer: PageLayerId;
  readonly nodeId: LayoutNodeId;
}

export interface PageLayers {
  readonly paintOrder: readonly PagePaintEntry[];
  readonly background: readonly PaintNode[];
  readonly behindText: readonly PaintNode[];
  readonly header: readonly PaintNode[];
  readonly body: readonly PaintNode[];
  readonly notes: readonly PaintNode[];
  readonly front: readonly PaintNode[];
  readonly footer: readonly PaintNode[];
}

/** One section-owned body-flow region on a physical page. A continuous section
 * may add another region below existing content without creating a new page. */
export interface PageSectionRegion {
  readonly id: string;
  readonly sectionOccurrenceId: string;
  /** Logical inline/block coordinates are retained independently of physical
   * x/y so vertical sections do not silently inherit horizontal Y-flow rules. */
  readonly coordinateSpace?: Readonly<{
    writingMode: WritingMode;
    logicalToPhysical: Matrix2DData;
  }>;
  readonly blockStartPt: number;
  readonly blockEndPt: number;
  readonly flowDomainIds: readonly string[];
  readonly section: DeepReadonly<SectionLayoutContext>;
}

export interface PageBookmarkStart {
  readonly name: string;
  readonly nodeId: LayoutNodeId;
  readonly sectionOccurrenceId: string;
}

export interface PageNumberMetadata {
  readonly displayNumber: number;
  readonly format: string;
  readonly sectionOccurrenceId: string;
}

export interface LayoutPage {
  readonly pageIndex: number;
  readonly geometry: PageGeometry;
  readonly flowDomains: readonly FlowDomain[];
  readonly section: DeepReadonly<SectionLayoutContext>;
  /** Transitional optionals keep pre-A6 producers compiling while the canonical
   * page factory becomes the sole producer; A6 removes that migration latitude. */
  readonly sectionOccurrenceId?: string;
  readonly parityBlank?: boolean;
  readonly bookmarkStarts?: readonly PageBookmarkStart[];
  readonly pageNumber?: PageNumberMetadata;
  /** Transitional until A6's canonical page producer is the only producer. */
  readonly sectionRegions?: readonly PageSectionRegion[];
  readonly layers: PageLayers;
  readonly readingOrder: readonly LayoutNodeId[];
}

export type LayoutDiagnosticCode =
  | 'FLOW_OVERLAP'
  | 'BOTTOM_MARGIN_INVASION'
  | 'FLOW_DOMAIN_INVASION'
  | 'INVALID_REFERENCE'
  | 'INVALID_GEOMETRY'
  | 'NON_CONVERGENCE'
  | 'UNSUPPORTED_FEATURE';

export interface LayoutDiagnostic {
  readonly code: LayoutDiagnosticCode;
  readonly severity: 'warning' | 'error';
  readonly source?: SourceRef;
  readonly message: string;
}

export interface DocumentLayout {
  readonly pages: readonly LayoutPage[];
  readonly diagnostics: readonly LayoutDiagnostic[];
}

export type CompatibilityEvidence =
  | Readonly<{ kind: 'microsoft-note'; reference: string }>
  | Readonly<{
      kind: 'office-observation';
      syntheticFixtureId: string;
      application: string;
      version: string;
      platform: string;
    }>;

export interface CompatibilityRule {
  readonly id: string;
  readonly evidence: CompatibilityEvidence;
  readonly description: string;
}

export interface LayoutServices {
  readonly text: TextLayoutService;
  readonly images: ImageMetadataService;
  readonly math: MathMetadataService;
}

/** Plain, parser-independent input for shaping a numbering marker. The renderer
 * boundary snapshots effective level rPr facts into this contract before the
 * retained layout service sees them. */
export interface NumberingMarkerShapeInput {
  readonly fontSizePt: number;
  readonly fonts: TextFontSlots;
  readonly themeFonts?: TextFontSlots;
  readonly themeFontPresence?: TextFontSlotPresence;
  readonly weight: number;
  readonly style: 'normal' | 'italic';
  readonly complexScript: boolean;
  readonly fontHint?: 'default' | 'eastAsia' | 'cs';
  readonly eastAsiaLanguage?: string;
  readonly kerning?: boolean;
}

export interface ParagraphLayoutInput {
  readonly kind: 'paragraph';
  readonly source: SourceRef;
}

export interface AcquiredParagraphLayoutInput {
  readonly kind: 'paragraph';
  readonly id: LayoutNodeId;
  readonly source: SourceRef;
  readonly flowDomainId: string;
  readonly ordinaryFlow: boolean;
  readonly styleId?: string | null;
  readonly flowBounds: LayoutRect;
  readonly inkBounds: LayoutRect;
  readonly clipBounds?: LayoutRect;
  readonly spacing: ParagraphSpacingLayout;
  readonly contextualSpacing?: boolean;
  readonly lines: readonly LineLayout[];
  readonly borders: readonly BorderSegment[];
  readonly shading?: FillPaint;
  readonly resources: readonly InlineResourceLayout[];
  readonly drawings: readonly DrawingLayout[];
  readonly textBoxes: readonly TextBoxLayout[];
  readonly events: readonly ParagraphFlowEvent[];
  readonly exclusions: readonly WrapExclusion[];
  readonly anchorFrames?: readonly AnchorFrameResult[];
  readonly paragraphMark?: ParagraphMarkLayout;
  readonly continuation?: Readonly<{
    lineStart: number;
    lineEnd: number;
    continuesFromPrevious: boolean;
    continuesOnNext: boolean;
  }>;
}

export interface TableBorderInput {
  readonly widthPt: number;
  readonly color: string;
  readonly authoredStyle: string;
}

export interface TableEdgeInputs {
  readonly top: TableBorderInput | null;
  readonly right: TableBorderInput | null;
  readonly bottom: TableBorderInput | null;
  readonly left: TableBorderInput | null;
  readonly insideH: TableBorderInput | null;
  readonly insideV: TableBorderInput | null;
}

export interface TableCellBlockInput {
  readonly layout: ParagraphLayout | TableLayout;
  /** Stable source index; continuation slices must not renumber field ownership. */
  readonly sourceBlockIndex: number;
  /** True when destination-page context can change the acquired child geometry. */
  readonly pageDependent?: boolean;
  /** The required empty paragraph after a nested table owns no row-height ink. */
  readonly structuralTrailing?: boolean;
}

export type TablePreferredWidthConstraint = Readonly<{
  kind: 'dxa' | 'pct';
  /** Points for dxa; a fraction whose OOXML owner defines the pct base. */
  value: number;
}>;

export interface TableColumnCellConstraint {
  readonly columnStart: number;
  readonly columnSpan: number;
  readonly preferredWidth: TablePreferredWidthConstraint | null;
  readonly minContentWidthPt: number;
  readonly maxContentWidthPt: number;
}

export interface TableSkippedColumnConstraint {
  readonly columnSpan: number;
  readonly preferredWidth: TablePreferredWidthConstraint | null;
}

export interface TableColumnRowConstraint {
  readonly before: TableSkippedColumnConstraint | null;
  readonly after: TableSkippedColumnConstraint | null;
  readonly cells: readonly TableColumnCellConstraint[];
}

/** Plain inputs for the §17.18.87 fixed/autofit column algorithm. */
export interface TableColumnLayoutInput {
  readonly layout: 'fixed' | 'autofit';
  readonly availableWidthPt: number;
  readonly gridWidthsPt: readonly number[];
  readonly tablePreferredWidthPt: number | null;
  readonly rows: readonly TableColumnRowConstraint[];
}

/** Parser/model-normalized row height. Authored-presence and OOXML lexical
 * units are resolved before this value crosses into layout. */
export interface TableRowHeightInput {
  readonly rule: 'auto' | 'atLeast' | 'exact';
  readonly valuePt: number | null;
}

export interface TableCellMarginsInput {
  readonly top: number;
  readonly bottom: number;
  readonly left: number;
  readonly right: number;
}

export interface TableRowExceptionInput {
  /** Whether tblPrEx/tblW was authored, including auto/nil/zero values which
   * deliberately shadow the parent table width without producing a length. */
  readonly preferredWidthAuthored: boolean;
  readonly preferredWidth: TablePreferredWidthConstraint | null;
  readonly layout: 'fixed' | 'autofit' | null;
  readonly justification: string | null;
  /** Whether tblPrEx/tblInd was authored. A nil/zero indent must suppress the
   * parent tblInd instead of being mistaken for omission. */
  readonly indentAuthored: boolean;
  readonly indentPt: number | null;
  readonly borders: import('../types.js').TableBorders | null;
}

export interface TableRowFormatInput {
  readonly height: TableRowHeightInput | null;
  /** Effective CT_OnOff values after table-style and direct row resolution. */
  readonly cantSplit: boolean;
  readonly repeatedHeader: boolean;
  readonly cellSpacingPt: number;
  readonly justification: string | null;
  readonly exception: TableRowExceptionInput | null;
  readonly cells: readonly { readonly marginsPt: TableCellMarginsInput }[];
}

/** Immutable parser/model projection consumed by table acquisition. */
export interface TableFormatInput {
  readonly effectiveStyleId: string | null;
  readonly ordinaryFlow: boolean;
  readonly positioning: FloatingTablePositionInput | null;
  readonly rows: readonly TableRowFormatInput[];
  readonly firstRowException: TableRowExceptionInput | null;
}

/** Parser-independent positioning facts retained from §17.4.57 `<w:tblpPr>`. */
export interface FloatingTablePositionInput {
  readonly leftFromTextPt: number;
  readonly rightFromTextPt: number;
  readonly topFromTextPt: number;
  readonly bottomFromTextPt: number;
  readonly horzAnchor: string;
  readonly horzSpecified: boolean;
  readonly vertAnchor: string;
  readonly xPt: number;
  readonly yPt: number;
  readonly xAlign?: string;
  readonly yAlign?: string;
}

export interface TableCellLayoutInput {
  readonly id: LayoutNodeId;
  readonly source: SourceRef;
  readonly columnStart: number;
  readonly columnSpan: number;
  readonly verticalMerge: 'none' | 'restart' | 'continue';
  readonly margins: Readonly<{
    topPt: number;
    rightPt: number;
    bottomPt: number;
    leftPt: number;
  }>;
  readonly vAlign: 'top' | 'center' | 'bottom';
  readonly background?: FillPaint;
  readonly borders: TableEdgeInputs;
  readonly blocks: readonly TableCellBlockInput[];
}

export interface TableRowLayoutInput {
  readonly id: LayoutNodeId;
  readonly source: SourceRef;
  readonly logicalRowIndex: number;
  readonly cantSplit: boolean;
  readonly heightPt: number | null;
  readonly heightRule: 'auto' | 'atLeast' | 'exact';
  /** Effective §17.4.43/.44/.45 dxa spacing for this row. */
  readonly cellSpacingPt: number;
  /** §17.4.39 table-level border exception for this row, if authored. */
  readonly exceptionBorders: TableEdgeInputs | null;
  /** Effective §17.4.27/.26 row alignment. */
  readonly alignment: 'left' | 'center' | 'right';
  /** Effective table indent after Word's first-row tblPrEx rule. */
  readonly indentPt: number;
  readonly cells: readonly TableCellLayoutInput[];
  readonly repeatedHeader: boolean;
}

export interface TableLayoutInput {
  readonly kind: 'table';
  readonly id: LayoutNodeId;
  readonly source: SourceRef;
  readonly flowDomainId: string;
  readonly ordinaryFlow: boolean;
  readonly alignment: 'left' | 'center' | 'right';
  readonly indentPt: number;
  readonly bidiVisual: boolean;
  readonly columnWidthsPt: readonly number[];
  readonly borders: TableEdgeInputs;
  readonly rows: readonly TableRowLayoutInput[];
}

export type FlowBlockInput = ParagraphLayoutInput | TableLayoutInput;

export interface FlowContainer extends FlowDomain {}

export interface FlowCursor extends PointPt {}

export interface FlowBlockPlacement {
  readonly container: FlowContainer;
  readonly cursor: FlowCursor;
  readonly availableBounds: LayoutRect;
}

export interface BlockLayoutResult<T extends ParagraphLayout | TableLayout> {
  readonly layout: T;
  readonly nextCursor: FlowCursor;
}

export interface FlowLayoutInput {
  readonly blocks: readonly FlowBlockInput[];
  readonly container: FlowContainer;
  readonly cursor: FlowCursor;
  readonly source: SourceRef;
}

export interface FlowLayout extends FlowOwnership {
  readonly source: SourceRef;
  readonly container: FlowContainer;
  readonly blocks: readonly (ParagraphLayout | TableLayout)[];
  readonly nextCursor: FlowCursor;
}

export interface BlockLayoutAlgorithms {
  layoutParagraph(
    input: ParagraphLayoutInput,
    placement: FlowBlockPlacement,
    services: LayoutServices,
  ): BlockLayoutResult<ParagraphLayout>;
  layoutTable(
    input: TableLayoutInput,
    placement: FlowBlockPlacement,
    services: LayoutServices,
  ): BlockLayoutResult<TableLayout>;
}
