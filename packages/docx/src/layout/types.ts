import type { SectionLayoutContext } from '../layout-context.js';
import type { DocxStorySource } from '../types.js';
import type {
  GlyphInkBounds,
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
  HyperlinkTarget,
} from '@silurus/ooxml-core';

export type { TextLayoutService } from './text.js';
export type { ImageMetadataService, MathMetadataService } from './resources.js';
export type { HyperlinkTarget } from '@silurus/ooxml-core';
export type { MathRenderer } from '@silurus/ooxml-core';
export type { MathLayoutResource, MathOccurrence } from './resources.js';

export type LayoutNodeId = string;

export type SourceRef = Readonly<DocxStorySource>;

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

export type LayoutCoordinateSpace =
  | 'column-local-logical'
  | 'logical-page-points'
  | 'upright-physical-page-points';

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
  /** Section-region coordinate owner for page-level stories. */
  readonly sectionRegionId?: string;
  /** logical-page-points; same space as retained node flowBounds/inkBounds. */
  readonly logicalBounds: LayoutRect;
  /** upright-physical-page-points. */
  readonly physicalBounds: LayoutRect;
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

export interface SectionRegionCoordinateSpace {
  readonly writingMode: WritingMode;
  readonly logicalToPhysical: Matrix2DData;
  readonly physicalToLogical: Matrix2DData;
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
  readonly sectionFlowOwnership?: 'host-flow' | 'page';
}

interface LayoutNodeBase extends FlowOwnership {
  readonly id: LayoutNodeId;
  readonly source: SourceRef;
  /** Recoverable acquisition facts owned by this retained node. */
  readonly diagnostics?: readonly LayoutDiagnostic[];
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
      /** A retained image resource clipped to the authored DrawingML shape. */
      kind: 'drawingml-image-fill';
      plan: DeepReadonly<DrawingMLShapePaintPlan>;
      resourceKey: string;
      fillRect?: Readonly<{ l: number; t: number; r: number; b: number }>;
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
      /** Keep a non-text graphic upright after the enclosing section-logical
       * frame is rotated into a vertical physical page. */
      orientation?: 'upright-physical';
    }>;

export interface DrawingLayout extends LayoutNodeBase {
  readonly kind: 'drawing';
  /** The commands and owned text boxes are authored in an upright, drawing-
   * local physical frame. `transform` maps that frame into section-logical
   * points, where the page's quarter turn cancels it during paint. */
  readonly orientation?: 'upright-physical';
  readonly transform?: Matrix2DData;
  readonly clip?: ClipPathData;
  readonly commands: readonly DrawingPaintCommand[];
  readonly anchorLayer?: Readonly<{
    occurrenceId: string;
    /** Stable identity from acquisition before occurrence-local re-keying. */
    acquisitionOccurrenceId?: string;
    behindDoc: boolean;
    relativeHeight: number;
    sourceOrder: number;
    horizontalOwnership: 'page' | 'host';
    verticalOwnership: 'page' | 'host';
    /** @internal This occurrence contributes to its owning table-cell extent. */
    cellContainment?: true;
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
  /** Acquisition-retained glyph-local block-axis scale. Used by
   * `eastAsianLayout@vertCompress` so paint never remeasures font metrics. */
  readonly scaleY?: number;
  readonly direction: TextDirection;
  readonly kerning: 'auto' | 'normal' | 'none';
  readonly writingMode: WritingMode;
  readonly glyphOrientation?: 'sideways' | 'upright' | 'rotate';
  /** Acquisition proved that the selected face exposes this code point through
   * OpenType `vert`; paint applies the feature without probing or measuring. */
  readonly verticalFeature?: boolean;
  /** Glyph-local offset retained from vertical-form ink/corner geometry. */
  readonly glyphOffsetPt?: PointPt;
  /** `kashida` permits acquisition-inserted U+0640 glyphs over one source range. */
  readonly sourceMapping?: 'identity' | 'kashida';
  /** Selected-face ink relative to this operation's alphabetic baseline.
   * Pagination may consume it, but paint must never reconstruct it. */
  readonly inkBounds?: GlyphInkBounds;
  /** Tight ink after applying this operation's retained glyph transform,
   * expressed on the paragraph's logical block axis relative to `offset.yPt`.
   * Vertical final-line admission consumes this instead of reinterpreting
   * alphabetic metrics for counter-rotated glyphs. */
  readonly blockAxisInkBounds?: Readonly<{
    startPt: number;
    endPt: number;
  }>;
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
  readonly noteReference?: Readonly<{ kind: 'footnote' | 'endnote'; id: string }>;
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
  /** Resolved ECMA-376 §17.16.22/§17.16.23 external or bookmark target. */
  readonly hyperlink?: HyperlinkTarget;
}

export interface TabPlacement {
  readonly kind: 'tab';
  readonly range: TextRange;
  readonly bounds?: LayoutRect;
  readonly advancePt: number;
  readonly leader: 'none' | 'dot' | 'hyphen' | 'underscore' | 'heavy' | 'middleDot';
  /** Fully repeated and positioned during acquisition; paint never measures. */
  readonly leaderGlyphs?: readonly RetainedGlyphPaintOperation[];
  /** Run formatting applies to the tab character itself (§17.3.1.37). */
  readonly decorations?: readonly TextDecorationLayout[];
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
  /** Authored image traversal order, retained across registry key sorting so
   * path-keyed decode deduplication preserves the first document occurrence. */
  documentOrder?: number;
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
  /** Parsed run occurrence retained for element-context source locators. */
  readonly sourceRunIndex?: number;
  readonly resourceKey: string;
  readonly resourceKind: InlineResourceKind;
  /** Keep a non-text graphic upright after the enclosing section-logical frame
   * is rotated into a vertical physical page. Absent means flow-relative. */
  readonly orientation?: 'upright-physical';
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
  /** §17.18.84 bar-tab vertical rules acquired independently of tab advances. */
  readonly barTabRules?: readonly BorderSegment[];
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
  readonly style: 'solid' | 'double' | 'compound' | 'dotted' | 'dashed' | 'wavy';
  /** Final ST_Border cadence in point-space; empty for continuous/double rails. */
  readonly dashPatternPt?: readonly number[];
}

export type FillPaint = Readonly<{ color: string }>;

export interface WrapExclusion {
  readonly id: string;
  readonly wrap: 'square' | 'tight' | 'through' | 'topAndBottom';
  readonly wrapSide?: 'bothSides' | 'left' | 'right' | 'largest';
  readonly bounds: LayoutRect;
  readonly polygon: readonly PointPt[];
  readonly anchorOccurrenceId?: string;
  readonly verticalOwnership?: 'page' | 'host';
}

/** @internal Occurrence-keyed DrawingML object bounds for §20.4.2.3.
 * Kept separate from text-wrap exclusions because wrapNone objects still
 * participate in object-to-object collision avoidance. */
export interface DrawingMLCollisionEntryPt {
  readonly occurrenceId: string;
  readonly bounds: LayoutRect;
  readonly horizontalOwnership: 'page' | 'host';
  readonly verticalOwnership: 'page' | 'host';
  /** Authored wp:anchor z-order, when the collision came from a retained
   * DrawingML anchor rather than a legacy compatibility float. */
  readonly relativeHeight?: number;
}

export interface ParagraphFlowEvent {
  readonly kind: 'break';
  readonly breakKind: 'line' | 'page' | 'column';
  readonly offset: number;
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
  /** Source `w14:paraId`; identity only, never interpreted by layout or paint. */
  readonly paragraphId?: string;
  readonly styleId?: string | null;
  /**
   * ECMA-376 §17.13.6.2 bookmark starts owned by this retained paragraph
   * fragment. The parser currently preserves paragraph ownership rather than
   * an inline character offset, so only the first page slice carries them.
   */
  readonly bookmarkStarts?: readonly string[];
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
  /** @internal Union of layoutInCell drawing frames owned by this fragment. */
  readonly cellContainmentBounds?: LayoutRect;
  /** @internal */
  readonly anchorCollisions?: readonly DrawingMLCollisionEntryPt[];
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

/** A point-space rectangular frame whose outer compound border segments are
 * painted as one joined unit. Layout owns rectangle recognition and segment
 * membership; paint owns only device-pixel rail projection. */
export interface CompoundBorderFrameLayout {
  readonly bounds: LayoutRect;
  readonly border: Pick<BorderSegment, 'authoredStyle' | 'color' | 'widthPt' | 'style'>;
  readonly segmentIndexes: readonly number[];
}

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
  readonly compoundBorderFrames?: readonly CompoundBorderFrameLayout[];
  readonly floatingTables?: readonly FloatingTablePlacementLayout[];
  readonly resolvedFloatingTables?: readonly ResolvedFloatingTablePlacementLayout[];
  /** Point space already owned by `resolvedFloatingTables`; occurrence projection
   * must not translate those final frames a second time. */
  readonly resolvedFloatingTableCoordinateSpace?: FloatRegistryCoordinateSpace;
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

interface FloatRegistryEntryCorePt {
  readonly occurrenceId: string;
  readonly paragraphId: number;
  readonly bounds: LayoutRect;
  readonly exclusionBounds: LayoutRect;
  /** DrawingML collision-axis ownership; present for parser-owned shapes. */
  readonly horizontalOwnership?: 'page' | 'host';
  readonly verticalOwnership?: 'page' | 'host';
  /** Stable retained-graph identity for exclusions owned by an accepted occurrence. */
  readonly exclusionId?: string;
  /** Parser-authored DrawingML wrap facts. Ordinary tables/frames omit these
   * and retain their established square exclusion semantics. */
  readonly wrap?: WrapExclusion['wrap'];
  readonly wrapSide?: string | null;
  readonly wrapDistances?: Readonly<{
    topPt: number;
    rightPt: number;
    bottomPt: number;
    leftPt: number;
  }>;
  readonly wrapPolygon?: readonly PointPt[];
}

/** Point-space snapshot used while final table-fragment float placement is
 * probed. A table entry must retain its §17.4.56 fact because a blocker-side
 * `never` constrains tables placed later in source order. */
export type FloatRegistryEntryPt =
  | Readonly<FloatRegistryEntryCorePt & {
      readonly kind: 'table';
      readonly overlap: 'never' | 'overlap';
    }>
  | Readonly<FloatRegistryEntryCorePt & {
      readonly kind: 'shape';
    }>
  | Readonly<FloatRegistryEntryCorePt & {
      readonly kind: 'frame';
    }>;

export type FloatRegistryCoordinateSpace = Exclude<
  LayoutCoordinateSpace,
  'column-local-logical'
>;

export interface FloatRegistrySnapshotPt {
  readonly coordinateSpace: FloatRegistryCoordinateSpace;
  readonly flowDomainId: string;
  readonly entries: readonly FloatRegistryEntryPt[];
  readonly nextParagraphId: number;
}

/** @internal Transaction-local delta. Cloned deltas intentionally lose lineage
 * and therefore cannot be committed. */
export interface FloatRegistryDeltaPt {
  readonly coordinateSpace: FloatRegistryCoordinateSpace;
  readonly flowDomainId: string;
  /** Exact immutable snapshot lineage; registry deltas never cross a clone boundary. */
  readonly baseEntries: FloatRegistrySnapshotPt['entries'];
  readonly baseNextParagraphId: number;
  readonly nextParagraphId: number;
  readonly entries: readonly FloatRegistryEntryPt[];
}

/** Accepted DrawingML object bounds in source-acceptance order.
 * Unlike the text float registry, this registry is never populated by page
 * prescan because §20.4.2.3 collision authority begins at source acceptance. */
export interface DrawingMLCollisionRegistrySnapshotPt {
  readonly coordinateSpace: FloatRegistryCoordinateSpace;
  readonly flowDomainId: string;
  readonly entries: readonly DrawingMLCollisionEntryPt[];
}

/** @internal Transaction-local delta. Cloned deltas intentionally lose lineage
 * and therefore cannot be committed. */
export interface DrawingMLCollisionRegistryDeltaPt {
  readonly coordinateSpace: FloatRegistryCoordinateSpace;
  readonly flowDomainId: string;
  /** Exact immutable snapshot lineage; registry deltas never cross a clone boundary. */
  readonly baseEntries: DrawingMLCollisionRegistrySnapshotPt['entries'];
  readonly baseEntryCount: number;
  readonly entries: readonly DrawingMLCollisionEntryPt[];
}

/** Page-local registries committed together after one body candidate is accepted.
 * The shared occurrence ID is a join key, not duplicate geometry authority. */
export interface BodyFlowRegistrySnapshotPt {
  readonly floats: FloatRegistrySnapshotPt;
  readonly drawingCollisions: DrawingMLCollisionRegistrySnapshotPt;
}

export interface BodyFlowRegistryDeltaPt {
  readonly floats?: FloatRegistryDeltaPt;
  readonly drawingCollisions?: DrawingMLCollisionRegistryDeltaPt;
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
  readonly story: StoryLayout;
  /** DrawingML fontRef color inherited by runs without an explicit w:color. */
  readonly defaultTextColor?: string;
  /** Acquisition-local story coordinates to the containing page/shape frame. */
  readonly transform: Matrix2DData;
  readonly writingMode: WritingMode;
  readonly verticalMode?: 'vert' | 'vert270' | 'eaVert' | 'mongolianVert';
  readonly contentBounds?: LayoutRect;
  readonly insets: Readonly<{ topPt: number; rightPt: number; bottomPt: number; leftPt: number }>;
}

/** Parser-owned placeholder for schema-permitted CT_TxbxContent members that
 * are not yet layout-capable. The structural path keeps their authored order
 * observable without widening the public ShapeRun contract. */
export interface ParsedUnsupportedTextBoxBlock {
  readonly type: 'unsupportedTextBoxBlock';
  readonly qName: string;
  readonly sourcePath: readonly number[];
}

export type StoryBlockInput = FlowBlockInput | ParsedUnsupportedTextBoxBlock;

export interface StoryLayoutInput {
  readonly source: SourceRef;
  readonly container: FlowContainer;
  readonly blocks: readonly StoryBlockInput[];
}

export interface StoryLayout {
  readonly story: SourceRef['story'];
  readonly flowBounds: LayoutRect;
  readonly inkBounds: LayoutRect;
  readonly clipBounds?: LayoutRect;
  readonly blocks: readonly PaintNode[];
  readonly advancePt: number;
  readonly diagnostics: readonly LayoutDiagnostic[];
}

export interface NoteLayout extends LayoutNodeBase {
  readonly kind: 'note';
  readonly separator: readonly BorderSegment[];
  readonly story: StoryLayout;
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

export interface PageLayerRoot {
  readonly layer: PageLayerId;
  readonly node: PaintNode;
  readonly coordinateSpace: 'section-logical' | 'upright-physical';
}

export type PagePaintFrame =
  | Readonly<{
      kind: 'transform';
      transform: Matrix2DData;
    }>
  | Readonly<{
      kind: 'clip';
      clip: LayoutRect;
    }>;

interface PagePaintEntryBase {
  /** Final semantic stacking layer. */
  readonly layer: PageLayerId;
  /** Story/root layer from which this paint operation was materialized. */
  readonly sourceLayer: PageLayerId;
  /** Top-level retained root represented by this operation. */
  readonly rootNodeId: LayoutNodeId;
  readonly coordinateSpace: 'section-logical' | 'upright-physical';
  /** Flow domain whose section-region transform owns this operation. */
  readonly flowDomainId: string;
}

export interface PagePaintNodeEntry extends PagePaintEntryBase {
  readonly kind: 'node';
  readonly node: PaintNode;
  /** Anchored drawings owned by this root are represented by drawing entries. */
  readonly omitAnchoredDrawings: boolean;
}

export interface PagePaintDrawingEntry extends PagePaintEntryBase {
  readonly kind: 'drawing';
  readonly layer: 'behindText' | 'front';
  readonly node: DrawingLayout;
  /** Diagnostic/ownership identity; paint does not traverse the owner graph. */
  readonly ownerNodeId?: LayoutNodeId;
  readonly textBoxes: readonly TextBoxLayout[];
  /** Plain-data Canvas frames from the root to the anchor encounter. */
  readonly frames: readonly PagePaintFrame[];
  /** Cumulative host placement used to undo page-owned anchor axes exactly once. */
  readonly layoutTranslationPt: PointPt;
}

export type PagePaintEntry = PagePaintNodeEntry | PagePaintDrawingEntry;

export interface PagePaintCapabilities {
  /** At least one retained text operation proved an OpenType `vert` form.
   * Canvas paint must therefore use an element-backed context when available,
   * because OffscreenCanvas cannot apply the proven feature. */
  readonly requiresElementBackedVerticalGlyphPaint: boolean;
}

export interface PageLayers {
  /** Top-level retained story roots in composition order. This is graph and
   * reading-order authority, not paint-order authority. */
  readonly roots: readonly PageLayerRoot[];
  /** Completed layout-owned order. Paint consumes it without graph discovery,
   * sorting, callback collection, or callback re-entry. */
  readonly paintOrder: readonly PagePaintEntry[];
  /** Target requirements derived once from the immutable retained paint plan. */
  readonly capabilities: PagePaintCapabilities;
  readonly background: readonly PaintNode[];
  readonly behindText: readonly PaintNode[];
  readonly header: readonly PaintNode[];
  readonly body: readonly PaintNode[];
  readonly notes: readonly PaintNode[];
  readonly front: readonly PaintNode[];
  readonly footer: readonly PaintNode[];
}

/** One section-owned body-flow region on a physical page. A continuous section
 * may add another region below existing content without creating a new page;
 * an occurrence owning only out-of-flow content has an empty region. */
export interface PageSectionRegion {
  readonly id: string;
  readonly sectionOccurrenceId: string;
  /** Logical inline/block coordinates are retained independently of physical
   * x/y so vertical sections do not silently inherit horizontal Y-flow rules. */
  readonly coordinateSpace: SectionRegionCoordinateSpace;
  readonly blockStartPt: number;
  readonly blockEndPt: number;
  readonly columnFlowDirection: 'ltr' | 'rtl';
  /** Authored §17.6.4 column indexes owned by this physical region. */
  readonly columnIndexes: readonly number[];
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

/** Paint-ready §17.6.10 page-border geometry in its owning section's logical
 * point space. Visibility, reference-box geometry, and line treatment are
 * resolved during document layout; Canvas paint applies only the retained page
 * transform and front/back ordering. */
export interface PageBorderLayout {
  readonly zOrder: 'front' | 'back';
  readonly logicalToPhysical: Matrix2DData;
  readonly segments: readonly BorderSegment[];
}

/** Physical page-space ink fixed during layout for ECMA-376 §17.6.4
 * `w:cols/@w:sep`. Paint may snap the retained endpoints to device pixels,
 * but it must not recover section-band geometry. */
export interface ColumnSeparatorLayout {
  readonly start: PointPt;
  readonly end: PointPt;
}

export interface LayoutPage {
  readonly pageIndex: number;
  readonly geometry: PageGeometry;
  readonly flowDomains: readonly FlowDomain[];
  readonly section: DeepReadonly<SectionLayoutContext>;
  readonly sectionOccurrenceId: string;
  readonly parityBlank: boolean;
  readonly bookmarkStarts: readonly PageBookmarkStart[];
  readonly pageNumber: PageNumberMetadata;
  readonly sectionRegions: readonly PageSectionRegion[];
  readonly columnSeparators: readonly ColumnSeparatorLayout[];
  readonly pageBorder: PageBorderLayout | null;
  readonly layers: PageLayers;
  readonly readingOrder: readonly LayoutNodeId[];
}

export type LayoutDiagnosticCode =
  | 'FLOW_OVERLAP'
  | 'BOTTOM_MARGIN_INVASION'
  | 'FLOW_DOMAIN_INVASION'
  | 'INVALID_REFERENCE'
  | 'INVALID_GEOMETRY'
  | 'INVALID_VALUE'
  | 'MISSING_RESOURCE'
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
  | Readonly<{ kind: 'regression-test'; reference: string }>
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
  /** Geometry-affecting vertical glyph acquisition capability. Kept separate
   * from horizontal text shaping so each service fingerprint stays truthful. */
  readonly verticalGlyphFingerprint?: string;
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
  /** Source `w14:paraId`; retained for text-run projection. */
  readonly paragraphId?: string;
  readonly flowDomainId: string;
  readonly ordinaryFlow: boolean;
  readonly styleId?: string | null;
  readonly bookmarkStarts?: readonly string[];
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
  /** @internal Union of layoutInCell drawing frames owned by this fragment. */
  readonly cellContainmentBounds?: LayoutRect;
  /** @internal */
  readonly anchorCollisions?: readonly DrawingMLCollisionEntryPt[];
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
  /** Physical occurrence ceiling. `null` means the containing frame imposes
   * no width ceiling (for example, an authored fixed table nested in a cell). */
  readonly availableWidthPt: number | null;
  readonly gridWidthsPt: readonly number[];
  readonly gridWidthKeys?: readonly (string | null)[];
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
  /** Parser-owned ECMA-376 §17.4.37 logical table membership. */
  readonly logicalSequenceId?: string | null;
  readonly logicalRowOffset?: number;
  readonly logicalTotalRows?: number;
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
  /** Effective table indent after Word's first-row tblPrEx rule. In an adjacent
   * §17.4.37 group this is re-oriented into the group frame by the union
   * builder, so no separate physical-indent field is retained here. */
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
  /** Exact grid topology; numeric widths remain the Canvas paint coordinate. */
  readonly columnWidthKeys?: readonly (string | null)[];
  readonly borders: TableEdgeInputs;
  readonly rows: readonly TableRowLayoutInput[];
}

export type FlowBlockInput = ParagraphLayoutInput | TableLayoutInput;

export interface FlowContainer {
  readonly id: string;
  readonly kind: FlowDomainKind;
  /** Acquisition-local logical bounds; never retained physical page bounds. */
  readonly bounds: LayoutRect;
  /** Text boxes lay out complete overflow before applying bodyPr clipping or
   * spAutoFit to the outer frame. Other stories remain capacity-bounded. */
  readonly capacity?: 'bounded' | 'unbounded';
}

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
