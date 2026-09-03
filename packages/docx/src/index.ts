export {
  DocxDocument,
  type CollectPageRunsOptions,
  type DocxPageCommentThreadsOptions,
  type LoadOptions,
  type RenderPageToBitmapOptions,
} from './document';
export { DocxViewer, type DocxViewerOptions } from './viewer';
export {
  FontProvider,
  FontProviderSession,
  GoogleFontsProvider,
  type FontAsset,
  type FontAssetSource,
  type FontFailure,
  type FontResolveOptions,
  type ResolvedFontFace,
  type ResolvedFonts,
} from '@silurus/ooxml-core';
export { DocxScrollViewer, type DocxScrollViewerOptions } from './scroll-viewer';
export type { DocxCommentsOptions } from './comment-margin';
export type {
  ViewerCommentConnectorOptions,
  ViewerCommentConnectorRoute,
  ViewerCommentConnectorStroke,
  ViewerCommentMessageContext,
  ViewerCommentThreadContext,
} from '@silurus/ooxml-core';
export { buildDocxTextLayer } from './text-layer';
export {
  readDocxTextSelectionContext,
  type DocxSelectionContext,
  type DocxTextSelectionContext,
  type DocxElementContext,
  type DocxCommentSelectionContext,
  type DocxPagePoint,
  type DocxSelectionContextOptions,
  type DocxSelectionSourceLocator,
  type DocxSelectionRunLocator,
} from './selection-context';
export type { DocxElementContextOptions } from './element-context';
export type {
  ChartAreaGroupDecorations,
  ChartBarGroupDecorations,
  ChartDataLabelOverride,
  ChartDataTable,
  ChartDataPointOverride,
  ChartDecorationLineStyle,
  ChartErrBars,
  ChartLabelBox,
  ChartLegendEntryOverride,
  ChartLineGroupDecorations,
  ChartManualLayout,
  ChartModel,
  ChartPlotGroup,
  ChartPlotGroupAxisSlot,
  ChartPlotGroupKind,
  ChartRect,
  ChartSeries,
  ChartSeriesDataLabels,
  ChartStockBarPaint,
  ChartStockUpDownBarStyle,
  ChartSurfaceBandFormat,
  ChartTextBox,
  ChartTextParagraph,
  ChartTextRun,
  ChartTrendline,
  ChartType,
  ChartOfPie,
  ChartThreeD,
  ChartThreeDPictureOptions,
  ChartThreeDSurface,
  ChartThreeDSeriesAxis,
  ChartThreeDRenderer,
  ChartRegionMapRenderer,
  ChartExRenderer,
  TiffRenderer,
  TiffRenderOptions,
  ChartExElementStyle,
  ChartLineDashSegment,
  ChartexHistogramBinning,
  ChartexBoxSeries,
  ChartexBoxWhisker,
  ChartexSunburst,
  ChartexSunburstRow,
  ChartexTreemap,
  ChartexRegionMap,
  ChartexRegionMapRow,
  ChartexGeography,
  ChartexRegionMapColors,
  ChartexValueColorStop,
  Duotone,
  DrawingMLCustomDashSegment,
  FillRect,
  GradientFill,
  ImageFill,
  LegendManualLayout,
  MathAccent,
  MathArray,
  MathBar,
  MathBorderBox,
  MathBox,
  MathDelimiter,
  MathFraction,
  MathFunc,
  MathGroup,
  MathGroupChr,
  MathLimit,
  MathRenderer,
  MathNary,
  MathNode,
  MathPhant,
  MathRadical,
  MathRun,
  MathScript,
  MathSPre,
  MathStyle,
  MathSvg,
  MatchRunSlice,
  PatternFill,
  NoFill,
  SolidFill,
  ChartDisplayUnits,
  ChartDisplayUnitsLabel,
  SecondaryValueAxis,
  SrcRect,
  TextSelectionContextOptions,
  TileInfo,
  ViewerContextMenuEvent,
  ZoomableViewer,
} from '@silurus/ooxml-core';
// IX2 find-in-document: the highlight overlay builder + the docx match-location
// shape. `FindMatch` / `FindMatchesOptions` come from core (shared across formats).
export {
  buildDocxHighlightLayer,
  type DocxHighlightMatch,
  type DocxHighlightColors,
} from './find-highlight-layer';
// ECMA-376 §17.13.4 comment data projections for application-owned review UIs.
export {
  resolveCommentAnchorRuns,
  resolveDocxCommentThreads,
  type CommentAnchorPoint,
  type CommentAnchorGeometryFallback,
  type CommentAnchorRange,
  type DocxCommentAnchorKind,
  type DocxCommentHighlightRect,
  type ResolvedDocxCommentAnchor,
  type ResolvedDocxCommentThread,
  type ResolveDocxCommentThreadsOptions,
} from './comments';
export {
  resolveRevisionAnchorRuns,
  type RevisionAnchorGeometryFallback,
  type RevisionAnchorRange,
} from './revisions';
export type { DocxStorySource } from './types';
export type { DocxMatchLocation } from './find';
export type { FindHighlightColors, FindMatch, FindMatchesOptions } from '@silurus/ooxml-core';
export { autoResize, type AutoResizeOptions } from '@silurus/ooxml-core';
// IX1 — the shared hyperlink target shape surfaced by `DocxViewerOptions.
// onHyperlinkClick`, `DocxTextRunInfo.hyperlink`, and the 5th arg of
// `buildDocxTextLayer`, plus the default "open in a new tab, sanitised" helper.
export { type HyperlinkTarget, openExternalHyperlink } from '@silurus/ooxml-core';
// Typed load-time error surfaced by DocxDocument.load (e.g. a password-protected
// or legacy-binary .doc file). Re-exported so `@silurus/ooxml/docx` consumers can
// narrow on `err.code`.
export {
  OoxmlError,
  OoxmlDecodedImageLimitError,
  OoxmlResourceLimitError,
  TiffDecodeError,
  isOoxmlDecodedImageLimitError,
  isTiffDecodeError,
  type OoxmlDecodedImageLimitMetric,
  type DecodedImageBudgetStrategy,
  type ImageResourceOptions,
  type OoxmlErrorCode,
  type OoxmlErrorStage,
  type OoxmlFormat,
  type OoxmlResourceLimit,
  type OoxmlResourceLimitErrorDetails,
  type OoxmlResourceLimits,
  type OoxmlResourceMetric,
  type OoxmlResourceMetrics,
  type OoxmlResourceMetricsCheckpoint,
  type OoxmlResourceName,
  type OoxmlResourcePolicySnapshot,
  type OoxmlResourceUsageSnapshot,
  type OoxmlResourceViolation,
} from '@silurus/ooxml-core';
export { noteText } from './types';
export type {
  DocxDocumentModel,
  DocSettings,
  // Embedded-font reference (reachable via DocxDocumentModel.embeddedFonts,
  // ECMA-376 §17.8.3.3-.6). The viewer de-obfuscates + registers these.
  EmbeddedFontRef,
  SectionProps,
  // Per-section page geometry (reachable via the BodyElement sectionBreak arm's
  // `geom` and canonical section-region geometry, ECMA-376 §17.6.13/§17.6.11).
  SectionGeom,
  // Per-section page-numbering settings (reachable via SectionProps.pageNumType
  // and the BodyElement sectionBreak arm's `pageNumType`, ECMA-376 §17.6.12).
  PageNumType,
  // Per-section page decorations (reachable via SectionProps): page borders
  // (§17.6.10) and line numbering (§17.6.8).
  PageBorders,
  PageBorderEdge,
  LineNumbering,
  // Multi-column section sub-types (reachable via SectionProps.columns).
  ColumnsSpec,
  ColSpec,
  HeadersFooters,
  HeaderFooter,
  NumberingInfo,
  BodyElement,
  DocParagraph,
  DocRun,
  // Absolute-position tab run (reachable via the DocRun union's `ptab` arm,
  // ECMA-376 §17.3.3.23).
  PTabRun,
  DocxTextRun,
  FieldRun,
  ImageRun,
  // DrawingML chart run (reachable via the DocRun union's `chart` arm,
  // ECMA-376 §21.2).
  ChartRun,
  AnchorHostMetrics,
  ShapeRun,
  // VML `<v:textpath>` watermark text (reachable via ShapeRun.textPath).
  TextPath,
  ShapeText,
  // Per-run shape-text formatting (reachable via ShapeText.runs).
  ShapeTextRun,
  RubyAnnotation,
  RenderPageOptions,
  RunRevision,
  DocRevision,
  DocComment,
  // Comment-anchor boundary (reachable via DocParagraph.commentMarks).
  DocxCommentMark,
  DocNote,
  NoteRef,
  // Paragraph / line-spacing sub-types.
  LineSpacing,
  // Text-frame / drop-cap properties (reachable via DocParagraph.framePr).
  FramePr,
  TabStop,
  ParagraphBorders,
  ParaBorderEdge,
  DocxRunBorder,
  // Table model (reachable via BodyElement table variant).
  DocTable,
  TblpPr,
  DocTableRow,
  DocTableCell,
  CellElement,
  TableBorders,
  CellBorders,
  BorderSpec,
  // Shape geometry / fill sub-types.
  PathCmd,
  ShapeFill,
  ShapeStrokeFill,
  GradientStop,
  LineEnd,
} from './types';
export type { DocxTextRunInfo } from './renderer';
