export {
  XlsxWorkbook,
  type LoadOptions,
  type RenderViewportToBitmapOptions,
} from './workbook.js';
export { XlsxViewer, XlsxSheetViewer } from './viewer.js';
// Resolved list-validation values (reachable via XlsxWorkbook.resolveValidationList).
export type { ResolvedList } from './validation-list.js';
export type {
  XlsxViewerOptions,
  XlsxSheetViewerOptions,
  XlsxViewportOffset,
  XlsxScrollToCellOptions,
  HiddenSheetMode,
  XlsxCopyResult,
} from './viewer.js';
export type {
  CellAddress,
  XlsxSelectionArea,
  XlsxSelectionContext,
  XlsxRangeSelectionContext,
  XlsxElementAnchorMarker,
  XlsxElementContext,
  XlsxSelectionContextCell,
  XlsxSelectionContextOptions,
  XlsxSelectionInput,
  XlsxSelectionState,
} from './selection.js';
export {
  MAX_SELECTION_AREAS,
  MAX_SELECTION_CONTEXT_CELLS,
  MAX_SELECTION_CONTEXT_TEXT_CHARACTERS,
} from './selection.js';
// IX2 find-in-document: the xlsx match-location shape (sheet + A1 cell ref).
// `FindMatch` / `FindMatchesOptions` come from core (shared across formats).
export type { XlsxMatchLocation } from './find.js';
export type { FindHighlightColors, FindMatch, FindMatchesOptions } from '@silurus/ooxml-core';
export type {
  ChartAreaGroupDecorations,
  ChartBarGroupDecorations,
  ChartDecorationLineStyle,
  ChartLabelBox,
  ChartLegendEntryOverride,
  ChartLineGroupDecorations,
  ChartPlotGroup,
  ChartPlotGroupAxisSlot,
  ChartPlotGroupKind,
  ChartRect,
  ChartTextBox,
  ChartTextParagraph,
  ChartTextRun,
  ChartTrendline,
  ChartType,
  ChartOfPie,
  ChartStockBarPaint,
  ChartStockUpDownBarStyle,
  ChartSurfaceBandFormat,
  ChartThreeD,
  ChartThreeDPictureOptions,
  ChartThreeDSurface,
  ChartThreeDSeriesAxis,
  ChartThreeDRenderer,
  ChartRegionMapRenderer,
  ChartExRenderer,
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
  DrawingMLCustomDashSegment,
  FillRect,
  ImageFill,
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
  ChartDisplayUnits,
  ChartDisplayUnitsLabel,
  SecondaryValueAxis,
  SrcRect,
  NoFill,
  SpaceLine,
  TileInfo,
  ViewerContextMenuEvent,
  ZoomableViewer,
} from '@silurus/ooxml-core';
export { autoResize, type AutoResizeOptions } from '@silurus/ooxml-core';
// IX1 — the shared hyperlink target shape surfaced by `XlsxViewerOptions.
// onHyperlinkClick`, plus the default "open in a new tab, sanitised" helper.
export { type HyperlinkTarget, openExternalHyperlink } from '@silurus/ooxml-core';
// Resolve `{type:'shared',si}` cells against a workbook's sharedStrings table
// (ECMA-376 §18.4.8). Exported so headless callers that parse a Worksheet
// directly (for example a bounded Node worksheet session) can concretize cell text.
export { resolveSharedStrings } from './shared-strings.js';
// Typed load-time error surfaced by XlsxWorkbook.load (e.g. a password-protected
// or legacy-binary .xls file). Re-exported so `@silurus/ooxml/xlsx` consumers can
// narrow on `err.code`.
export {
  OoxmlError,
  OoxmlDecodedImageLimitError,
  OoxmlResourceLimitError,
  isOoxmlDecodedImageLimitError,
  type OoxmlDecodedImageLimitMetric,
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
export type {
  Workbook,
  SheetMeta,
  SheetVisibility,
  Worksheet,
  WorksheetCellRange,
  // Outline (row/column grouping) display flags, reachable via Worksheet.outlinePr.
  OutlinePr,
  Row,
  Cell,
  CellValue,
  Styles,
  CellFont,
  CellFill,
  Border,
  BorderEdge,
  CellXf,
  NumFmt,
  MergeCell,
  ParsedWorkbook,
  ViewportRange,
  XlsxTextRunInfo,
  XlsxRenderViewportOptions,
  // Rich-text run sub-types (reachable via Cell rich-text values).
  Run,
  RunFont,
  SharedString,
  // Phonetic hints / furigana (ECMA-376 §18.4.6 / §18.4.3), reachable via
  // Cell text values and SharedString.
  PhoneticRun,
  PhoneticProperties,
  PhoneticType,
  PhoneticAlignment,
  // Differential / gradient style sub-types (reachable via Styles).
  Dxf,
  GradientFillSpec,
  // Conditional formatting (reachable via Worksheet.conditionalFormats).
  ConditionalFormat,
  CfRule,
  CfValue,
  CfStop,
  CfIcon,
  // Workbook-level metadata.
  DefinedName,
  Hyperlink,
  // Cell comments / notes (reachable via Worksheet.comments).
  XlsxComment,
  // Data validation rules (reachable via Worksheet.dataValidations).
  DataValidation,
  // Excel tables (reachable via Worksheet.tables).
  TableInfo,
  TableColumnInfo,
  // Slicers.
  SlicerAnchor,
  SlicerItem,
  SlicerStyle,
  SlicerElementStyle,
  // Metadata-only pivot facts (reachable via Worksheet.pivotTables).
  PivotTableMetadata,
  PivotLocation,
  PivotPageField,
  PivotDataField,
  PivotCacheSource,
  PivotMetadataStatus,
  PivotPartialReason,
  PivotDiagnostic,
  // Sparklines (reachable via Worksheet sparkline groups).
  SparklineGroup,
  Sparkline,
  // Drawings / shapes (reachable via Worksheet drawings).
  ImageAnchor,
  Duotone,
  ChartAnchor,
  ShapeAnchor,
  ShapeInfo,
  ShapeFill,
  ShapeGeom,
  ShapeText,
  ShapeParagraph,
  ShapeTextRun,
  PathInfo,
  PathCmd,
  ArrowEnd,
  GradientFill,
  GradientStop,
  PatternFill,
  SolidFill,
  // Canonical chart model (shared with core / pptx). `ChartAnchor.chart` is a
  // `ChartModel`.
  ChartModel,
  ChartSeries,
  ChartSeriesDataLabels,
  ChartDataLabelOverride,
  ChartDataTable,
  ChartDataPointOverride,
  ChartErrBars,
  ChartManualLayout,
  LegendManualLayout,
} from './types.js';
