export type {
  ArrowEnd,
  Bullet,
  DrawingMLCustomDashSegment,
  EquationRun,
  Fill,
  FillRect,
  Glow,
  GradientFill,
  GradientStop,
  ImageFill,
  LineBreak,
  NoFill,
  Paragraph,
  PathCmd,
  PatternFill,
  Reflection,
  RenderOptions,
  Shadow,
  SoftEdge,
  SolidFill,
  SpaceLine,
  Stroke,
  TabStop,
  TextBody,
  TextOutline,
  TextRun,
  TextRunData,
  TileInfo,
} from './types/common';
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
  ChartDisplayUnits,
  ChartDisplayUnitsLabel,
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
  LegendManualLayout,
  SecondaryValueAxis,
} from './types/chart';
export type {
  LoadOptions,
  OoxmlResourceLimit,
  OoxmlResourceLimits,
  ProgressiveLayoutProgress,
  ProgressiveLayoutPartial,
} from './types/load-options';
export type {
  OoxmlResourceMetrics,
  OoxmlResourceMetricsCheckpoint,
  OoxmlResourcePolicySnapshot,
} from './types/resource-metrics';
export type { TextSelectionContextOptions, ViewerContextMenuEvent } from './selection-context';
export type {
  ViewerCommentsOptions,
  ViewerCommentConnectorOptions,
  ViewerCommentConnectorRoute,
  ViewerCommentConnectorStroke,
  ViewerCommentMessageContext,
  ViewerCommentThreadContext,
} from './comment-ui';
// Typed load-time error (PD4 seed): the `load()` factories throw this with a
// stable `code` for container-level failures detected on the main thread.
export {
  OoxmlError,
  OoxmlResourceLimitError,
  type OoxmlErrorCode,
  type OoxmlErrorStage,
  type OoxmlFormat,
  type OoxmlResourceMetric,
  type OoxmlResourceName,
  type OoxmlResourceLimitErrorDetails,
  type OoxmlResourceUsageSnapshot,
  type OoxmlResourceViolation,
} from './errors/ooxml-error';
// CFB (OLE2) container sniffer: the `load()` factories call this on the raw
// bytes before touching the parser worker, so a password-protected or legacy
// .doc/.xls/.ppt file becomes a typed OoxmlError instead of an opaque zip error.
export { sniffCfb, type CfbKind } from './errors/cfb-sniff';
// Shared load() guard: throws the right OoxmlError when the bytes are a CFB
// container (encrypted / legacy-binary / other) instead of an OOXML ZIP.
// `resolveOoxmlContainer` is the decrypt-aware superset the load() factories
// call: it returns plaintext ZIP bytes, decrypting an Agile-encrypted file when
// a password is supplied ([MS-OFFCRYPTO], PD8).
export { assertNotCfbContainer, resolveOoxmlContainer, toArrayBuffer } from './errors/cfb-guard';
// Agile Encryption decryption ([MS-OFFCRYPTO]): `decryptOoxml` turns an
// encrypted CFB + password into plaintext ZIP bytes. Lower-level primitives
// (key derivation, EncryptionInfo parse) are exported for testing / advanced use.
export {
  decryptOoxml,
  parseEncryptionInfo,
  AgileDecryptError,
  type DecryptResult,
  type DecryptFailure,
  type EncryptionInfoKind,
  type AgileEncryptionDescriptor,
} from './crypto';
export { readCfbStream } from './errors/cfb-read';
export { preloadGoogleFonts, unloadGoogleFonts, type FontPreloadEntry } from './fonts/preload';
// Embedded-font registration: docx `.odttf` (§17.8.1 obfuscated) + pptx
// `.fntdata` (raw sfnt) faces turned into FontFace objects in the active set.
export {
  registerEmbeddedFonts,
  unregisterEmbeddedFonts,
  deobfuscateOdttf,
  type EmbeddedFontFace,
} from './fonts/embedded';
// Shared Office-font → Google-Fonts substitute registry (Calibri → Carlito,
// Cambria → Caladea, popular web fonts, Arabic Noto fallbacks). Each package
// spreads this into its own map; script-fallback Noto faces live in
// SCRIPT_GOOGLE_FONTS below.
export { GOOGLE_FONT_SUBSTITUTES } from './fonts/google-fonts';
export { canvasFontString, createCanvasFontRoute, type CanvasFontRoute } from './fonts/canvas-route';
export {
  classifyCjkFont,
  classifyFontGeneric,
  isComplexScriptCodePoint,
  cjkFallbackChain,
  NON_CJK_SANS_FALLBACKS,
  NON_CJK_SERIF_FALLBACKS,
  SCRIPT_GOOGLE_FONTS,
  SCRIPT_PRELOAD_NAMES,
  scriptPreloadNamesForText,
  type CjkLang,
  type FontGenericClass,
  type FontVariant,
} from './fonts/scripts';
export {
  SYMBOL_MAP,
  WINGDINGS_MAP,
  symbolFontToUnicode,
  isSymbolFontFamily,
  symbolTextToUnicodeSegments,
  type SymbolTextSegment,
} from './fonts/symbol-font';
export { renderChart } from './chart/renderer';
export { formatLocalizedExcelShortDate } from './chart/chart-number-format';
export { autoResize, type AutoResizeOptions } from './autoResize';
export {
  buildCustomPath,
  getCustomGeometryBounds,
  type CustomGeometryBounds,
} from './shape/custGeom';
export {
  getCustGeomEndpoints,
  type CustGeomEndpoint,
  type CustGeomEndpoints,
} from './shape/custgeom-endpoints';
export { hexToRgba, relativeLuma, autoContrastColor, resolveFill, applyStroke } from './shape/paint';
export { buildShapePath, drawStar, drawPolygon, ooxmlArcTo } from './shape/preset';
export {
  paintDrawingMLShape,
  clipDrawingMLShape,
  withDrawingMLShapeTransform,
  type DrawingMLShapeFill,
  type DrawingMLShapeGeometry,
  type DrawingMLShapePaintPlan,
} from './shape/drawingml-shape';
export {
  drawArrowHead,
  lineEndPaintExtent,
  lineEndRetract,
  retractLineEndpoint,
} from './shape/arrow';
// Shared embedded-SVG decoder (Microsoft asvg:svgBlip extension) — used by all
// three renderers to prefer the vector original over the raster fallback.
// Path-keyed for the lazy byte-on-demand pipeline: fetches SVG bytes via a
// caller-supplied fetchImage(path, mimeType) and owns the object-URL lifecycle.
export {
  getCachedSvgImageByPath,
  dropSvgImageCache,
  type SvgImageSource,
  type SvgImageDecodeOptions,
} from './image/svg-image-by-path';
export type { SvgBlobDecoder, SvgDecodeTarget } from './worker/svg-decode-bridge';
// Shared per-document decoded-raster owner: one weighted LRU, decode gate and
// render-pass lease covers base blips, media posters, and derived effects across
// all three renderers. `fetchImage` remains the compatibility owner by default,
// while the general primitive accepts an explicit document owner object.
// `withBitmapCacheLease` serializes image-bearing paints for one document owner
// and pins that paint's decoded bitmaps. While the lease is held, LRU evictions
// and drops defer their GPU close to the last release, so a pass resolving more
// images than the cap never draws a closed bitmap. The lower-level acquire API
// remains available for callers that already own admission sequencing.
export {
  captureDecodedBitmapCacheEpoch,
  dropCachedDerivedBitmapNamespace,
  getCachedBitmapByPath,
  getCachedDecodedBitmap,
  getCachedDerivedBitmap,
  inspectCachedRasterSource,
  resolvedCachedBitmapVariantKey,
  peekCachedBitmapByPath,
  dropDecodedBitmapCache,
  releaseOwnedBitmap,
  dropBitmapCacheByPath,
  acquireBitmapCacheLease,
  withBitmapCacheLease,
  deferBitmapCloseWhileLeased,
  type DecodedBitmapCacheEpoch,
  type DecodedBitmapCacheOwner,
  type CachedBitmapOptions,
} from './image/bitmap-image-by-path';
export {
  isDecodeTargetResizableRasterFormat,
  type RasterBlobInspection,
  type RasterFormat,
} from './image/raster-blob-inspection';
export {
  chartImageFillKey,
  chartImageFillUsageSize,
  collectChartImageFillUsages,
  collectChartImageFillUsagesForCharts,
  collectChartMarkerImageFills,
  collectChartMarkerImageFillsForCharts,
  type ChartImageFillUsage,
  type ChartImageFillUsageSize,
  type ChartImageLookup,
} from './chart/image-fill';
// Shared WMF (Windows Metafile) player.
// The docx-specific cosmetic window/device-frame suppression is gated behind
// `suppressBoundaryFrame` (default off = spec-clean).
export {
  isWmf,
  isEmf,
  isMetafileMime,
  playWmf,
  renderWmfToBitmap,
  wmfRasterTarget,
} from './image/wmf';
// Cross-format raster/metafile admission and decode boundary.
export {
  decodeRasterOrMetafile,
  type DecodeRasterOptions,
} from './image/raster-or-metafile';
export {
  TiffDecodeError,
  isTiffDecodeError,
  type TiffRenderer,
  type TiffRenderOptions,
} from './image/tiff-contract';
// Raster pixel-dimension budget + header sniff (decode-bomb guard, RB1). Shared
// caps live in `./image/pixel-budget`; `decodeRasterOrMetafile` uses the sniff to
// refuse an over-budget PNG/JPEG/GIF/BMP/WEBP before `createImageBitmap`.
export {
  HARD_MAX_DECODED_IMAGE_BYTES,
  MAX_CONCURRENT_IMAGE_DECODES,
  MAX_DECODED_IMAGE_BYTES,
  MAX_RASTER_DIMENSION,
  MAX_RASTER_SOURCE_DIMENSION,
  MAX_RASTER_PIXELS,
  MAX_RASTER_SOURCE_PIXELS,
  OoxmlDecodedImageLimitError,
  isOoxmlDecodedImageLimitError,
  type OoxmlDecodedImageLimitMetric,
} from './image/pixel-budget';
export {
  normalizeImageResourceOptions,
  planDecodedImageTargets,
  type DecodedImageBudgetStrategy,
  type DecodedImageTargetDemand,
  type DecodedImageTargetPlan,
  type ImageResourceOptions,
  type NormalizedImageResourceOptions,
} from './image/adaptive-image-budget';
export {
  sniffRasterDimensions,
  rasterExceedsBudget,
  sourceRasterExceedsBudget,
  rasterHeaderExceedsBudget,
  type RasterDimensions,
} from './image/raster-dimensions';
// Shared `<a:srcRect>` crop (§20.1.8.55) for all three renderers: the source-rect
// math, the full-frame raster size for a cropped metafile, and the draw wrapper.
export {
  cropSourceRect,
  drawImageCropped,
  imageNaturalSize,
  metafileRasterSize,
  sourceRasterTargetSize,
  type SrcRect,
} from './image/crop';
// Shared DrawingML duotone image effect (§20.1.8.23): recolour a decoded image
// along a `clr1`→`clr2` luminance ramp. The pure pixel transform + a canvas
// wrapper that returns a new ImageBitmap, cached by the renderers per (path +
// colours). Consumed by all three formats (xlsx first; pptx/docx pictures can
// carry duotone too).
export {
  applyDuotone,
  duotoneImageData,
  hex6ToRgb,
  luminance601,
  defaultOffscreenFactory,
  type Duotone,
  type RgbaBuffer,
  type OffscreenFactory,
  type OffscreenSurface,
} from './image/duotone';
// Second-layer path-keyed cache that decodes a blip (shared base cache) then
// applies its `<a:duotone>` recolour once per (path + colours). Shared by the
// docx and pptx renderers so a duotone picture decodes + recolours once and is
// reused across page/slide revisits. xlsx keeps its own worksheet-scoped map.
export {
  getCachedDuotoneBitmapByPath,
  duotoneCacheKey,
  dropDuotoneBitmapCache,
} from './image/duotone-bitmap-by-path';
export { decodedBitmapTargetResizeOptions } from './image/raster-target';
// Shared vector-vs-raster blip gate: prefer the Microsoft asvg:svgBlip vector
// original except when an <a:srcRect> crop is present (then the raster's native
// pixel grid is required for the fractional crop math). Used by all three
// renderers' picture-decode paths.
export { preferVectorBlip } from './image/blip-gate';
// ECMA-376 §20.1.9 spec-driven preset geometry engine (presets.json from
// presetShapeDefinitions.xml). Coexists with the legacy hand-rolled
// `buildShapePath` above, which the pptx renderer still uses as a silhouette /
// fallback codepath. Consolidating the two is intentionally out of scope here.
export {
  renderPresetShape,
  hasPreset,
  buildPresetGeometryPath,
  buildPresetGeometryFillPath,
  getPresetGeometryBounds,
  getConnectorAnchors,
} from './shape/preset-geometry';
export { type PresetPath } from './shape/preset-geometry/path-executor';
// ECMA-376 §20.1.9.19 WordArt text-warp envelopes (presetTextWarpDefinitions.xml).
// Reuses the same guide-formula evaluator as the preset-geometry engine; drives
// the pptx renderer's per-glyph WordArt path.
export {
  hasTextWarp,
  isSingleEdgeWarp,
  buildWarpEnvelope,
  samplePolyline,
  warpGlyphTransform,
  warpArcLength,
  followPathUScale,
  type WarpEnvelope,
  type WarpGlyphTransform,
  type Polyline,
} from './shape/text-warp';
export {
  applyOuterShadow,
  applyInnerShadow,
  applySoftEdge,
  applyReflection,
  createAuxCanvas,
  type PaintShape,
  type EffectBBox,
} from './shape/effects';
// DrawingML 3D camera perspective projection (planar homography).
// ECMA-376 §20.1.5.5 (camera) / §20.1.5.11 (rot). sp3d bevel / extrusion / light
// shading lives in ./shape/bevel-shading (exported below).
export {
  computeScene3dQuad,
  isScene3dNonIdentity,
  computeDepthOffset,
  type CameraInput,
  type RotInput,
  type Scene3dQuad,
  type Vec2,
} from './shape/scene3d-camera';
export { drawProjected, expandProjectedQuad, projectQuadPoint } from './shape/scene3d-draw';
// DrawingML 3D bevel shading (Phase B). ECMA-376 §20.1.5.12 (sp3d) /
// §20.1.5.3 (bevel) / §20.1.10.9 (ST_BevelPresetType) / §20.1.5.9 (lightRig).
// `edt1d` / `shadePixel` / `shadeParamsFor` / `fillDirFromKey` are internal
// bevel-shading helpers with no cross-package consumer — only bevel-shading.test.ts
// exercises them, and it deep-imports them from './bevel-shading' directly, so they
// are intentionally NOT re-exported here (they stay `export`ed from their module for
// that deep import). `materialClass` / `lightDirFromRig` ARE re-exported: the pptx
// renderer consumes them through this barrel.
export {
  applyBevelShading,
  applyExtrusion,
  type ExtrusionInput,
  type BevelRegion,
  computeBevelNormals,
  bevelHeightProfile,
  distanceToEdge,
  lightDirFromRig,
  materialClass,
  type BevelMaterial,
  type BevelShadeParams,
  type LightRigRot,
  type BevelInput,
  type BevelCtx,
  type Vec3,
} from './shape/bevel-shading';
export {
  renderSparkline,
  type SparklineKind,
  type SparklineModel,
  type SparklineRect,
} from './sparkline/renderer';
export {
  THREE_D_MAX_SHAPE_FACES_PER_DATUM,
  type ChartThreeDRenderer,
} from './chart/three-d-contract';
export type { ChartRegionMapRenderer } from './chart/region-map-contract';
export type { ChartExRenderer } from './chart/chart-ex-contract';
export { workerRendererDescriptors } from './worker/renderer-module-contract';
export {
  mathToMathML,
  svgExtents,
  recolorSvg,
  MATH_RASTER_PX_PER_EM,
  rasterizeMathSvg,
  sizeMathSvgForRaster,
  tintMathRaster,
  type MathSvg,
  type MathRenderer,
  type RasterizedMathSvg,
} from './math';
export type {
  MathAccent,
  MathArray,
  MathBar,
  MathBorderBox,
  MathBox,
  MathDelimiter,
  MathFormula,
  MathFraction,
  MathFunc,
  MathGroup,
  MathGroupChr,
  MathLimit,
  MathNary,
  MathNode,
  MathPhant,
  MathRadical,
  MathRun,
  MathScript,
  MathSPre,
  MathStyle,
} from './types/math';
export { EMU_PER_INCH, EMU_PER_PT, EMU_PER_PX, PT_TO_PX } from './units';
export { isHTMLCanvas, defaultDpr } from './canvas/env';
export { crispOffset } from './canvas/crisp';
// Canvas backing-store size clamp (browser-limit guard, RB5): bound a requested
// canvas size to per-axis + total-area limits every engine honors, preserving
// aspect ratio, so a pathological page/slide size can't produce a blank canvas.
export {
  clampCanvasSize,
  MAX_CANVAS_DIMENSION,
  MAX_CANVAS_AREA,
  type ClampedCanvasSize,
} from './canvas/clamp';
// Shared border / line dash-pattern core (§17.18.2 ST_Border / §18.18.3
// ST_BorderStyle / §20.1.10.49 ST_PresetLineDashVal shape borders /
// §20.1.10.82 ST_TextUnderlineType run underlines). The
// [on, off, …].map(x => x*unit) helper + per-format relative tables; each
// format keeps its own multipliers. pptxPresetDashArray is intentionally not
// exported — applyStroke consumes it internally.
export {
  dashArray,
  docxBorderDashArray,
  xlsxBorderDashArray,
  pptxUnderlineDashArray,
  type RelativeDashPattern,
} from './draw/dash';
// Shared `double` border rail geometry (§17.18.2 / §18.18.3): floored-thirds
// device-pixel rail/gap/rail bands + a fill-based painter.
export { doubleRailGeometry, fillDoubleBorder } from './draw/double-border';
// Shared run-underline painter (§20.1.10.82 ST_TextUnderlineType). Single source
// of truth for underline geometry across the pptx / docx renderers; docx
// normalizes its §17.18.99 ST_Underline vocabulary to this DrawingML vocabulary
// before calling in.
export { drawUnderline } from './text/underline';
export {
  WorkerBridge,
  type WorkerLike,
  type WorkerBridgeOptions,
  type WorkerBridgeTransport,
  type WorkerRequestOptions,
  decodeDataUrl,
  WasmParserHost,
  WasmTrapError,
  isWasmTrap,
  type WasmTrapErrorCode,
  type WasmInit,
  type WasmInitOptions,
  type WasmReinit,
  type WasmInitInput,
  type WasmParserHostOptions,
} from './worker';
export {
  toVisualSegments,
  resolveBaseDirection,
  getDefaultBidiEngine,
  setBidiEngine,
  resetBidiEngine,
  RTL_GATE,
  hasStrongRtl,
  OBJECT_PLACEHOLDER,
  buildVisualOrder,
  type BidiEngine,
  type BaseDirection,
  type BidiClass,
  type BidiLevels,
  type StyledRun,
  type VisualSegment,
  type SegmentPart,
} from './text/bidi';
export type { KinsokuRules } from './text/kinsoku';
export {
  resolveKinsokuRules,
  DEFAULT_KINSOKU_RULES,
  kinsokuAdjustedSplit,
  crossRunKinsokuRetract,
} from './text/kinsoku';
export { isCjkBreakChar, isLatinWordCodePoint } from './text/cjk-ranges';
export {
  lineBreakClass,
  isUax14NoBreakPair,
  LINE_BREAK_UNICODE_VERSION,
  type LBClass,
} from './text/line-break';
// Dictionary-based line breaking for no-inter-word-space SEA scripts (Thai / Lao
// / Khmer, issue #797): the ICU-backed word-break offset enumerator plus the
// shared greedy whole-word fit kernel, consumed by all three renderers' wrap
// loops. Graceful fallback to cluster/character wrap when Intl.Segmenter is
// unavailable. `setSeaWordSegmenterForTest`/`resetSeaSegmenterForTest` are test
// seams only.
export {
  type SeaScript,
  type SeaWordSegmenter,
  type SeaMixedKinsoku,
  isSeaScriptCodePoint,
  isSeaGraphemeExtend,
  containsSeaScript,
  isGraphemeFillText,
  isDictionarySeaText,
  seaWordBreakOffsets,
  seaTransitionOffsets,
  seaMixedBreakOffsets,
  fitSeaWordPrefix,
  graphemeClusterOffsets,
  setSeaWordSegmenterForTest,
  resetSeaSegmenterForTest,
} from './text/sea-break';
// UAX #50 Vertical_Orientation (vo): how a code point orients in vertical text
// (tbRl / eaVert). Consumed by the vertical-text draw paths across packages.
export {
  verticalOrientation,
  verticalFormSubstitute,
  verticalBracketFormSubstitute,
  verticalTrUprightFallback,
  verticalTrLongMark,
  VO_UNICODE_VERSION,
} from './text/vertical-orientation';
export type { VerticalOrientation } from './text/vertical-orientation';
export {
  measureVerticalVertGlyph,
  verticalVertGlyphReachable,
  withVertFeature,
  withVertFeatureCanvasScope,
} from './text/vertical-vert-feature';
export type { VerticalGlyphCellMetrics } from './text/vertical-vert-feature';
// Shared Excel serial-date → UTC `Date` conversion (ECMA-376 §18.17.4.1),
// with the 1900 Lotus leap-year-bug compat and 1900/1904 date-system select.
// Used by the xlsx cell formatter and the core chart date formatter.
export { excelSerialToUtcDate, utcDateToExcelSerial } from './excel-date';
export { highlightBox } from './text/highlight-box';
export {
  distributeLineSlack,
  type DistributeSeg,
  type DistributeResult,
  type DistributeOptions,
  type SegStretch,
} from './text/line-distribute';
export { justifiedPiecePositions, type JustifiedPiece } from './text/justify-positions';
// ECMA-376 §17.18.59 ST_NumberFormat rendering + §17.16.4.3.1 field-format switch
// parsing — the shared numbering kernel (page numbers, list markers, …).
export { formatOrdinalNumber, type NumberFormat } from './text/number-format';
// Office-style decimal (round-half-up) formatting for fixed-precision display —
// Excel/Word/PowerPoint round `.xx5` up where `toFixed` rounds the binary double
// down (2.675 → "2.68"). Shared by xlsx number-format + chart label formatting.
export { roundDecimalHalfUp } from './text/round-decimal';
export { parseFieldFormatSwitch } from './text/field-format-switch';
// ECMA-376 §17.16.4.1 date-and-time formatting ("picture") switch — evaluates a
// DATE/TIME field's `\@ "…"` picture against a given instant (shared by docx/pptx).
export { formatDateTimePicture, parseDateTimePictureSwitch } from './text/date-time-picture';
// Format-agnostic index navigation for hidden-item "skip" mode (pptx hidden
// slides, xlsx hidden sheets): pure math over an isHidden(i) callback.
export { nextVisibleIndex, resolveVisibleIndex, countVisible } from './nav/visible-index';
// Internal-hyperlink target resolution (IX-nav): OPC part-name normalization
// (TS mirror of Rust resolve_target, so a pptx slide-rel target resolves to the
// same part name the parser keys slides by) + the pptx relative slide-show jump
// verbs (firstslide/lastslide/next/previous, ECMA-376 §21.1.2.3.5). Pure — the
// docx bookmark→page and pptx slidePart→index maps live in each viewer.
export {
  resolveOpcPartName,
  parseRelativeSlideJump,
  resolveRelativeSlideJump,
  type RelativeSlideJump,
} from './nav/internal-target';
// Virtualization range math for the continuous-scroll viewers (DocxScrollViewer /
// PptxScrollViewer): pure prefix-sum + binary-search over per-item heights. No DOM.
export { computeVisibleRange, type VisibleRange, type VisibleRangePad } from './layout/virtual-scroll';
// Shared exponential wheel/pinch zoom step (Ctrl/⌘+wheel). Pure — the caller
// clamps to its own [zoomMin, zoomMax]. Used by XlsxViewer + the scroll viewers.
// `anchoredZoomOffset` is the companion pointer-anchored ("zoom toward the
// cursor") scroll correction: given the pointer position it returns the new
// scroll offset that keeps the content point under the cursor fixed across a zoom.
export { zoomStepScale, anchoredZoomOffset, type ScrollClamp } from './interaction/zoom';
// IX9 — the shared zoom API contract for all five viewers: the ZoomableViewer
// interface (getScale/setScale/zoomIn/zoomOut/fitWidth/fitPage) plus its pure
// support logic (the discrete zoom-step ladder, fit-to-content scale math, clamp).
// One definition of "what a zoom factor means / what +/- steps are" across
// docx / pptx / xlsx so a host drives any viewer through the same calls.
export {
  type ZoomableViewer,
  type FitInput,
  type FitMode,
  ZOOM_STEP_LADDER,
  nextZoomStep,
  prevZoomStep,
  clampScale,
  fitScale,
} from './interaction/zoomable';
// Shared hyperlink model + URL scheme-allowlist sanitiser (IX1). One
// HyperlinkTarget shape + one sanitizeHyperlinkUrl predicate for docx / pptx /
// xlsx so the external-URL safety policy is defined once, not per format.
export {
  type HyperlinkTarget,
  DEFAULT_ALLOWED_HYPERLINK_SCHEMES,
  hyperlinkUrlScheme,
  sanitizeHyperlinkUrl,
  openExternalHyperlink,
} from './interaction/hyperlink';
// Format-agnostic font design line-metrics (OS/2 win / hhea sums) for faces the
// browser substitutes with different metrics — shared so docx (Word's design
// line box), pptx and xlsx can size line boxes / floor single-line height
// uniformly instead of each under-measuring a substituted Meiryo/Sakkal face.
export {
  fontWinLineHeightRatio,
  intendedSingleLinePx,
  correctLineMetrics,
} from './text/line-metrics';
// Exact local-font metric probing. Shared contract for docx/xlsx/pptx; format
// packages supply only their evidence-backed Office line-height policy.
export {
  loadLocalFontMetrics,
  unloadLocalFontMetrics,
  normalizeLocalFontMetricFamily,
  type LocalFontMetricRequest,
  type ResolvedLocalFontMetric,
  type LoadedLocalFontMetrics,
} from './fonts/local-metrics';
// Format-agnostic same-font Canvas-vs-Word line-fit bias. Consumers keep their
// layout/paint wiring local, while the metric provenance and normalized family
// matching remain shared data.
export { fontAdvanceBiasEm } from './text/font-advance-metrics';
// IX2 in-document text search (findText). Format-agnostic index + match →
// run-slice resolution (buildTextIndex/findMatches), the pure highlight-extent
// helper (sliceHorizontalExtent), the active-match cursor arithmetic behind
// findNext/findPrev (nextActive/prevActive/clampActive), and the shared public
// FindMatch result shape. Each viewer supplies the run stream + turns a slice
// into a pixel rect / DOM box in its own geometry; core owns the string math.
export {
  buildTextIndex,
  findMatches,
  type SearchRun,
  type TextIndex,
  type MatchRunSlice,
  type TextMatch,
  type FindMatchesOptions,
} from './search/text-index';
export { sliceHorizontalExtent, overlayPercent } from './search/highlight-rect';
export { nextActive, prevActive, clampActive } from './search/find-cursor';
export type { FindMatch } from './search/find-match';
export type { FindHighlightColors } from './search/find-highlight';
