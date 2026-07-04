export type {
  ArrowEnd,
  Bullet,
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
  ChartDataLabelOverride,
  ChartDataPointOverride,
  ChartErrBars,
  ChartManualLayout,
  ChartModel,
  ChartRect,
  ChartSeries,
  ChartSeriesDataLabels,
  ChartType,
  LegendManualLayout,
  SecondaryValueAxis,
} from './types/chart';
export type { LoadOptions } from './types/load-options';
// Typed load-time error (PD4 seed): the `load()` factories throw this with a
// stable `code` for container-level failures detected on the main thread.
export { OoxmlError, type OoxmlErrorCode } from './errors/ooxml-error';
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
export { preloadGoogleFonts, type FontPreloadEntry } from './fonts/preload';
// Shared Office-font → Google-Fonts substitute registry (Calibri → Carlito,
// Cambria → Caladea, popular web fonts, Arabic Noto fallbacks). Each package
// spreads this into its own map; script-fallback Noto faces live in
// SCRIPT_GOOGLE_FONTS below.
export { GOOGLE_FONT_SUBSTITUTES } from './fonts/google-fonts';
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
export { autoResize, type AutoResizeOptions } from './autoResize';
export { buildCustomPath } from './shape/custGeom';
export {
  getCustGeomEndpoints,
  type CustGeomEndpoint,
  type CustGeomEndpoints,
} from './shape/custgeom-endpoints';
export { hexToRgba, relativeLuma, autoContrastColor, resolveFill, applyStroke } from './shape/paint';
export { buildShapePath, drawStar, drawPolygon, ooxmlArcTo } from './shape/preset';
export { drawArrowHead, lineEndRetract, retractLineEndpoint } from './shape/arrow';
// Shared embedded-SVG decoder (Microsoft asvg:svgBlip extension) — used by all
// three renderers to prefer the vector original over the raster fallback.
// Path-keyed for the lazy byte-on-demand pipeline: fetches SVG bytes via a
// caller-supplied fetchImage(path, mimeType) and owns the object-URL lifecycle.
export { getCachedSvgImageByPath, dropSvgImageCache } from './image/svg-image-by-path';
// Sibling of the SVG cache for raster/metafile blips: a per-document (keyed by
// `fetchImage`), path-keyed LRU of decoded `ImageBitmap`s shared by all three
// renderers, so a picture is decoded once per document and reused across
// re-renders / page revisits instead of re-decoding every frame. `peek…` serves
// the synchronous picture-bullet draw; `drop…` closes a document's bitmaps on
// its viewer's `destroy()`.
export {
  getCachedBitmapByPath,
  peekCachedBitmapByPath,
  dropBitmapCacheByPath,
  type CachedBitmapOptions,
} from './image/bitmap-image-by-path';
// Shared WMF (Windows Metafile) player + the raster/metafile decoder all three
// renderers route through, so a WMF/EMF blip (which `createImageBitmap` cannot
// decode) rasterizes or is skipped gracefully instead of throwing and vanishing.
// The docx-specific cosmetic window/device-frame suppression is gated behind
// `suppressBoundaryFrame` (default off = spec-clean).
export {
  isWmf,
  isEmf,
  isMetafileMime,
  playWmf,
  renderWmfToBitmap,
  wmfRasterTarget,
  decodeRasterOrMetafile,
  type DecodeRasterOptions,
} from './image/wmf';
// Shared `<a:srcRect>` crop (§20.1.8.55) for all three renderers: the source-rect
// math, the full-frame raster size for a cropped metafile, and the draw wrapper.
export {
  cropSourceRect,
  drawImageCropped,
  imageNaturalSize,
  metafileRasterSize,
  type SrcRect,
} from './image/crop';
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
  getConnectorAnchors,
} from './shape/preset-geometry';
export { type PresetPath } from './shape/preset-geometry/path-executor';
export {
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
export { drawProjected, expandProjectedQuad } from './shape/scene3d-draw';
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
  mathToMathML,
  svgExtents,
  recolorSvg,
  type MathSvg,
  type MathRenderer,
} from './math';
export type { MathNode, MathFormula, MathStyle } from './types/math';
export { EMU_PER_INCH, EMU_PER_PT, EMU_PER_PX, PT_TO_PX } from './units';
export { isHTMLCanvas, defaultDpr } from './canvas/env';
export { crispOffset } from './canvas/crisp';
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
  type WorkerRequestOptions,
  decodeDataUrl,
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
// Trailing-whitespace-robust text advance: `measureText` excludes a run's
// trailing/lone whitespace on Firefox & WebKit/Safari but includes it on
// Chrome/Blink & skia-canvas. Shared so docx/pptx/xlsx keep inter-word spaces
// from collapsing on trimming engines (measure==draw on every engine).
export { measureAdvanceWithTrailingSpace } from './text/measure-space';
// Format-agnostic index navigation for hidden-item "skip" mode (pptx hidden
// slides, xlsx hidden sheets): pure math over an isHidden(i) callback.
export { nextVisibleIndex, resolveVisibleIndex, countVisible } from './nav/visible-index';
// Virtualization range math for the continuous-scroll viewers (DocxScrollViewer /
// PptxScrollViewer): pure prefix-sum + binary-search over per-item heights. No DOM.
export { computeVisibleRange, type VisibleRange, type VisibleRangePad } from './layout/virtual-scroll';
// Shared exponential wheel/pinch zoom step (Ctrl/⌘+wheel). Pure — the caller
// clamps to its own [zoomMin, zoomMax]. Used by XlsxViewer + the scroll viewers.
export { zoomStepScale } from './interaction/zoom';
// Format-agnostic font design line-metrics (OS/2 win / hhea sums) for faces the
// browser substitutes with different metrics — shared so docx (Word's design
// line box), pptx and xlsx can size line boxes / floor single-line height
// uniformly instead of each under-measuring a substituted Meiryo/Sakkal face.
export {
  fontWinLineHeightRatio,
  intendedSingleLinePx,
  correctLineMetrics,
} from './text/line-metrics';
