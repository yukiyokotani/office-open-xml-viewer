// Node entry: every parse goes through an owned asynchronous session
// (`open*`) or an asynchronous materializer built on that same session. There
// is deliberately no synchronous wrapper; Worker threads and `Atomics.wait`
// must not be used to imitate one.
export {
  materializePptxPresentation,
  openPptxPresentation,
  type OpenPptxPresentationOptions,
  type PptxPresentationSession,
  type PptxSessionRenderOptions,
} from './pptx';
export {
  materializeDocxDocument,
  openDocxDocument,
  type DocxDocumentSession,
  type DocxPageRenderOptions,
  type DocxRenderedPage,
  type OpenDocxDocumentOptions,
} from './docx';
export {
  materializeXlsxWorkbook,
  materializeXlsxWorkbookIndex,
  materializeXlsxWorksheet,
  openXlsxWorkbook,
  type OpenXlsxWorkbookOptions,
  type DeepReadonly,
  type MaterializedXlsxWorkbook,
  type ReadonlyParsedWorkbook,
  type XlsxWorkbookSession,
  type XlsxWorksheetRowChunk,
} from './xlsx';
export {
  renderSlideNode,
  installImageBitmapShim,
  installOffscreenCanvasShim,
  type NodeCanvasLike,
  type NodeCanvasFactory,
  type NodeImageLike,
} from './render';
export type { OoxmlNodeSessionOptions } from './session-options';
export type {
  OoxmlResourceMetrics,
  OoxmlResourceMetricsCheckpoint,
  OoxmlResourcePolicySnapshot,
} from '@silurus/ooxml-core';
// Application-supplied model sources (each format's opt-in source option).
export type {
  ModelSource,
  ModelSourceConfig,
  ModelSourceConfigValue,
  ModelSourceLoad,
  ModelSourceModule,
  ModelSourceModuleDescriptor,
  ModelSourceOpenContext,
  ModelSourceResourcePolicy,
  ModelSourceTarget,
  ModelSourceWorksheetLimits,
  OpenedModelSource,
} from '@silurus/ooxml-core';
export {
  OoxmlDecodedImageLimitError,
  OoxmlError,
  type OoxmlErrorCode,
  OoxmlResourceLimitError,
  isOoxmlDecodedImageLimitError,
  type OoxmlDecodedImageLimitMetric,
  type OoxmlResourceLimit,
  type OoxmlResourceLimitErrorDetails,
  type OoxmlUnsupportedContentDetails,
  type OoxmlFontSourceQuota,
  type OoxmlFontSourceViolation,
  type OoxmlResourceLimits,
  type OoxmlResourceUsageSnapshot,
  type XlsxWorksheetLimits,
} from '@silurus/ooxml-core';
