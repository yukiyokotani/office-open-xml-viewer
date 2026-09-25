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
// Application-supplied model sources (OoxmlNodeSessionOptions.modelSources).
export type {
  ModelSource,
  ModelSourceConfig,
  ModelSourceConfigValue,
  ModelSourceLoad,
  ModelSourceModule,
  ModelSourceModuleDescriptor,
  ModelSourceTarget,
  OpenedModelSource,
} from '@silurus/ooxml-core';
export {
  OoxmlDecodedImageLimitError,
  OoxmlResourceLimitError,
  isOoxmlDecodedImageLimitError,
  type OoxmlDecodedImageLimitMetric,
  type OoxmlResourceLimit,
  type OoxmlResourceLimitErrorDetails,
  type OoxmlResourceLimits,
  type OoxmlResourceUsageSnapshot,
} from '@silurus/ooxml-core';
