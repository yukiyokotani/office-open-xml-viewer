/**
 * Format-owned internal entry point for the Node facade. This is deliberately
 * not part of the public browser API: it exposes the canonical DOCX acquisition,
 * normalization, layout, and paint pipeline without letting Node reconstruct it
 * through arbitrary source-file deep imports.
 */
export { normalizeDocxDocumentModel } from '../parser-model.js';
export {
  materializeDocumentPullAdapterSession,
  materializeDocumentPullLayoutSession,
  materializeDocumentPullSession,
} from '../document-pull-client.js';
export { DocumentPullWorker, type DocxDocumentCursorArchive } from '../document-pull-worker.js';
export { createLayoutServices } from '../layout-runtime.js';
export { errorPageLayoutReady } from '../layout/document-layout-variants.js';
export { retainRenderWorkerDocumentLayout } from '../render-worker-layout.js';
export { renderLayoutSourceToCanvas } from '../renderer.js';
export { normalizeLayoutOptions } from '../layout/options.js';
export {
  acquireDocxNodeDocument,
  type AcquiredDocxNodeDocument,
  type DocxNodeAcquisitionOptions,
  type DocxNodeArchive,
  type DocxNodeSessionArchive,
  type DocxNodePullIdentity,
  type DocxNodePullOptions,
  type DocxNodePullTransport,
} from './node-acquisition.js';
export { normalizeDocxParserError } from './parser-errors.js';
