/** Canonical XLSX workbook/worksheet coordinator entry point consumed by Node. */
export { GridGeometry } from './grid-geometry.js';
export {
  isXlsxWorksheetPullResponse,
  XlsxWorksheetPullClient,
  type XlsxWorksheetPullClientOptions,
  type XlsxWorksheetPullUnit,
} from '../worksheet-pull-client.js';
export {
  WorksheetPullWorker,
  XLSX_WORKSHEET_PULL_BYTES,
  type WorksheetWireChunk,
} from '../worksheet-pull-worker.js';
export {
  addWorksheetCacheUsage,
  addWorksheetUsage,
  assertWorksheetCacheUsage,
  assertWorksheetJsonBytes,
  assertWorksheetModelUsage,
  completeWorksheetUsage,
  measureRows,
  measureWorksheet,
  type WorksheetCacheUsage,
  type WorksheetModelUsage,
} from '../worksheet-resource-limits.js';
export {
  acquireXlsxNodeSession,
  acquireXlsxSessionFromArchive,
  type XlsxNodeAcquisition,
  type XlsxNodeAcquisitionOptions,
  type XlsxNodeArchive,
} from './node-acquisition.js';
