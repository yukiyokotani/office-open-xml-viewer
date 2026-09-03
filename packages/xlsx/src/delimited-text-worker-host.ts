import DelimitedTextWorker from './delimited-text-worker.ts?worker&inline';

/** Loaded only when XlsxSheetViewer is asked to preview delimited text. */
export function createDelimitedTextWorker(): Worker {
  return new DelimitedTextWorker();
}
