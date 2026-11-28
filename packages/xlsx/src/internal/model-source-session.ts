/** Optional Node model-source acquisition; imported only for configured sources. */
export {
  validateXlsxModelSourceArchive,
  validateXlsxModelSourceViewDefaults,
  type XlsxModelSourceArchive,
} from './worker-worksheet-source.js';
export { configureHostLayout, type HostLayoutFont } from './host-layout.js';
export { MDW_FALLBACK } from './grid-metrics.js';
export {
  acquireXlsxSessionFromArchive,
  type XlsxOwnedArchiveSource,
} from './node-model-source-acquisition.js';
