/** ESM sidecar the PPTX parse workers import when font sources are
 * configured (font-demand-collector.ts). An inline (Blob) worker cannot
 * resolve the lazily emitted collector chunk, so without this module the
 * collector would be inlined into the worker source in the static graph. */
export { collectSlideFontDemand } from '../font-preload-demand.js';
