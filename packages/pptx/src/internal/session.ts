/** Canonical PPTX acquisition/projector entry point consumed by Node. */
export { normalizePresentationBootstrap } from '../presentation-preflight.js';
export { PptxSlidePullClient } from '../slide-pull-client.js';
export {
  readPptxSlideCursorUsage,
  type PptxSlideCursorArchive,
} from '../slide-cursor-operation.js';
export { SlidePullWorker } from '../slide-pull-worker.js';
export type { PresentationBootstrap } from '../worker-protocol.js';
export { renderSlide } from '../renderer.js';
export {
  acquirePptxNodeSession,
  acquirePptxSessionFromArchive,
  type PptxNodeAcquisition,
  type PptxNodeAcquisitionOptions,
  type PptxNodeArchive,
} from './node-acquisition.js';
