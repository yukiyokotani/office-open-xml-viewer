// Worker mode with a `fontResolver`: the font-source module together with the
// main-side bridge host, from one import() (workbook.ts), so the viewer realm
// and the worker bridge load or fail together. Main mode and documents
// without a resolver import font-source-load.ts alone.
export * from './font-source-load.js';
export { FontResolverBridgeHost } from '@silurus/ooxml-core/internal/font-sources';
