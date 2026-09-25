// The Node facade and canvas helpers the legacy end-to-end tests drive. Every
// test under src/node/ imports the OOXML Node package through this one file.
export {
  materializeDocxDocument,
  materializePptxPresentation,
  materializeXlsxWorkbook,
  openDocxDocument,
  openPptxPresentation,
  openXlsxWorkbook,
  renderSlideNode,
} from '../../../node/src/index.ts';
export type { NodeCanvasFactory, NodeCanvasLike } from '../../../node/src/render.ts';
// XLSX worksheet painting and grid geometry, as the Node XLSX tests drive them.
export { installImageBitmapShim, installOffscreenCanvasShim } from '../../../node/src/render.ts';
export { renderWorksheetViewport } from '../../../xlsx/src/render-orchestrator.ts';
export { GridGeometry } from '../../../xlsx/src/internal/grid-geometry.ts';
import type { NodeCanvasFactory } from '../../../node/src/render.ts';
import { loadSkiaForTests } from '../../../node/src/test-imports.ts';

/** skia-canvas, or null (local runs without the binding skip; CI requires it). */
export const skia = await loadSkiaForTests();

/** A Node canvas factory backed by skia-canvas; call only when `skia` is set. */
export function skiaFactory(): NodeCanvasFactory {
  if (!skia) throw new Error('skia-canvas is unavailable');
  const { Canvas, loadImage } = skia;
  return {
    createCanvas: (width, height) => new Canvas(width, height) as unknown as ReturnType<NodeCanvasFactory['createCanvas']>,
    loadImage: (async (bytes: ArrayBuffer | Uint8Array) => loadImage(Buffer.from(new Uint8Array(bytes)))) as unknown as NodeCanvasFactory['loadImage'],
  };
}
