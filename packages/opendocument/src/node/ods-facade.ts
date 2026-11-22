// The Node XLSX facade and canvas helpers the OpenDocument spreadsheet tests
// drive: workbook materialization, cell display text and sheet painting.
export { materializeXlsxWorkbook } from '../../../node/src/index.ts';
export type { NodeCanvasFactory } from '../../../node/src/render.ts';
export { installImageBitmapShim, installOffscreenCanvasShim } from '../../../node/src/render.ts';
export { renderWorksheetViewport } from '../../../xlsx/src/render-orchestrator.ts';
export { formatCellValue } from '../../../xlsx/src/number-format.ts';
export { resolveSharedStrings } from '../../../xlsx/src/shared-strings.ts';
import type { NodeCanvasFactory } from '../../../node/src/render.ts';
import { loadSkiaForTests } from '../../../node/src/test-imports.ts';

/** skia-canvas, or null (local runs without the binding skip). */
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
