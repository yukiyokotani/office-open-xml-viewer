import { overlayPercent } from '@silurus/ooxml-core';
import type { PptxTextRunInfo } from './renderer.js';
import { pptxRunFrameTransform } from './run-frame-transform.js';

/** Run frame shared by selection and find. Only `cellHorzOverflow` admits the
 * physical cell x clip, which the renderer emits for horizontal bodies alone;
 * other runs, including vertical table cells, keep the shape-frame overflow.
 */
export function createPptxRunFrame(
  ownerDocument: Document, run: PptxTextRunInfo, cssWidth: number, cssHeight: number,
  pointerEvents: 'all' | 'none', shapeOverflow: 'visible' | 'hidden',
): { div: HTMLDivElement; w: number; h: number } {
  const frame = ownerDocument.createElement('div');
  const cell = run.cellHorzOverflow !== undefined;
  frame.style.cssText = `position:absolute;` +
    `left:${overlayPercent(run.shapeX, cssWidth)};top:${overlayPercent(run.shapeY, cssHeight)};` +
    `width:${overlayPercent(run.shapeW, cssWidth)};height:${overlayPercent(run.shapeH, cssHeight)};` +
    `pointer-events:${pointerEvents};overflow:${cell ? 'visible' : shapeOverflow};`;
  if (run.cellHorzOverflow === 'clip') {
    // CSS Overflow 3 §3.1: clip + visible retains single-axis clipping, whereas
    // hidden + visible would compute the other axis to auto.
    // https://www.w3.org/TR/css-overflow-3/#overflow-properties
    frame.style.overflowX = 'clip';
    frame.style.overflowY = 'visible';
  }
  const transform = pptxRunFrameTransform(run);
  if (transform) {
    frame.style.transformOrigin = 'center center';
    frame.style.transform = transform;
  }
  return { div: frame, w: run.shapeW, h: run.shapeH };
}
