import type { PptxTextRunInfo } from './renderer.js';

/**
 * CSS transform for a PPTX run's local frame. ECMA-376 Part 1 DrawingML
 * `CT_Transform2D` defines the graphic-frame `rot`, `flipH`, and `flipV` values;
 * those outer transforms precede the text body's vertical-text rotation, so the
 * order cannot be reduced to one summed rotation when a flip exists. Kept in
 * the PPTX package because this is PresentationML/DrawingML frame geometry, not
 * a format-independent canvas primitive.
 */
export function pptxRunFrameTransform(run: PptxTextRunInfo): string {
  const textBodyRotation = run.textBodyRotation ?? 0;
  const flipH = run.shapeFlipH === true;
  const flipV = run.shapeFlipV === true;
  if (flipH || flipV) {
    return `rotate(${run.rotation}deg) ` +
      `scale(${flipH ? -1 : 1}, ${flipV ? -1 : 1})` +
      (textBodyRotation === 0 ? '' : ` rotate(${textBodyRotation}deg)`);
  }
  const totalRotation = run.rotation + textBodyRotation;
  return totalRotation === 0 ? '' : `rotate(${totalRotation}deg)`;
}

/** Stable grouping key shared by selection and find-highlight overlays. */
export function pptxRunFrameKey(run: PptxTextRunInfo, transform: string): string {
  return `${run.shapeX},${run.shapeY},${run.shapeW},${run.shapeH},` +
    transform;
}
