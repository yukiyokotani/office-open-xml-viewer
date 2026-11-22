import type { DocxTextRunInfo } from './renderer';

/** Advance width in CSS px of a substring of `run`, measured the way the run
 * was laid out and drawn: with its reference-em font scaled onto the run when
 * the projection supplies one (core reference-advances.ts), else with
 * `run.font` itself. `measureForFont` returns a width-measurer primed with a
 * Canvas font shorthand. */
export function runOverlayMeasure(
  run: DocxTextRunInfo,
  measureForFont: (font: string) => (s: string) => number,
): (s: string) => number {
  const reference = run.referenceFont;
  if (!reference) return measureForFont(run.font);
  const measure = measureForFont(reference.font);
  return (s) => measure(s) * reference.scale;
}
