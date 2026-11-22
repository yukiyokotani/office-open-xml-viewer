import type { DocxTextRunInfo } from './renderer';

/** Advance width in CSS px of a substring of `run`, measured the way the run
 * was laid out and drawn: with its reference-em font scaled onto the run when
 * the projection supplies one (core reference-advances.ts), else with
 * `run.font` itself, in the run's Canvas language. `measureForFont` returns a
 * width-measurer primed with a Canvas font shorthand and language. */
export function runOverlayMeasure(
  run: DocxTextRunInfo,
  measureForFont: (font: string, lang?: string) => (s: string) => number,
): (s: string) => number {
  const reference = run.referenceFont;
  if (!reference) return measureForFont(run.font, run.lang);
  const measure = measureForFont(reference.font, run.lang);
  return (s) => measure(s) * reference.scale;
}
