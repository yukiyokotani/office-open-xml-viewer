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
  // Untagged runs call the factory with the font alone, as callbacks that
  // inspect their argument count expect.
  const measurerFor = (font: string) =>
    run.lang === undefined ? measureForFont(font) : measureForFont(font, run.lang);
  if (!reference) return measurerFor(run.font);
  const measure = measurerFor(reference.font);
  return (s) => measure(s) * reference.scale;
}
