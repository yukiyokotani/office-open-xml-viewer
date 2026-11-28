import type { VerticalGlyphMeasurementService } from '../layout/measurement-capabilities.js';

/** The vo=Tr fallback can under-report an upright glyph's inline extent.
 * Prefix fitting and committed line measurement use the same correction. */
export function verticalRunInkExtra(
  text: string,
  verticalRun: boolean,
  measurement?: VerticalGlyphMeasurementService,
  /** The route language the text is measured in (CanvasFontRoute.lang). */
  lang?: string,
): number {
  if (!verticalRun) return 0;
  if (!measurement) {
    throw new Error('Vertical glyph measurement capability is required for vertical text');
  }
  return measurement.measureRunInkExtra(text, lang);
}
