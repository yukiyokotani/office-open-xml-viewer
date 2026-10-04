import { powerPointSymbolCoverage, powerPointCjkCoverage } from './powerpoint-line-metrics.js';
import { isReferenceSymbolCodePoint } from './font-resource-catalogue.js';
import { isCjkFallbackGlyph, EAST_ASIAN_SYMBOL_FALLBACK_FACES } from './east-asian-default.js';

// Per-cut attribution belongs to rendering. Preflight imports only the shared
// lightweight stack/tier selector and cannot load this resource catalogue.
/**
 * The face that draws one East Asian-slot glyph of an empty-slot run, when
 * the renderer can know it; null when only the platform's glyph fallback can
 * (decision B, and owner decision (c): the browser does not report which
 * installed face draws a fallback glyph).
 *
 * - A recorded symbol belongs to the first face in the painting stack whose
 *   real/synthetic cut's cmap covers it. Stop at unknown coverage: that face
 *   might draw it. Missing symbols continue through Calibri, Cambria Math and
 *   the CJK faces, just as painting does. The catalogue's Cambria Math lacks
 *   U+25C6 although the #1689 PDF resource drew it; no cmap presence is invented
 *   to emulate that different resource. Unrecorded scalars keep the existing
 *   selected-face model. This is font-data routing, not a new Office heuristic.
 * - CJK uses the same per-glyph rule, including S and the symbol faces
 *   before the CJK tiers. An OS/2 Far-East code page or some basic Han in
 *   a family's cmap selects a fallback chain; neither proves that this cut
 *   paints a particular glyph. Unknown earlier coverage stops attribution
 *   (decision c), and a missing glyph after decision B's first fallback
 *   remains unknown. This changes attribution only, not the painting stack.
 */
export function emptyEastAsianDrawingFace(
  selectedFace: string | null,
  cjkFaces: readonly string[],
  ch: string,
  bold = false,
  italic = false,
  symbolCoverage = powerPointSymbolCoverage,
  cjkCoverage = powerPointCjkCoverage,
): string | null {
  const cp = ch.codePointAt(0) ?? 0;
  const cjk = isCjkFallbackGlyph(ch);
  // Coverage outside the symbol sweep is unknown for every resource. Keep
  // the previous model there; CJK outside its catalogue domain stays unknown.
  if (!cjk && !isReferenceSymbolCodePoint(cp)) return selectedFace;
  const coverageFor = cjk ? cjkCoverage : symbolCoverage;
  for (const face of [selectedFace, ...EAST_ASIAN_SYMBOL_FALLBACK_FACES, ...cjkFaces]) {
    if (!face) continue;
    const coverage = coverageFor(face, bold, italic, cp);
    if (coverage === undefined) return null;
    if (coverage) return face;
  }
  return null;
}
