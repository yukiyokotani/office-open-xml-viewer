import { readOpenTypeLineMetrics, type OpenTypeLineMetrics } from '../fonts/open-type-metrics.js';
import { parseFontSupportFacts, takeFontErasureGlyphs } from './font-support-facts.js';
import { retainFontSupportFacts } from './font-support-registry.js';
/** PPTX/generation opt into shaping support over the same validated sfnt,
 * maxp and all-map cmap engine as ordinary resource metrics. Analysis and
 * erasure projection are synchronous; only immutable facts survive, never
 * DataViews, bytes, glyph graphs or callbacks. Registration remains separate. */
export function parseOpenTypeResourceWithSupport(bytes: Uint8Array, faceIndex?: number): OpenTypeLineMetrics | null {
  return readOpenTypeLineMetrics(bytes, faceIndex, true, (view, tables, safeCoverage) => {
    let support = parseFontSupportFacts(view, tables);
    const erasing = takeFontErasureGlyphs(support);
    if (erasing) support = Object.freeze({ ...support, erasureSafeRanges: safeCoverage(erasing) });
    return owner => retainFontSupportFacts(owner, support);
  });
}
