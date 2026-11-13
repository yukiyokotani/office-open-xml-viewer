/**
 * The OpenType metric tables of a source face's bytes
 * (`LoadedSourceFace.resourceMetrics`), for the formats that lay out with
 * them (DOCX). A WOFF / WOFF2 face is inflated first (`web-font-sfnt.ts`);
 * the tables of a variable font carry its variation data (fvar / MVAR).
 * Kept apart from `source-session.ts`, which takes it as an option, so a
 * format that never reads the tables (PPTX, XLSX) neither runs nor bundles
 * the decoders.
 */
import { parseOpenTypeResourceMetrics, type OpenTypeLineMetrics } from './open-type-metrics.js';
import { metricSfnt } from './web-font-sfnt.js';

/** The tables of `data`, or null when they cannot be read. */
export async function readFontResourceMetrics(data: ArrayBuffer): Promise<OpenTypeLineMetrics | null> {
  try {
    const sfnt = await metricSfnt(data);
    return sfnt instanceof Uint8Array ? parseOpenTypeResourceMetrics(sfnt, undefined, { variations: true }) : null;
  } catch {
    return null;
  }
}
